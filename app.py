# -*- coding: utf-8 -*-
"""模型输入输出测试工具
启动: python app.py  然后浏览器打开 http://127.0.0.1:5000
"""
import base64
import html as html_mod
import io
import re
import socket
import threading
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor

import requests
from flask import Flask, jsonify, render_template, request

# 兼容中文 Windows 主机名：getfqdn 可能因编码问题抛 UnicodeDecodeError
_original_getfqdn = socket.getfqdn
def _safe_getfqdn(name=""):
    try:
        return _original_getfqdn(name)
    except Exception:
        return "localhost"
socket.getfqdn = _safe_getfqdn

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 允许最大 200MB 的请求体（多模态附件为 base64）

IMAGE_EXTS = {"png", "jpg", "jpeg", "gif", "webp", "bmp"}
AUDIO_EXTS = {"wav", "mp3"}
TEXT_EXTS = {"txt", "md", "csv", "json", "log", "xml", "html", "yaml", "yml",
             "py", "js", "ts", "java", "c", "cpp", "h", "go", "rs", "sql", "ini", "bat"}
MAX_TEXT_CHARS = 100000  # 提取出的文档文本上限，超出则截断


def extract_docx_text(data_b64: str) -> str:
    """从 docx（zip 包）中解出 word/document.xml 并提取纯文本"""
    raw = base64.b64decode(data_b64)
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    xml = re.sub(r"</w:p>", "\n", xml)  # 段落换行
    text = re.sub(r"<[^>]+>", "", xml)  # 去掉所有 XML 标签
    return html_mod.unescape(text).strip()


def extract_text_file(data_b64: str) -> str:
    return base64.b64decode(data_b64).decode("utf-8", "ignore").strip()


_COM_LOCK = threading.Lock()  # COM 转换同一时刻只允许一个，WPS 不支持并发自动化

def docx_to_pdf_b64(data_b64: str) -> str:
    """调用本机 WPS/Word（COM）将 docx 转为 PDF，返回 PDF 的 base64。失败抛异常。"""
    import os
    import shutil
    import tempfile
    import win32com.client

    tmpdir = tempfile.mkdtemp(prefix="mt_conv_")
    src = os.path.join(tmpdir, "input.docx")
    dst = os.path.join(tmpdir, "input.pdf")
    try:
        with open(src, "wb") as f:
            f.write(base64.b64decode(data_b64))
        last_err = None
        for progid in ("KWPS.Application", "Word.Application"):  # WPS 优先，兼容真实 MS Word
            try:
                import pythoncom
                pythoncom.CoInitialize()  # Flask 每个请求一个线程，COM 需按线程初始化
                try:
                    with _COM_LOCK:
                        word = win32com.client.DispatchEx(progid)  # 每次新建独立实例，避免状态污染
                        try:
                            doc = word.Documents.Open(os.path.abspath(src))
                            doc.SaveAs2(os.path.abspath(dst), FileFormat=17)  # 17 = PDF
                            doc.Close()
                        finally:
                            word.Quit()
                finally:
                    pythoncom.CoUninitialize()
                with open(dst, "rb") as f:
                    return base64.b64encode(f.read()).decode()
            except Exception as e:
                last_err = e
                print(f"[warn] COM 转换（{progid}）失败: {e}", flush=True)
        raise RuntimeError(f"docx 转 PDF 失败: {last_err}")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def preprocess_attachments(payload: dict):
    """请求分发前统一处理：docx 转为 PDF 原样发送，转换失败则保留原样（后续走文本提取兜底）"""
    out = []
    for att in payload.get("attachments") or []:
        name = att.get("name") or "file"
        ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if ext == "docx":
            try:
                pdf_b64 = docx_to_pdf_b64(att.get("data") or "")
                out.append({"name": name + ".pdf", "mime": "application/pdf", "size": att.get("size"),
                            "data": pdf_b64})
                continue
            except Exception as e:
                print(f"[warn] docx 转 PDF 失败（{name}），改用文本提取: {e}", flush=True)
        out.append(att)
    payload["attachments"] = out


def build_content_parts(attachments):
    """把前端上传的附件转为 OpenAI 多模态 content parts"""
    parts = []
    for att in attachments or []:
        name = att.get("name") or "file"
        mime = (att.get("mime") or "application/octet-stream").lower()
        data = att.get("data") or ""
        ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if mime.startswith("image/") or ext in IMAGE_EXTS:
            if not mime.startswith("image/"):
                mime = "image/" + ("jpeg" if ext == "jpg" else ext)
            parts.append({"type": "image_url", "image_url": {"url": f"data:{mime};base64,{data}"}})
        elif mime.startswith("audio/") or ext in AUDIO_EXTS:
            fmt = ext if ext in AUDIO_EXTS else ("mp3" if "mpeg" in mime else "wav")
            parts.append({"type": "input_audio", "input_audio": {"data": data, "format": fmt}})
        elif ext == "docx" or mime == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            # docx 接口的 file 类型不支持，改为提取文本后作为文本部分发送
            try:
                text = extract_docx_text(data)
            except Exception as e:
                parts.append({"type": "text", "text": f"[附件 {name} 解析失败：{e}]"})
                continue
            if not text:
                parts.append({"type": "text", "text": f"[附件 {name} 内容为空]"})
                continue
            if len(text) > MAX_TEXT_CHARS:
                text = text[:MAX_TEXT_CHARS] + f"\n...(文档过长，已截断，共发送前 {MAX_TEXT_CHARS} 字符)"
            parts.append({"type": "text", "text": f"以下是附件 {name} 的正文内容：\n\n{text}"})
        elif ext in TEXT_EXTS:
            text = extract_text_file(data)
            if len(text) > MAX_TEXT_CHARS:
                text = text[:MAX_TEXT_CHARS] + f"\n...(文件过长，已截断，共发送前 {MAX_TEXT_CHARS} 字符)"
            parts.append({"type": "text", "text": f"以下是附件 {name} 的内容：\n\n{text}"})
        else:
            # pdf 走 file 部分（OpenAI 格式仅支持 pdf），不兼容的服务会在结果中报错提示
            parts.append({"type": "file", "file": {"filename": name, "file_data": f"data:{mime};base64,{data}"}})
    return parts


TIMEOUT = 300  # 单次请求超时（秒）


def normalize_base_url(base_url: str) -> str:
    """补全 base_url：去掉末尾斜杠；若末尾不是 /v1、/v4 这类版本号则追加 /v1"""
    url = (base_url or "").strip().rstrip("/")
    if not url:
        raise ValueError("请填写 Base URL")
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    if not re.search(r"/v\d+$", url):
        url += "/v1"
    return url


def get_headers(api_key: str) -> dict:
    return {
        "Authorization": f"Bearer {api_key.strip()}",
        "Content-Type": "application/json",
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/models", methods=["POST"])
def list_models():
    data = request.get_json(force=True)
    try:
        base = normalize_base_url(data.get("base_url", ""))
        resp = requests.get(base + "/models", headers=get_headers(data.get("api_key", "")), timeout=30)
        resp.raise_for_status()
        payload = resp.json()
        models = [m["id"] for m in payload.get("data", []) if m.get("id")]
        return jsonify({"ok": True, "models": sorted(models)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 200


def call_one_model(payload: dict, model: str) -> dict:
    """调用单个模型，返回统一结构的结果"""
    body = {
        "model": model,
        "messages": [],
        "temperature": payload.get("temperature", 0.7),
        "max_tokens": payload.get("max_tokens", 1024),
    }
    if payload.get("top_p") is not None:
        body["top_p"] = payload["top_p"]
    if payload.get("system_prompt", "").strip():
        body["messages"].append({"role": "system", "content": payload["system_prompt"]})

    user_text = payload.get("user_prompt", "")
    parts = build_content_parts(payload.get("attachments"))
    if parts:
        content = []
        if user_text:
            content.append({"type": "text", "text": user_text})
        content.extend(parts)
        body["messages"].append({"role": "user", "content": content})
    else:
        body["messages"].append({"role": "user", "content": user_text})

    start = time.time()
    try:
        resp = requests.post(
            payload["base_url"] + "/chat/completions",
            headers=get_headers(payload["api_key"]),
            json=body,
            timeout=TIMEOUT,
        )
        elapsed = round(time.time() - start, 2)
        data = resp.json()
        if resp.status_code != 200:
            msg = data.get("error", {})
            err = msg.get("message") if isinstance(msg, dict) else str(msg or data)
            return {"model": model, "ok": False, "error": err or f"HTTP {resp.status_code}", "elapsed": elapsed}
        choice = (data.get("choices") or [{}])[0]
        message = choice.get("message") or {}
        content = message.get("content")
        if content is None and message.get("reasoning_content"):
            content = message["reasoning_content"]
        return {
            "model": model,
            "ok": True,
            "content": content if content is not None else "(空响应)",
            "reasoning": message.get("reasoning_content") or "",
            "elapsed": elapsed,
            "usage": data.get("usage") or {},
        }
    except Exception as e:
        return {"model": model, "ok": False, "error": f"{type(e).__name__}: {e}", "elapsed": round(time.time() - start, 2)}


@app.route("/api/test", methods=["POST"])
def test():
    payload = request.get_json(force=True)
    try:
        payload["base_url"] = normalize_base_url(payload.get("base_url", ""))
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 200

    models = [m for m in payload.get("models", []) if m.strip()]
    if not models:
        return jsonify({"ok": False, "error": "请至少选择一个模型"}), 200
    if not payload.get("user_prompt", "").strip() and not payload.get("attachments"):
        return jsonify({"ok": False, "error": "请填写用户输入提示词，或添加附件"}), 200

    preprocess_attachments(payload)  # docx → PDF（一次性处理，多模型并发时复用）

    with ThreadPoolExecutor(max_workers=min(len(models), 8)) as pool:
        results = list(pool.map(lambda m: call_one_model(payload, m), models))
    return jsonify({"ok": True, "results": results})


if __name__ == "__main__":
    print("请在浏览器打开 http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)
