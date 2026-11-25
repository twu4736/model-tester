# Model Tester — 模型输入输出测试工具

一个轻量级的本地 Web 工具，用于测试大语言模型（LLM）的输入输出。支持自定义系统提示词、用户提示词，选择/对比多个模型，并支持**多模态输入**（图片、音频、PDF、Word 文档）。

## 功能特性

- 🔌 **OpenAI 兼容接口**：支持 OpenAI、智谱 GLM、DeepSeek、通义、Kimi、Ollama 等所有兼容 OpenAI 格式的服务
- 📋 **模型列表**：一键拉取可用模型，点击选择；单击卡片选中（可多选对比），双击复制模型名
- ✍️ **提示词自定义**：系统提示词 + 用户提示词，支持调节 `temperature` / `max_tokens` / `top_p`
- 🖼️ **多模态输入**：
  - 图片（png/jpg/gif/webp/bmp）→ `image_url` 格式
  - 音频（wav/mp3）→ `input_audio` 格式
  - PDF → `file` 格式（OpenAI 兼容接口文件输入的标准格式）
  - **Word 文档（.docx）** → 自动调用本机 WPS/Word 转成 PDF 后原样发送，模型可见完整排版；转换失败自动降级为文本提取
  - 文本文件（txt/md/csv/json/代码文件等）→ 自动提取文本发送
- ⚡ **多模型并行对比**：多个模型同时请求，结果并排展示，含耗时与 token 用量
- 💾 **配置持久化**：提示词、接口配置自动保存在浏览器本地

## 环境要求

- Python 3.8+（Windows 下已处理中文主机名导致的 Flask 启动问题）
- Flask、requests（`pip install flask requests`）
- （可选）Word 文档转 PDF 功能需要本机安装 WPS Office 或 Microsoft Word

## 快速开始

```bash
cd model-tester
python app.py
```

然后浏览器打开 http://127.0.0.1:5000

1. 填写 Base URL 和 API Key（API Key 仅保存在您浏览器本地，不会上传到任何服务器）
2. 点击「获取模型列表」，选择要测试的模型
3. 填写系统提示词和用户提示词（可拖入图片/音频/文档作为附件）
4. 点击「运行测试」，查看并对比各模型的输出

## 项目结构

```
model-tester/
├── app.py              # Flask 后端：接口转发、多模态消息组装、docx→PDF 转换
└── templates/
    └── index.html      # 前端页面（单文件，无构建依赖）
```

## 注意事项

- 附件总大小限制 100MB；base64 编码后体积会膨胀约 1/3，建议图片单张控制在几 MB 内
- 多模态输入要求所选模型支持对应模态，不支持的模型会在结果卡片中显示具体报错，不影响其他模型
- `file` 类型输入（PDF）是否可用取决于具体服务商：OpenAI 官方支持，部分第三方服务不支持（会返回报错）
- Word 转 PDF 每次约需 1~2 秒（启动 WPS/Word 进程）；多模型并发测试时只转换一次

## License

[MIT](LICENSE)
