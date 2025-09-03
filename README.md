# File Converter Tool

跨平台本地文件转换与编辑工具（Node + Express + Web UI）

功能
- HTML → PDF（Puppeteer）
- HTML → DOCX（html-to-docx）
- DOCX → HTML（mammoth）
- Markdown → HTML（marked）
- Markdown → PDF（Puppeteer）
- 文本/HTML 查找替换
- 批量 ZIP 转换（.html → pdf/docx、.md → html/pdf、.docx → html）

快速开始
1. 安装依赖并启动
```
npm install
npm start
```
2. 打开浏览器访问 `http://localhost:3456`

使用说明
- 可上传文件，或直接在文本框粘贴内容
- 转换完成后将自动下载输出文件
- 批量：上传压缩包（zip），结果将打包 zip 返回

系统要求
- Node.js 18+
- macOS/Windows/Linux 皆可运行

PDF 样式与页眉/页脚
- 可在 UI 中输入 Theme CSS、Header/ Footer HTML
- 默认 Footer 包含页码：`<span class="pageNumber"></span>/<span class="totalPages"></span>`

目录结构
- `server.js` 后端服务
- `ui.html` 简易 Web UI
- `uploads/` 临时上传目录（自动创建）
- `outputs/` 转换结果目录（自动创建）

开发
- 端口默认 `3456`
- 可扩展新增路由实现更多格式（如 Markdown → PDF 等）

可选鉴权
- 通过环境变量启用：`export FILE_TOOL_TOKEN=your_token`
- 所有接口将要求 Header：`Authorization: Bearer your_token`

Docker
```
docker build -t file-converter .
docker run -p 3456:3456 -e FILE_TOOL_TOKEN=your_token file-converter
```


