# File Converter

Windows 本地文件转换工具 — Node.js Web UI + Word/WPS COM + Python PDF 引擎。

## 功能

| 转换 | 引擎 | 说明 |
|------|------|------|
| HTML → PDF | Word / WPS COM | 完美保真 |
| HTML → DOCX | html-to-docx | 纯 Node.js |
| DOCX → HTML | mammoth | 纯 Node.js |
| DOCX → PDF | Word / WPS COM | 完美保真 |
| PDF → DOCX | Python (pdf2docx / docling) | 自适应路由 + 字体保留 |
| Markdown → HTML | marked | 纯 Node.js |
| Markdown → PDF | marked + Word / WPS COM | 完美保真 |
| 批量 ZIP | 以上全部 | 上传 .zip，返回 .zip |
| 文本查找替换 | RegExp | 上传或粘贴 |

## 安装

```bash
# 1. Node.js
npm install

# 2. Python (PDF → DOCX 需要)
cd pdf_engine
pip install -r requirements.txt
```

## 启动

```bash
npm start
# → http://localhost:3456
```

## 使用

浏览器打开 `http://localhost:3456`，三个标签页：

- **Convert** — 所有格式转换，上传文件或粘贴内容
- **Batch** — 上传 .zip 批量转换，结果打包下载
- **Settings** — 鉴权 Token、PDF 主题 CSS / 页眉页脚

## PDF → DOCX 引擎

Python 引擎自动分析 PDF 类型并选择最佳策略：

| 类型 | 引擎 | 特点 |
|------|------|------|
| 普通电子 PDF | pdf2docx | 毫秒级，矢量字体保留 |
| 扫描件 / 纯图片 | docling | 本地视觉模型 OCR |
| 多栏 / 公式 / 表格 | docling | 语义结构提取 |

自动检测 PDF 主字体并映射系统可用字体（SimSun → 宋体、Helvetica → Arial 等）。

## 鉴权（可选）

```bash
set FILE_TOOL_TOKEN=your_token   # Windows CMD
$env:FILE_TOOL_TOKEN=your_token  # PowerShell
```

启用后所有接口需携带 `Authorization: Bearer your_token`。

## 命令行（Python）

```bash
cd pdf_engine

# 一个命令，自动识别格式
python convert.py input.pdf  output.docx    # PDF → DOCX
python convert.py input.docx output.pdf     # DOCX → PDF (Word COM)
python convert.py input.html output.pdf     # HTML → PDF (Word COM)
python convert.py input.pdf                 # 自动命名 output.docx
```

## 依赖

- Windows 10/11
- Node.js 18+
- Microsoft Word 或 WPS（HTML→PDF、DOCX→PDF 需要）
- Python 3.11+（PDF→DOCX 需要）
  - `pip install -r pdf_engine/requirements.txt`

## 项目结构

```
filechange/
├── server.js                       # Express 后端（110 行）
├── ui.html                         # Web UI
├── package.json
├── pdf_engine/
│   ├── convert.py                  # CLI 统一入口
│   ├── universal_pdf_converter.py  # PDF → DOCX 引擎
│   ├── docx_to_pdf.py             # DOCX → PDF (Word COM)
│   ├── html_to_pdf.py             # HTML → PDF (Word COM)
│   ├── bridge.js                   # Node.js ↔ Python
│   └── requirements.txt
├── uploads/                        # 临时上传（自动创建）
└── outputs/                        # 转换结果（自动创建）
```
