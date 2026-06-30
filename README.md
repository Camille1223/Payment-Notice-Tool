<div align="center">

# 付款通知书生成器 · Payment Notice Generator

**批量生成付款通知书，Word + PDF 打包下载，无需安装任何软件**  
**Batch-generate payment notices as Word + PDF, packed into a single ZIP — no installation required**

[![Deploy on Render](https://img.shields.io/badge/Backend-Render-46E3B7?logo=render&logoColor=white)](https://payment-notice-tool.onrender.com)
[![GitHub Pages](https://img.shields.io/badge/Frontend-GitHub%20Pages-blue?logo=github)](https://camille1223.github.io/Payment-Notice-Tool/)
[![Python](https://img.shields.io/badge/Python-3.12-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

</div>

---

## 目录 · Contents

- [功能特点 · Features](#功能特点--features)
- [在线使用 · Live Demo](#在线使用--live-demo)
- [工作原理 · How It Works](#工作原理--how-it-works)
- [快速上手 · Quick Start](#快速上手--quick-start)
- [本地运行 · Local Development](#本地运行--local-development)
- [Docker 部署 · Docker Deployment](#docker-部署--docker-deployment)
- [项目结构 · Project Structure](#项目结构--project-structure)

---

## 功能特点 · Features

| 功能 | Feature |
|------|---------|
| 📄 上传 `.docx` 模板，自动替换 `{字段名}` 占位符 | Upload a `.docx` template; `{field}` placeholders are replaced automatically |
| 📊 读取 `.xlsx` 数据表，每行生成一份文件 | Read an `.xlsx` spreadsheet; one file is generated per row |
| 📦 Word + PDF 双格式打包为单个 ZIP 下载 | Word + PDF output bundled into a single ZIP download |
| 🌐 纯浏览器前端，无需后端即可生成 Word | Pure-browser frontend; Word generation requires no backend |
| 🖨️ PDF 转换由云端 LibreOffice 服务完成，中文字体完整 | PDF conversion via cloud LibreOffice with full CJK font support |
| 🚀 无需安装 Word / WPS，任意操作系统均可使用 | No Word or WPS needed — works on any OS |
| 🔒 文件在本地或云端临时处理，不持久化存储 | Files are processed locally or in ephemeral cloud storage — never persisted |

---

## 在线使用 · Live Demo

直接访问 GitHub Pages，零配置开箱即用：

> **[https://camille1223.github.io/Payment-Notice-Tool/](https://camille1223.github.io/Payment-Notice-Tool/)**

Open the GitHub Pages link above — no sign-up, no installation.

---

## 工作原理 · How It Works

```
用户浏览器 (GitHub Pages)               云端后端 (Render)
──────────────────────────             ──────────────────
  ① 上传 .docx 模板                      ④ 接收 .docx
  ② 上传 .xlsx 数据表                    ⑤ LibreOffice 转换 PDF
  ③ docxtemplater 批量生成 Word  ──────► ⑥ 返回 .pdf 二进制流
  ⑦ JSZip 打包 Word/ + PDF/
  ⑧ 浏览器触发下载 付款通知书.zip
```

**前端 (Frontend)**  — `docs/index.html`（纯静态，托管于 GitHub Pages）
- [docxtemplater](https://docxtemplater.com/) 在浏览器内渲染 Word 模板
- [SheetJS (xlsx)](https://sheetjs.com/) 解析 Excel 数据
- [JSZip](https://stuk.github.io/jszip/) 打包最终 ZIP

**后端 (Backend)** — `app.py`（Flask，托管于 Render）
- 接收 `.docx`，调用 `LibreOffice --headless` 转换为 PDF
- Docker 镜像内置 `fonts-noto-cjk`，确保中文字符正确渲染

---

## 快速上手 · Quick Start

### 1. 准备 Word 模板 · Prepare the Word Template

在 `.docx` 文件中，将需要替换的内容改为 `{字段名}` 格式：

```
尊敬的 {公司名称}：

根据合同约定，请于 {付款日期} 前支付款项 ¥{金额}。

收款账户：{收款账户}
```

> **字段名必须与 Excel 表头完全一致（区分大小写）。**  
> Field names must exactly match Excel column headers (case-sensitive).

### 2. 准备 Excel 数据表 · Prepare the Excel Spreadsheet

Sheet1 第一行为列名，每一行对应一份通知书：

| 公司名称 | 付款日期 | 金额 | 收款账户 |
|----------|----------|------|----------|
| 示例科技有限公司 | 2025-07-31 | 50,000.00 | 工行 6222... |
| 测试贸易有限公司 | 2025-08-15 | 12,800.00 | 招行 6225... |

### 3. 上传并生成 · Upload and Generate

1. 打开 [在线工具](https://camille1223.github.io/Payment-Notice-Tool/)
2. 上传 `.docx` 模板和 `.xlsx` 数据表
3. 点击 **"生成并下载 ZIP"**
4. 解压后即得 `Word/` 和 `PDF/` 两个文件夹

---

## 本地运行 · Local Development

### 前提 · Prerequisites

- Python 3.10+
- LibreOffice（用于 PDF 转换）
- `pip install -r requirements.txt`

### 启动 · Start

**Windows（双击）：**
```
启动.bat
```

**命令行：**
```bash
pip install -r requirements.txt
python app.py
```

服务默认监听 `http://localhost:8000`。  
The service listens on `http://localhost:8000` by default.

前端打开 `docs/index.html`，将页面内的 `PDF_API` 改为 `http://localhost:8000/pdf` 即可本地调试。  
Open `docs/index.html` and change `PDF_API` to `http://localhost:8000/pdf` for local testing.

---

## Docker 部署 · Docker Deployment

镜像已内置 LibreOffice 及 CJK 字体，开箱即用：  
The image bundles LibreOffice and CJK fonts out of the box:

```bash
# 构建 · Build
docker build -t payment-notice-tool .

# 运行 · Run
docker run -p 8000:8000 payment-notice-tool
```

推荐使用 [Render](https://render.com/) 免费部署：将本仓库连接至 Render，选择 Docker 环境，自动完成构建与发布。  
Recommended free hosting: connect this repo to [Render](https://render.com/), select Docker runtime, and deploy automatically.

---

## 项目结构 · Project Structure

```
Payment-Notice-Tool/
├── docs/
│   └── index.html          # 前端页面 (GitHub Pages) · Frontend (GitHub Pages)
├── app.py                  # Flask PDF 转换后端 · Flask PDF conversion backend
├── requirements.txt        # Python 依赖 · Python dependencies
├── Dockerfile              # 含 LibreOffice + CJK 字体 · Includes LibreOffice + CJK fonts
├── 启动.bat                # Windows 一键启动 · Windows one-click launcher
└── README.md
```

---

## API 参考 · API Reference

### `POST /pdf`

将 `.docx` 文件转换为 PDF。  
Converts a `.docx` file to PDF.

**请求 · Request**

```
Content-Type: multipart/form-data
Field: file  (.docx binary)
```

**响应 · Response**

| 状态码 | 说明 |
|--------|------|
| `200 OK` | PDF 二进制流 · PDF binary stream |
| `400 Bad Request` | 缺少文件字段 · Missing file field |
| `500 Internal Server Error` | 转换失败 / LibreOffice 未安装 · Conversion failed / LibreOffice not found |

---

## License

MIT © [Camille1223](https://github.com/Camille1223)
