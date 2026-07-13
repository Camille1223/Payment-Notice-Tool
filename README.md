# 付款通知书生成器 · Payment Notice Generator

[English](#english) | [中文](#中文)

🌐 **在线使用 · Live Demo：[https://payment-notice-tool.onrender.com](https://payment-notice-tool.onrender.com)**

---

## 中文

### 简介

上传一份 Word 模板和一张 Excel 数据表，自动批量生成每位客户的付款通知书（Word + PDF），打包成 ZIP 一键下载。无需安装任何软件，浏览器直接使用。

### 功能特点

- 支持 `{字段名}` 占位符，自动填充 Excel 数据
- 同时输出 `.docx` 和 `.pdf`，格式与模板完全一致
- 实时进度条，生成过程可见
- 同一公司多行数据自动加 `_2`、`_3` 后缀，避免文件名冲突
- 无需安装，浏览器直接使用

### 使用方法

**第一步：准备 Word 模板**

将需要替换的内容改为 `{字段名}` 格式，例如：

```
致：{公司名称}

请于到期日前支付金额：{金额} 元
维护区间：{维护区间}
```

> 模板中的图片、印章、Logo 均会完整保留。

**第二步：准备 Excel 数据表**

- 使用第一个 Sheet
- 第一行为列名，列名须与模板中的字段名完全一致（不含花括号）
- 每行对应一份付款通知书

示例：

| 公司名称 | 金额 | 维护区间 |
|---------|------|---------|
| ABC 有限公司 | 100,000.00 | 2025–2026 |
| XYZ 股份公司 | 50,000.00 | 2025–2026 |

**第三步：生成下载**

1. 打开 [https://payment-notice-tool.onrender.com](https://payment-notice-tool.onrender.com)
2. 上传 Word 模板（`.docx`）
3. 上传数据表（`.xlsx`）
4. 点击「生成并下载 ZIP」

下载的 ZIP 结构：

```
付款通知书.zip
├── Word/
│   ├── ABC有限公司_付款通知书.docx
│   └── XYZ股份公司_付款通知书.docx
└── PDF/
    ├── ABC有限公司_付款通知书.pdf
    └── XYZ股份公司_付款通知书.pdf
```

### 注意事项

- Word 模板格式：`.docx`（不支持旧版 `.doc`）
- Excel 格式：`.xlsx` 或 `.xls`
- 首次访问可能需要等待约 30 秒（免费服务器冷启动），之后正常
- 文件大小上限：50 MB

### 本地运行

```bash
git clone https://github.com/Camille1223/Payment-Notice-Tool.git
cd Payment-Notice-Tool
pip install -r requirements.txt
python app.py
# 访问 http://127.0.0.1:5000
```

> 本地运行生成 PDF 需要安装 [LibreOffice](https://www.libreoffice.org/download/download/)。

### 技术栈

- **后端**：Python · Flask · python-docx · unoserver
- **前端**：原生 HTML/CSS/JS（无框架依赖）
- **部署**：Docker · Render

---

## English

### Overview

Upload a Word template and an Excel data sheet to automatically generate batch payment notices (Word + PDF) for each customer, packaged into a single ZIP download. No installation required — works entirely in the browser.

### Features

- `{FieldName}` placeholders automatically filled from Excel data
- Outputs both `.docx` and `.pdf` with formatting identical to the template
- Real-time progress bar during generation
- Duplicate company names automatically suffixed with `_2`, `_3`, etc.
- No installation needed — browser-based

### How to Use

**Step 1: Prepare the Word Template**

Replace dynamic content with `{FieldName}` placeholders, for example:

```
To: {CompanyName}

Please remit payment of {Amount} CNY before the due date.
Maintenance period: {Period}
```

> Images, stamps, and logos in the template are fully preserved.

**Step 2: Prepare the Excel Data Sheet**

- Use the first sheet
- First row must be column headers, matching the placeholder names exactly (without curly braces)
- Each subsequent row generates one payment notice

Example:

| CompanyName | Amount | Period |
|-------------|--------|--------|
| ABC Co. Ltd | 100,000.00 | 2025–2026 |
| XYZ Corp | 50,000.00 | 2025–2026 |

**Step 3: Generate and Download**

1. Open [https://payment-notice-tool.onrender.com](https://payment-notice-tool.onrender.com)
2. Upload the Word template (`.docx`)
3. Upload the data sheet (`.xlsx`)
4. Click **Generate & Download ZIP**

The downloaded ZIP is structured as:

```
payment-notices.zip
├── Word/
│   ├── ABC_Co_Ltd_payment_notice.docx
│   └── XYZ_Corp_payment_notice.docx
└── PDF/
    ├── ABC_Co_Ltd_payment_notice.pdf
    └── XYZ_Corp_payment_notice.pdf
```

### Notes

- Word template must be `.docx` (legacy `.doc` is not supported)
- Excel file: `.xlsx` or `.xls`
- First visit may take ~30 seconds to load (free-tier cold start); subsequent requests are fast
- Maximum file size: 50 MB

### Run Locally

```bash
git clone https://github.com/Camille1223/Payment-Notice-Tool.git
cd Payment-Notice-Tool
pip install -r requirements.txt
python app.py
# Open http://127.0.0.1:5000
```

> PDF generation requires [LibreOffice](https://www.libreoffice.org/download/download/) to be installed locally.

### Tech Stack

- **Backend**: Python · Flask · python-docx · unoserver
- **Frontend**: Vanilla HTML/CSS/JS (no framework dependencies)
- **Deployment**: Docker · Render
