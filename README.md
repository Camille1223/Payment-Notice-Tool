# 付款通知书生成器 · Payment Notice Generator

[English](#english) | [中文](#中文)

🌐 **在线使用 · Live Demo：[https://payment-notice-tool.onrender.com](https://payment-notice-tool.onrender.com)**

> 更新代码后需重新部署；网站首页显示「付款通知书 PDF 模板 (.pdf)」即表示已启用新版上传流程。Mac 可双击 `启动.command`，打开 http://127.0.0.1:5050 。PDF 模板模式不需要自动化控制 Word。
>
> 修复版提供明确的断线/任务丢失提示，限制 HTML→PDF 单份转换为 60 秒，串行处理生成任务以降低内存压力；PDF 转换不完整时会提示警告，ZIP 仍保留全部 Word 文件。

### 推荐：PDF 模板原版模式

在 Word 中将**带 `{字段名}` 占位符的原始模板**导出为 PDF，再上传 PDF + Excel。固定文字、Logo、印章、表格与页面尺寸保留，不再经过 HTML 重排。PDF 模式只输出 PDF；动态字段缺少原字体字形时会提示兼容字体，内容放不下时明确报错。

网页上传区与 `/start` 接口现已只接受 PDF 模板。旧 `/generate` 同步接口保留 Word 兼容处理，但不再默认启用会改变版式的 HTML 转换。

完整说明和部署注意事项见 [PDF模板使用说明.md](PDF模板使用说明.md)，包括 PyMuPDF 授权确认要求。

---

## 中文

### 简介

将带占位符的 Word 模板导出为 PDF，再上传 PDF 模板和 Excel 数据表，批量生成每位客户的 PDF 付款通知书，打包成 ZIP 一键下载。PDF 原版模式保留固定页面内容，不再通过 HTML 重排。

### 功能特点

- 支持 `{字段名}` 占位符，自动填充 Excel 数据
- 网页使用 PDF 模板，输出 `.pdf`；Logo、印章、固定文字和表格版式保留
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

> 在 Word 中将模板导出为 PDF，检查图片、印章和 Logo 位置，然后上传这份 PDF。占位符文字需能选中。

**第二步：准备 Excel 数据表**

- 使用当前选中的工作表
- 第一行为列名，列名须与模板中的字段名完全一致（不含花括号）
- 每行对应一份付款通知书

示例：

| 公司名称 | 金额 | 维护区间 |
|---------|------|---------|
| ABC 有限公司 | 100,000.00 | 2025–2026 |
| XYZ 股份公司 | 50,000.00 | 2025–2026 |

**第三步：生成下载**

1. 打开 [https://payment-notice-tool.onrender.com](https://payment-notice-tool.onrender.com)
2. 上传从 Word 导出的 PDF 模板（`.pdf`，保留占位符）
3. 上传数据表（`.xlsx`）
4. 点击「生成并下载 ZIP」

下载的 ZIP 结构：

```
付款通知书.zip
└── PDF/
    ├── 付款通知_ABC有限公司.pdf
    └── 付款通知_XYZ股份公司.pdf
```

### 注意事项

- 网页模板格式：`.pdf`（请先在 Word 中导出；不接受 Word 文件、扫描件或成品通知书）
- Excel 格式：`.xlsx`（旧 `.xls` 请先另存为 `.xlsx`）
- 首次访问可能需要等待约 30 秒（免费服务器冷启动），之后正常
- 文件大小上限：50 MB

### 本地运行

```bash
git clone https://github.com/Camille1223/Payment-Notice-Tool.git
cd Payment-Notice-Tool
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python app.py
# 访问 http://127.0.0.1:5050；可用 PORT 环境变量更改端口
```

> 本地 PDF 转换优先使用 Microsoft Word（Windows/macOS）或 [LibreOffice](https://www.libreoffice.org/download/download/)（跨平台）。Mac 需允许自动化控制 Word；HTML 备用引擎还需要 Pango 等系统库。未能生成 PDF 时仍可下载 Word 文件。

### 技术栈

- **后端**：Python · Flask · python-docx · unoserver
- **前端**：原生 HTML/CSS/JS（无框架依赖）
- **部署**：Docker · Render

---

## English

### Overview

Export your Word template to a text-based PDF with placeholders, then upload the PDF and Excel data to generate payment notices as a ZIP. The PDF template's fixed layout and artwork remain unchanged. The web upload accepts PDF templates only; legacy Word handling remains in the synchronous API.

### Features

- `{FieldName}` placeholders automatically filled from Excel data
- Outputs PDF files while retaining the PDF template's fixed text, artwork and tables
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

> Export the original Word template to PDF, keeping selectable placeholders and checking the artwork before uploading.

**Step 2: Prepare the Excel Data Sheet**

- Use the active worksheet
- First row must be column headers, matching the placeholder names exactly (without curly braces)
- Each subsequent row generates one payment notice

Example:

| CompanyName | Amount | Period |
|-------------|--------|--------|
| ABC Co. Ltd | 100,000.00 | 2025–2026 |
| XYZ Corp | 50,000.00 | 2025–2026 |

**Step 3: Generate and Download**

1. Open [https://payment-notice-tool.onrender.com](https://payment-notice-tool.onrender.com)
2. Upload the exported PDF template (`.pdf`)
3. Upload the data sheet (`.xlsx`)
4. Click **Generate & Download ZIP**

The downloaded ZIP is structured as:

```
payment-notices.zip
└── PDF/
    ├── 付款通知_ABC_Co_Ltd.pdf
    └── 付款通知_XYZ_Corp.pdf
```

### Notes

- The web template must be a text-based `.pdf` with placeholders, not a scan or an already-filled notice
- Excel file: `.xlsx` (save legacy `.xls` files as `.xlsx` first)
- First visit may take ~30 seconds to load (free-tier cold start); subsequent requests are fast
- Maximum file size: 50 MB

### Run Locally

```bash
git clone https://github.com/Camille1223/Payment-Notice-Tool.git
cd Payment-Notice-Tool
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python app.py
# Open http://127.0.0.1:5050
```

> PDF generation prefers Microsoft Word (Windows/macOS) or LibreOffice. Allow Word automation on macOS. The HTML fallback requires native Pango libraries and may change layout. All Word files remain downloadable if PDF conversion fails. The hosted service must be redeployed to pick up these fixes.

### Tech Stack

- **Backend**: Python · Flask · python-docx · unoserver
- **Frontend**: Vanilla HTML/CSS/JS (no framework dependencies)
- **Deployment**: Docker · Render
