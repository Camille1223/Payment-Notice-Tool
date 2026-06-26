# 付款通知书生成器

**在线工具（点开即用）：https://camille1223.github.io/Payment-Notice-Tool/**

上传付款通知书 Word 模板和 Excel 数据表，自动批量生成 Word + PDF 文件，打包成 ZIP 下载。**无需安装任何软件，浏览器直接使用。**

---

## 使用方法

### 第一步：准备 Word 模板

将模板中的合并字段改为 `{字段名}` 格式，例如：

| 原格式（Mail Merge） | 新格式 |
|---|---|
| `«公司名称»` | `{公司名称}` |
| `«金额»` | `{金额}` |
| `«日期»` | `{日期}` |

### 第二步：准备 Excel 数据表

- 使用 **Sheet1**
- **第一行为列名**，列名需与模板字段完全一致
- 每一行对应一份付款通知书

示例：

| 公司名称 | 金额 | 日期 |
|---|---|---|
| ABC 贸易有限公司 | 50,000 | 2026-06-26 |
| XYZ 科技有限公司 | 30,000 | 2026-06-26 |

### 第三步：上传并生成

1. 打开 https://camille1223.github.io/Payment-Notice-Tool/
2. 上传 Word 模板（`.docx`）
3. 上传数据表（`.xlsx`）
4. 点击「生成并下载 ZIP」

### 输出结果

ZIP 包含两个文件夹：

```
Word/  <- 每行数据一个 .docx
PDF/   <- 每行数据一个 .pdf
```

同一公司有多行时，自动添加 `_2`、`_3` 后缀区分。

---

## 技术架构

| 组件 | 说明 |
|---|---|
| 前端 | 纯静态 HTML + JS，托管于 GitHub Pages |
| Word 生成 | docxtemplater，在浏览器本地完成 |
| PDF 转换 | 后端 Flask API + LibreOffice，托管于 Render |
| 数据读取 | SheetJS，在浏览器本地完成 |

---

## 本地开发

如需本地运行后端：

```bash
pip install flask flask-cors gunicorn
python app.py
```

需要系统安装 LibreOffice（用于 DOCX 转 PDF）。
