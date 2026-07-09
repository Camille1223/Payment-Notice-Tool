# 付款通知书生成器

网页 Dashboard：上传付款通知书 Word 模板和 Excel 数据表，自动生成对应数量的 Word + PDF 文件，打包成 ZIP 下载。

## 依赖环境

- Python 3.8+
- Microsoft Word（用于生成 PDF）
- Windows 系统

## 安装依赖

```bash
pip install flask openpyxl docx-mailmerge2 pywin32
```

## 启动方式

双击 `启动.bat`，或命令行运行：

```bash
python app.py
```

然后浏览器打开：http://127.0.0.1:5000

## 使用方法

1. 上传付款通知书 Word 模板（`.docx`，含 Mail Merge 字段）
2. 上传数据表（`.xlsx`，Sheet1 第一行为列名，需与模板字段对应）
3. 点击「生成并下载 ZIP」

ZIP 结构：
```
Word/  ← 每行数据一个 .docx
PDF/   ← 每行数据一个 .pdf
```

同一家公司有多行时，自动加 `_2`、`_3` 后缀。
