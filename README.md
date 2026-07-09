# 付款通知书生成器

🌐 **在线使用（无需安装）：[https://payment-notice-tool.onrender.com](https://payment-notice-tool.onrender.com)**

上传付款通知书 Word 模板和 Excel 数据表，自动批量生成 Word + PDF，打包成 ZIP 下载。

## 使用方法

1. 打开 [https://payment-notice-tool.onrender.com](https://payment-notice-tool.onrender.com)
2. 上传付款通知书 Word 模板（`.docx`）
3. 上传数据表（`.xlsx`，Sheet1 第一行为列名）
4. 点击「生成并下载 ZIP」

ZIP 结构：
```
Word/  ← 每行数据一个 .docx
PDF/   ← 每行数据一个 .pdf
```

## 模板字段格式

将 Word 模板中的合并字段改为 `{字段名}`，例如 `{公司名称}`、`{金额}`。Excel 第一行列名需与模板字段一致。

同一家公司有多行时，自动加 `_2`、`_3` 后缀。

---

> 首次访问可能需要等待约 30 秒（免费服务器冷启动），之后正常。
