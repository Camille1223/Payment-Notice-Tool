import os
import io
import zipfile
import tempfile

from flask import Flask, request, send_file, render_template_string, jsonify
from mailmerge import MailMerge
import openpyxl

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

HTML = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>付款通知书生成器</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    background: #f0f2f5;
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 24px;
  }
  .card {
    background: #fff;
    border-radius: 12px;
    box-shadow: 0 2px 16px rgba(0,0,0,0.1);
    padding: 40px;
    width: 100%;
    max-width: 560px;
  }
  .logo { font-size: 26px; font-weight: 700; color: #1a73e8; margin-bottom: 6px; }
  .subtitle { color: #666; font-size: 14px; margin-bottom: 28px; }
  .upload-area {
    border: 2px dashed #d0d7de;
    border-radius: 8px;
    padding: 20px;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.2s, background 0.2s;
    margin-bottom: 14px;
    position: relative;
  }
  .upload-area:hover, .upload-area.drag-over {
    border-color: #1a73e8;
    background: #f6f9ff;
  }
  .upload-area input[type=file] {
    position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%;
  }
  .upload-icon { font-size: 28px; margin-bottom: 6px; }
  .upload-label { font-size: 14px; color: #444; font-weight: 500; }
  .upload-hint { font-size: 12px; color: #999; margin-top: 3px; }
  .file-badge {
    display: none;
    margin-top: 8px;
    padding: 3px 10px;
    background: #e8f0fe;
    border-radius: 4px;
    font-size: 12px;
    color: #1a73e8;
    word-break: break-all;
  }
  .btn {
    width: 100%;
    padding: 13px;
    background: #1a73e8;
    color: #fff;
    border: none;
    border-radius: 8px;
    font-size: 15px;
    font-weight: 600;
    cursor: pointer;
    margin-top: 6px;
    transition: background 0.2s;
  }
  .btn:hover:not(:disabled) { background: #1557b0; }
  .btn:disabled { background: #a8c7fa; cursor: not-allowed; }
  .progress-bar {
    height: 4px; background: #e8eaed; border-radius: 2px;
    margin-top: 14px; overflow: hidden; display: none;
  }
  .progress-fill {
    height: 100%; background: #1a73e8; border-radius: 2px;
    transition: width 0.4s ease; width: 0%;
  }
  .status {
    margin-top: 14px; padding: 11px 14px; border-radius: 8px;
    font-size: 13px; display: none; line-height: 1.5;
  }
  .status.info    { background: #e8f0fe; color: #1a73e8; }
  .status.success { background: #e6f4ea; color: #137333; }
  .status.error   { background: #fce8e6; color: #c5221f; }
  hr { border: none; border-top: 1px solid #eee; margin: 24px 0; }
</style>
</head>
<body>
<div class="card">
  <div class="logo">付款通知书生成器</div>
  <div class="subtitle">上传模板与数据表，自动生成 Word + PDF，打包下载</div>
  <hr>

  <div class="upload-area" id="wordArea">
    <input type="file" id="wordFile" accept=".docx" onchange="setFile('word',this)">
    <div class="upload-icon">📄</div>
    <div class="upload-label">付款通知书模板 (.docx)</div>
    <div class="upload-hint">点击或拖拽上传</div>
    <div class="file-badge" id="wordName"></div>
  </div>

  <div class="upload-area" id="excelArea">
    <input type="file" id="excelFile" accept=".xlsx,.xls" onchange="setFile('excel',this)">
    <div class="upload-icon">📊</div>
    <div class="upload-label">付款通知数据表 (.xlsx)</div>
    <div class="upload-hint">使用 Sheet1，第一行为列名</div>
    <div class="file-badge" id="excelName"></div>
  </div>

  <button class="btn" id="btn" disabled onclick="generate()">生成并下载 ZIP</button>

  <div class="progress-bar" id="bar"><div class="progress-fill" id="fill"></div></div>
  <div class="status" id="status"></div>
</div>
<script>
function setFile(type, input) {
  const f = input.files[0];
  if (!f) return;
  const badge = document.getElementById(type + 'Name');
  badge.textContent = f.name;
  badge.style.display = 'inline-block';
  document.getElementById('btn').disabled =
    !(document.getElementById('wordFile').files[0] && document.getElementById('excelFile').files[0]);
}

['wordArea','excelArea'].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('drag-over'); });
  ['dragleave','drop'].forEach(ev => el.addEventListener(ev, () => el.classList.remove('drag-over')));
});

function showStatus(msg, type) {
  const s = document.getElementById('status');
  s.innerHTML = msg;
  s.className = 'status ' + type;
  s.style.display = 'block';
}

async function generate() {
  const btn = document.getElementById('btn');
  const bar = document.getElementById('bar');
  const fill = document.getElementById('fill');
  btn.disabled = true;
  btn.textContent = '生成中，请稍候…';
  bar.style.display = 'block';
  fill.style.width = '20%';
  showStatus('正在上传并处理文件…', 'info');

  const form = new FormData();
  form.append('template', document.getElementById('wordFile').files[0]);
  form.append('data', document.getElementById('excelFile').files[0]);

  try {
    fill.style.width = '55%';
    const resp = await fetch('/generate', { method: 'POST', body: form });
    fill.style.width = '85%';

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({ error: resp.statusText }));
      throw new Error(err.error || '服务器错误');
    }

    const disp = resp.headers.get('Content-Disposition') || '';
    const m = disp.match(/filename\\*?=(?:UTF-8'')?[\"']?([^\"'\\n]+)/i);
    const filename = m ? decodeURIComponent(m[1]) : '付款通知书.zip';

    const blob = await resp.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);

    fill.style.width = '100%';
    showStatus('✅ 生成成功！ZIP 已开始下载。', 'success');
  } catch(e) {
    showStatus('❌ ' + e.message, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = '生成并下载 ZIP';
    setTimeout(() => { bar.style.display = 'none'; fill.style.width = '0%'; }, 1500);
  }
}
</script>
</body>
</html>
"""


def excel_rows(xlsx_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb["Sheet1"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Sheet1 为空")
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = []
    for row in rows[1:]:
        if not any(v is not None for v in row):
            continue
        record = {}
        for h, v in zip(headers, row):
            if h:
                record[h] = str(v).strip() if v is not None else ""
        data.append(record)
    return data


def safe_filename(name: str) -> str:
    return "".join(c for c in name if c not in r'\/:*?"<>|').strip() or "未知"


def batch_docx_to_pdf(pairs: list[tuple[str, str]]):
    """Convert a list of (docx_path, pdf_path) pairs using a single Word instance.
    Initializes COM on the current thread (required when called from a Flask worker thread).
    """
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    errors = []
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        for docx_path, pdf_path in pairs:
            doc = None
            try:
                doc = word.Documents.Open(os.path.abspath(docx_path))
                doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
            except Exception as e:
                errors.append((os.path.basename(docx_path), str(e)))
            finally:
                if doc:
                    doc.Close(False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()
    return errors


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/generate", methods=["POST"])
def generate():
    if "template" not in request.files or "data" not in request.files:
        return jsonify(error="缺少上传文件"), 400

    template_bytes = request.files["template"].read()
    data_bytes = request.files["data"].read()

    try:
        rows = excel_rows(data_bytes)
    except Exception as e:
        return jsonify(error=f"读取 Excel 失败：{e}"), 400

    if not rows:
        return jsonify(error="Sheet1 没有数据行"), 400

    name_count: dict[str, int] = {}
    zip_buffer = io.BytesIO()

    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_pairs: list[tuple[str, str]] = []

        # Generate all Word files first
        for i, row in enumerate(rows, 1):
            company = row.get("公司名称", f"客户{i}")
            base = safe_filename(company) + "_付款通知书"
            name_count[base] = name_count.get(base, 0) + 1
            count = name_count[base]
            unique_name = base if count == 1 else f"{base}_{count}"

            docx_out = os.path.join(tmpdir, f"{unique_name}.docx")
            with MailMerge(io.BytesIO(template_bytes)) as doc:
                doc.merge(**{k: str(v) for k, v in row.items()})
                doc.write(docx_out)

            pdf_out = os.path.join(tmpdir, f"{unique_name}.pdf")
            pdf_pairs.append((docx_out, pdf_out))

        # Convert all to PDF in one Word session
        pdf_errors = batch_docx_to_pdf(pdf_pairs)
        if pdf_errors:
            for name, err in pdf_errors:
                app.logger.warning("PDF error %s: %s", name, err)

        # Pack everything into ZIP
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for docx_out, pdf_out in pdf_pairs:
                base = os.path.basename(docx_out)
                name_no_ext = base[:-5]
                zf.write(docx_out, f"Word/{base}")
                if os.path.exists(pdf_out):
                    zf.write(pdf_out, f"PDF/{name_no_ext}.pdf")

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name="付款通知书.zip",
    )


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    print("=" * 40)
    print("Dashboard: http://127.0.0.1:5000")
    print("=" * 40)
    app.run(debug=False, port=5000)
