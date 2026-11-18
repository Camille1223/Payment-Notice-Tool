from __future__ import annotations

import os
import io
import re
import zipfile
import tempfile
import threading
import uuid
import time
import json
import shutil
import subprocess
import sys
from functools import lru_cache
from copy import deepcopy

from flask import Flask, request, send_file, render_template_string, jsonify, Response, stream_with_context
from docx import Document
from docx.oxml.ns import qn
import openpyxl

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

# ──────────────────────────────────────────────
# File-based job store — survives across gunicorn
# workers and process restarts.  Each job is a
# small JSON file in JOB_DIR; zip payload written
# alongside as <job_id>.zip.
# ──────────────────────────────────────────────
_default_job_dir = (os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs", "jobs")
                    if sys.platform == "darwin"
                    else os.path.join(tempfile.gettempdir(), "pnt_jobs"))
JOB_DIR = os.environ.get("JOB_DIR", _default_job_dir)
os.makedirs(JOB_DIR, mode=0o700, exist_ok=True)

_meta_lock = threading.Lock()  # guards same-process concurrent writes only
_conversion_lock = threading.Lock()  # avoid simultaneous PDF engines exhausting memory


def _write_meta(job_id: str, data: dict):
    data["updated_at"] = time.time()
    # Readers must never see a half-written JSON file.
    with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=JOB_DIR,
                                     delete=False) as f:
        json.dump(data, f, ensure_ascii=False)
        tmp_path = f.name
    os.replace(tmp_path, _meta_path(job_id))


def _meta_path(job_id: str) -> str:
    return os.path.join(JOB_DIR, f"{job_id}.json")


def _zip_path(job_id: str) -> str:
    return os.path.join(JOB_DIR, f"{job_id}.zip")


def job_create(job_id: str, **kw):
    data = {"pct": 5, "label": "已收到文件…", "done": False}
    data.update(kw)
    with _meta_lock:
        _write_meta(job_id, data)


def job_update(job_id: str, **kw):
    import json
    path = _meta_path(job_id)
    with _meta_lock:
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
        data.update(kw)
        # never persist raw zip bytes into JSON — stored in separate file
        data.pop("zip_data", None)
        _write_meta(job_id, data)


def job_read(job_id: str) -> dict:
    import json
    try:
        with open(_meta_path(job_id), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


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
  .tip { background: #fffbe6; border: 1px solid #ffe58f; border-radius: 6px; padding: 10px 14px; font-size: 12px; color: #7d6608; margin-bottom: 14px; line-height: 1.6; }
  .tip code { background: #fff3d0; padding: 1px 5px; border-radius: 3px; font-family: monospace; }
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
  .progress-wrap { margin-top: 14px; display: none; }
  .progress-bar {
    height: 6px; background: #e8eaed; border-radius: 3px;
    overflow: hidden;
  }
  .progress-fill {
    height: 100%; background: #1a73e8; border-radius: 3px;
    transition: width 0.3s ease; width: 0%;
  }
  .progress-text { font-size: 12px; color: #666; margin-top: 5px; text-align: right; }
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
  <div class="subtitle">上传 PDF 模板与 Excel 数据，保留原版式，批量生成 PDF</div>
  <hr>
  <div class="tip">
    <b>模板字段格式：</b>将 Word 模板中需要替换的内容改为 <code>{字段名}</code>，例如 <code>{公司名称}</code>、<code>{金额}</code>。
    Excel 当前工作表第一行列名需与模板字段名一致（不含花括号）。<br>
    <b>原版模式：</b>在 Word 中将带占位符的模板导出为 PDF，再上传 PDF + Excel。Logo、印章、颜色和表格不再经过 HTML 重排。请使用文字可选中的 PDF，而非扫描件。<br>
    <b>上传格式：</b>本页面只接受 PDF 模板。请先在 Word 中导出 PDF，不要直接上传 Word，也不要上传已填好数据的通知书。
  </div>

  <div class="upload-area" id="templateArea">
    <input type="file" id="templateFile" accept=".pdf,application/pdf" onchange="setFile('template',this)">
    <div class="upload-icon">📄</div>
    <div class="upload-label">付款通知书 PDF 模板 (.pdf)</div>
    <div class="upload-hint">PDF 需保留 {字段名} 占位符；原版模式只输出 PDF</div>
    <div class="file-badge" id="templateName"></div>
  </div>

  <div class="upload-area" id="excelArea">
    <input type="file" id="excelFile" accept=".xlsx" onchange="setFile('excel',this)">
    <div class="upload-icon">📊</div>
    <div class="upload-label">付款通知数据表 (.xlsx)</div>
    <div class="upload-hint">使用当前工作表，第一行为列名；旧 .xls 请先另存为 .xlsx</div>
    <div class="file-badge" id="excelName"></div>
  </div>

  <button class="btn" id="btn" disabled onclick="generate()">生成并下载 ZIP</button>

  <div class="progress-wrap" id="progressWrap">
    <div class="progress-bar"><div class="progress-fill" id="fill"></div></div>
    <div class="progress-text" id="progressText">0%</div>
  </div>
  <div class="status" id="status"></div>
</div>
<script>
function setFile(type, input) {
  const f = input.files[0];
  if (!f) return;
  if (type === 'template' && !f.name.toLowerCase().endsWith('.pdf')) {
    input.value = '';
    document.getElementById('templateName').style.display = 'none';
    document.getElementById('btn').disabled = true;
    showStatus('请上传 PDF 模板（.pdf）。Word 模板需先导出为 PDF。', 'error');
    return;
  }
  const badge = document.getElementById(type + 'Name');
  badge.textContent = f.name;
  badge.style.display = 'inline-block';
  document.getElementById('btn').disabled =
    !(document.getElementById('templateFile').files[0] && document.getElementById('excelFile').files[0]);
}

['templateArea','excelArea'].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('drag-over'); });
  ['dragleave','drop'].forEach(ev => el.addEventListener(ev, () => el.classList.remove('drag-over')));
});

function showStatus(msg, type) {
  const s = document.getElementById('status');
  s.textContent = msg;
  s.className = 'status ' + type;
  s.style.display = 'block';
}

function setProgress(pct, text) {
  document.getElementById('fill').style.width = pct + '%';
  document.getElementById('progressText').textContent = text || (pct + '%');
}

let _pollTimer = null;
let _resetTimer = null;

function startPolling(jobId) {
  let failures = 0;
  async function poll() {
    try {
      const r = await fetch('/progress/' + jobId, {signal: AbortSignal.timeout(15000)});
      if (r.status === 404) {
        showStatus('❌ 生成任务已丢失，服务器可能已重启。请重新生成，或使用本地版本。', 'error');
        resetBtn();
        return;
      }
      if (!r.ok) throw new Error('连接服务器失败');
      const d = await r.json();
      failures = 0;
      if (d.pct !== undefined) setProgress(d.pct, d.label);
      if (d.done) {
        if (d.error) {
          showStatus('❌ ' + d.error, 'error');
          resetBtn();
        } else {
          setProgress(100, '完成');
          showStatus((d.warning ? '⚠️ ' + d.warning : '✅ ' + (d.summary || '生成成功！')) + ' ZIP 已开始下载。', d.warning ? 'info' : 'success');
          const a = document.createElement('a');
          a.href = '/download/' + jobId;
          a.download = '付款通知书.zip';
          a.click();
          resetBtn();
        }
        return;
      }
    } catch(e) {
      failures += 1;
      if (failures >= 5) {
        showStatus('❌ 多次无法连接服务器。请检查网络或使用本地版本；稍后可刷新页面重试。', 'error');
        resetBtn();
        return;
      }
      showStatus('连接中断，正在重试（' + failures + '/5）…', 'info');
    }
    _pollTimer = setTimeout(poll, 1500);
  }
  poll();
}

function resetBtn() {
  clearTimeout(_pollTimer);
  const btn = document.getElementById('btn');
  btn.disabled = false;
  btn.textContent = '生成并下载 ZIP';
  _resetTimer = setTimeout(() => { document.getElementById('progressWrap').style.display = 'none'; setProgress(0,''); }, 2000);
}

async function generate() {
  clearTimeout(_resetTimer);
  const btn = document.getElementById('btn');
  btn.disabled = true;
  btn.textContent = '生成中，请稍候…';
  document.getElementById('progressWrap').style.display = 'block';
  setProgress(5, '上传中…');
  showStatus('正在上传文件…', 'info');

  const form = new FormData();
  form.append('template', document.getElementById('templateFile').files[0]);
  form.append('data', document.getElementById('excelFile').files[0]);

  try {
    const resp = await fetch('/start', { method: 'POST', body: form, signal: AbortSignal.timeout(120000) });
    if (!resp.ok) {
      const err = await resp.json().catch(() => ({ error: resp.statusText }));
      throw new Error(err.error || '服务器错误');
    }
    const { job_id } = await resp.json();
    showStatus('正在生成文件，请稍候…', 'info');
    startPolling(job_id);
  } catch(e) {
    showStatus('❌ ' + e.message, 'error');
    resetBtn();
  }
}
</script>
</body>
</html>
"""


# ──────────────────────────────────────────────
# Excel parsing
# ──────────────────────────────────────────────

def excel_rows(xlsx_bytes: bytes) -> list[dict]:
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb.active  # use active sheet instead of hardcoded "Sheet1"
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("表格为空")
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


# ──────────────────────────────────────────────
# DOCX merge  –  handles {field} across split runs
# ──────────────────────────────────────────────

def _run_has_drawing(run) -> bool:
    """Return True if this run contains an image/drawing element (must not be touched)."""
    return bool(run._r.findall('.//' + qn('w:drawing')))


def _merge_runs_in_paragraph(para):
    """Coalesce consecutive text-only runs so split {field} placeholders can be matched.
    Runs that contain drawings are left completely alone."""
    runs = para.runs
    if len(runs) < 2:
        return

    # Build groups of consecutive text-only runs; flush each group into its first run.
    group_start = None
    group_text = ""
    for run in runs:
        if _run_has_drawing(run):
            # flush any open group, then skip this run entirely
            if group_start is not None:
                group_start.text = group_text
                group_start = None
                group_text = ""
            continue
        if group_start is None:
            group_start = run
            group_text = run.text or ""
        else:
            group_text += run.text or ""
            run.text = ""
    if group_start is not None:
        group_start.text = group_text


def _replace_in_paragraph(para, mapping: dict):
    # Edit only placeholder text nodes. Coalescing every run loses unrelated
    # fonts, colors, emphasis and can change the template even before PDF export.
    nodes, pieces, position = [], [], 0
    for node in para._p.iter():
        if node.tag not in (qn("w:t"), qn("w:tab"), qn("w:br"), qn("w:cr")):
            continue
        ancestors = list(node.iterancestors())
        if any(a.tag in (qn("w:drawing"), qn("w:pict")) for a in ancestors):
            continue
        text = (node.text or "") if node.tag == qn("w:t") else ("\t" if node.tag == qn("w:tab") else "\n")
        if node.tag == qn("w:t") and text:
            nodes.append((node, position, position + len(text)))
        pieces.append(text)
        position += len(text)
    text = "".join(pieces)
    pattern = re.compile(r"\{([^{}\r\n\t]+)\}")
    for match in reversed(list(pattern.finditer(text))):
        if match.group(1) not in mapping:
            continue
        affected = [(node, start, end) for node, start, end in nodes
                    if start < match.end() and end > match.start()]
        if not affected:
            continue
        first, start, _ = affected[0]
        last, last_start, _ = affected[-1]
        raw_value = mapping[match.group(1)]
        value = str(raw_value if raw_value is not None else "")
        prefix = first.text[:match.start() - start]
        suffix = last.text[match.end() - last_start:]
        if first is last:
            first.text = prefix + value + suffix
        else:
            first.text = prefix + value
            for node, _, _ in affected[1:-1]:
                node.text = ""
            last.text = suffix
        for node, _, _ in affected:
            if node.text and (node.text[0].isspace() or node.text[-1].isspace()):
                node.set(qn("xml:space"), "preserve")


def _replace_in_table(table, mapping: dict):
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                _replace_in_paragraph(para, mapping)
            for tbl in cell.tables:
                _replace_in_table(tbl, mapping)


def fill_template(template_bytes: bytes, mapping: dict) -> bytes:
    doc = Document(io.BytesIO(template_bytes))
    for para in doc.paragraphs:
        _replace_in_paragraph(para, mapping)
    for table in doc.tables:
        _replace_in_table(table, mapping)
    # headers / footers
    for section in doc.sections:
        for hdr in (section.header, section.footer):
            for para in hdr.paragraphs:
                _replace_in_paragraph(para, mapping)
            for table in hdr.tables:
                _replace_in_table(table, mapping)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ──────────────────────────────────────────────
# PDF conversion  –  LibreOffice (preferred) or
#                    weasyprint (fallback)
# ──────────────────────────────────────────────

def _pdf_batch_win32com(docx_paths: list[str], out_dir: str) -> dict[str, str]:
    """Convert docx→PDF via Word COM automation (Windows only, highest fidelity)."""
    try:
        import win32com.client
    except ImportError:
        return {}
    results = {}
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        for docx_path in docx_paths:
            base = os.path.splitext(os.path.basename(docx_path))[0]
            pdf_path = os.path.join(out_dir, base + ".pdf")
            try:
                abs_docx = os.path.abspath(docx_path)
                abs_pdf  = os.path.abspath(pdf_path)
                doc = word.Documents.Open(abs_docx)
                doc.SaveAs(abs_pdf, FileFormat=17)  # 17 = wdFormatPDF
                doc.Close(False)
                if os.path.exists(pdf_path):
                    results[docx_path] = pdf_path
            except Exception as e:
                app.logger.warning("win32com PDF failed for %s: %s", docx_path, e)
    except Exception as e:
        app.logger.warning("win32com Word init failed: %s", e)
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass
    return results


def _pdf_batch_libreoffice(docx_paths: list[str], out_dir: str) -> dict[str, str]:
    """Convert a list of docx files to PDF in one LibreOffice call. Returns {docx_path: pdf_path}."""
    import subprocess
    if not docx_paths:
        return {}
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if not executable and os.path.exists("/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        executable = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if not executable:
        return {}
    # Timeout scales with number of files; cap at 180 s
    timeout = min(30 + 5 * len(docx_paths), 180)
    try:
        with tempfile.TemporaryDirectory() as profile:
            from pathlib import Path
            subprocess.run(
                [executable, "-env:UserInstallation=" + Path(profile).as_uri(),
                 "--headless", "--convert-to", "pdf", "--outdir", out_dir] + docx_paths,
                capture_output=True, timeout=timeout, check=True
            )
    except Exception as e:
        app.logger.warning("LibreOffice batch conversion failed: %s", e)
    result = {}
    for docx_path in docx_paths:
        base = os.path.splitext(os.path.basename(docx_path))[0]
        pdf_path = os.path.join(out_dir, base + ".pdf")
        if os.path.exists(pdf_path):
            result[docx_path] = pdf_path
    return result


def _pdf_one_libreoffice(docx_path: str, out_dir: str) -> str | None:
    """Convert a single docx to PDF via LibreOffice (per-file fallback with short timeout)."""
    return _pdf_batch_libreoffice([docx_path], out_dir).get(docx_path)


def _pdf_batch_macos_word(docx_paths: list[str], out_dir: str) -> dict[str, str]:
    """Use installed Microsoft Word on macOS, with a bounded subprocess."""
    if sys.platform != "darwin" or not docx_paths:
        return {}
    if not os.path.exists("/Applications/Microsoft Word.app"):
        return {}
    script = r"""
function run(args) {
    const word = Application('Microsoft Word');
    word.launch();
    const entries = JSON.parse(args[0]);
    for (const entry of entries) {
        word.open(Path(entry.input));
        // Word may return undefined from open(), even when opening succeeds.
        // Use a unique staged name, never the user's currently active document.
        const name = entry.input.split('/').pop();
        const doc = word.documents.byName(name);
        if (doc.name() !== name) throw new Error('Generated document not found');
        try {
            doc.saveAs({fileName: entry.output, fileFormat: 'format PDF'});
        } finally {
            doc.close({saving: 'no'});
        }
    }
}
"""
    try:
        # Control Word only, without System Events or accessibility permissions.
        # Keep Word running and never close documents other than our generated ones.
        with tempfile.TemporaryDirectory(prefix="pnt-word-", dir=out_dir) as staged_dir:
            entries = []
            for i, path in enumerate(docx_paths):
                staged = os.path.abspath(os.path.join(staged_dir, f"pnt-{uuid.uuid4().hex}-{i}.docx"))
                shutil.copy2(path, staged)
                pdf = os.path.abspath(os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".pdf"))
                entries.append({"input": staged, "output": pdf})
            subprocess.run(
                ["/usr/bin/osascript", "-l", "JavaScript", "-e", script,
                 json.dumps(entries)],
                capture_output=True, timeout=min(60 + 10 * len(docx_paths), 600), check=True
            )
    except Exception as e:
        detail = e.stderr.decode("utf-8", errors="replace") if isinstance(e, subprocess.CalledProcessError) else str(e)
        app.logger.warning("macOS Word PDF conversion failed: %s", detail)
    return {path: os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
            for path in docx_paths
            if os.path.exists(os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".pdf"))}


def _pdf_via_unoconvert(docx_path: str, out_dir: str) -> str | None:
    """Convert a single docx to PDF via unoserver (resident LibreOffice process)."""
    import subprocess
    if not shutil.which("unoconvert"):
        return None
    pdf_path = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    try:
        r = subprocess.run(
            ["unoconvert", "--convert-to", "pdf", "--port", "2003", docx_path, pdf_path],
            capture_output=True, timeout=30
        )
        if r.returncode == 0 and os.path.exists(pdf_path):
            return pdf_path
    except Exception:
        pass
    return None


_WEASYPRINT_CSS = """
@font-face {
  font-family: 'CJK';
  src: local('Noto Sans CJK SC'), local('Noto Sans SC'), local('Source Han Sans CN'),
       local('Microsoft YaHei'), local('SimHei'), local('WenQuanYi Micro Hei');
}
@page { margin: 2cm; size: A4; }
body {
  font-family: 'CJK', 'Microsoft YaHei', SimHei, Arial, sans-serif;
  font-size: 10.5pt; line-height: 1.6; color: #000;
}
p { margin: 0 0 4pt 0; }
table { border-collapse: collapse; width: 100%; margin: 8pt 0; }
td, th { border: 1px solid #999; padding: 5pt 8pt; vertical-align: top; }
img { max-width: 100%; }
strong { font-weight: bold; }
"""

@lru_cache(maxsize=1)
def _html_pdf_engine():
    try:
        import mammoth
        from weasyprint import HTML
        return mammoth, HTML
    except Exception as e:
        app.logger.warning("HTML PDF engine unavailable: %s", e)
        return None


def _render_html_pdf(docx_path: str, out_dir: str) -> str | None:
    """Convert docx → rich HTML (with embedded images) → PDF via weasyprint.
    Pure Python, no external programs needed. Works on Linux (Render) with
    apt packages: fonts-noto-cjk libpango-1.0-0 libpangocairo-1.0-0 libcairo2."""
    try:
        engine = _html_pdf_engine()
        if not engine:
            return None
        mammoth, WH = engine
        with open(docx_path, "rb") as f:
            result = mammoth.convert_to_html(f)
        html_body = result.value
        html_full = (
            "<!DOCTYPE html><html><head><meta charset='UTF-8'>"
            f"<style>{_WEASYPRINT_CSS}</style></head><body>"
            f"{html_body}</body></html>"
        )
        pdf_path = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
        WH(string=html_full).write_pdf(pdf_path)
        return pdf_path if os.path.exists(pdf_path) else None
    except Exception as e:
        app.logger.warning("mammoth+weasyprint failed for %s: %s", docx_path, e)
        return None


def _pdf_via_mammoth_weasyprint(docx_path: str, out_dir: str) -> str | None:
    # A slow or memory-heavy conversion must not hang the background thread.
    if not _html_pdf_engine():
        return None
    try:
        subprocess.run(
            [sys.executable, os.path.abspath(__file__), "--render-pdf", docx_path, out_dir],
            capture_output=True, timeout=60, check=True
        )
    except Exception as e:
        app.logger.warning("HTML PDF conversion failed: %s", e)
        return None
    pdf_path = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    return pdf_path if os.path.exists(pdf_path) else None


def docx_to_pdf_one(docx_path: str, out_dir: str) -> str | None:
    """Convert a single docx to PDF, trying all available backends in priority order."""
    # 1. win32com (Windows + Word installed)
    r = _pdf_batch_win32com([docx_path], out_dir)
    if r:
        return r[docx_path]
    r = _pdf_batch_macos_word([docx_path], out_dir)
    if r:
        return r[docx_path]
    # The HTML conversion is lossy and must never be a silent fallback.
    pdf = (_pdf_via_mammoth_weasyprint(docx_path, out_dir)
           if os.environ.get("ALLOW_APPROXIMATE_PDF") == "1" else None)
    if pdf:
        return pdf
    # 3. unoserver
    pdf = _pdf_via_unoconvert(docx_path, out_dir)
    if pdf:
        return pdf
    # 4. LibreOffice one-shot
    pdf = _pdf_one_libreoffice(docx_path, out_dir)
    return pdf


# ──────────────────────────────────────────────
# Background worker
# ──────────────────────────────────────────────

def _run_job(job_id: str, template_bytes: bytes, data_bytes: bytes):
    try:
        job_update(job_id, label="等待生成任务…")
        while not _conversion_lock.acquire(timeout=10):
            job_update(job_id, label="正在排队，等待上一批生成完成…")
        try:
            _process_job(job_id, template_bytes, data_bytes)
        finally:
            _conversion_lock.release()
    except Exception:
        app.logger.exception("Generation failed for job %s", job_id)
        job_update(job_id, done=True, error="生成失败，请重试或检查服务器日志；建议减少单次生成数量。")


def _process_job(job_id: str, template_bytes: bytes, data_bytes: bytes):
    try:
        rows = excel_rows(data_bytes)
    except Exception as e:
        job_update(job_id, done=True, error=f"读取 Excel 失败：{e}")
        return

    if not rows:
        job_update(job_id, done=True, error="当前工作表没有数据行")
        return

    total = len(rows)
    if template_bytes.startswith(b"%PDF-"):
        _process_pdf_job(job_id, template_bytes, rows)
        return
    name_count: dict[str, int] = {}

    # macOS Word can require extra file-access prompts for system temp folders.
    # Stage files beside the private job store instead.
    with tempfile.TemporaryDirectory(dir=JOB_DIR) as tmpdir:
        # Phase 1: generate all Word files
        docx_entries = []  # list of (unique_name, docx_path)
        for i, row in enumerate(rows, 1):
            company = row.get("公司名称", f"客户{i}")
            base = "付款通知_" + safe_filename(company)
            name_count[base] = name_count.get(base, 0) + 1
            cnt = name_count[base]
            unique = base if cnt == 1 else f"{base}_{cnt}"

            try:
                docx_bytes = fill_template(template_bytes, row)
            except Exception as e:
                job_update(job_id, done=True, error=f"第 {i} 行填充模板失败：{e}")
                return

            docx_path = os.path.join(tmpdir, f"{unique}.docx")
            with open(docx_path, "wb") as f:
                f.write(docx_bytes)
            docx_entries.append((unique, docx_path))

            pct = int(10 + 40 * i / total)
            job_update(job_id, pct=pct, label=f"生成 Word {i}/{total}…")

        # Phase 2: convert docx → PDF
        # Try win32com batch first (Windows), then fall back per-file for any remaining
        job_update(job_id, pct=55, label="转换 PDF…")
        docx_paths = [p for _, p in docx_entries]
        pdf_map: dict[str, str] = {}

        # Attempt batch win32com (fast on Windows)
        batch = _pdf_batch_win32com(docx_paths, tmpdir)
        if not batch:
            batch = _pdf_batch_macos_word(docx_paths, tmpdir)
        if not batch:
            batch = _pdf_batch_libreoffice(docx_paths, tmpdir)
        pdf_map.update(batch)
        remaining = [p for p in docx_paths if p not in pdf_map]

        # Per-file fallback for any not converted yet
        for i, docx_path in enumerate(remaining, 1):
            pct = int(55 + 35 * (len(batch) + i) / total)
            job_update(job_id, pct=pct, label=f"转换 PDF {len(batch)+i}/{total}…")
            # try mammoth+weasyprint → unoserver → libreoffice
            pdf = ((_pdf_via_mammoth_weasyprint(docx_path, tmpdir)
                    if os.environ.get("ALLOW_APPROXIMATE_PDF") == "1" else None)
                   or _pdf_via_unoconvert(docx_path, tmpdir)
                   or _pdf_one_libreoffice(docx_path, tmpdir))
            if pdf:
                pdf_map[docx_path] = pdf

        pdf_entries = []
        for unique, docx_path in docx_entries:
            pdf_path = pdf_map.get(docx_path)
            if pdf_path:
                pdf_entries.append((unique, pdf_path))

        # Always preserve Word files, even when PDF conversion fails.
        job_update(job_id, pct=92, label="打包 ZIP…")
        pdf_count = len(pdf_entries)
        zip_out = _zip_path(job_id)
        with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as zf:
            for unique, docx_path in docx_entries:
                zf.write(docx_path, f"Word/{unique}.docx")
            for unique, pdf_path in pdf_entries:
                zf.write(pdf_path, f"PDF/{unique}.pdf")

    pdf_note = f"共生成 Word {total} 份、PDF {pdf_count}/{total} 份"
    warning = None
    if pdf_count < total:
        warning = (f"PDF 仅生成 {pdf_count}/{total} 份，完整 Word 文件已保留。"
                   "推荐把原 Word 模板导出为带占位符的 PDF，再使用 PDF 原版模式。Word 模式需要 Word / LibreOffice。")
    elif os.environ.get("ALLOW_APPROXIMATE_PDF") == "1":
        warning = "已启用近似 HTML 转换，PDF 的字体、图片位置和表格可能与 Word 不同。保留原版式请改用 PDF 模板模式。"
    job_update(job_id, pct=100, label="完成", done=True,
               summary=pdf_note, warning=warning)


def _process_pdf_job(job_id: str, template_bytes: bytes, rows: list[dict]):
    from pdf_template import PDFTemplate
    try:
        template = PDFTemplate(template_bytes, set(rows[0]))
    except ValueError as e:
        job_update(job_id, done=True, error=str(e))
        return
    used, substitutions = set(), set()
    with tempfile.TemporaryDirectory(dir=JOB_DIR) as tmpdir:
        zip_tmp = os.path.join(tmpdir, "notices.zip")
        with zipfile.ZipFile(zip_tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, row in enumerate(rows, 1):
                base = "付款通知_" + safe_filename(row.get("公司名称") or f"客户{i}")
                unique, suffix = base, 2
                while unique in used:
                    unique = f"{base}_{suffix}"
                    suffix += 1
                used.add(unique)
                try:
                    pdf, font_changes = template.render(row)
                except ValueError as e:
                    job_update(job_id, done=True, error=f"第 {i} 行：{e}")
                    return
                substitutions.update(font_changes)
                zf.writestr(f"PDF/{unique}.pdf", pdf)
                job_update(job_id, pct=int(10 + 80 * i / len(rows)),
                           label=f"按 PDF 原版填充 {i}/{len(rows)}…")
        os.replace(zip_tmp, _zip_path(job_id))
    warning = None
    if substitutions:
        warning = ("原 PDF 的固定版式已保留；部分动态字段使用兼容中文字体（原 PDF 的嵌入字体缺少新文字）："
                   + "、".join(sorted(substitutions)))
    job_update(job_id, done=True, pct=100, label="完成",
               summary=f"共生成 PDF {len(rows)}/{len(rows)} 份（PDF 模板原版模式）", warning=warning)


def safe_filename(name: str) -> str:
    return "".join(c for c in name if c not in r'\/:*?"<>|').strip() or "未知"


# ──────────────────────────────────────────────
# Routes
# ──────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/start", methods=["POST"])
def start():
    if "template" not in request.files or "data" not in request.files:
        return jsonify(error="缺少上传文件"), 400

    filename = (request.files["template"].filename or "").lower()
    if not filename.endswith(".pdf"):
        return jsonify(error="本页面只接受 PDF 模板（.pdf），请先在 Word 中导出带占位符的 PDF"), 400
    if not request.files["data"].filename.lower().endswith(".xlsx"):
        return jsonify(error="数据表必须为 .xlsx，请将旧 .xls 另存为 .xlsx"), 400

    template_bytes = request.files["template"].read()
    data_bytes = request.files["data"].read()
    if not template_bytes or not data_bytes:
        return jsonify(error="上传文件为空"), 400
    if filename.endswith(".pdf") and not template_bytes.startswith(b"%PDF-"):
        return jsonify(error="文件不是有效的 PDF，请从 Word 导出真正的 PDF 文件"), 400

    job_id = str(uuid.uuid4())
    job_create(job_id)

    t = threading.Thread(target=_run_job, args=(job_id, template_bytes, data_bytes), daemon=True)
    t.start()
    return jsonify(job_id=job_id)


@app.route("/progress/<job_id>")
def progress(job_id: str):
    try:
        uuid.UUID(job_id)
    except ValueError:
        return jsonify(error="job not found"), 404
    job = job_read(job_id)
    if not job:
        return jsonify(error="job not found"), 404
    if not job.get("done") and time.time() - job.get("updated_at", 0) > 900:
        job = dict(job, done=True, error="任务长时间未更新，服务器可能已重启或转换超时。请减少生成数量后重试，或使用本地版本。")
    return jsonify(job)


@app.route("/download/<job_id>")
def download(job_id: str):
    try:
        uuid.UUID(job_id)
    except ValueError:
        return jsonify(error="job not found"), 404
    job = job_read(job_id)
    if not job or not job.get("done"):
        return jsonify(error="未就绪"), 404
    if job.get("error"):
        return jsonify(error=job["error"]), 500
    zip_file = _zip_path(job_id)
    if not os.path.exists(zip_file):
        return jsonify(error="无数据"), 500
    return send_file(
        zip_file,
        mimetype="application/zip",
        as_attachment=True,
        download_name="付款通知书.zip",
    )


@app.route("/generate", methods=["POST"])
def generate_sync():
    """Legacy synchronous endpoint kept for backward compatibility."""
    if "template" not in request.files or "data" not in request.files:
        return jsonify(error="缺少上传文件"), 400
    template_bytes = request.files["template"].read()
    data_bytes = request.files["data"].read()

    job_id = str(uuid.uuid4())
    job_create(job_id)
    _run_job(job_id, template_bytes, data_bytes)
    return download(job_id)


@app.errorhandler(413)
def upload_too_large(error):
    return jsonify(error="上传文件总大小不能超过 50 MB"), 413


@app.route("/health")
def health():
    return jsonify(status="ok", template_format="pdf",
                   version=os.environ.get("RENDER_GIT_COMMIT", "local")[:12])


if __name__ == "__main__":
    import sys
    if len(sys.argv) == 4 and sys.argv[1] == "--render-pdf":
        sys.exit(0 if _render_html_pdf(sys.argv[2], sys.argv[3]) else 1)
    sys.stdout.reconfigure(encoding="utf-8")
    print("=" * 40)
    port = int(os.environ.get("PORT", "5050"))
    print(f"Dashboard: http://127.0.0.1:{port}")
    print("=" * 40)
    if os.environ.get("OPEN_BROWSER") == "1":
        import webbrowser
        timer = threading.Timer(1, webbrowser.open, args=(f"http://127.0.0.1:{port}",))
        timer.daemon = True
        timer.start()
    app.run(debug=False, host="127.0.0.1", port=port)
