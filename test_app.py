import io
import json
import os
import tempfile
import time
import unittest
import uuid
import zipfile
from unittest.mock import patch

from docx import Document
from openpyxl import Workbook

import app as tool


def sample_files(count=2):
    doc = Document()
    doc.add_paragraph("致：{公司名称}，金额：{金额}")
    template = io.BytesIO()
    doc.save(template)
    wb = Workbook()
    wb.active.append(["公司名称", "金额"])
    for _ in range(count):
        wb.active.append(["测试公司", 123])
    data = io.BytesIO()
    wb.save(data)
    return template.getvalue(), data.getvalue()


class GenerationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.job_dir = patch.object(tool, "JOB_DIR", self.tmp.name)
        self.job_dir.start()
        self.client = tool.app.test_client()

    def tearDown(self):
        self.job_dir.stop()
        self.tmp.cleanup()

    def run_job(self):
        job_id = str(uuid.uuid4())
        tool.job_create(job_id)
        tool._run_job(job_id, *sample_files())
        return job_id, tool.job_read(job_id)

    def test_missing_pdf_preserves_word_and_warns(self):
        with patch.object(tool, "_pdf_batch_win32com", return_value={}), \
             patch.object(tool, "_pdf_batch_macos_word", return_value={}), \
             patch.object(tool, "_pdf_batch_libreoffice", return_value={}), \
             patch.object(tool, "_pdf_via_mammoth_weasyprint", return_value=None), \
             patch.object(tool, "_pdf_via_unoconvert", return_value=None):
            job_id, job = self.run_job()
        self.assertTrue(job["done"])
        self.assertIn("0/2", job["warning"])
        response = self.client.get("/download/" + job_id)
        with zipfile.ZipFile(io.BytesIO(response.data)) as zf:
            self.assertEqual(len(zf.namelist()), 2)
            self.assertIn("Word/付款通知_测试公司_2.docx", zf.namelist())
            output = Document(io.BytesIO(zf.read("Word/付款通知_测试公司.docx")))
            self.assertIn("测试公司", output.paragraphs[0].text)
            self.assertNotIn("{金额}", output.paragraphs[0].text)
        response.close()

    def test_successful_batch_contains_word_and_pdf(self):
        def fake_batch(paths, out_dir):
            self.assertEqual(os.path.dirname(out_dir), tool.JOB_DIR)
            result = {}
            for path in paths:
                pdf = os.path.splitext(path)[0] + ".pdf"
                with open(pdf, "wb") as f:
                    f.write(b"%PDF-1.4\nmock")
                result[path] = pdf
            return result
        with patch.object(tool, "_pdf_batch_win32com", side_effect=fake_batch):
            job_id, job = self.run_job()
        self.assertIsNone(job["warning"])
        with zipfile.ZipFile(tool._zip_path(job_id)) as zf:
            self.assertEqual(len(zf.namelist()), 4)

    def test_unexpected_exception_finishes_job(self):
        with patch.object(tool, "_process_job", side_effect=RuntimeError("test failure")):
            job_id, job = self.run_job()
        self.assertTrue(job["done"])
        self.assertIn("error", job)
        self.assertEqual(self.client.get("/download/" + job_id).status_code, 500)
        self.assertTrue(tool._conversion_lock.acquire(blocking=False))
        tool._conversion_lock.release()

    def test_invalid_excel_reports_failure(self):
        job_id = str(uuid.uuid4())
        tool.job_create(job_id)
        tool._run_job(job_id, b"bad", b"bad")
        self.assertTrue(tool.job_read(job_id)["done"])
        self.assertIn("读取 Excel", tool.job_read(job_id)["error"])

    def test_upload_validation(self):
        self.assertEqual(self.client.post("/start").status_code, 400)
        response = self.client.post("/start", data={
            "template": (io.BytesIO(b"%PDF-test"), "template.pdf"),
            "data": (io.BytesIO(b"a"), "data.xls"),
        })
        self.assertEqual(response.status_code, 400)
        self.assertIn(".xlsx", response.json["error"])

    def test_stale_job_is_not_left_spinning(self):
        job_id = str(uuid.uuid4())
        with open(tool._meta_path(job_id), "w") as f:
            json.dump({"done": False, "updated_at": time.time() - 1000}, f)
        response = self.client.get("/progress/" + job_id)
        self.assertTrue(response.json["done"])
        self.assertIn("error", response.json)
        self.assertEqual(self.client.get("/progress/not-a-uuid").status_code, 404)

    def test_html_conversion_timeout_is_handled(self):
        with patch.object(tool, "_html_pdf_engine", return_value=True), \
             patch.object(tool.subprocess, "run", side_effect=tool.subprocess.TimeoutExpired("test", 60)):
            self.assertIsNone(tool._pdf_via_mammoth_weasyprint("test.docx", self.tmp.name))

    def test_macos_word_uses_unique_names_without_overwriting_docx(self):
        source = os.path.join(self.tmp.name, "notice.docx")
        with open(source, "wb") as f:
            f.write(b"original Word data")
        staged_paths = []
        exists = os.path.exists

        def fake_run(command, **kwargs):
            script = command[command.index("-e") + 1]
            self.assertIn("word.documents.byName(name)", script)
            self.assertNotIn("word.activeDocument", script)
            entries = json.loads(command[-1])
            for entry in entries:
                staged_paths.append(entry["input"])
                self.assertNotEqual(entry["input"], source)
                self.assertTrue(entry["output"].endswith(".pdf"))
                self.assertNotEqual(entry["output"], source)
                with open(entry["input"], "rb") as f:
                    self.assertEqual(f.read(), b"original Word data")
                with open(entry["output"], "wb") as f:
                    f.write(b"%PDF-1.4\nmock")

        with patch.object(tool.sys, "platform", "darwin"), \
             patch.object(tool.os.path, "exists", side_effect=lambda p: p == "/Applications/Microsoft Word.app" or exists(p)), \
             patch.object(tool.subprocess, "run", side_effect=fake_run):
            result = tool._pdf_batch_macos_word([source], self.tmp.name)
        self.assertEqual(result, {source: os.path.join(self.tmp.name, "notice.pdf")})
        self.assertFalse(any(exists(p) for p in staged_paths))
        with open(source, "rb") as f:
            self.assertEqual(f.read(), b"original Word data")

    def test_async_upload_download(self):
        def fake_process(job_id, template, data):
            with zipfile.ZipFile(tool._zip_path(job_id), "w") as zf:
                zf.writestr("test.txt", "ok")
            tool.job_update(job_id, done=True, pct=100)
        with patch.object(tool, "_process_job", side_effect=fake_process):
            from test_pdf_template import sample_pdf
            _, data = sample_files()
            template = sample_pdf()
            response = self.client.post("/start", data={
                "template": (io.BytesIO(template), "template.pdf"),
                "data": (io.BytesIO(data), "data.xlsx"),
            })
            job_id = response.json["job_id"]
            deadline = time.monotonic() + 3
            while time.monotonic() < deadline:
                if tool.job_read(job_id).get("done"):
                    break
                time.sleep(0.01)
        self.assertTrue(tool.job_read(job_id)["done"])
        download = self.client.get("/download/" + job_id)
        self.assertEqual(download.status_code, 200)
        download.close()


if __name__ == "__main__":
    unittest.main()
