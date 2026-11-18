import io
import json
import os
import tempfile
import unittest
import uuid
import zipfile
from unittest.mock import patch

from PIL import Image, ImageChops
import pymupdf as fitz
from docx import Document
from docx.shared import RGBColor

import app
from pdf_template import PDFTemplate
from test_app import sample_files


def sample_pdf():
    with fitz.open() as doc:
        page = doc.new_page(width=595, height=842)
        page.insert_text((40, 70), "SAP notice", fontsize=13, color=(0, 0.4, 0.7))
        page.insert_text((40, 115), "致：{公司名称}", fontname="china-s", fontsize=11)
        page.draw_rect(fitz.Rect(70, 150, 200, 185), color=(0, 0, 0))
        # Center the four placeholder characters in this visible table cell.
        page.insert_text((113, 172), "{金额}", fontname="china-s", fontsize=11, color=(0, 0.4, 0.7))
        page.insert_text((40, 350), "Account: 123456789012", fontsize=11)
        for color, rect in [((0, 130, 200), (430, 30, 560, 90)),
                            ((210, 20, 40), (310, 410, 410, 510))]:
            image = Image.new("RGB", (50, 50), color)
            stream = io.BytesIO()
            image.save(stream, format="PNG")
            page.insert_image(fitz.Rect(rect), stream=stream.getvalue())
        return doc.tobytes()


class PDFTemplateTests(unittest.TestCase):
    def test_preserves_static_artwork_page_and_table(self):
        source = sample_pdf()
        template = PDFTemplate(source, {"公司名称", "金额"})
        rendered, changes = template.render({"公司名称": "测试公司", "金额": "10598.22"})
        with fitz.open(stream=source, filetype="pdf") as before, \
             fitz.open(stream=rendered, filetype="pdf") as after:
            self.assertEqual(after[0].rect, before[0].rect)
            self.assertEqual(after[0].get_drawings(), before[0].get_drawings())
            self.assertEqual(len(after[0].get_images()), 2)
            text = after[0].get_text()
            self.assertIn("测试公司", text)
            self.assertIn("10598.22", text)
            self.assertIn("Account: 123456789012", text)
            self.assertNotIn("{公司名称}", text)
            self.assertNotIn("{金额}", text)
            original = Image.frombytes("RGB", (595, 842), before[0].get_pixmap().samples)
            actual = Image.frombytes("RGB", (595, 842), after[0].get_pixmap().samples)
            difference = ImageChops.difference(original, actual)
            # Only the dynamic text regions may differ; all other pixels match.
            for field in template.fields:
                difference.paste((0, 0, 0), (int(field.left), int(field.rect[1]) - 8,
                                            int(field.right) + 1, int(field.rect[3]) + 8))
            self.assertIsNone(difference.getbbox())

    def test_table_value_remains_centered_and_same_size_and_color(self):
        template = PDFTemplate(sample_pdf(), {"公司名称", "金额"})
        rendered, _ = template.render({"公司名称": "测试", "金额": "10598.22"})
        with fitz.open(stream=rendered, filetype="pdf") as doc:
            amount = doc[0].search_for("10598.22")[0]
            self.assertAlmostEqual((amount.x0 + amount.x1) / 2, 135, delta=0.2)
            spans = [span for block in doc[0].get_text("dict")["blocks"]
                     for line in block.get("lines", []) for span in line["spans"]
                     if "10598.22" in span["text"]]
            self.assertEqual(spans[0]["size"], 11)
            self.assertEqual(spans[0]["color"], next(f.color for f in template.fields if f.name == "金额"))

    def test_overflow_fails_instead_of_shrinking_or_overlapping(self):
        template = PDFTemplate(sample_pdf(), {"公司名称", "金额"})
        with self.assertRaisesRegex(ValueError, "内容过长"):
            template.render({"公司名称": "测试", "金额": "1" * 100})

    def test_filled_pdf_is_not_accepted_as_template(self):
        template = PDFTemplate(sample_pdf(), {"公司名称", "金额"})
        rendered, _ = template.render({"公司名称": "测试", "金额": "1"})
        with self.assertRaisesRegex(ValueError, "找不到"):
            PDFTemplate(rendered, {"公司名称", "金额"})

    def test_missing_column_and_multiline_values_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "没有对应列"):
            PDFTemplate(sample_pdf(), {"公司名称"})
        template = PDFTemplate(sample_pdf(), {"公司名称", "金额"})
        with self.assertRaisesRegex(ValueError, "换行"):
            template.render({"公司名称": "测试\n公司", "金额": "1"})

    def test_blank_values_remove_placeholders_without_ghost_text(self):
        template = PDFTemplate(sample_pdf(), {"公司名称", "金额"})
        rendered, changes = template.render({"公司名称": "", "金额": None})
        self.assertFalse(changes)
        with fitz.open(stream=rendered, filetype="pdf") as doc:
            self.assertNotIn("{", doc[0].get_text())
            self.assertEqual(len(doc[0].get_images()), 2)

    def test_repeated_fields_and_letter_page_size_are_retained(self):
        with fitz.open() as doc:
            for _ in range(2):
                page = doc.new_page(width=612, height=792)
                page.insert_text((40, 80), "{Company}")
                page.insert_text((40, 120), "{Company}")
            template = PDFTemplate(doc.tobytes(), {"Company"})
        rendered, _ = template.render({"Company": "Example"})
        with fitz.open(stream=rendered, filetype="pdf") as doc:
            self.assertEqual(len(doc), 2)
            for page in doc:
                self.assertEqual(tuple(page.rect), (0, 0, 612, 792))
                self.assertEqual(page.get_text().count("Example"), 2)

    def test_pdf_job_download_contains_only_valid_pdfs(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(app, "JOB_DIR", tmp):
            client = app.app.test_client()
            _, xlsx = sample_files()
            # Run synchronously here to validate the complete worker and ZIP.
            job_id = str(uuid.uuid4())
            app.job_create(job_id)
            app._run_job(job_id, sample_pdf(), xlsx)
            job = app.job_read(job_id)
            self.assertTrue(job["done"])
            self.assertNotIn("error", job)
            self.assertIn("2/2", job["summary"])
            download = client.get("/download/" + job_id)
            self.assertEqual(download.status_code, 200)
            with zipfile.ZipFile(io.BytesIO(download.data)) as zf:
                self.assertEqual(len(zf.namelist()), 2)
                self.assertTrue(all(n.endswith(".pdf") for n in zf.namelist()))
                self.assertIn("PDF/付款通知_测试公司_2.pdf", zf.namelist())
                for name in zf.namelist():
                    with fitz.open(stream=zf.read(name), filetype="pdf") as doc:
                        self.assertIn("测试公司", doc[0].get_text())
            download.close()

    def test_pdf_upload_is_accepted(self):
        with patch.object(app.threading.Thread, "start"):
            _, xlsx = sample_files()
            response = app.app.test_client().post("/start", data={
                "template": (io.BytesIO(sample_pdf()), "template.pdf"),
                "data": (io.BytesIO(xlsx), "data.xlsx"),
            })
        self.assertEqual(response.status_code, 200)
        self.assertIn("job_id", response.json)

    def test_page_and_start_endpoint_only_accept_pdf_templates(self):
        client = app.app.test_client()
        page = client.get("/").data.decode()
        self.assertIn('id="templateFile" accept=".pdf,application/pdf"', page)
        self.assertIn("付款通知书 PDF 模板 (.pdf)", page)
        self.assertNotIn('accept=".pdf,.docx"', page)
        template, data = sample_files()
        response = client.post("/start", data={
            "template": (io.BytesIO(template), "template.docx"),
            "data": (io.BytesIO(data), "data.xlsx"),
        })
        self.assertEqual(response.status_code, 400)
        self.assertIn("只接受 PDF", response.json["error"])

    def test_compatibility_radical_in_pdf_field_matches_excel_column(self):
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((40, 80), "{⾦额}", fontname="china-s")
            template = PDFTemplate(doc.tobytes(), {"金额"})
        self.assertEqual(template.fields[0].name, "金额")
        rendered, _ = template.render({"金额": "10598.22"})
        with fitz.open(stream=rendered, filetype="pdf") as doc:
            self.assertIn("10598.22", doc[0].get_text())
            self.assertNotIn("{", doc[0].get_text())

    def test_ambiguous_normalized_excel_columns_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "规范化后重名"):
            PDFTemplate(sample_pdf(), {"公司名称", "金额", "⾦额"})

    def test_docx_replacement_preserves_unrelated_run_styles(self):
        doc = Document()
        p = doc.add_paragraph()
        label = p.add_run("致：")
        label.bold = True
        first = p.add_run("{公司")
        first.font.color.rgb = RGBColor(0, 102, 178)
        p.add_run("名称}")
        suffix = p.add_run(" 正文不变")
        suffix.italic = True
        label_xml, suffix_xml = label._r.xml, suffix._r.xml
        app._replace_in_paragraph(p, {"公司名称": "测试公司"})
        self.assertEqual(p.text, "致：测试公司 正文不变")
        self.assertEqual(label._r.xml, label_xml)
        self.assertEqual(suffix._r.xml, suffix_xml)
        self.assertEqual(first.font.color.rgb, RGBColor(0, 102, 178))

    def test_docx_multiple_fields_and_zero(self):
        doc = Document()
        p = doc.add_paragraph("{金额} / {金额} / {公司名称}")
        app._replace_in_paragraph(p, {"金额": 0, "公司名称": "测试"})
        self.assertEqual(p.text, "0 / 0 / 测试")


if __name__ == "__main__":
    unittest.main()
