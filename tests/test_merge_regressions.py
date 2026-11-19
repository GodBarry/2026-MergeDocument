import hashlib
from io import BytesIO
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from lxml import etree
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main"
PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000d49444154789c63f8cfc0f01f00050001ff89993d1d0000000049454e44ae426082"
)


def paragraph(text):
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


def table(rows):
    joined_rows = "".join(
        "<w:tr>" + "".join(f"<w:tc>{paragraph(text)}</w:tc>" for text in row) + "</w:tr>"
        for row in rows
    )
    return f'<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>{joined_rows}</w:tbl>'


def write_docx(path, blocks, embed_png=False, declare_png=True, image_bytes=PNG,
               image_id="rId1", link_target=None, link_id="rId2"):
    default_png = '<Default Extension="png" ContentType="image/png"/>' if embed_png and declare_png else ""
    with ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", f"""
            <Types xmlns="{CONTENT_TYPES}">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Override PartName="/word/document.xml"
                ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
              {default_png}
            </Types>""")
        package.writestr("_rels/.rels", f"""
            <Relationships xmlns="{PACKAGE_REL}">
              <Relationship Id="rId1" Type="{OFFICE_REL}/officeDocument" Target="word/document.xml"/>
            </Relationships>""")
        package.writestr("word/document.xml", f"""
            <w:document xmlns:w="{WORD_NS}" xmlns:r="{OFFICE_REL}" xmlns:a="{DRAWING}"
                        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
              <w:body>{''.join(blocks)}<w:sectPr/></w:body>
            </w:document>""")
        if embed_png or link_target:
            rels = etree.Element(f"{{{PACKAGE_REL}}}Relationships")
            if embed_png:
                etree.SubElement(rels, f"{{{PACKAGE_REL}}}Relationship", Id=image_id,
                                 Type=f"{OFFICE_REL}/image", Target="media/image1.png")
                package.writestr("word/media/image1.png", image_bytes)
            if link_target:
                etree.SubElement(rels, f"{{{PACKAGE_REL}}}Relationship", Id=link_id,
                                 Type=f"{OFFICE_REL}/hyperlink", Target=link_target, TargetMode="External")
            package.writestr("word/_rels/document.xml.rels", etree.tostring(rels))



class MergeRegressions(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.folder = Path(directory.name)
        self.copies = self.folder / "copies"
        self.copies.mkdir()
        self.base = self.folder / "base.docx"
        self.output = self.folder / "merged.docx"
        self.report = self.folder / "merged_conflicts.txt"
        self.first = paragraph("Alpha department completed 10 projects.")
        self.second = paragraph("Bravo department updated 20 records.")
        write_docx(self.base, [self.first, self.second])

    def run_merge(self, output=None):
        return subprocess.run(
            [sys.executable, "-X", "utf8", str(ROOT / "run_merge.py"), str(self.base),
             str(self.copies), "-o", str(output or self.output)],
            capture_output=True, encoding="utf-8", timeout=30,
        )

    def assert_success(self, result):
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def read_document(self):
        with ZipFile(self.output) as package:
            return etree.fromstring(package.read("word/document.xml"))

    def test_bad_copy_aborts_without_replacing_existing_output(self):
        self.output.write_bytes(b"previous output")
        bad_copy = self.copies / "broken.docx"
        bad_copy.write_bytes(b"not a Word file")
        result = self.run_merge()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(bad_copy.name, result.stdout)
        self.assertEqual(self.output.read_bytes(), b"previous output")
        self.assertFalse(self.report.exists())

    def test_output_cannot_overwrite_copy(self):
        copy = self.copies / "archive.docx"
        write_docx(copy, [self.first.replace("10", "11"), self.second])
        original = copy.read_bytes()
        result = self.run_merge(output=copy)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(copy.read_bytes(), original)

    def test_conflict_report_distinguishes_same_and_different_paragraphs(self):
        a = self.copies / "a.docx"
        b = self.copies / "b.docx"
        write_docx(a, [self.first.replace("10", "11"), self.second])
        write_docx(b, [self.first, self.second.replace("20", "21")])
        self.assert_success(self.run_merge())
        self.assertIn("冲突数量: 0", self.report.read_text(encoding="utf-8"))

        write_docx(b, [self.first.replace("10", "12"), self.second])
        self.assert_success(self.run_merge())
        self.assertIn("冲突数量: 1", self.report.read_text(encoding="utf-8"))
        result_text = "".join(self.read_document().itertext())
        self.assertIn("11 projects", result_text)
        self.assertIn("12 projects", result_text)

    def test_table_row_addition_and_removal(self):
        header = ["Team", "Count"]
        existing = ["Alpha", "10"]
        for rows in ([header, existing, ["Bravo", "20"]], [header], [header, existing, ["", ""]]):
            with self.subTest(rows=rows):
                write_docx(self.base, [self.first, table([header, existing])])
                write_docx(self.copies / "a.docx", [self.first, table(rows)])
                self.assert_success(self.run_merge())
                out = self.read_document()
                actual = [
                    ["".join(cell.itertext()).strip() for cell in row.findall(f"{{{WORD_NS}}}tc")]
                    for row in out.findall(f".//{{{WORD_NS}}}tr")
                ]
                self.assertEqual(actual, rows)

    def test_table_conflicts_keep_both_submissions(self):
        header = ["Team", "Count"]
        write_docx(self.base, [self.first, table([header, ["Alpha", "10"]])])
        write_docx(self.copies / "a.docx", [self.first, table([header, ["Alpha", "15"]])])
        write_docx(self.copies / "b.docx", [self.first, table([header, ["Alpha", "10"], ["Bravo", "20"]])])
        self.assert_success(self.run_merge())
        self.assertIn("冲突数量: 1", self.report.read_text(encoding="utf-8"))
        output_tables = ["".join(table_node.itertext()) for table_node in self.read_document().iter(f"{{{WORD_NS}}}tbl")]
        self.assertTrue(any("Alpha15" in text for text in output_tables))
        self.assertTrue(any("Bravo20" in text for text in output_tables))

    def test_delete_and_edit_keep_original_marked_and_edited_version(self):
        write_docx(self.copies / "a.docx", [self.second])
        write_docx(self.copies / "b.docx", [self.first.replace("10", "11"), self.second])
        self.assert_success(self.run_merge())
        self.assertIn("冲突数量: 1", self.report.read_text(encoding="utf-8"))
        output = self.read_document()
        self.assertTrue(output.findall(f".//{{{WORD_NS}}}strike"))
        text = "".join(output.itertext())
        self.assertIn("10 projects", text)
        self.assertIn("11 projects", text)

    def test_format_only_change_and_benign_resave(self):
        copy = self.copies / "a.docx"
        resaved = ('<w:p w:rsidR="00A0FA00"><w:r><w:rPr/><w:t xml:space="preserve">'
                   'Alpha department </w:t></w:r><w:r><w:t>completed 10 projects.</w:t></w:r></w:p>')
        write_docx(copy, [resaved, self.second])
        self.assert_success(self.run_merge())
        self.assertIn("自动合并变更数: 0", self.report.read_text(encoding="utf-8"))

        bold = self.first.replace('<w:r>', '<w:r><w:rPr><w:b/></w:rPr>', 1)
        justified = self.second.replace('<w:p>', '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>', 1)
        write_docx(copy, [bold, justified])
        self.assert_success(self.run_merge())
        document = self.read_document()
        self.assertEqual(len(document.findall(f".//{{{WORD_NS}}}b")), 1)
        self.assertEqual(document.find(f".//{{{WORD_NS}}}jc").get(f"{{{WORD_NS}}}val"), "center")

    def test_hyperlink_target_compares_across_relationship_ids(self):
        link = '<w:hyperlink r:id="{rid}"><w:r><w:t> Link</w:t></w:r></w:hyperlink>'
        base_with_link = self.first.replace("</w:p>", link.format(rid="rId2") + "</w:p>")
        copy_with_link = self.first.replace("</w:p>", link.format(rid="linkA") + "</w:p>")
        write_docx(self.base, [base_with_link, self.second], link_target="https://example.invalid/old")
        copy = self.copies / "a.docx"
        write_docx(copy, [copy_with_link, self.second], link_target="https://example.invalid/old", link_id="linkA")
        self.assert_success(self.run_merge())
        self.assertIn("自动合并变更数: 0", self.report.read_text(encoding="utf-8"))

        for original_blocks in ([base_with_link, self.second], [self.first, self.second]):
            with self.subTest(original_has_link=original_blocks[0] == base_with_link):
                old_target = "https://example.invalid/old" if original_blocks[0] == base_with_link else None
                write_docx(self.base, original_blocks, link_target=old_target)
                write_docx(copy, [copy_with_link, self.second],
                           link_target="https://example.invalid/new", link_id="linkA")
                self.assert_success(self.run_merge())
                document = self.read_document()
                rid = document.find(f".//{{{WORD_NS}}}hyperlink").get(f"{{{OFFICE_REL}}}id")
                with ZipFile(self.output) as package:
                    rels = etree.fromstring(package.read("word/_rels/document.xml.rels"))
                targets = [rel.get("Target") for rel in rels if rel.get("Id") == rid]
                self.assertEqual(targets, ["https://example.invalid/new"])

    def test_replaced_image_with_same_id_and_filename(self):
        picture = ('<w:drawing><wp:inline><a:graphic><a:graphicData>'
                   '<a:blip r:embed="rId1"/></a:graphicData></a:graphic></wp:inline></w:drawing>')
        text_and_picture = self.first.replace("</w:r>", picture + "</w:r>")
        write_docx(self.base, [text_and_picture, self.second], embed_png=True)
        replacement = BytesIO()
        Image.new("RGB", (2, 2), (11, 22, 33)).save(replacement, format="PNG")
        write_docx(self.copies / "a.docx", [text_and_picture, self.second],
                   embed_png=True, image_bytes=replacement.getvalue())
        self.assert_success(self.run_merge())
        document = self.read_document()
        rid = document.find(f".//{{{DRAWING}}}blip").get(f"{{{OFFICE_REL}}}embed")
        self.assertNotIn("[image 插入]", "".join(document.itertext()))
        with ZipFile(self.output) as package:
            rels = etree.fromstring(package.read("word/_rels/document.xml.rels"))
            target = next(rel.get("Target") for rel in rels if rel.get("Id") == rid)
            self.assertEqual(package.read(f"word/{target}"), replacement.getvalue())

    def test_two_departments_can_use_the_same_drawing_id(self):
        picture = ('<w:r><w:drawing><wp:inline><wp:docPr id="1" name="Picture 1"/>'
                   '<a:graphic><a:graphicData><a:blip r:embed="rId1"/></a:graphicData>'
                   '</a:graphic></wp:inline></w:drawing></w:r>')
        first = self.first.replace("</w:p>", picture + "</w:p>")
        second = self.second.replace("</w:p>", picture + "</w:p>")
        write_docx(self.copies / "a.docx", [first, self.second], embed_png=True)
        write_docx(self.copies / "b.docx", [self.first, second], embed_png=True)
        self.assert_success(self.run_merge())
        drawings = self.read_document().xpath('//*[local-name()="docPr"]')
        self.assertEqual(len(drawings), 2)
        self.assertEqual(len({node.get("id") for node in drawings}), 2)

    def test_tracked_insertion_and_deletion_are_preserved_with_unique_ids(self):
        tracked = ('<w:p><w:r><w:t xml:space="preserve">Alpha department completed </w:t></w:r>'
                   '<w:del w:id="1" w:author="Alice"><w:r><w:delText>10</w:delText></w:r></w:del>'
                   '<w:ins w:id="2" w:author="Alice"><w:r><w:t>11</w:t></w:r></w:ins>'
                   '<w:r><w:t xml:space="preserve"> projects.</w:t></w:r></w:p>')
        addition = '<w:ins w:id="1" w:author="Bob"><w:r><w:t> Confirmed.</w:t></w:r></w:ins>'
        write_docx(self.copies / "a.docx", [tracked, self.second])
        write_docx(self.copies / "b.docx", [self.first, self.second.replace("</w:p>", addition + "</w:p>")])
        self.assert_success(self.run_merge())
        document = self.read_document()
        deleted = document.findall(f".//{{{WORD_NS}}}del")
        inserted = document.findall(f".//{{{WORD_NS}}}ins")
        self.assertEqual(["".join(node.itertext()) for node in deleted], ["10"])
        self.assertEqual(["".join(node.itertext()) for node in inserted], ["11", " Confirmed."])
        self.assertEqual([node.get(f"{{{WORD_NS}}}author") for node in inserted], ["Alice", "Bob"])
        self.assertEqual(len({node.get(f"{{{WORD_NS}}}id") for node in deleted + inserted}), 3)

    def test_first_png_has_content_type_and_missing_type_aborts(self):
        image = (
            '<w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData>'
            '<a:blip r:embed="rId1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
        )
        copy = self.copies / "a.docx"
        write_docx(copy, [self.first, self.second, image], embed_png=True)
        self.assert_success(self.run_merge())
        with ZipFile(self.output) as package:
            types = etree.fromstring(package.read("[Content_Types].xml"))
            matches = [entry.get("PartName") for entry in types if entry.get("ContentType") == "image/png"]
            self.assertEqual(len(matches), 1)
            self.assertEqual(package.read(matches[0].lstrip("/")), PNG)

        old_output = hashlib.sha256(self.output.read_bytes()).hexdigest()
        write_docx(copy, [self.first, self.second, image], embed_png=True, declare_png=False)
        result = self.run_merge()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("资源缺少内容类型", result.stdout)
        self.assertEqual(hashlib.sha256(self.output.read_bytes()).hexdigest(), old_output)


if __name__ == "__main__":
    unittest.main()
