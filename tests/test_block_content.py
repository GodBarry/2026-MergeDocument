from io import BytesIO
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from lxml import etree
from PIL import Image

from test_merge_regressions import DRAWING, OFFICE_REL, ROOT, WORD_NS, paragraph, write_docx


W = f"{{{WORD_NS}}}"


def make_table(*cells):
    return (
        '<w:tbl><w:tblGrid>' + '<w:gridCol w:w="2400"/>' * len(cells) + '</w:tblGrid>'
        '<w:tr>' + ''.join(f'<w:tc>{content}</w:tc>' for content in cells) + '</w:tr></w:tbl>'
    )


def control(*paragraphs):
    return (
        '<w:sdt><w:sdtPr><w:tag w:val="quantity"/><w:id w:val="7"/></w:sdtPr>'
        '<w:sdtContent>' + ''.join(paragraphs) + '</w:sdtContent></w:sdt>'
    )


def text_of(element):
    return ''.join(node.text or '' for node in element.iter(W + 't')).strip()


class BlockContentTests(unittest.TestCase):
    def setUp(self):
        work = TemporaryDirectory()
        self.addCleanup(work.cleanup)
        self.folder = Path(work.name)
        self.original = self.folder / 'base.docx'
        self.copies = self.folder / 'copies'
        self.copies.mkdir()
        self.output = self.folder / 'merged.docx'

    def merge(self):
        command = [sys.executable, '-X', 'utf8', '-m', 'mddm.run_merge',
                   str(self.original), str(self.copies), '-o', str(self.output)]
        result = subprocess.run(command, cwd=ROOT, capture_output=True, encoding='utf-8', timeout=30)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        with ZipFile(self.output) as package:
            return etree.fromstring(package.read('word/document.xml')).find(W + 'body')

    def report(self):
        return self.output.with_name('merged_conflicts.txt').read_text(encoding='utf-8')

    def test_multi_paragraph_cells_keep_structure_and_independent_edits(self):
        description = paragraph('First line') + paragraph('Budget 20')
        memo = paragraph('Unchanged')
        edited = paragraph('First line') + '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Budget 30</w:t></w:r></w:p>'
        write_docx(self.original, [make_table(description, memo)])
        write_docx(self.copies / 'a.docx', [make_table(edited, memo)])
        write_docx(self.copies / 'b.docx', [make_table(description, paragraph('Approved'))])
        cells = list(self.merge().iter(W + 'tc'))
        self.assertEqual([[text_of(p) for p in cell.findall(W + 'p')] for cell in cells],
                         [['First line', 'Budget 30'], ['Approved']])
        self.assertIsNotNone(cells[0].find('.//' + W + 'b'))
        self.assertIn('冲突数量: 0', self.report())
        self.assertIn('自动合并变更数: 2', self.report())

    def test_table_image_compares_target_bytes_and_ignores_local_ids(self):
        picture = '<w:p><w:r><w:drawing><a:blip r:embed="{rid}"/></w:drawing></w:r></w:p>'
        write_docx(self.original, [make_table(picture.format(rid='rId1'))], embed_png=True)
        copy = self.copies / 'a.docx'
        write_docx(copy, [make_table(picture.format(rid='rId8'))], embed_png=True, image_id='rId8')
        self.merge()
        self.assertIn('自动合并变更数: 0', self.report())

        changed = BytesIO()
        Image.new('RGB', (2, 2), 'green').save(changed, format='PNG')
        write_docx(copy, [make_table(picture.format(rid='rId1'))], embed_png=True,
                   image_bytes=changed.getvalue())
        blip = self.merge().find('.//' + f'{{{DRAWING}}}blip')
        reference = blip.get(f'{{{OFFICE_REL}}}embed')
        with ZipFile(self.output) as package:
            rels = etree.fromstring(package.read('word/_rels/document.xml.rels'))
            target = next(rel.get('Target') for rel in rels if rel.get('Id') == reference)
            self.assertEqual(package.read('word/' + target), changed.getvalue())
        self.assertIn('自动合并变更数: 1', self.report())

    def test_control_preserves_metadata_and_multiple_paragraphs(self):
        original = control(paragraph('Delivery'), paragraph('10 units'))
        edited = control(paragraph('Delivery'), paragraph('40 units'))
        write_docx(self.original, [original])
        write_docx(self.copies / 'a.docx', [edited])
        block, = self.merge().findall(W + 'sdt')
        self.assertEqual(block.find(W + 'sdtPr').find(W + 'id').get(W + 'val'), '7')
        self.assertEqual([text_of(p) for p in block.iter(W + 'p')], ['Delivery', '40 units'])
        self.assertIn('自动合并变更数: 1', self.report())

    def test_multiple_append_blocks_keep_source_order_and_final_section_properties(self):
        for original in ([], [paragraph('A'), paragraph('B')]):
            with self.subTest(original=original):
                additions = [paragraph('C'), make_table(paragraph('D')), control(paragraph('E'))]
                write_docx(self.original, original)
                write_docx(self.copies / 'a.docx', original + additions)
                body = self.merge()
                self.assertEqual([text_of(block) for block in body[:-1]],
                                 (['A', 'B'] if original else []) + ['C', 'D', 'E'])
                self.assertEqual(body[-1].tag, W + 'sectPr')

    def test_conflicting_control_edits_remain_separate_controls(self):
        write_docx(self.original, [control(paragraph('10 units'))])
        write_docx(self.copies / 'a.docx', [control(paragraph('40 units'))])
        write_docx(self.copies / 'b.docx', [control(paragraph('50 units'))])
        controls = self.merge().findall(W + 'sdt')
        self.assertTrue({'40 units', '50 units'}.issubset({text_of(node) for node in controls}))
        ids = [node.find(W + 'sdtPr').find(W + 'id').get(W + 'val') for node in controls]
        self.assertEqual(len(ids), len(set(ids)))
        self.assertIn('冲突数量: 1', self.report())


if __name__ == '__main__':
    unittest.main()
