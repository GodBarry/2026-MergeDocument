import os
from pathlib import Path
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from lxml import etree

from mddm.output_batch import publish_outputs
from test_merge_regressions import ROOT, WORD_NS, paragraph, table, write_docx


class OutputAndHeadingTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.folder = Path(temp.name)
        self.copies = self.folder / "副本"
        self.copies.mkdir()
        self.base = self.folder / "原稿.docx"
        self.outputs = [self.folder / name for name in
                        ("合稿.docx", "合稿_conflicts.txt", "合稿_merged_changes_log.txt")]

    def merge(self, *options):
        return subprocess.run(
            [sys.executable, "-X", "utf8", "-m", "mddm.run_merge", str(self.base),
             str(self.copies), "-o", str(self.outputs[0]), *options],
            cwd=ROOT, capture_output=True, encoding="utf-8", timeout=30,
        )

    def test_heading_after_blank_and_table_preserves_body_and_bookmarks(self):
        heading = ('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
                   '<w:bookmarkStart w:id="1" w:name="WorkPlan"/>'
                   '<w:r><w:t>第四章 </w:t></w:r>'
                   '<w:r><w:rPr><w:b/></w:rPr><w:t>工程安排</w:t></w:r>'
                   '<w:bookmarkEnd w:id="1"/></w:p>')
        blocks = [paragraph(""), table([["负责人"]]), paragraph("本周验收预算为十二万元。"),
                  heading, paragraph("项目已验收。")]
        write_docx(self.base, blocks)
        shutil.copyfile(self.base, self.copies / "部门.docx")
        result = self.merge("--auto-fix-headings")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        with ZipFile(self.outputs[0]) as package:
            root = etree.fromstring(package.read("word/document.xml"))
        paragraphs = root.xpath("w:body/w:p", namespaces={"w": WORD_NS})
        self.assertEqual(["".join(p.itertext()) for p in paragraphs],
                         ["", "本周验收预算为十二万元。", "第一章 工程安排", "项目已验收。"])
        heading = paragraphs[2]
        for local_name in ("bookmarkStart", "bookmarkEnd", "b"):
            self.assertEqual(len(heading.findall(f".//{{{WORD_NS}}}{local_name}")), 1)

    @unittest.skipUnless(os.name == "nt", "Windows delete-sharing semantics")
    def test_locked_output_aborts_entire_batch_and_retry_succeeds(self):
        write_docx(self.base, [paragraph("旧条款")])
        write_docx(self.copies / "部门.docx", [paragraph("新条款")])
        previous = [f"previous {i}".encode() for i in range(len(self.outputs))]
        for locked in self.outputs:
            with self.subTest(locked=locked.name):
                for path, old in zip(self.outputs, previous):
                    path.write_bytes(old)
                with locked.open("rb"):
                    result = self.merge()
                self.assertNotEqual(result.returncode, 0)
                self.assertNotIn("处理完成!", result.stdout)
                self.assertIn(locked.name, result.stdout + result.stderr)
                self.assertEqual([p.read_bytes() for p in self.outputs], previous)
                retry = self.merge()
                self.assertEqual(retry.returncode, 0, retry.stdout + retry.stderr)
                self.assertTrue(all(p.read_bytes() != old for p, old in zip(self.outputs, previous)))
                self.assertFalse(list(self.folder.glob(".mddm-*")))

    def test_publication_error_rolls_back_existing_and_new_files(self):
        staging = self.folder / "staging"
        staging.mkdir()
        for p in self.outputs:
            (staging / p.name).write_bytes(b"new")
        self.outputs[1].write_bytes(b"old conflict report")
        replace = os.replace

        def fail_last(source, target):
            if Path(source) == staging / self.outputs[2].name:
                raise OSError("simulated write error")
            return replace(source, target)

        with patch("mddm.output_batch.os.replace", side_effect=fail_last):
            with self.assertRaisesRegex(OSError, "simulated"):
                publish_outputs(staging / self.outputs[0].name, self.outputs[0])
        self.assertFalse(self.outputs[0].exists())
        self.assertEqual(self.outputs[1].read_bytes(), b"old conflict report")
        self.assertFalse(self.outputs[2].exists())
        self.assertFalse(list(self.folder.glob(".mddm-*")))

    def test_retry_recovers_batch_interrupted_after_first_new_file(self):
        staging = self.folder / "stage"
        staging.mkdir()
        for path in self.outputs:
            path.write_bytes(b"previous")
            (staging / path.name).write_bytes(b"updated")
        # Simulate process termination without a chance to execute rollback.
        command = """
import os, sys
from pathlib import Path
from mddm.output_batch import publish_outputs
original_replace = os.replace
staged, destination = map(Path, sys.argv[1:])
def terminate_after_first_file(source, target):
    original_replace(source, target)
    if Path(source) == staged:
        os._exit(73)
os.replace = terminate_after_first_file
publish_outputs(staged, destination)
"""
        staged = staging / self.outputs[0].name
        result = subprocess.run([sys.executable, "-c", command, str(staged), str(self.outputs[0])],
                                cwd=ROOT, timeout=10)
        self.assertEqual(result.returncode, 73)
        with self.assertRaises(FileNotFoundError):
            publish_outputs(staged, self.outputs[0])
        self.assertEqual([p.read_bytes() for p in self.outputs], [b"previous"] * 3)
        staged.write_bytes(b"updated")
        publish_outputs(staged, self.outputs[0])
        self.assertEqual([p.read_bytes() for p in self.outputs], [b"updated"] * 3)
