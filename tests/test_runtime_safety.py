import copy
import unittest
import warnings
from types import SimpleNamespace
from unittest import mock

from docx import Document

import thesis_format_cli
import word_postprocess
from thesis_config import DEFAULT_CONFIG
from thesis_formatter import headings


class _FakeStdin:
    def __init__(self, is_tty=True, exc=None):
        self._is_tty = is_tty
        self._exc = exc

    def isatty(self):
        if self._exc is not None:
            raise self._exc
        return self._is_tty


class RuntimeSafetyTests(unittest.TestCase):
    def test_should_prompt_before_exit_only_for_interactive_frozen_exe(self):
        with mock.patch.object(thesis_format_cli.sys, "stdin", _FakeStdin(True)):
            with mock.patch.object(thesis_format_cli.sys, "frozen", True, create=True):
                self.assertTrue(thesis_format_cli.should_prompt_before_exit())

            with mock.patch.object(thesis_format_cli.sys, "frozen", False, create=True):
                self.assertFalse(thesis_format_cli.should_prompt_before_exit())

        with mock.patch.object(thesis_format_cli.sys, "stdin", _FakeStdin(exc=RuntimeError("boom"))):
            with mock.patch.object(thesis_format_cli.sys, "frozen", True, create=True):
                self.assertFalse(thesis_format_cli.should_prompt_before_exit())

    def test_terminate_process_targets_specific_pid(self):
        with mock.patch.object(word_postprocess.subprocess, "run", return_value=SimpleNamespace(returncode=0)) as run_mock:
            self.assertTrue(word_postprocess._terminate_process(4321))

        run_mock.assert_called_once_with(
            ["taskkill", "/F", "/PID", "4321"],
            capture_output=True,
            text=True,
            timeout=5,
        )

    def test_normalize_heading_spacing_handles_english_chapter_without_warning(self):
        doc = Document()
        para = doc.add_paragraph("Chapter 2 Introduction", style="Heading 1")

        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            headings.normalize_heading_spacing(doc, copy.deepcopy(DEFAULT_CONFIG))

        self.assertEqual(para.text, "Chapter 2  Introduction")
        self.assertFalse(
            any(issubclass(item.category, DeprecationWarning) for item in caught),
            caught,
        )

    def test_renumber_headings_accepts_skip_para_ids(self):
        doc = Document()
        preserved = doc.add_paragraph("第9章 保留", style="Heading 1")
        body = doc.add_paragraph("第7章 正文", style="Heading 1")

        changes = headings.renumber_headings(
            doc,
            copy.deepcopy(DEFAULT_CONFIG),
            skip_para_ids={id(preserved._element)},
        )

        self.assertEqual(preserved.text, "第9章 保留")
        self.assertEqual(body.text, "第1章 正文")
        self.assertTrue(any("第7章 正文" in change for change in changes), changes)


if __name__ == "__main__":
    unittest.main()