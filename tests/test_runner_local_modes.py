import tempfile
import unittest
from pathlib import Path
from unittest import mock

import yaml
from docx import Document

from thesis_config import dump_default_config
from thesis_runner import run_format


def _make_config():
    cfg = yaml.safe_load(dump_default_config())
    cfg["cover"]["enabled"] = False
    cfg["toc"]["enabled"] = False
    return cfg


def _build_input_doc(path):
    doc = Document()
    doc.add_paragraph("测试文档")
    doc.save(path)


class RunnerLocalModesTests(unittest.TestCase):
    @mock.patch("thesis_runner.apply_format", return_value=[])
    @mock.patch("thesis_runner.postprocess")
    def test_page_number_only_mode_skips_postprocess(self, mock_postprocess, _mock_apply):
        cfg = _make_config()
        cfg["page_numbers"]["only_insert"] = True
        logs = []

        with tempfile.TemporaryDirectory() as tmp:
            input_path = Path(tmp) / "input.docx"
            output_path = Path(tmp) / "output.docx"
            _build_input_doc(input_path)

            ok = run_format(str(input_path), str(output_path), logs.append, config=cfg)

        self.assertTrue(ok)
        mock_postprocess.assert_not_called()

    @mock.patch("thesis_runner.apply_format", return_value=[])
    @mock.patch("thesis_runner.postprocess")
    def test_header_only_mode_with_chapter_title_runs_fields_only(self, mock_postprocess, _mock_apply):
        cfg = _make_config()
        cfg["header_footer"]["enabled"] = True
        cfg["header_footer"]["only_insert"] = True
        cfg["header_footer"]["odd_page_text"] = "{chapter_title}"
        logs = []

        with tempfile.TemporaryDirectory() as tmp:
            input_path = Path(tmp) / "input.docx"
            output_path = Path(tmp) / "output.docx"
            _build_input_doc(input_path)

            ok = run_format(str(input_path), str(output_path), logs.append, config=cfg)

        self.assertTrue(ok)
        mock_postprocess.assert_called_once_with(str(output_path), config=cfg, mode="fields_only")


if __name__ == "__main__":
    unittest.main()
