import io
import sys
import unittest
from unittest import mock


class CliOptionTests(unittest.TestCase):
    def test_no_postprocess_option_is_rejected(self):
        import thesis_format_cli

        stderr = io.StringIO()
        argv = ["thesis_format_cli.py", "--no-postprocess"]

        with mock.patch.object(sys, "argv", argv), \
                mock.patch("sys.stderr", stderr), \
                mock.patch("thesis_format_cli.FormatterGUI") as gui_cls:
            with self.assertRaises(SystemExit) as exc:
                thesis_format_cli.main()

        self.assertEqual(exc.exception.code, 2)
        self.assertIn("--no-postprocess", stderr.getvalue())
        gui_cls.assert_not_called()


if __name__ == "__main__":
    unittest.main()
