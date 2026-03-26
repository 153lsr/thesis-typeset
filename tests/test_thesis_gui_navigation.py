import sys
import tempfile
import tkinter
import unittest
from pathlib import Path
from unittest import mock


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


_ORIGINAL_MAINLOOP = tkinter.Misc.mainloop


class FormatterGUINavigationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        tkinter.Misc.mainloop = lambda self, n=0: None

    @classmethod
    def tearDownClass(cls):
        tkinter.Misc.mainloop = _ORIGINAL_MAINLOOP

    def setUp(self):
        import thesis_gui

        self.gui = thesis_gui.FormatterGUI(theme="sandstone")

    def tearDown(self):
        self.gui._root.destroy()

    def test_on_cat_select_is_safe_without_list_widget(self):
        self.assertFalse(hasattr(self.gui, "_cat_list"))

        self.gui._cur_panel = "页面"
        self.gui._on_cat_select()

        self.assertEqual(self.gui._cur_panel, "页面")

    def test_start_passes_log_callback_not_skip_flag_to_run_format(self):
        with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
            self.gui._v_in.set(tmp.name)
            self.gui._v_out.set(str(Path(tmp.name).with_name("out.docx")))
            self.gui._v_skip.set(True)

            calls = []

            def fake_run_format(*args, **kwargs):
                calls.append((args, kwargs))
                return True

            class ImmediateThread:
                def __init__(self, target=None, daemon=None):
                    self._target = target

                def start(self):
                    if self._target:
                        self._target()

            with mock.patch("thesis_gui.run_format", side_effect=fake_run_format), \
                    mock.patch("thesis_gui.threading.Thread", ImmediateThread):
                self.gui._start()

            self.assertEqual(len(calls), 1)
            args, kwargs = calls[0]
            self.assertEqual(args[0], tmp.name)
            self.assertEqual(args[1], self.gui._v_out.get())
            self.assertEqual(args[2], self.gui._append_log)
            self.assertFalse(kwargs["config"]["toc"]["enabled"])

    def test_standalone_modes_are_mutually_exclusive_and_clear_skip(self):
        self.gui._v_skip.set(True)
        self.gui._v_toc_only.set(True)
        self.gui._v_cover_only.set(True)
        self.gui._v_pn_only.set(True)
        self.gui._v_hf_only.set(True)

        self.gui._on_pn_only_toggle()

        self.assertFalse(self.gui._v_skip.get())
        self.assertFalse(self.gui._v_toc_only.get())
        self.assertFalse(self.gui._v_cover_only.get())
        self.assertTrue(self.gui._v_pn_only.get())
        self.assertFalse(self.gui._v_hf_only.get())

    def test_start_logs_header_only_mode(self):
        with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
            self.gui._v_in.set(tmp.name)
            self.gui._v_out.set(str(Path(tmp.name).with_name("out.docx")))
            self.gui._v_hf_only.set(True)

            calls = []
            logs = []

            def fake_run_format(*args, **kwargs):
                calls.append((args, kwargs))
                return True

            class ImmediateThread:
                def __init__(self, target=None, daemon=None):
                    self._target = target

                def start(self):
                    if self._target:
                        self._target()

            with mock.patch("thesis_gui.run_format", side_effect=fake_run_format), \
                    mock.patch("thesis_gui.threading.Thread", ImmediateThread), \
                    mock.patch.object(self.gui, "_append_log", side_effect=logs.append):
                self.gui._start()

            self.assertEqual(len(calls), 1)
            self.assertIn("单独处理: 仅更新页眉", logs)


if __name__ == "__main__":
    unittest.main()
