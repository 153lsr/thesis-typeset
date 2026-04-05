import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


from word_postprocess import PostprocessError


class WordPostprocessSmokeTests(unittest.TestCase):
    def test_postprocess_error_inherits_runtime_error(self):
        self.assertTrue(issubclass(PostprocessError, RuntimeError))


if __name__ == "__main__":
    unittest.main()
