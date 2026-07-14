import tempfile
import unittest
from pathlib import Path

from src.utils.exceptions import PathSecurityError
from src.utils.path_validator import validate_safe_path


class TestPathValidator(unittest.TestCase):
    def test_normalizes_parent_components_and_hidden_paths(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            hidden = root / '.output' / '..' / '.output' / 'result.docx'

            result = validate_safe_path(str(hidden))

            self.assertEqual((root / '.output' / 'result.docx').resolve(), result)

    def test_rejects_paths_outside_an_explicit_base_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)

            with self.assertRaises(PathSecurityError):
                validate_safe_path(str(root.parent / 'outside.docx'), base_dir=str(root))

    def test_rejects_absolute_paths_when_disabled(self):
        with self.assertRaises(PathSecurityError):
            validate_safe_path('/tmp/output.docx', allow_absolute=False)


if __name__ == '__main__':
    unittest.main()
