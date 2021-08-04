# std
import os
import subprocess
from pathlib import Path
from unittest import TestCase


class TestPyCodeStyle(TestCase):
    """
    Lint the code using pycodestyle.
    """
    def pycodestyle(self, path):
        p = subprocess.run(['pycodestyle', Path(__file__).parent.parent / path])
        self.assertEqual(p.returncode, 0, f'pycodestyle linter failed (path)')

    def test_pycodestyle_excelr(self):
        self.pycodestyle('excelr.py')

    def test_pycodestyle_tests(self):
        self.pycodestyle('tests')
