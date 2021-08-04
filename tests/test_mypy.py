import subprocess
import unittest
from pathlib import Path


class TestMyPy(unittest.TestCase):
    """
    Test that runs the mypy type checker.
    """
    def test_mypy(self):
        p = subprocess.run(['mypy', Path(__file__).parent.parent/'excelr.py'])
        self.assertEqual(p.returncode, 0, 'mypy type checking failed')
