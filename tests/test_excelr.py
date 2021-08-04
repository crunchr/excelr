# std
from xml.etree import ElementTree
from zipfile import ZipFile
from io import BytesIO
from unittest import TestCase
# excelr
from excelr import to_excel


class Tests(TestCase):

    def test_strings_should_be_inline(self):
        """
        Verify that strings are written as inlineStr.
        """
        with BytesIO() as io:
            to_excel(io, ['a'])
            io.seek(0)
            with ZipFile(io) as z:
                with z.open('xl/worksheets/sheet1.xml') as f:
                    tree = ElementTree.fromstring(f.read().decode())
                    self.assertEqual(tree[4][0][0].attrib['t'], 'inlineStr')
