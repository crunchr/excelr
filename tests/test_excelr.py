# std
from xml.etree import ElementTree
from zipfile import ZipFile
from io import BytesIO
from unittest import TestCase
# excelr
from excelr import to_excel


class Tests(TestCase):

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        # create an excel file and parse the styles and worksheets
        with BytesIO() as io:
            to_excel(io, ['abc', [1, 2, 3]], {0: '0.00%', 2: '0.00%'})
            io.seek(0)
            with ZipFile(io) as z:
                with z.open('xl/styles.xml') as f:
                    cls.styles = ElementTree.fromstring(f.read().decode())
                with z.open('xl/worksheets/sheet1.xml') as f:
                    cls.sheet1 = ElementTree.fromstring(f.read().decode())

    def test_custom_format_code_should_be_in_styles(self):
        """
        Verify that styles.xml is generated correctly when a custom
        format is given.
        """
        num_fmts = self.styles[0]
        self.assertEqual(num_fmts.attrib['count'], "2")
        num_fmt = self.styles[0][0]
        self.assertEqual(num_fmt.attrib['numFmtId'], "164")
        self.assertEqual(num_fmt.attrib['formatCode'], 'General')
        num_fmt = self.styles[0][1]
        self.assertEqual(num_fmt.attrib['numFmtId'], "165")
        self.assertEqual(num_fmt.attrib['formatCode'], '0.00%')
        cell_xfs = self.styles[5]
        self.assertEqual(cell_xfs.attrib['count'], "2")
        self.assertEqual(cell_xfs[0].attrib['numFmtId'], "164")
        self.assertEqual(cell_xfs[1].attrib['numFmtId'], "165")

    def test_strings_should_be_inline(self):
        """
        Verify that strings are written as inlineStr.
        """
        self.assertEqual(self.sheet1[4][0][0].attrib['t'], 'inlineStr')

    def test_custom_format_code_should_be_selected_in_worksheet(self):
        """
        Verify that custom format codes are properly selected in
        sheet.xml
        """
        # first column should have custom format
        self.assertEqual(self.sheet1[4][0][0].attrib['s'], "1")
        # second column should not have custom format
        self.assertEqual(self.sheet1[4][0][1].attrib['s'], "0")
        # third column should have custom format
        self.assertEqual(self.sheet1[4][0][2].attrib['s'], "1")
