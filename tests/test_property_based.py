# std
from xml.etree import ElementTree

from zipfile import ZipFile

from io import StringIO, BytesIO

from itertools import zip_longest
from tempfile import NamedTemporaryFile
from unittest import TestCase
# 3rd party
from hypothesis import given, strategies as st
from openpyxl import load_workbook
# excelr
from excelr import to_excel


class Tests(TestCase):

    @classmethod
    def _test_generated_excel_should_equal_input(cls, draw, st_cells):
        """
        Common code for test_generated_excel_should_equal_input functions which:

        - draws some example data to write to an excel file
        - creates an excel file using excelr
        - reads the file back in using openpyxl
        """
        # draw some example cells
        data = draw(st.integers(min_value=1, max_value=16).flatmap(
            lambda num_cols:
            st.lists(
                st.lists(st_cells, min_size=num_cols, max_size=num_cols),
                min_size=1,
                max_size=16,
            )
        ))

        # generate an excel file with using exclr
        with NamedTemporaryFile('wb', delete=False, suffix='.xlsx') as s:
            to_excel(s, data)

        # use openpyxl to read the generated excel file
        actual = [list(x) for x in load_workbook(s.name)['Sheet1'].values]

        # We expect None values to have been converted to '-'
        expected = [['-' if col is None else col for col in row] for row in data]

        return expected, actual

    @given(st.data())
    def test_generated_excel_should_equal_input(self, data):
        """
        Check the property that the generated excel file should be readable
        by openpyxl and the data read from that file should be the same as the
        input data.
        """
        st_cells = st.one_of(
            st.none(),
            st.integers(),
            st.text(st.characters(blacklist_categories=['Cs', 'Cc']), min_size=1),
            st.dates(),
            st.booleans(),
        )
        expected, actual = self._test_generated_excel_should_equal_input(data.draw, st_cells)
        self.assertEqual(expected, actual)

    @given(st.data())
    def test_generated_excel_should_equal_input_for_float(self, data):
        """
        Check the property that the generated excel file should be readable
        by openpyxl and the data read from that file should be the same as the
        input data (we do floats separately since we need to use almostEqual)
        """
        st_cells = st.floats(allow_infinity=False, allow_nan=False)
        expected, actual = self._test_generated_excel_should_equal_input(data.draw, st_cells)
        for expected_row, actual_row in zip_longest(expected, actual):
            for expected_cell, actual_cell in zip_longest(expected_row, actual_row):
                self.assertAlmostEqual(expected_cell, actual_cell)

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
