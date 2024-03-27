# std
import string
import warnings
from datetime import date, datetime, time
from itertools import zip_longest
from tempfile import NamedTemporaryFile
from unittest import TestCase
# 3rd party
from hypothesis import given, strategies as st, settings
from openpyxl.reader.excel import load_workbook

# excelr
from excelr import to_excel


class Tests(TestCase):

    @classmethod
    def _test_generated_excel_should_equal_input(cls, draw, cell_strategies):
        """
        Common code for test_generated_excel_should_equal_input functions which:

        - draws some example data to write to an excel file
        - creates an excel file using excelr
        - reads the file back in using pandas
        """
        num_cols = draw(st.integers(min_value=1, max_value=16))
        num_rows = draw(st.integers(min_value=1, max_value=16))
        column_strategies = [draw(st.sampled_from(cell_strategies)) for _ in range(num_cols)]
        data_types = [type(draw(s)) for s in column_strategies]

        # draw some example cells
        data = [[draw(s) for s in column_strategies] for _ in range(num_rows)]
        headers = [[string.ascii_letters[i] for i in range(num_cols)]]
        headers_and_data = headers + data

        column_format_codes = {}
        for i, t in enumerate(data_types):
            if t is datetime:
                column_format_codes[i] = "YYYY-MM-DD HH:MM:SS"
            elif t is time:
                column_format_codes[i] = "HH:MM:SS"

        # generate an excel file with using exclr
        with NamedTemporaryFile('wb', delete=False, suffix='.xlsx') as s:
            to_excel(
                s,
                headers_and_data,
                column_format_codes=column_format_codes,
                coalesce='-',
            )
            s.close()

        warnings.simplefilter("ignore", ResourceWarning)
        actual = [list(x) for x in load_workbook(s.name)['Sheet1'].values]

        # We expect None values to have been converted to '-'
        expected = [['-' if col is None else col for col in row] for row in headers_and_data]

        return expected, actual

    @given(st.data())
    @settings(deadline=None)
    def test_generated_excel_should_equal_input(self, data):
        """
        Check the property that the generated excel file should be readable
        by openpyxl and the data read from that file should be the same as the
        input data.
        """
        cell_strategies = [
            st.none(),
            st.integers(min_value=-2_147_483_647, max_value=2_147_483_647),
            st.dates(),
            st.times().map(lambda t: t.replace(microsecond=0)),
            st.datetimes().map(lambda t: t.replace(microsecond=0)),
            st.booleans(),
            st.text(st.characters(blacklist_categories=['Cs', 'Cc']), min_size=1),
        ]
        expected, actual = self._test_generated_excel_should_equal_input(data.draw, cell_strategies)
        self.assertEqual(expected, actual)

    @given(st.data())
    @settings(deadline=None)
    def test_generated_excel_should_equal_input_for_float(self, data):
        """
        Check the property that the generated excel file should be readable
        by openpyxl and the data read from that file should be the same as the
        input data (we do floats separately since we need to use almostEqual)
        """
        cell_strategies = [st.floats(allow_infinity=False, allow_nan=False)]
        expected, actual = self._test_generated_excel_should_equal_input(data.draw, cell_strategies)
        for expected_row, actual_row in zip_longest(expected, actual):
            for expected_cell, actual_cell in zip_longest(expected_row, actual_row):
                self.assertAlmostEqual(expected_cell, actual_cell)
