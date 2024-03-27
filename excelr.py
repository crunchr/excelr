import os
import shutil
import string
import typing
from datetime import date, datetime, time
from itertools import product
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Union, IO, Iterable, Optional
from zipfile import ZipFile, ZIP_DEFLATED


src_path = Path(__file__).parent / 'xlsx_template'


# collect all the files from the excel template that should be copied
# when generating a new excel file.
files = [
    Path(root).relative_to(src_path) / file
    for root, dirs, files in os.walk(src_path)
    for file in files
]


# collect all the directories from the excel template that should be copied
# when generating a new excel file.
dirs = [
    Path(root).relative_to(src_path) / dir
    for root, dirs, _ in os.walk(src_path)
    for dir in dirs
]

TYPE_MAP = {bool: 'b', float: 'n', int: 'n', date: 'd', time: 'n', datetime: 'n'}

# enough for 18278 columns
COLUMN_COORDINATES = [
    ''.join(x)
    for i in range(1, 4)
    for x in product(string.ascii_uppercase, repeat=i)
]


StrPath = Union[str, os.PathLike]
Value = Optional[Union[bool, float, int, date, str]]
Rows = Iterable[Iterable[Value]]


def to_excel(
    output: Union[StrPath, IO[bytes]],
    rows: Rows,
    column_format_codes: dict[int, str] = None,
    coalesce: Union[str, None] = None
) -> Union[StrPath, IO[bytes]]:
    """
    Create a simple excel file from the given rows, writing to the desired
    output file.

    :param output: path to file, or file like object.
    :param rows: Iterable of Iterables containing rows / columns to export
    :param column_format_codes: Dictionary mapping column indexes to format codes, see
                                https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
                                for information about supported format codes.
    :param coalesce: string to coalesce `None` values to.
    """
    with TemporaryDirectory() as d:

        # the bare bones of the excel file are contained in the directory,
        # xlsx_template - copy this to the temporary directory and use this as
        # a starting point for the excel file
        dst_path = Path(d) / 'excelr'
        for dir in dirs:
            os.makedirs(dst_path / dir)
        for file in files:
            shutil.copy(src_path / file, dst_path / file)

        # No need to repeat format codes when they occur for multiple columns
        _column_format_codes = column_format_codes or {}
        format_codes = list(set(_column_format_codes.values()))

        # generate the indexes which will be used to reference the actual styles
        style_indexes = {
            column_index: format_codes.index(format_code) + 1  # +1 to skip 'General'
            for column_index, format_code in _column_format_codes.items()
        }

        # Here we read the styles file into memory, however the file is small
        num_fmts = _get_num_fmts_xml(format_codes)
        cell_xfs = _get_cell_xfs(len(format_codes))

        styles_path = Path(dst_path / 'xl' / 'styles.xml')
        styles_path.write_text(
            styles_path
            .read_text()
            .replace("{{ numFmts }}", num_fmts)
            .replace("{{ cellXfs }}", cell_xfs)
        )

        # the sheet1.xml file is incomplete, fill it in now by writing
        # output cells to the worksheet (template already contains <sheetData>
        # and <worksheet> xml opening tags)
        with open(dst_path / 'xl' / 'worksheets' / 'sheet1.xml', 'a') as f:

            for row, row_values in enumerate(rows, 1):

                f.write(f'<row r="{row}">')

                for i, (col, value) in enumerate(zip(COLUMN_COORDINATES, row_values)):

                    data_type = TYPE_MAP.get(type(value), 'inlineStr')
                    tag_start, tag_end = '<v>', '</v>'

                    if value is None:
                        value = coalesce

                    if isinstance(value, (datetime, time)):
                        value = _to_excel_serial(value)

                    if type(value) is bool:
                        value = int(value)

                    # use inlineStr so that we don't need to keep track of
                    # shared_strings.xml which would increase the memory
                    # we need. Since xlsx is a compressed format this shouldn't
                    # affect the size of the generated file too much.
                    elif data_type == 'inlineStr':
                        value = _xml_escape(str(value))
                        tag_start, tag_end = '<is><t>', '</t></is>'

                    f.write(f'<c r="{col}{row}" '
                            f's="{style_indexes.get(i, 0)}" '
                            f't="{data_type}">{tag_start}{value}{tag_end}</c>')

                f.write(f'</row>')

            f.write("</sheetData></worksheet>")

        # convert to an xlsx file which is actually just a zip file containing
        # the various xml files we generated.
        with ZipFile(output, 'w', compression=ZIP_DEFLATED) as zf:
            for file in files:
                zf.write(dst_path / file, file)

        return output


def _get_num_fmts_xml(column_format_codes: list[str]) -> str:
    """
    Get the xml fragment for the numFmts part of styles.xml
    """
    _column_format_codes = ['General'] + column_format_codes
    return ''.join((
        f'<numFmts count="{len(_column_format_codes)}">',
        *(
            f'<numFmt numFmtId="{num_fmt_id}" formatCode="{format_code}"/>'
            for num_fmt_id, format_code in enumerate(_column_format_codes, start=164)
        ),
        f'</numFmts>',
    ))


def _get_cell_xfs(num_column_format_codes: int) -> str:
    """
    Get the xml fragment for the cellXfs part of styles.xml
    """
    _num_column_format_codes = num_column_format_codes + 1  # + 1 for General
    return ''.join((
        f'<cellXfs count="{_num_column_format_codes}">',
        *(
            f"""
                <xf numFmtId="{num_fmt_id + 164}" fontId="0" fillId="0" borderId="0" xfId="0"
                    applyFont="false" applyBorder="false" applyAlignment="false"
                    applyProtection="false">
                    <alignment horizontal="general" vertical="bottom" textRotation="0"
                               wrapText="false" indent="0" shrinkToFit="false"/>
                    <protection locked="true" hidden="false"/>
                </xf>
            """
            for num_fmt_id in range(_num_column_format_codes)
        ),
        f'</cellXfs>',
    ))


def _xml_escape(value: str) -> str:
    """
    Escape special xml characters.

    :param value: Value to escape.

    :return: Escaped value.
    """
    return value.replace("&", "&amp;").replace(">", "&gt;").replace("<", "&lt;")


def _to_excel_serial(value: typing.Union[date, time, datetime]) -> float:
    """
    Dates and times are stored as numbers in excel, whole part is the number of
    days since 1900. NOTE excel incorrectly assumes 1900 is a leap year.
    Fractional part is the time component (seconds / seconds_in_day).

    See:
        https://support.microsoft.com/en-us/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487
        https://en.wikipedia.org/wiki/Year_1900_problem

    :param value: The date value to convert.

    :return: number representing the given date.
    """
    base_date = date(1899, 12, 31)  # Adjusted base date to December 31, 1899
    if isinstance(value, datetime):
        return _to_excel_serial(value.date()) + _to_excel_serial(value.time())
    elif isinstance(value, date):
        # Calculate delta considering base date as December 31, 1899
        delta = value - base_date
        return delta.days + 1  # Adding 1 because Excel starts counting from 1
    elif isinstance(value, time):
        # Calculate the fractional day for time values
        seconds_in_day = 24 * 60 * 60
        fractional_day = (value.hour * 3600 + value.minute * 60 + value.second) / seconds_in_day
        return fractional_day
    else:
        raise TypeError(type(value))
