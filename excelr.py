import os
import shutil
import string
from datetime import date
from itertools import product
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Union, IO, Iterable, Optional
from zipfile import ZipFile


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


TYPE_MAP = {bool: 'b', float: 'n', int: 'n', date: 'd'}


# enough for 18278 columns
COLUMN_COORDINATES = [
    ''.join(x)
    for i in range(1, 4)
    for x in product(string.ascii_uppercase, repeat=i)
]


StrPath = Union[str, os.PathLike]
Value = Optional[Union[bool, float, int, date, str]]
Rows = Iterable[Iterable[Value]]


def to_excel(output: Union[StrPath, IO[bytes]], rows: Rows):
    """
    Create a simple excel file from the given rows, writing to the desired
    output file.

    :param output: path to file, or file like object.
    :param rows: Iterable of Iterables containing rows / columns to export
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

        # the sheet1.xml file is incomplete, fill it in now by writing
        # output cells to the worksheet (template already contains <sheetData>
        # and <worksheet> xml opening tags)
        with open(dst_path / 'xl' / 'worksheets' / 'sheet1.xml', 'a') as f:

            for row, row_values in enumerate(rows, 1):

                f.write(f'<row r="{row}">')

                for col, value in zip(COLUMN_COORDINATES, row_values):

                    data_type = TYPE_MAP.get(type(value), 'inlineStr')
                    tag_start, tag_end = '<v>', '</v>'

                    if value is None:
                        value = '-'

                    if data_type == 'b':
                        assert isinstance(value, bool)
                        value = int(value)

                    # use inlineStr so that we don't need to keep track of
                    # shared_strings.xml which would increase the memory
                    # we need. Since xlsx is a compressed format this shouldn't
                    # affect the size of the generated file too much.
                    elif data_type == 'inlineStr':
                        assert isinstance(value, str)
                        value = _xml_escape(value)
                        tag_start, tag_end = '<is><t>', '</t></is>'

                    f.write(f'<c r="{col}{row}" t="{data_type}">{tag_start}{value}{tag_end}</c>')

                f.write(f'</row>')

            f.write("</sheetData></worksheet>")

        # convert to an xlsx file which is actually just a zip file containing
        # the various xml files we generated.
        with ZipFile(output, 'w') as zf:
            for file in files:
                zf.write(dst_path / file, file)


def _xml_escape(value: str) -> str:
    """
    Escape special xml characters.

    :param value: Value to escape.

    :return: Escaped value.
    """
    return value.replace("&", "&amp;").replace(">", "&gt;").replace("<", "&lt;")
