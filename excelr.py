import os
import shutil
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import pyximport; pyximport.install()  # TODO: build so that we don't need cython as dependency
from _excelr import write_rows


src_path = Path(__file__).parent / 'xlsx_template'


files = [
    Path(root).relative_to(src_path) / file
    for root, dirs, files in os.walk(src_path)
    for file in files
]

dirs = [
    Path(root).relative_to(src_path) / dir
    for root, dirs, _ in os.walk(src_path)
    for dir in dirs
]


def to_excel(output, rows):
    """
    Create a simple excel file from the given rows, writing to the given file.

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

        # output cells to the worksheet (template already contains <sheetData>
        # and <worksheet> xml opening tags)
        with open(dst_path / 'xl' / 'worksheets' / 'sheet1.xml', 'a') as f:
            write_rows(f, rows)
            f.write("</sheetData></worksheet>")

        # convert to an xlsx file which is actually just a zip file containing
        # the various xml files we generated.
        with ZipFile(output, 'w') as zf:
            for file in files:
                zf.write(dst_path / file, file)
