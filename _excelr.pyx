import string
from datetime import date
from itertools import product


type_map = {bool: 'b', float: 'n', int: 'n', date: 'd'}


column_coordinates = [
    ''.join(x)
    for i in range(1, 4)
    for x in product(string.ascii_uppercase, repeat=i)
]


cdef _write_cell(f, value, str col, int row):
    data_type = type_map.get(type(value), 't')
    if data_type == 'b':
        value = int(value)
    elif data_type == 't':
        value = value.replace("&", "&amp;").replace(">", "&gt;").replace("<", "&lt;")
    f.write(f'<c r="{col}{row}" t="{data_type}"><v>{value}</v></c>')


def write_rows(f, rows):
    for row, row_values in enumerate(rows, 1):
        f.write(f'<row r="{row}">')
        for col, value in zip(column_coordinates, row_values):
            _write_cell(f, value, col, row)
        f.write(f'</row>')
