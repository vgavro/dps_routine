#!/usr/bin/env python

import xml.etree.cElementTree as ET
from datetime import datetime
import os.path

import xlrd


DECLARHEAD_FIELDS = ('TIN,C_DOC,C_DOC_SUB,C_DOC_VER,C_DOC_TYPE,C_DOC_CNT,'
                     'C_REG,C_RAJ,PERIOD_MONTH,PERIOD_TYPE,PERIOD_YEAR,'
                     'C_STI_ORIG,C_DOC_STAN,D_FILL'.split(','))
# Assuming any other field will be appended to body

DEFAULTS = {
    'C_DOC': 'F01',
    'C_DOC_TYPE': 0,
    'C_DOC_VER': 5,
    'C_DOC_CNT': 1,
}


def create_xml(data):
    root = ET.Element('DECLAR')
    head = ET.SubElement(root, 'DECLARHEAD')
    body = ET.SubElement(root, 'DECLARBODY')

    data_ = DEFAULTS.copy()
    data_.update(data)

    for key, value in data_.items():
        if callable(value):
            value = value(data_)

        tag = key.split(' ')[0]
        tag_str = '<{}></{}>'.format(key, tag)
        try:
            e = ET.fromstring(tag_str)
        except ET.ParseError:
            raise RuntimeError('Invalid field {}: could not parse {}'
                               .format(key, tag_str))

        if isinstance(value, float) and value == int(value):
            value = int(value)
        e.text = str(value)

        if tag in DECLARHEAD_FIELDS:
            head.append(e)
        else:
            body.append(e)

    return ET.ElementTree(root)


def write_xml(data, filename=None, encoding='windows-1251'):
    filename = filename or (datetime.now().isoformat() + '.xml')
    create_xml(data).write(filename, encoding)


def main(xlsx_filename='Книга1.xlsx', sheet_index=0,
         fields_row_index=1, data_start_row_index=2, supress_exc=False):

    book = xlrd.open_workbook(xlsx_filename)
    sheet = book.sheet_by_index(sheet_index)
    fields = sheet.row_values(fields_row_index)

    output_dir = os.path.basename(xlsx_filename) + '_xml'
    output_dir = os.path.join(os.path.dirname(xlsx_filename), output_dir)
    try:
        os.mkdir(output_dir)
    except OSError:
        pass

    for i in range(data_start_row_index, sheet.nrows):
        data = dict(zip(fields, sheet.row_values(i)))
        try:
            write_xml(data, os.path.join(output_dir, '{:02d}.xml'.format(i)))
        except Exception as exc:
            if not supress_exc:
                raise
            print('SKIPPED {}: {}: {}'.format(i, data, exc))


if __name__ == '__main__':
    from sys import argv
    (len(argv) > 1) and main(argv[1]) or main()
