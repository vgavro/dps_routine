#!/usr/bin/env python

import xml.etree.cElementTree as ET
from datetime import datetime
import os.path

import xlrd


FILENAME_TEMPLATE = '{C_STI_ORIG}{TIN}F010330510000000012032017{C_STI_ORIG}.xml'

DECLARHEAD_FIELDS = ('TIN,C_DOC,C_DOC_SUB,C_DOC_VER,C_DOC_TYPE,C_DOC_CNT,'
                     'C_REG,C_RAJ,PERIOD_MONTH,PERIOD_TYPE,PERIOD_YEAR,'
                     'C_STI_ORIG,C_DOC_STAN,D_FILL'.split(','))
# Assuming any other field will be appended to body

DEFAULTS = {
    'C_DOC': 'F01',
    'C_DOC_SUB': '033',
    'C_DOC_TYPE': 0,
    'C_DOC_VER': 5,
    'C_DOC_CNT': 1,
    'HZ': 1,
}


def create_xml(data):
    root = ET.Element('DECLAR', {'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
                                 'xsi:noNamespaceSchemaLocation': 'F0103305.xsd'})
    head = ET.SubElement(root, 'DECLARHEAD')
    ET.SubElement(head, 'LINKED_DOCS', {'xsi:nil': 'true'})
    ET.SubElement(head, 'SOFTWARE', {'xsi:nil': 'true'})

    body = ET.SubElement(root, 'DECLARBODY')

    data_ = DEFAULTS.copy()
    data_.update(data)

    for key, value in data_.items():
        if callable(value):
            value = value(data_)

        if not value and key not in ['C_DOC_TYPE']:
            continue

        tag = key.split(' ')[0]
        tag_str = '<{}></{}>'.format(key, tag)
        try:
            e = ET.fromstring(tag_str)
        except ET.ParseError:
            raise RuntimeError('Invalid field {}: could not parse {}'
                               .format(key, tag_str))
        e.text = str(value)

        if tag in DECLARHEAD_FIELDS:
            head.append(e)
        else:
            body.append(e)

    return ET.ElementTree(root)


def write_xml(data, output_dir='./', encoding='windows-1251'):
    filename = os.path.join(output_dir, FILENAME_TEMPLATE.format(**data))
    print('Creating {}'.format(filename))
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

    def parse_value(value):
        if isinstance(value, float) and value == int(value):
            value = int(value)
        return value

    for i in range(data_start_row_index, sheet.nrows):
        data = dict(zip(fields, map(parse_value, sheet.row_values(i))))
        try:
            write_xml(data, output_dir)
        except Exception as exc:
            if not supress_exc:
                raise
            print('SKIPPED {}: {}: {}'.format(i, data, exc))


if __name__ == '__main__':
    from sys import argv
    (len(argv) > 1) and main(argv[1]) or main()
    input('DONE. press any key to close')
