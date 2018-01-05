#!/usr/bin/env python

import xml.etree.cElementTree as ET
from datetime import datetime
import os.path
from collections import OrderedDict

import xlrd


FILENAME_TEMPLATE = ('{C_STI_ORIG}{TIN}'
                     '{C_DOC}{C_DOC_SUB}{C_DOC_VER:02d}'
                     '1{PERIOD_TYPE}{PERIOD_MONTH:02d}{PERIOD_YEAR}'
                     '{C_STI_ORIG}.xml')

DECLARHEAD_FIELDS = ('TIN,C_DOC,C_DOC_SUB,C_DOC_VER,C_DOC_TYPE,C_DOC_CNT,'
                     'C_REG,C_RAJ,PERIOD_MONTH,PERIOD_TYPE,PERIOD_YEAR,'
                     'C_STI_ORIG,C_DOC_STAN,D_FILL'.split(','))
# Assuming any other field will be appended to body

DEFAULTS = OrderedDict([
    ('TIN', None),  # should always be first
    ('C_DOC', 'F01'),
    ('C_DOC_SUB', '033'),
    ('C_DOC_TYPE', 0),
    ('C_DOC_VER', 6),
    ('C_DOC_CNT', 1),
    ('HZ', 1),
    ('HNACTL', 0),
])


def create_xml(data, output_dir='./', encoding='windows-1251'):
    root = ET.Element('DECLAR', {'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
                                 'xsi:noNamespaceSchemaLocation': 'F0103306.xsd'})
    head = ET.SubElement(root, 'DECLARHEAD')
    body = ET.SubElement(root, 'DECLARBODY')

    data_ = DEFAULTS.copy()
    data_.update(data)
    data_['PERIOD_MONTH'] = int(data_['PERIOD_MONTH'])  # for proper filename formatting

    for key, value in data_.items():
        if callable(value):
            value = value(data_)

        if not value and key not in ['C_DOC_TYPE', 'HNACTL']:
            continue

        if isinstance(value, float):
            value = '{:.2f}'.format(value)

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

    ET.SubElement(head, 'LINKED_DOCS', {'xsi:nil': 'true'})
    ET.SubElement(head, 'SOFTWARE', {'xsi:nil': 'true'})

    filename = os.path.join(output_dir, FILENAME_TEMPLATE.format(**data_))
    ET.ElementTree(root).write(filename, encoding)
    print('Created {}'.format(filename))


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
        data = OrderedDict(zip(fields, map(parse_value, sheet.row_values(i))))
        try:
            create_xml(data, output_dir)
        except Exception as exc:
            if not supress_exc:
                raise
            print('SKIPPED {}: {}: {}'.format(i, data, exc))


if __name__ == '__main__':
    from sys import argv
    try:
        if len(argv) > 1:
            main(argv[1])
        else:
            filename = input('Enter filename: [default="Книга1.xlsx"]')
            if not filename:
                main()
            else:
                main(filename)
    except Exception as e:
        print('Error', repr(e))
    input('DONE. press any key to close')