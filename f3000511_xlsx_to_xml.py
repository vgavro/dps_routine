#!/usr/bin/env python

import xml.etree.cElementTree as ET
import os.path
from collections import OrderedDict
import traceback

import xlrd


def _extend_dict(x, y):
    rv = x.copy()
    rv.update(y)
    return rv


FILENAME_TEMPLATE = (
    '{C_STI_ORIG}{TIN}'
    '{C_DOC}{C_DOC_SUB}{C_DOC_VER}{C_DOC_STAN}'
    '000000000'
    '{PERIOD_TYPE}{PERIOD_MONTH:02d}{PERIOD_YEAR}'
    '{C_STI_ORIG}.xml'
)

DECLARHEAD_FIELDS = ('TIN,C_DOC,C_DOC_SUB,C_DOC_VER,C_DOC_TYPE,C_DOC_CNT,'
                     'C_REG,C_RAJ,PERIOD_MONTH,PERIOD_TYPE,PERIOD_YEAR,'
                     'C_STI_ORIG,C_DOC_STAN,D_FILL'.split(','))
# Assuming any other field will be appended to body

DEFAULTS = OrderedDict([
    ('TIN', None),  # should always be first
    ('C_DOC', 'F30'),
    ('C_DOC_SUB', None),
    ('C_DOC_TYPE', 0),
    ('C_DOC_VER', 11),
    ('C_DOC_CNT', 1),
    ('C_REG', lambda d: str(d['C_STI_ORIG'])[:2]),
    ('C_RAJ', lambda d: str(d['C_STI_ORIG'])[-2:]),
    ('PERIOD_MONTH', 12),
    ('PERIOD_TYPE', 5),
    ('PERIOD_YEAR', 2017),
    ('C_DOC_STAN', 1),
    ('HZY', 2017),
    ('HZB', 1),
    ('HTIN', lambda d: d['TIN']),
])

HEAD_DEFAULTS = _extend_dict(DEFAULTS, [
    ('C_DOC_SUB', '005'),
    ('HKSTI', lambda d: d['C_STI_ORIG']),
    ('HFILL', lambda d: d['D_FILL']),
    ('_linked_doc_type', 2),
])

_SUBREPORT_DEFAULTS = _extend_dict(DEFAULTS, [
    ('_linked_doc_type', 1),
    ('HLNAME', lambda d: d['HBOS'].split()[0]),
    ('HPNAME', lambda d: d['HBOS'].split()[1]),
    ('HFNAME', lambda d: d['HBOS'].split()[2]),
])

SUBREPORT_MONTH_VALUES = {
    '501': OrderedDict([
        ('R0{:02d}G2', 3200),
        ('R0{:02d}G3', 3200),
        ('R0{:02d}G4', 22),
        ('R0{:02d}G5', 704),
    ]),
    '502': OrderedDict([
        ('R0{:02d}G2', 3200),
        ('R0{:02d}G3', 22),
        ('R0{:02d}G4', 704),
    ]),
}

SUBREPORT_DEFAULTS = {
    '501': _extend_dict(_SUBREPORT_DEFAULTS, [
        ('C_DOC_SUB', '501'),
        ('R01G2', lambda d: sum(d.get('R0{:02d}G2'.format(x), 0) for x in range(1, 13))),
        ('R01G3', lambda d: sum(d.get('R0{:02d}G3'.format(x), 0) for x in range(1, 13))),
        ('R01G5', lambda d: sum(d.get('R0{:02d}G5'.format(x), 0) for x in range(1, 13))),
        ('R02G1', 22),
        ('R02G2', lambda d: d['R01G5']),

    ]),
    '502': _extend_dict(_SUBREPORT_DEFAULTS, [
        ('C_DOC_SUB', '502'),
        ('R01G2', lambda d: sum(d.get('R0{:02d}G2'.format(x), 0) for x in range(1, 13))),
        ('R01G4', lambda d: sum(d.get('R0{:02d}G4'.format(x), 0) for x in range(1, 13))),
        ('R02G1', 22),
        ('R02G2', lambda d: d['R01G4']),
    ]),
}


def create_element(key, value):
    """Allows key to have attributes instead of raw tag"""
    tag = key.split(' ')[0]
    tag_str = '<{}></{}>'.format(key, tag)
    try:
        e = ET.fromstring(tag_str)
    except ET.ParseError:
        raise RuntimeError('Invalid field {}: could not parse {}'
                           .format(key, tag_str))
    e.text = str(value)
    return e


def create_filename(data):
    data['PERIOD_MONTH'] = int(data['PERIOD_MONTH'])  # for proper filename formatting
    print(data)
    return FILENAME_TEMPLATE.format(**data)


def create_xml(data, linked_data=[], output_dir='./', encoding='windows-1251'):
    version = data['C_DOC'] + data['C_DOC_SUB'] + str(data['C_DOC_VER'])
    root = ET.Element('DECLAR', {'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
                                 'xsi:noNamespaceSchemaLocation': '{}.xsd'.format(version)})
    head = ET.SubElement(root, 'DECLARHEAD')
    body = ET.SubElement(root, 'DECLARBODY')

    for key, value in data.items():
        if not key:
            continue

        if callable(value):
            data[key] = value(data)
            value = data[key]

        if ((not value and key not in ['C_DOC_TYPE']) or
           key.startswith('_')):
            continue

        if isinstance(value, float):
            value = '{:.2f}'.format(value)

        e = create_element(key, value)

        if e.tag in DECLARHEAD_FIELDS:
            head.append(e)
        else:
            body.append(e)

    if not linked_data:
        ET.SubElement(head, 'LINKED_DOCS', {'xsi:nil': 'true'})
    else:
        linked_docs = ET.SubElement(head, 'LINKED_DOCS')
        for i, data_ in enumerate(linked_data):
            doc = ET.SubElement(linked_docs, 'DOC')
            doc.attrib['TYPE'] = str(data_['_linked_doc_type'])
            doc.attrib['NUM'] = str(i + 1)

            for k in ['C_DOC', 'C_DOC_SUB', 'C_DOC_VER', 'C_DOC_TYPE',
                      'C_DOC_CNT', 'C_DOC_STAN']:
                doc.append(create_element(k, data_[k]))

            doc.append(create_element('FILENAME', create_filename(data_)))

    ET.SubElement(head, 'SOFTWARE', {'xsi:nil': 'true'})

    filename = os.path.join(output_dir, create_filename(data))
    ET.ElementTree(root).write(filename, encoding)
    print('Created {}'.format(filename))


def main(xlsx_filename='f3000511.xlsx',
         fields_row_index=0, data_start_row_index=2, supress_exc=False):

    output_dir = os.path.basename(xlsx_filename) + '_xml'
    output_dir = os.path.join(os.path.dirname(xlsx_filename), output_dir)
    try:
        os.mkdir(output_dir)
    except OSError:
        pass

    def parse_value(value):
        if value == 'YES':
            return True
        elif value == 'NO':
            return False
        elif isinstance(value, float) and value == int(value):
            value = int(value)
        return value

    def map_sheet_by_tin(book, index):
        rv = {}
        sheet = book.sheet_by_index(index)
        fields = sheet.row_values(fields_row_index)

        for i in range(data_start_row_index, sheet.nrows):
            data = OrderedDict(zip(fields, map(parse_value, sheet.row_values(i))))
            rv[data['TIN']] = data
        return rv

    book = xlrd.open_workbook(xlsx_filename)

    linked_data_map = OrderedDict([
        ('501', map_sheet_by_tin(book, 1)),
        ('502', map_sheet_by_tin(book, 2)),
    ])

    sheet = book.sheet_by_index(0)
    fields = sheet.row_values(fields_row_index)

    for i in range(data_start_row_index, sheet.nrows):
        data = HEAD_DEFAULTS.copy()
        data.update(OrderedDict(zip(fields, map(parse_value, sheet.row_values(i)))))

        linked_data = []
        for c_doc_sub, map_ in linked_data_map.items():
            if data['TIN'] in map_:
                data_ = SUBREPORT_DEFAULTS[c_doc_sub].copy()
                data_.update({
                    'HBOS': data.get('HNAME'),
                })
                data_.update(map_[data['TIN']])
                for k in ['C_STI_ORIG']:
                    data_[k] = data[k]
                for month in range(1, 13):
                    flag = data_.pop('_' + str(month), None)
                    if flag:
                        for fld, value in SUBREPORT_MONTH_VALUES[c_doc_sub].items():
                            data_[fld.format(month)] = value
                linked_data.append(data_)

                # Add corresponding field to HEAD
                fld = {'501': 'R001G3', '502': 'R002G3'}.get(c_doc_sub)
                data[fld] = 1

    try:
        create_xml(data, linked_data, output_dir)
        for data_ in linked_data:
            create_xml(data_, [data], output_dir)
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
            filename = input('Enter filename: [default="f3000511.xlsx"]')
            if not filename:
                main()
            else:
                main(filename)
    except Exception as e:
        traceback.print_exc()
        print('Error', repr(e))
    input('DONE. press any key to close')
