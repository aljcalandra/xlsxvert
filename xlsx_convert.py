import argparse
import csv
import os
import re
import xml.etree.ElementTree as ElementTree
from zipfile import ZipFile


SHARED_STRING_NAME = r'xl/sharedStrings.xml'
DELIMITER_DICT = {
    'tab': '\t'
}

def extract_zip(input_zip):
    input_zip = ZipFile(input_zip)
    return {name: input_zip.read(name) for name in input_zip.namelist()}


def get_sheet_names(workbook):
    sheet_re = r'.+worksheets/(.*).xml'
    return {re.search(sheet_re, item).group(1):item for item in workbook.keys() if 'worksheets' in item}


def get_sheet(workbook, sheet_name):
    return workbook[sheet_name]


def get_string_list(workbook):
    string_xml = workbook[SHARED_STRING_NAME]
    root = ElementTree.fromstring(string_xml)
    return [element.text for element in root.findall(".//{*}t")]


def get_rows(sheet_data, str_data):
    row_data = []
    root = ElementTree.fromstring(sheet_data)
    row_nodes = root.findall(".//{*}row")
    for row in row_nodes:
        cell_list = []
        for cell in row.findall(".//{*}c"):
            if cell.attrib.get('t') == 's':
                cell_list.append(str_data[int(cell[0].text)])
            else:
                cell_list.append(cell[0].text)
        row_data.append(cell_list)
    return row_data


def write_csv(file_name, data, delimiter=','):
    with open(f'{file_name}.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile, delimiter=delimiter)
        for row in data:
            writer.writerow(row)
    pass


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='convert xlsx to csv, outputs to files based on sheet names')
    parser.add_argument('-i', '--infile', dest='infile', required=True, help='Filename')
    parser.add_argument('-o', '--output_directory', dest='dest_path', required=False, help='Output Path', default='')
    parser.add_argument('-d', '--delimiter', dest='delimiter', required=False, default=',')
    args = parser.parse_args()

    delimiter = DELIMITER_DICT.get(args.delimiter, args.delimiter)

    wb = extract_zip(args.infile)
    string_list = get_string_list(wb)
    sheet_names = get_sheet_names(wb)
    for sheet, sheet_literal in sheet_names.items():
        data = get_sheet(wb, sheet_literal)
        row_data = get_rows(data, string_list)
        write_csv(os.path.join(args.dest_path, sheet), row_data, delimiter)
    pass
