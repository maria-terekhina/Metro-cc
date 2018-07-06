import os
from collections import deque
from lxml import etree
import xlsxwriter
import schedule
import time

DB = {'new': deque(),
      'parsed': list(),
      'failed': list()}


def look_for_orders():
    for _, _, files in os.walk(r'C:\Users\maria.terekhina\Documents\Orders_new'):
        for f in files:
            DB['new'].append(f)
    return


def _extract_meta(tree):
    cells = [['GTIN', 'QUANTITY']]

    for item in tree.findall('.//lineItem'):
        line = list()
        for tag in item.getchildren():
            if tag.tag == 'gtin':
                line.append(tag.text)
            if tag.tag == 'requestedQuantity':
                line.append(tag.text)
        cells.append(line)

    creation_time = tree.find('.//creationDateTimeBySender').text
    client_id = tree.find('.//contractIdentificator').attrib['number']

    return cells, client_id, creation_time


def _write_xlsx(data, filename):
    if 'Orders_xlsx' not in os.listdir(r'C:\Users\maria.terekhina\Documents'):
        os.mkdir(r'C:\Users\maria.terekhina\Documents\Orders_xlsx')

    workbook = xlsxwriter.Workbook(r'C:\Users\maria.terekhina\Documents\Orders_xlsx\%s.xlsx' % (filename))
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0
    for line in data:
        worksheet.write(row, col, line[0])
        worksheet.write(row, col + 1, line[1])
        row += 1
    workbook.close()
    return


def _open_file(file):
    with open(r'C:\Users\maria.terekhina\Documents\Orders_new\%s' % (file), 'r', encoding='utf-8') as f:
        tree = etree.parse(f)
        f.close()
    return tree


def parse_order():
    while len(DB['new']) != 0:
        to_parse = DB['new'].pop()
        try:
            tree = _open_file(to_parse)
            data, client_id, creation_time = _extract_meta(tree)
            name = '%s_%s_%s' % (to_parse.split('.')[0], client_id, str(time.time()).split('.')[0])
            _write_xlsx(data, name)
            DB['parsed'].append(to_parse)

            if 'Orders_parsed' not in os.listdir(r'C:\Users\maria.terekhina\Documents'):
                os.mkdir(r'C:\Users\maria.terekhina\Documents\Orders_parsed')

            os.replace(r'C:\Users\maria.terekhina\Documents\Orders_new\%s' % (to_parse),
                       r'C:\Users\maria.terekhina\Documents\Orders_parsed\%s' % (to_parse))
        except:
            DB['failed'].append(to_parse)
    return


def main():
    schedule.every(1).minutes.do(look_for_orders)
    schedule.every(1).minutes.do(parse_order)

    while True:
        schedule.run_pending()
        time.sleep(1)