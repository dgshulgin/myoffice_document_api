#
# Copyright (c) MyOffice Hub of Knowledge, NPO, 2022
#
# You can not use the contents of the file in any way without MyOffice Hub of Knowledge, NPO written permission.
# To obtain such a permit, you should contact MyOffice Hub of Knowledge, NPO at contact@myofficehub.ru
#

#
# Практическая работа "Формирование табличного документа с помощью MyOffice SDK Python"
# Занятие №????
#

import sys
import os
import csv
from MyOfficeSDKDocumentAPI import DocumentAPI as dapi

this = sys.modules[__name__]
this.app = None
this.doc = None

def _print_help():
    print("Usage: python3.8 tabel.py [template] [csv-file]")

def _make_new_sheet(rows, columns, title):
    return this.doc.getRange().getEnd().insertTable(rows,columns,title)

def _set_data(table, data, start=0):
    for (row, rec) in enumerate(data):
        for idx in range(0,36):
            col_name = 'd{}'.format(idx+1)
            table.getCell( dapi.CellPosition(start+row,idx) ).setText(rec[col_name])

# Получает в формате имя таблицы!диапазон
def _copy_range(src_addr, dst_addr):
    src_sheet = this.doc.getBlocks().getTable(src_addr.split('!')[0])
    src_range = src_sheet.getCellRange(src_addr.split('!')[1])

    dst_sheet = this.doc.getBlocks().getTable(dst_addr.split('!')[0])
    dst_range = dst_sheet.getCellRange(dst_addr.split('!')[1])

    cells = []
    for c in src_range.getEnumerator():
        cells.append( ( c.getFormattedValue(),
                        c.getFormat(),
                        c.getCellProperties(),
                        c.getRange().getTextProperties(),
                        c.getParagraphProperties(),
                        c.getBorders() ))

    idx = 0
    for c in dst_range.getEnumerator():
        props = cells[idx]
        c.setFormattedValue(            props[0] )
        c.setFormat(                    props[1] )
        c.setCellProperties(            props[2] )
        c.getRange().setTextProperties( props[3] )
        c.setParagraphProperties(       props[4] )
        c.setBorders(                   props[5] )
        idx += 1

def process(template, data) -> int:
    if this.app is None:
        this.app = dapi.Application()
    
    # Создать новый табличный документ.
    if this.doc is None:
        this.doc = this.app.loadDocument(template)
    
    # Итоговый табель размещается на новом листе.
    sheet = _make_new_sheet(40,40,"Tabel") #"Табель учета")

    # Копирование заголовочной части табеля
    # Разработчик на VBA  использовал бы для этой цели Selection,
    # но на уровне SDK (ядра) такая возможность не поддерживается.
    header = "A1:AK22"
    _copy_range(  'header!{}'.format(header), 'Tabel1!{}'.format(header)) #'Табель учета1!{}'.format(header))

    # Занесение данных на лист
    # Заголовочная часть табеля занимает 22 строки
    starting = 22
    _set_data(sheet, data, starting)

    # Копирование завершающей части табеля
    # Здесь сложнее, надо вычислять целевой диапазон
    src_footer = "A1:AL12"
    dst_footer = 'A{}:AL{}'.format( starting + len(data) +1,
                                        starting + len(data) + 1 + 11)
    _copy_range(  'footer!{}'.format(src_footer), 'Tabel1!{}'.format(dst_footer))

    # Сохранить табличный документ в формате OOXML, с расширением XLSX.
    # Тип документа определяется из расширения в имени файла.
    head, tail = os.path.split(template)
    filename = 'final-{}'.format(tail)
    this.doc.saveAs(filename)

    return 0


def main() -> int:
    num_args = len(sys.argv)
    if num_args == 3:
        with open(sys.argv[2]) as csv_file:
            data = csv.DictReader(csv_file, delimiter=';')
            retcode = process( sys.argv[1], list(data) )
            return retcode
    #
    _print_help()
    return 0

if __name__ == '__main__':
    sys.exit(main())