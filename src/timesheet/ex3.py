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

def _set_data(table, data):
    for (row, rec) in enumerate(data):
        print(rec)
        for idx in range(0,36):
            col_name = 'd{}'.format(idx+1)
            table.getCell( dapi.CellPosition(row,idx) ).setText(rec[col_name])

def process(template, data) -> int:
    if this.app is None:
        this.app = dapi.Application()
    
    # Создать новый табличный документ.
    if this.doc is None:
        this.doc = this.app.loadDocument(template)
    
    # Итоговый табель размещается на новом листе.
    sheet = _make_new_sheet(40,40,"Табель учета")

    # Занесение данных на лист
    _set_data(sheet, data)

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
            retcode = process( sys.argv[1], data )
            return retcode
    #
    _print_help()
    return 0

if __name__ == '__main__':
    sys.exit(main())