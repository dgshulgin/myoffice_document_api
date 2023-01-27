#
# Copyright (c) MyOffice Hub of Knowledge, NPO, 2022
#
# You can not use the contents of the file in any way without MyOffice Hub of Knowledge, NPO written permission.
# To obtain such a permit, you should contact MyOffice Hub of Knowledge, NPO at contact@myofficehub.ru
#

#
# Практическая работа "Формирование текстового документа с помощью MyOffice SDK Python"
# Занятие №4. Заполнение таблицы
#

import sys
import os
import csv
from MyOfficeSDKDocumentAPI import DocumentAPI as dapi

this = sys.modules[__name__]
this.app = None
this.doc = None

def _print_help():
    print("Usage: python3.8 praktika.py [template] [csv-file]")

def _set_table_data(table, data):
    for (row, rec) in enumerate(data):
        table.getCell( dapi.CellPosition(row,0) ).setText(rec['num'])
        table.getCell( dapi.CellPosition(row,1) ).setText(rec['name_long'])
        table.getCell( dapi.CellPosition(row,2) ).setText(rec['alma_mater'])
        table.getCell( dapi.CellPosition(row,3) ).setText(rec['supervisor'])

def process(template, data) -> int:
    if this.app is None:
        this.app = dapi.Application()
    
    # Загрузить шаблон
    if this.doc is None:
        this.doc = this.app.loadDocument(template)

    # Вставка разрыва страниц 
    this.doc.getRange().getEnd().insertPageBreak()

    caption =  "Приложение №1\nк приказу Руководителя\nГородского Совета\nот\n"
    this.doc.getRange().getEnd().insertText(caption)

    hdr1 = "СПИСОК"
    hdr2 = "студентов ##var01##, направляемых в Городской Совет для прохождения ##var19## практики с ##var03## по ##var04## года\n"
    this.doc.getRange().getEnd().insertText(hdr1)
    this.doc.getRange().getEnd().insertText(hdr2)

    # Вставка таблицы
    t_students = this.doc.getRange().getEnd().insertTable(4,4,"Практиканты")
    _set_table_data(t_students, data)

    # Сохранить текстовый документ в формате OOXML, с расширением DOCX.
    # Тип документа определяется из расширения в имени файла.
    head, tail = os.path.split(template)
    filename = 'final-{}'.format(tail)
    this.doc.saveAs(filename)
    #
    return 0

def main() -> int:
    num_args = len(sys.argv)
    if num_args == 3:
        with open(sys.argv[2]) as csv_file:
            data = csv.DictReader(csv_file)
            # Добавление названий для столбцов
            header = {'num':'№ п.п', \
                'name_short':'ФИО', \
                'name_long':'ФИО полное', \
                'start_date':'Дата начала практики', \
                'end_date':'Дата окончания практики', \
                'practice_type':'Вид практики', \
                'alma_mater':'ВУЗ', \
                'alma_mater_nom':'ВУЗ', \
                'supervisor':'ФИО руководителя', \
                'authority':'Основание'}
            table_rows = []
            table_rows.append(header)
            # Добавление данных
            for r in data:
                table_rows.append(r)
            retcode = process( sys.argv[1], table_rows )
            return retcode
    #
    _print_help()
    return 0

if __name__ == '__main__':
   sys.exit(main())
