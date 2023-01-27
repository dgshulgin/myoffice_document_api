#
# Copyright (c) MyOffice Hub of Knowledge, NPO, 2022
#
# You can not use the contents of the file in any way without MyOffice Hub of Knowledge, NPO written permission.
# To obtain such a permit, you should contact MyOffice Hub of Knowledge, NPO at contact@myofficehub.ru
#

#
# Практическая работа "Формирование текстового документа с помощью MyOffice SDK Python"
# Занятие №6. Форматирование документа
#

import sys
import os
import csv
from datetime import date
from MyOfficeSDKDocumentAPI import DocumentAPI as dapi

this = sys.modules[__name__]
this.app    = None
this.doc    = None
this.search = None

def _print_help():
    print("Usage: python3.8 praktika.py [template] [csv-file]")

def _search_replace(mark, text):
    ranges = this.search.findText(mark)
    if ranges is not None:
        for occ in ranges:
            occ.replaceText(text)

def _set_table_data(table, data):
    for (row, rec) in enumerate(data):
        table.getCell( dapi.CellPosition(row,0) ).setText(rec['num'])
        table.getCell( dapi.CellPosition(row,1) ).setText(rec['name_long'])
        table.getCell( dapi.CellPosition(row,2) ).setText(rec['alma_mater'])
        table.getCell( dapi.CellPosition(row,3) ).setText(rec['supervisor'])

def _format_table_header(table):
    header_line = table.getCellRange(dapi.CellRangePosition(0,0,0,3))
    for cell in header_line:
        # Цвет фона ячейки
        cell_props = cell.getCellProperties()
        cell_props.backgroundColor = dapi.ColorRGBA(212, 212, 212, 255)
        cell.setCellProperties(cell_props)
        # Выравнивание по центру
        para_props = cell.getParagraphProperties()
        para_props.alignment = dapi.Alignment_Center
        cell.setParagraphProperties(para_props)
        # Печать жирным (bold) шрифтом
        text_props = cell.getRange().getTextProperties()
        text_props.bold = True
        cell.getRange().setTextProperties(text_props)

def _set_template_data(data):
    # Поиск и замена маркеров согласно таблице.
    # Для простоты считаем, что все студенты из одного ВУЗ и проходят
    # практику в одно и то же время.
    #
    # Определить окончание "а" или "ов", в зависимости от кл-ва студентов.
    # Кол-во студентов: длина списка data за вычетом заголовка.
    var12 = ('а', 'ов')[(len(data)-1) > 1] # (False, True)
    var13 = ('указанного студента','указанных студентов')[(len(data)-1) > 1] # (False, True)
    # Построить список коротких и длинных ФИО
    names_short = [];  names_long = []
    for i in data:
        names_short.append(i['name_short'])
        names_long.append(i['name_long'])
    names_short.remove('ФИО') # убрать поле заголовка
    names_long.remove('ФИО полное') # убрать поле заголовка
    #
    marks_values =[]
    marks_values.append({'mark':'##date##',  'value':date.today().strftime('%Y')})
    marks_values.append({'mark':'##var01##', 'value':data[1]['alma_mater']})
    marks_values.append({'mark':'##var18##', 'value':data[1]['alma_mater_nom']})
    marks_values.append({'mark':'##var19##', 'value':data[1]['practice_type']})
    marks_values.append({'mark':'##var12##', 'value':var12})
    marks_values.append({'mark':'##var10##', 'value':', '.join(names_short)})
    marks_values.append({'mark':'##var17##', 'value':', '.join(names_long)})
    marks_values.append({'mark':'##var03##', 'value':data[1]['start_date']})
    marks_values.append({'mark':'##var04##', 'value':data[1]['end_date']})
    marks_values.append({'mark':'##var05##', 'value':data[1]['supervisor']})
    marks_values.append({'mark':'##var13##', 'value':var13})
    for (row, rec) in enumerate(marks_values):
        _search_replace(rec['mark'], rec['value'])

def _insert_text_with_props(text, props = None):
    this.doc.getRange().getEnd().insertText(text)
    if props is not None:
        count = -1
        for (r,p) in enumerate(this.doc.getBlocks().getParagraphsEnumerator()):
            #print(r, p.getRange().extractText())
            count += 1
        #print(count)
        para = this.doc.getBlocks().getParagraph(count)
        para.setParagraphProperties(props) 

def process(template, data) -> int:
    if this.app is None:
        this.app = dapi.Application()
    
    # Создать новый текстовый документ.
    if this.doc is None:
        this.doc = this.app.loadDocument(template)
    
    if this.search is None:
      this.search = dapi.createSearch(this.doc)

    # Вставка разрыва страниц 
    this.doc.getRange().getEnd().insertPageBreak()

    para_props = dapi.ParagraphProperties()
    para_props.alignment = dapi.Alignment_Right
    _insert_text_with_props("Приложение №1", para_props)
    _insert_text_with_props("к приказу Руководителя", para_props)
    _insert_text_with_props("Городского Совета", para_props)
    _insert_text_with_props("от ________", para_props)

    _insert_text_with_props("\n")
    para_props.alignment = dapi.Alignment_Center
    _insert_text_with_props("СПИСОК", para_props)
    _insert_text_with_props("\n")
    t = ['студентов ##var01##, направляемых в',
        'Городской Совет, для прохождения ##var19##',
        'с ##var03## по ##var04## года']
    _insert_text_with_props(' '.join(t), para_props)
    _insert_text_with_props("\n")

    # Вставка таблицы
    t_students = this.doc.getRange().getEnd().insertTable(4,4,"Практиканты")
    _set_table_data(t_students, data)
    # Форматирование заголовков столбцов
    _format_table_header(t_students)

    # Поиск и замена маркеров согласно таблице.
    _set_template_data(data)

    #
    #
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
            data = csv.DictReader(csv_file, delimiter=';')
            # Добавление заголовков столбцов
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
