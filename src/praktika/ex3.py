#
# Copyright (c) MyOffice Hub of Knowledge, NPO, 2022
#
# You can not use the contents of the file in any way without MyOffice Hub of Knowledge, NPO written permission.
# To obtain such a permit, you should contact MyOffice Hub of Knowledge, NPO at contact@myofficehub.ru
#

#
# Практическая работа "Формирование текстового документа с помощью MyOffice SDK Python"
# Занятие №3. Вставка таблицы в шаблон документа
#

import sys
import os
from MyOfficeSDKDocumentAPI import DocumentAPI as dapi

this = sys.modules[__name__]
this.app = None
this.doc = None

def _print_help():
    print("Usage: python3.8 praktika.py [template]")

def process(template) -> int:
    if this.app is None:
        this.app = dapi.Application()
    
    # Загрузить шаблон документа
    if this.doc is None:
        this.doc = this.app.loadDocument(template)

    # Вставка разрыва страниц 
    this.doc.getRange().getEnd().insertPageBreak()

    # Вставка таблицы
    t_students = this.doc.getRange().getEnd().insertTable(4,4,"Практиканты")
    # Заголовок таблицы   
    t_students.getCell( dapi.CellPosition(0,0) ).setText("№ п.п")
    t_students.getCell( dapi.CellPosition(0,1) ).setText("ФИО студента")
    t_students.getCell( dapi.CellPosition(0,2) ).setText("Наименование ВУЗ")
    t_students.getCell( dapi.CellPosition(0,0) ).setText("ФИО руководителя")

    # Сохранить текстовый документ в формате OOXML, с расширением DOCX.
    # Тип документа определяется из расширения в имени файла.
    head, tail = os.path.split(template)
    filename = 'final-{}'.format(tail)
    this.doc.saveAs(filename)
    #
    return 0

def main() -> int:
    num_args = len(sys.argv)
    if num_args > 2:
        retcode = process( sys.argv[1] )
        return retcode
    #
    _print_help()
    return 0

if __name__ == '__main__':
   sys.exit(main())
