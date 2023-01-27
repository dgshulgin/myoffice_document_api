#
# Copyright (c) MyOffice Hub of Knowledge, NPO, 2022
#
# You can not use the contents of the file in any way without MyOffice Hub of Knowledge, NPO written permission.
# To obtain such a permit, you should contact MyOffice Hub of Knowledge, NPO at contact@myofficehub.ru
#

#
# Практическая работа "Формирование табличного документа с помощью MyOffice SDK Python"
# Занятие №2. Разработка базового приложения
#

import sys
from MyOfficeSDKDocumentAPI import DocumentAPI as dapi

this = sys.modules[__name__]
this.app = None
this.doc = None

def process() -> int:
    if this.app is None:
        this.app = dapi.Application()
    
    # Создать новый табличный документ.
    if this.doc is None:
        this.doc = this.app.createDocument(dapi.DocumentType_Workbook)
    
    # Сразу же после создания табличный документ не содержит ни одного листа.
    # Первый лист необходимо создать.
    pos = this.doc.getRange().getBegin()
    sheet1 = pos.insertTable(20,20,"Табель учета")
    sheet1.getCell("A1").setNumber(10)

    # Сохранить документ в формате OOXML, с расширением XLSX.
    # Тип документа определяется из расширения в имени файла.
    this.doc.saveAs("basic.xlsx")
    return 0


def main() -> int:
    return process()

if __name__ == '__main__':
    sys.exit(main())