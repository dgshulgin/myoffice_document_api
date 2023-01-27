#
# Copyright (c) MyOffice Hub of Knowledge, NPO, 2022
#
# You can not use the contents of the file in any way without MyOffice Hub of Knowledge, NPO written permission.
# To obtain such a permit, you should contact MyOffice Hub of Knowledge, NPO at contact@myofficehub.ru
#

#
# Практическая работа "Формирование текстового документа с помощью MyOffice SDK Python"
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
    
    # Создать новый текстовый документ.
    if this.doc is None:
        this.doc = this.app.createDocument(dapi.DocumentType_Text)
    
    # Вставить фрагмент текста в начало документа.
    this.doc.getRange().getBegin().insertText("Фрагмент текста")

    # Сохранить текстовый документ в формате OOXML, с расширением DOCX.
    # Тип документа определяется из расширения в имени файла.
    this.doc.saveAs("basic.docx")
    return 0


def main() -> int:
    return process()

if __name__ == '__main__':
    sys.exit(main())