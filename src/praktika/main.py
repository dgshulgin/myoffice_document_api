import pandas
import csv

import sys

from MyOfficeSDKDocumentAPI import DocumentAPI as sdk



document = sys.argv[1]
csv_file = sys.argv[2]
delimiter = ';'

def main(document, csv_file):
    data_students = get_data_students(csv_file=csv_file, delimiter=delimiter)

    replace_all_matches(document, data_students)

def get_data_students(csv_file, delimiter):
    tags = get_csv_tags(csv_file)
    filled_data = get_filled_data(csv_file, tags, delimiter)
    return filled_data

def get_csv_tags(file):
    tags = get_csv_file_tags(file, delimiter)
    return tags


def get_csv_file_tags(file, delimeter):
    csv_file = open_csv_file(file, delimeter)
    file_reader = csv.reader(csv_file, delimiter=delimiter)
    tags = next(file_reader)
    csv_file.close()
    return tags

def get_filled_data(csv_file, tags, delimiter):
    tags_data = get_tags_data(csv_file, tags, delimiter)
    students_data = dict()
    students_data["##date##"] = [tags_data["end_date"][0][6:]]
    students_data["##var01##"] = [tags_data["alma_mater"][0]]
    students_data["##var18##"] = [tags_data["alma_mater_nom"][0]]
    students_data["##var19##"] = [tags_data["practice_type"][0]]
    if len(tags_data["name_short"]) == 1:
        students_data["##var12##"] = ["а"]
        students_data["##var13##"] = ["указанного студента"]
    else:
        students_data["##var13##"] = ["указанных студентов"]
        students_data["##var12##"] = ["ов"]
    students_data["##var10##"] = tags_data["name_short"]
    students_data["##var17##"] = tags_data["name_long"]
    students_data["##var03##"] = [str(tags_data["start_date"][0])]
    students_data["##var04##"] = [tags_data["end_date"][0]]
    students_data["##var05##"] = [tags_data["supervisor"][0]]
    return students_data

def get_tags_data(csv_file, tags, delimiter):
    csv_pandas_file = open_csv_file_pandas(csv_file, delimiter)
    filled_tags_info = dict()

    for tag in tags:
        filled_tags_info[tag] = list(csv_pandas_file[tag])
    return filled_tags_info

def open_csv_file_pandas(file, delimiter):
    try:
        file = pandas.read_csv(csv_file, delimiter=';')
    except FileNotFoundError:
        print(f"Файла {csv_file} не существует! ")
        return 1
    return file

def open_csv_file(file, delimeter):
    try:
        csv_file = open(file, "r")
    except FileNotFoundError:
        print(f"Файла {file} не существует! ")
        return 1
    return csv_file

def replace_all_matches(file, data_students):
    application = sdk.Application()
    try:
        document = application.loadDocument(file)
    except sdk.UnknownError:
        print("Документ не найден!")
        return 1
    for key in data_students:
        collect = get_values_found(document, key)
        for tag in collect:
            tag.replaceText(", ".join(data_students[key]))

    document.saveAs("filled_document.docx")

def open_document(file):
    application = sdk.Application()
    try:
        document = application.loadDocument(file)
    except sdk.UnknownError:
        print("Документ не найден!")
        return 1
    print(type(document))
    return document

def get_values_found(file, text):
    textSearch = sdk.createSearch(file)
    collect = textSearch.findText(text)
    return collect

def document_save(document):
    document.saveAs("filled_document.docx")

if __name__ == '__main__':
    main(document, csv_file)

