#!/usr/bin/python3
from docx import Document
import pandas
import xlrd

def doc():
    document = Document('test_doc.docx')

    line = 'Объем дисциплины '
    new_line = 'Объем дисциплины '

    for paragraph in document.paragraphs:
        #if line in paragraph.text:
        print(paragraph.text)
            #paragraph.text = new_line
            #print(paragraph.text)

    #document.save('test.docx')

#openpyxl
def xls():
    file = pandas.ExcelFile('IST4.xls')

    sheet = file.sheet_names

    df = pandas.read_excel(file, 'План')
    volume = df.iat[20,14]

    return volume
    
def discipline_volume(discipline):

    education_level = "Бакалавриат"
    course_discipline = "Информатика"
    course_code = "09.03.02"
    course = "Информационные системы и технологии"
    education_start_year = "2018"
    graduate_department = "ГИС"
    developer_department = "ГИС"
    programm_variant = "5"

   
    xls_file = pandas.ExcelFile('IST4.xls')
    df = pandas.read_excel(xls_file, 'План')
    doc_file = Document('test_doc.docx')

###   Рабочая программа дисциплини
    search_line = 'РАБОЧАЯ ПРОГРАММА ДИСЦИПЛИНЫ'
    paragraph_index = 0
    for paragraph in doc_file.paragraphs:
        if paragraph_index == 1:
            paragraph.text += '   ' + "РПД1"
            break
        if search_line in paragraph.text:
            paragraph_index = 1

###   Рабочая программа дисциплини 2
    search_line = 'РАБОЧАЯ ПРОГРАММА ДИСЦИПЛИНЫ'
    paragraph_index = 0
    for paragraph in doc_file.paragraphs:
        if paragraph_index == 3:
            paragraph.text += '   ' + "РПД2"
            break
        if paragraph_index == 2:
            paragraph_index += 1
        if search_line in paragraph.text:
            paragraph_index += 1

###   База

    search_line = 'Направление подготовки'
    for paragraph in doc_file.paragraphs:
        if search_line in paragraph.text:
            paragraph.text += '   ' + course_code + ' ' + course
            break

    search_line = 'Направленность: '
    for paragraph in doc_file.paragraphs:
        if search_line in paragraph.text:
            paragraph.text += '   ' + "НАПРАВЛЕННОСТЬ"
            break

    search_line = 'Год начала подготовки '
    for paragraph in doc_file.paragraphs:
        if search_line in paragraph.text:
            paragraph.text += '   ' + education_start_year
            break

    search_line = 'Выпускающая кафедра'
    for paragraph in doc_file.paragraphs:
        if search_line in paragraph.text:
            paragraph.text += '   ' + graduate_department
            break

    search_line = 'Кафедра-разработчик'
    for paragraph in doc_file.paragraphs:
        if search_line in paragraph.text:
            paragraph.text += '   ' + developer_department
            break

###   Промежуточная аттестация
    for i in range(df.shape[0]):
        if df.iat[i,3] == discipline:
            exam = df.iat[i,6]
            score_test = df.iat[i,7]
            test = df.iat[i,8]

    if type(exam) is str:
        attestation = 'экзамен'
    elif type(score_test) is str:
        attestation = 'зачет с оценкой'
    elif type(test) is str:
        attestation = 'зачет'
    else:
        print("ERROR")
        return 1

    search_line = 'Промежуточная аттестация'
    for paragraph in doc_file.paragraphs:
        if search_line in paragraph.text:
            paragraph.text += '   ' + attestation
            break



###   Объем дисциплины
    for i in range(df.shape[0]):
        if df.iat[i,3] == discipline:
            discipline_volume_amount = df.iat[i,14]

    search_line = 'Объем дисциплины '
    #new_line = f"jksjkdjf {volume}"
    for paragraph in doc_file.paragraphs:
        if search_line in paragraph.text:
            paragraph.text += discipline_volume_amount



    doc_file.save('test_output.docx')


discipline_volume("Технологии программирования")

