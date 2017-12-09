#! python2

import re
import csv
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO

def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = file(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text


def get_string_from_user(question):

    """asks the user a question and returns the response as a string

        args: question(str): Question to ask the user

        returns: string
    """

    return raw_input("Please input your file location to be extracted: ")

    return string_to_return
    

def get_start_words():

    """Gets the keywords for data sections to be extracted

        returns(list): list of all key words

    """

    number_of_columns = raw_input("How many sub-sections of data do you wish to extract: ")

    number_of_columns = int(number_of_columns)

    start_words_line_numbers = []    

    #gets the start words as well as how many lines they usually occur after section keyword
    for i in range(number_of_columns):
        start_word = raw_input("What is the start word or start phrase for data section: " + str(i + 1))
        line_number = raw_input("How many lines does start word occur after the section keyword: ")
        start_words_line_numbers.append([start_word, line_number])



    return start_words_line_numbers

def locate_data_section(heading, file_document):
    """Locates the sections of specific of the document
        where data will be then adds line number to list

    args: heading(str): the word or phrase used to location data

    returns list of ints

    """
    list_of_line_numbers = []
    lines=file_document.split("\n")

    
    for i, line in enumerate(lines):
        if heading in line:
            list_of_line_numbers.append(i)


    return list_of_line_numbers
        
    

def extract_data(section_start, line_extract, file_document):
    """returns a list of extracted datastring of the extracted data

    """

    list_of_extracted_data = []
    #sub_list = []
    lines=file_document.split("\n")
    
    
    for i in line_extract:
        for k, line in enumerate(lines):
            if (int(i[1])+section_start) == k:
                
                
                line_extract = line[len(i[0]):]
               
        list_of_extracted_data.append(line_extract)
        
                    

    return list_of_extracted_data


                

def format_extracted_data():

    total_rows_to_format = len(data_lists)
    formated_list = []
    sub_lists = []
    x = 0

    while x < total_rows_to_format:
        for i, line in enumerate(data_lists):
            formated_list.append(data_lists[i][x])
            x += 1

    return formated_list


def insert_data(data_lists, file_to_extract_to):
    
    with open(file_to_extract_to, 'wb') as f:
        w=csv.writer(f,delimiter=',')
        for i in data_lists:
            print('running extraction')
            w.writerow(i)
    print('extraction complete')
    
def main():
    
   
    #get location of PDF document
    location_pdf = raw_input('please input your file location to be extracted: ')

    #get location of CSV file to put data
    location_csv = raw_input('please input the location of the CSV file where data is to be extracted to: ')

    #get key word to locate data sections
    section_key_word = raw_input("please enter key word or key phrase to identifies a section of data: ")

    #get list of key words for sub sections
    start_words = get_start_words()

    #get information from PDF
    document=convert_pdf_to_txt(location_pdf)

    #locate lines where data each section begins    
    line_numbers=locate_data_section(section_key_word, document)

    #extra sub sections of data and return as list
    extracted_data = []
    for i in line_numbers:
        extracted_data_sub_list = extract_data(i, start_words, document)
        extracted_data.append(extracted_data_sub_list)

    #insert extracted data into excel file.
    insert_data(extracted_data, location_csv)
    raw_input()

if __name__ == "__main__":
    main()

#H:/Documents/Exceldatapython/docudatastriper/Register of Company Directors.pdf
#H:/Documents/Exceldatapython/docudatastriper/enterdata.csv
#insert_data(extract_data, '/Users/gordonowens/Documents/docudatastripper/testnumbers.csv')

