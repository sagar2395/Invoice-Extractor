import pytesseract
from PIL import Image
import numpy as np
import cv2
import pytesseract
from pytesseract import Output
from pytesseract import image_to_string
import pandas as pd
import pdf2image
import time
import re

# File Locations
full_file_name = 'OnePage_G1.pdf'

file_name_list = full_file_name.split('.')
extension = file_name_list[1]
file_name = file_name_list[0]


file = r'Invoices/'+ full_file_name
tabledata_log_file = r'Error Logs/' + file_name + '_logs.txt' 
full_invoice_text = r'Text Invoice Outputs/' + file_name + '_fullinvoice.txt' 

totals = []

#Constants

DPI = 500
FORMAT = 'png'
THREAD_COUNT = 1

# This method convert pdf file to image format to gain information about whitespaces
def pdf_to_image():
    
    start_time = time.time()
    pil_images = pdf2image.convert_from_path(file, dpi=DPI, fmt=FORMAT, thread_count=THREAD_COUNT)
    print ("Time taken : " + str(time.time() - start_time))
    return pil_images

#This method helps in converting the images in PIL Image file format to the required image format
def save_images(pil_images):
    global file
    index = 1
    for image in pil_images:
        image.save(r'Invoice Image Outputs/'+ file_name + str(index) + ".png")
        index += 1
        
    file = r'Invoice Image Outputs/'+ file_name + "1" + ".png"

#Converts image to text format maintaining whitespaces
def image_to_text_whitespaces(path):
    img = cv2.imread(path , cv2.COLOR_BGR2GRAY)
    gauss = cv2.GaussianBlur(img, (3, 3), 0)
    text = ""
    custom_config = r'--oem 1 --psm 6 -c preserve_interword_spaces=1'
    d = pytesseract.image_to_data(gauss, config=custom_config, output_type=Output.DICT)
    df = pd.DataFrame(d)
    
    # clean up blanks
    df1 = df[(df.conf != '-1') & (df.text != ' ') & (df.text != '')]

    # sort blocks vertically
    sorted_blocks = df1.groupby('block_num').first().sort_values('top').index.tolist()
    for block in sorted_blocks:
        curr = df1[df1['block_num'] == block]
        sel = curr[curr.text.str.len() > 3]
        char_w = (sel.width / sel.text.str.len()).mean()
        prev_par, prev_line, prev_left = 0, 0, 0
        for ix, ln in curr.iterrows():
            # add new line when necessary
            if prev_par != ln['par_num']:
                text += '\n'
                prev_par = ln['par_num']
                prev_line = ln['line_num']
                prev_left = 0
            elif prev_line != ln['line_num']:
                text += '\n'
                prev_line = ln['line_num']
                prev_left = 0

            added = 0  # num of spaces that should be added
            if ln['left'] / char_w > prev_left + 1:
                added = int((ln['left']) / char_w) - prev_left
                text += ' ' * added
            text += ln['text'] + ' '
            prev_left += len(ln['text']) + added + 1
    text += '\n'
    return text

def save_invoice_to_text(full_invoice):
    invoice_text_file = open(full_invoice_text , 'w')
    invoice_text_file.write(full_invoice)
    invoice_text_file.close()

def extract_all_tables(full_invoice):
    i = 0
    remaining_data = full_invoice
    table_data = []
    while True:
        value , remaining_data = extract_table_data(remaining_data)
        table_data.append(value)
        if table_data[i] == '':
            break
        i += 1
    del table_data[-1]

    return table_data

# Fetches table data from full invoice
def extract_table_data(invoicetext):
    filetext = invoicetext.split('\n')
    tableStart = ['Description' , 'Content' ,'Product','QTY', 'Quantity' ,'Work', 'Unit Price' ,'TOTAL','Date','Material Number', 'UPC Code' 'Extended Amount' , 'VAT' , 'Price'  ,'Rate' , 'Price','S No.']
    tableEnd = ['Total' , 'Grand Total' , 'Amount']
    datatobetaken = False
    tableData = ''
    headers = {}
    linenumber = 0
    table_value_row = False
    remaining_data = ''
    table_taken = False

    for line in filetext:
        linenumber +=1
        count = 0

        if table_taken is False:
            for value in tableEnd:
                if value.lower() in line.lower() and datatobetaken:
                    datatobetaken = False
                    table_taken = True
                    get_totals(line) 
            
            for value in tableStart:
                if value.lower() in line.lower():
                    count += 1  

        if count > 3 and datatobetaken is False:
            datatobetaken = True

        if datatobetaken:
            tableData += line + '\n'
        else:
            remaining_data += line + '\n'

    return tableData , remaining_data

def get_totals(value):
    total_items = dict()
    value_list = []
    totals_line = replace_all(value , ['   ', '|', ')' , '(', '!' , '_'] , '%&' )
    totals_value_list = totals_line.split('%&')
    while '' in totals_value_list:
        totals_value_list.remove('')

    number_of_totals = len(totals_value_list)

    if number_of_totals == 2:
        total_items[totals_value_list[0]] = totals_value_list[1] 

    elif number_of_totals == 3:
        total_items[totals_value_list[0]] = totals_value_list[2]
        total_items['Quantity'] = totals_value_list[1]

    value_list.append(total_items)
    totals.append(value_list)

def get_table_dataframes(table_data):
    table_headers = []
    table_values = []
    dataset = []

    for i in range(len(table_data)):
        table_headers.append(get_headers(table_data[i]))
        table_values.append(get_table_values(len(table_headers[i]) , table_data[i]))
        dataset.append(get_pandas_dataframe(table_headers[i] , table_values[i]))

    return dataset


# Fetches headers of table 
def get_headers(table_data):
    headers = table_data.split('\n')
    char_in_line = len(headers[0])

    for header in headers:
        print(str(len(header)) + " " + header)

    header_list = replace_all(headers[0] , ['|','(','   ',')','!'] , ',').split(',')
    size = 0
    while ',' in header_list:
        header_list.remove(',')

    while '' in header_list:
        header_list.remove('')

    while ' ' in header_list:
        header_list.remove(' ')

    header_list_final = []

    for element in header_list:
        r = re.search("^[ —_]*$" , element)

        if not r:
            header_list_final.append(element)

    
    print(header_list_final)
    final_headers = []

    for value in header_list_final:
        templist = value.split()
        if len(templist) <= 2:
            final_headers.append(value.strip())
        else:
            if len(templist) == 3 :
                final_headers.append(templist[0].strip() + " " + templist[1].strip())
                final_headers.append(templist[2].strip())
            elif len(templist) == 4:
                final_headers.append(templist[0].strip() + " " + templist[1].strip())
                final_headers.append(templist[2].strip() + " " + templist[3].strip())
    print(final_headers)
    print(headers[0])

    return final_headers


#Fetches values of table
#Saves irregular data to log file
def get_table_values(number_of_columns , table_data):
    print(table_data)
    index = 0
    log_file = open(tabledata_log_file , 'w')
    current_data = []    
    table_values = table_data.split('\n')
    final_table_values = []
    error_id = 0
    improper_data = False

    for line in table_values[1:]:

        table_value_line = replace_all(line , ['   ', '|', ')' , '(', '!' , '_'] , '%&' )
        table_value_list = table_value_line.split('%&')
        while '' in table_value_list:
            table_value_list.remove('')

        values_in_row = len(table_value_list)

        print(str(table_value_list) + str(len(table_value_list)))

        if values_in_row == number_of_columns:
            current_data = table_value_list
            table_value_list.append('NaN')
            final_table_values.append(table_value_list)
            index += 1
            improper_data = False

        if values_in_row > number_of_columns:
            error_id += 1
            temp_list = []
            for x in range(number_of_columns):
                temp_list.append("NaN")

            temp_list.append("There is some anamoly in data . Please check log with id :" + str(error_id))
            final_table_values.append(temp_list)
            index += 1
            log_file.write("\n--------------------------------\n")
            log_file.write("Infomation related to id " + str(error_id) + ":\n")
            log_file.write("   ".join(table_value_list))
            current_data = temp_list
            improper_data = True

        if values_in_row < number_of_columns:
            if improper_data is True:
                log_file.write('\n')
                log_file.write("   ".join(table_value_list))
            else:
                if index > 0:
                    if current_data[-1] == 'NaN':
                        final_table_values[index-1][-1] = " ".join(table_value_list)
                    else:
                        final_table_values[index-1][-1] = final_table_values[index-1][-1] + " ".join(table_value_list)
                else:
                    log_file.write('------Table Start Information------\n')
                    log_file.write("   ".join(table_value_list)) 
        print(final_table_values)

    log_file.close()

    for values in final_table_values:
        print(len(values))   

    if final_table_values[-1][-1] == '':
            final_table_values[-1][-1] = 'NaN'

    print(final_table_values)
    return final_table_values

# This string method replaces all old elements in a string to new element
def replace_all(txt, old, new):
    for word in old:
        txt = txt.replace(word, new)
    return txt
        
# Using table headers and table values , this method gets pandas dataframe 
# Also a new column named additional information is added
# Saves Dataframe to excel file
def get_pandas_dataframe(table_headers , table_values):
    table_headers.append('Additional Information')
    df = pd.DataFrame(table_values , columns = table_headers)
    return df

def save_tables_to_excel(dataset):
    k = 1
    for df in dataset:
        df.to_excel(r'Excel Dataset Outputs/' + file_name + '_dataset_' + str(k) + ".xlsx")  
        k += 1  



if extension == 'pdf':
    pil_images = pdf_to_image()
    save_images(pil_images)

full_invoice = image_to_text_whitespaces(file)
save_invoice_to_text(full_invoice)
table_data = extract_all_tables(full_invoice)
dataset = get_table_dataframes(table_data)
save_tables_to_excel(dataset)
print(totals)






