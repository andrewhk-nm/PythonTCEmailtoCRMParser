'''
Take the emailed Excel xls TC sheet as input (command line args?)
Parse the needed info (Name (Last, First), Last Conversion Date)
Create a CSV file as output, with the columns needed to import into CRM
Maybe use API to auto import

TDD?
'''

import os
import glob
import xlrd
# import re #I used string methods instead

def transform_xl_to_list_of_dict():

    filename = 'C:\\Users\\perm7158\\Documents\\Projects\\Call RE Term Conversions\\Script\\06525_TC_1488954325929.xls'

    header_row_in_worksheet = 9 # Header row is row 9 in the emailed spreadsheet
    first_data_row_in_worksheet = 10 # First data row is right after that

    workbook = xlrd.open_workbook(filename)
    workbook = xlrd.open_workbook(filename, on_demand = True)
    worksheet = workbook.sheet_by_index(0)
    first_row = [] # The row where we stock the name of the column
    for col in range(worksheet.ncols):
        first_row.append( worksheet.cell_value(header_row_in_worksheet-1,col) )
        #print("Col {}, first_row: {}".format(col, first_row))
    #input()
    # tronsform the workbook to a list of dictionnary
    data =[]
    for row in range(first_data_row_in_worksheet-1, worksheet.nrows):
        elm = {}
        for col in range(worksheet.ncols):
            elm[first_row[col]]=worksheet.cell_value(row,col)
        data.append(elm)
    #print(data)
    return data

def importXLS(workbook):
    '''
    take a workbook filename
    open it
    output the rows of TC info.
    '''
    pass


def output_list_of_insureds(xl_list_of_dict):
    list_of_insureds = [] # Blank list to hold the names of the insureds
    for row in xl_list_of_dict:
        name = row.get("Insured")
        if name:
            list_of_insureds.append(name)
    print ('list_of_insureds: {}'.format(list_of_insureds))
    return list_of_insureds


if __name__ == "__main__":
    # Script run thru debugger
    xl_list_of_dict = transform_xl_to_list_of_dict()
        
    # output a list of the insureds names
    list_of_insureds = output_list_of_insureds(xl_list_of_dict)

    # Convert "First M  Last" to "Last, First"
    strinng = ''
    for name in list_of_insureds:
        double_space = len(name) - (name.find("  ")+2) 
        first_space = name.find(" ")
        first_name = name[0:first_space]
        print('first_name, len: {}, {}'.format(first_name, len(first_name)))
        last_name = name[-double_space:]
        print('last_name, len: {}, {}'.format(last_name, len(last_name)))
        lastcommafirst = last_name + ", " + first_name
        print(lastcommafirst)



