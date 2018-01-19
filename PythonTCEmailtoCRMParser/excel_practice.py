""" Practice with the xlrd module
"""

import xlrd
# Noll recommended I check out this library
import csv
import openpyxl

def transform_xl_to_list_of_dict(filename, header_row=9, data_row=10):
    '''
    takes the filename of the excel document to parse.
    returns a list of dictionaries representing the worksheet.
    Starts at the header row (defaults to row 9), and the data row is the next one (row 10)
    '''

    header_row_in_worksheet = header_row # Header row is row 9 in the emailed spreadsheet
    first_data_row_in_worksheet = data_row # First data row is right after that

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

    # Have xlrd close the workbook
    workbook.release_resources()
            
    # Move the file to the "/Processed" subfolder
    move_to_processed_folder(filename)
    
    return data


if __name__ == '__main__':
    """ Practicing with Excel and python.
    """

    xlrd.