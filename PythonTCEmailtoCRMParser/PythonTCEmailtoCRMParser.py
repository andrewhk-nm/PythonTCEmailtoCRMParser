'''
Take the emailed Excel xls TC sheet as input (command line args?)
Parse the needed info (Name (Last, First), Last Conversion Date)
Create a CSV file as output, with the columns needed to import into CRM
Maybe use API to auto import
'''

# TODO
# TDD?
# default to the directory the script is run in?
#   definitely need to take out the hard coded paths at the very least
# Process APB files
# Add "Last Conversion Date" to comments of Phone Call
#
# DONE
# 2017-12-06 Move processed files to completed directory
# Created the release branch
# Created the develop branch


import os
import glob
import xlrd
# import re #I used string methods instead
import datetime
# Noll recommended I check out this library
import csv

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

def move_to_processed_folder(path_filename):
    '''
    take a filename and path 'path_filename' and move it to 'processed/filename'. 
    create the 'processed' folder if not yet created.
    return nothing
    '''
    path, filename = os.path.split(path_filename)
    
    processed_path = os.path.join(path + '/processed')
    try:
        os.mkdir(processed_path)
        print('created directory: {}'.format(processed_path))
    except FileExistsError:
        # Folder already exists, so it doesn't need to be created. Do nothing.
        print('directory exists: {}'.format(processed_path))
    # Move the file to the 'processed' directory
    os.rename(path_filename,os.path.join(processed_path, filename))

    


def output_list_of_insureds(xl_list_of_dict):
    '''
    Takes a list of dictionaries representing an excel spreadsheet
    returns a list of the insureds names, in "First M  Last" format
    '''
    list_of_insureds = [] # Blank list to hold the names of the insureds
    for row in xl_list_of_dict:
        name = row.get("Insured")
        if name:
            list_of_insureds.append(name)
    print ('list_of_insureds: {}'.format(list_of_insureds))
    return list_of_insureds

def convert_fml_to_lcf(list_of_insureds):
    '''
    Takes a list of names in the format "First M  Last" and converts to "Last, First"
    Returns that list
    '''
    list_of_insureds_lcf = []
    for name in list_of_insureds:
        double_space = len(name) - (name.find("  ") + 2) 
        first_space = name.find(" ")
        first_name = name[0:first_space]
        #print('first_name, len: {}, {}'.format(first_name, len(first_name)))
        last_name = name[-double_space:]
        #print('last_name, len: {}, {}'.format(last_name, len(last_name)))
        lastcommafirst = last_name + ", " + first_name
        print(lastcommafirst)
        list_of_insureds_lcf.append(lastcommafirst)

    return list_of_insureds_lcf

def output_CSV_file(set_of_insureds_lcf):
    '''
    Takes a set of insured's names in last, first format, and outputs a CSV file for import into CRM.
    Returns nothing
    '''
    #create a list of the correct CRM headers from an example file
    CRM_Header_filename = r"C:\Users\perm7158\Documents\_Josh\Projects\Call RE Term Conversions\CRM_Headers.csv"
    #(Do Not Modify) Phone Call,(Do Not Modify) Row Checksum,(Do Not Modify) Modified On,Due,Recipient,Assigned To,Subject,Regarding
    with open(CRM_Header_filename,encoding='utf-8') as crm_header_file:

        crm_header_string = crm_header_file.read()
        crm_list_of_headers = crm_header_string.split(sep=",")

        print(crm_header_string)
        print(crm_list_of_headers)
    

    # create a list of strings. The first list will be the row number (index 0 = header row 1), and that will contain the row that should be printed
    output_row_list = []
    output_row_list.append(crm_header_string)
    

    # Create data rows
    datetime_today = datetime.datetime.today()
    due = datetime_today.strftime("%m/%d/%Y") + " 8:00:00 AM"
    assigned_to = "Rang, Joshua"
    subject = "TC"
    on_behalf_of_team = "Rang, Joshua David 006525"
    for insured in set_of_insureds_lcf:
        data_row_string = '{DUE},"{RECIPIENT}","{ASSIGNED_TO}","{SUBJECT}","{REGARDING}","{ON_BEHALF_OF_TEAM}"\n'.format(DUE=due,
                                                                                            RECIPIENT=insured,
                                                                                            ASSIGNED_TO=assigned_to,
                                                                                            SUBJECT=subject,
                                                                                            REGARDING=insured,
                                                                                            ON_BEHALF_OF_TEAM=on_behalf_of_team)
        output_row_list.append(data_row_string)
    


    # Print the header to the Output file, and then the list of insureds along with call details
    CRM_output_filename = r"C:\Users\perm7158\Documents\_Josh\Projects\Call RE Term Conversions\Import_this_file_into_CRM.csv"
    #(Do Not Modify) Phone Call,(Do Not Modify) Row Checksum,(Do Not Modify) Modified On,Due,Recipient,Assigned To,Subject,Regarding
    #,,,11/30/2017 8:00:00 AM,"Aardvark, Aaron","Rang, Joshua",TC - Minimal Test,"Aardvark, Aaron"
    with open(CRM_output_filename,encoding='utf-8',mode='w') as b_file:
        # write the header row to a file
        #$b_file.write(crm_header_string)
        for row in output_row_list:
            b_file.write(row)

def yield_TC_filenames(default_directory = 'C:\\Users\\perm7158\\Documents\\_Josh\\Projects\\Call RE Term Conversions'):
    '''
    Generator Function
    Gets a list of all file names in the default directory that match the TC sheet pattern.
    Yields one file name at a time.
    '''
    # Get the list of all TC files
    dir_files = glob.glob(default_directory + '*_TC_*.xls')
    # Extend the list to include the APB files
    # Not right now, the sheet is formatted differently so I'll have to modify it to work (just different starting rows)
    #dir_files.extend glob.glob(default_directory + '*_APB_*.xls')
    print(dir_files)
    for filename in dir_files:
        yield filename



if __name__ == "__main__":

    ##test file
    #filename = 'C:\\Users\\perm7158\\Documents\\_Josh\\Projects\\Call RE Term Conversions\\Script\\06525_TC_1488954325929.xls'

    # Get list of all file names in the default directory
    filename_generator = yield_TC_filenames()
    
    # Make sure each list (or set) is empty to start
    xl_list_of_dict = []
    list_of_insureds = []
    list_of_insureds_lcf = []
    set_of_insureds_lcf = {}

    # Add the results of each file name parsing to the final list via extend.
    for filename in filename_generator:
        # Output a list of dictionaries based on the workbook
        xl_list_of_dict = transform_xl_to_list_of_dict(filename)

        # output a list of the insureds names
        list_of_insureds = output_list_of_insureds(xl_list_of_dict)

        # Convert "First M  Last" to "Last, First"
        #list_of_insureds_lcf.
        list_of_insureds_lcf.extend(convert_fml_to_lcf(list_of_insureds))
        


    set_of_insureds_lcf = set(list_of_insureds_lcf)
    # output a CSV file in the correct format for import into CRM
    output_CSV_file(set_of_insureds_lcf)

    


        

