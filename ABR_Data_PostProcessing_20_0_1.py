#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
About this code: This code is written to post-process the recorded ABR CSV data from the Bio-Sig software and convert it
to more readable excel format in terms of WAVE I and WAVE II
"""

# importing different python modules required for this code
import pandas as pd
import sys
import openpyxl
import os
import os.path
import tkinter
import tkinter.filedialog
import time
from colorama import init, Fore, Style

# Copyright Information
__author__ = "Ankush Bhayekar"
__copyright__ = "Copyright (C) 2020, Ankush Bhayekar"
__credits__ = ["Ankush Bhayekar"]
__license__ = "Public Domain"
__version__ = "2020.0.1"
__maintainer__ = "Ankush Bhayekar"
__email__ = ["bhayekarav@gmail.com"]

# Colored text codes
init(autoreset=True)

refer_excel_cwd = os.getcwd()

# opening the dialog box to select the CSV files
root = tkinter.Tk()
select_file = tkinter.filedialog.askopenfilenames(parent=root, title='Select CSV files', filetypes=(("csv files", "*.csv"), ("all files", "*.*")))
file_list = list(root.tk.splitlist(select_file))
only_csvs = []
iter = 1
for fl in file_list:
    split_filepath = fl.split('/')
    only_csvs.append(split_filepath[-1])
    if iter == 1:
        csv_fileloc = (split_filepath[:-1])
        csv_fileloc = '/'.join(csv_fileloc)

os.chdir(csv_fileloc)
start_time = time.time()

'''
counting the number of given arguments on the command line
The user should run this python code below format
python arf_file.py file1.CSV file2.CSV file3.CSV .........
the CSV file naming convention should be - ExpID-MouseNo-Frequency/Click-Age/M/F-RecordDate
'''
n_of_files = len(only_csvs)
print('\nTotal Number of Input CSV Files = {}\n'.format(n_of_files))

'''
For loop to create/append the data from each CSV file to the excel sheet
'''
for file in range(0, n_of_files):
    print('Writing CSV file - {}'.format(only_csvs[file]))
    csv_file = only_csvs[file][:-4]

    # for loop for finding out the sheet name in which the processed data will be stored
    try:
        if '-' in csv_file:
            split_filename = csv_file.split('-')
        elif '_' in csv_file:
            split_filename = csv_file.split('_')
    except NameError:
        print('The file naming convention is not correct\nThe CSV file naming convention should be - '
              'ExpID-MouseNo-Frequency/Click-Age/M/F-RecordDate')
        
    if split_filename[0] == 'pg2xnefl' :
        split_filename[0] = split_filename[0]+'-'+split_filename[1]
        del split_filename[1]
        
    mouse_num = split_filename[1]
    expID = split_filename[0]
    
    sheet_n = split_filename[2]
    if sheet_n == 'click':
        sheet_n = 'Click'
    else:
        sheet_n = sheet_n.lower()
        
    if expID[0:5] == 'nidtr':
        # MM/DD/YYYY
        dob = expID[9:11] + '/' + expID[11:] + '/' + expID[5:9]
    elif expID[0:6] == 'POU4F3' or expID[0:6] == 'pou4f3':
        # MM/DD/YYYY
        dob = expID[10:12] + '/' + expID[12:] + '/' + expID[6:10]
    else:
        # MM/DD/YYYY
        dob = expID[-4:-2] + '/' + expID[-2:] + '/' + expID[-8:-4]
    mouse_age = split_filename[3]
    record_date = split_filename[4]
    L_R = 'L'
    new_output_excel = 'abr first analyze ' + expID.lower() + '_' + mouse_age.lower() + '_' + record_date

    output_excel = pd.read_excel(refer_excel_cwd+'/'+'Reference_Excel_Format.xls', sheet_name=sheet_n)
    output_excel = pd.DataFrame(output_excel)
    output_excel = output_excel.iloc[:5, ]

    arf_data = pd.read_csv(csv_file + '.csv')
    arf_data = pd.DataFrame(arf_data)
    v12_data = pd.DataFrame((arf_data['V2(nv)'] - arf_data['V1(nv)']) / 1000)
    v34_data = pd.DataFrame((arf_data['V4(nv)'] - arf_data['V3(nv)']) / 1000)

    new_col_ind = arf_data.columns.get_loc('V2(nv)') + 1

    arf_data.insert(loc=new_col_ind, column='V12', value=v12_data)
    arf_data.insert(loc=new_col_ind + 1, column='T2', value=arf_data['T2(ms)'])

    new_col_ind = arf_data.columns.get_loc('V4(nv)') + 1

    arf_data.insert(loc=new_col_ind, column='V34', value=v34_data)
    arf_data.insert(loc=new_col_ind + 1, column='T4', value=arf_data['T4(ms)'])

    def wave_data(input1, input2, input3='Level(dB)'):
        """
        This function generates the WAVE I and WAVE II data from the V1, V2, V3 and V4 data
        :param input1: V12 or V34
        :param input2: T2 or T4
        :param input3: Level(dB) data frame from the CSV file
        :return: WAVE data and click threshold values in dB
        """
        new_row = []
        ind = 0

        # if else statements to check the file type if it is click or frequency - This will define the order of
        # Level(dB) column
        if sheet_n != 'Click':
            if not (arf_data[input1] < 0.2).all():
                for i in arf_data[input1][::-1]:
                    if i >= 0.2:
                        if ind == 0:
                            # iloc[::-1] below reverses the arf data since it is Puretone and dB levels are higher to
                            # lower in order
                            for inew, tnew in zip(arf_data[input1][:].iloc[::-1], arf_data[input2][:].iloc[::-1]):
                                new_row.append(round(inew, 7))
                                new_row.append(round(tnew, 7))
                        else:
                            for inew, tnew in zip(arf_data[input1][:-ind].iloc[::-1], arf_data[input2][:-ind].iloc[::-1]):
                                new_row.append(round(inew, 7))
                                new_row.append(round(tnew, 7))
                        dB_val = (arf_data[input3]).iloc[(-1 - ind)]
                        db2_index = ind
                        break
                    ind += 1
            elif (arf_data[input1] < 0.2).all():
                dB_val = 'No dB'
                new_row = []
                db2_index = 0
        elif sheet_n == 'Click':
            if not (arf_data[input1] < 0.2).all():
                for i in arf_data[input1]:
                    if i >= 0.2:
                        if ind == 0:
                            for inew, tnew in zip(arf_data[input1][:], arf_data[input2][:]):
                                new_row.append(round(inew, 7))
                                new_row.append(round(tnew, 7))
                        else:
                            for inew, tnew in zip(arf_data[input1][ind:], arf_data[input2][ind:]):
                                new_row.append(round(inew, 7))
                                new_row.append(round(tnew, 7))
                        dB_val = arf_data[input3][ind]
                        db2_index = ind
                        break
                    ind += 1
            elif (arf_data[input1] < 0.2).all():
                dB_val = 'No dB'
                new_row = []
                db2_index = 0

        return pd.DataFrame(new_row), dB_val


    wave1_data, dB_val1 = wave_data('V12', 'T2')
    wave2_data, dB_val2 = wave_data('V34', 'T4')

    if len(wave1_data) != len(wave2_data) and len(wave1_data) < len(wave2_data):
    	new_wave1 = []
    	wave1_data = pd.DataFrame()
    	if sheet_n != 'Click':
    		wave1_v12 = arf_data['V12'][::-1]
    		wave1_t2 = arf_data['T2'][::-1]
    		level_db = arf_data['Level(dB)'][::-1]
    		diff = int(len(wave1_v12) - (len(wave2_data) / 2))
    		dB_val1 = level_db.iloc[diff]
    		for inew, tnew in zip(wave1_v12.iloc[diff:], wave1_t2[diff:]):
    			new_wave1.append(round(inew, 7))
    			new_wave1.append(round(tnew, 7))
    	elif sheet_n == 'Click':
    		wave1_v12 = arf_data['V12']
    		wave1_t2 = arf_data['T2']
    		level_db = arf_data['Level(dB)']
    		diff = int(len(wave1_v12) - (len(wave2_data) / 2))
    		dB_val1 = level_db.iloc[diff]
    		for inew, tnew in zip(wave1_v12.iloc[diff:], wave1_t2[diff:]):
    			new_wave1.append(round(inew, 7))
    			new_wave1.append(round(tnew, 7))

    	wave1_data = pd.DataFrame(new_wave1)
    
    '''
    Below code block is used to identify the index numbers of the columns for the given dB level values
    The db_ind_arr is the basically the list of two index numbers - WAVE I and WAVE II
    '''
    id_of_db = pd.Index(output_excel.iloc[2])
    id_tem = 0
    db_ind_arr = []
    if dB_val1 == 'No dB' or dB_val2 == 'No dB':
        print(Fore.YELLOW + 'WARNING: No dB values in this ABR data OR all values for WAVE II are below 0.2')
        db_ind_arr.append(5)
        db_ind_arr.append(6)
        wave1_data = pd.DataFrame()
        wave2_data = pd.DataFrame()
    else:
        for i in id_of_db.fillna(0):
            if id_tem <= 43:
                if i == dB_val1:
                    db_ind_arr.append(id_tem)
            elif id_tem > 43:
                if i == dB_val2:
                    db_ind_arr.append(id_tem)
            id_tem += 1

    adr_data = pd.DataFrame(
        [expID + '-' + mouse_num, dob, L_R, record_date[:2] + '/' + record_date[2:4] + '/' + record_date[4:], dB_val2])
    adr_data = pd.DataFrame.transpose(adr_data)
    
    if mouse_num[0]=='a':
        mouse_num = mouse_num[1:]
    elif mouse_num[0]=='b':
        mouse_num = int(mouse_num[1:])+12
    elif mouse_num[0]=='c':
        mouse_num = int(mouse_num[1:])+24
    elif mouse_num[0]=='d':
        mouse_num = int(mouse_num[1:])+36
    else:
        mouse_num = mouse_num

    def write_to_excel():
        """
        Function to write the data to the excel sheet
        :return: Excel sheet with the Wave I and Wave II data added
        - output_excel is the data frame reference from the existing post-processed data excel file
        - adr_data is the data of the ExpID, mouse number, DoB etc.
        """
        output_excel.to_excel(writer, sheet_name=sheet_n, index_label=None, index=False, header=False)
        adr_data.to_excel(writer, sheet_name=sheet_n, startrow=4 + int(mouse_num), startcol=0, index=False,
                          index_label=None, header=False)
        (wave1_data.transpose()).to_excel(writer, sheet_name=sheet_n, startrow=4 + int(mouse_num),
                                          startcol=db_ind_arr[0], index=False, index_label=None, header=False)
        (wave2_data.transpose()).to_excel(writer, sheet_name=sheet_n, startrow=4 + int(mouse_num),
                                          startcol=db_ind_arr[1], index=False, index_label=None, header=False)


    if os.path.isfile(new_output_excel + '.xlsx'):
        '''
        This is the if statement if the excel file is already present in the current working directory
        os.path.isfile searches the given filename in the CWD
        '''
        book = openpyxl.load_workbook(new_output_excel + '.xlsx')
        writer = pd.ExcelWriter(new_output_excel + '.xlsx', engine='openpyxl')
        writer.book = book
        '''
        ExcelWriter for some reason uses writer.sheets to access the sheet.
        If you leave it empty it will not know that sheet Main is already there
        and will create a new sheet.
        '''
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        write_to_excel()
        # below code lines are not required anymore but keeping it just in case for the reference
        # with pd.ExcelWriter(new_output_excel+'.xlsx', engine='openpyxl', mode='a') as writer:
        #    write_to_excel()
    else:
        '''
        If the excel file is not already present in the current folder then this else statement will create it
        '''
        writer = pd.ExcelWriter(new_output_excel + '.xlsx', engine='openpyxl')
        write_to_excel()

    writer.save()
    writer.close()
    print(Style.BRIGHT + '......Done......')

print(Fore.BLUE + Style.BRIGHT + '\nSuccessfully Post-processed the abr data from {} CSV files......\n'.format(n_of_files))
print("--- Total Execution Time = %.2f seconds ---\n" % (time.time() - start_time))
