'''
Created on 2015-05-27

@author: legnain
'''
import csv
import os
import sys
from xlsxwriter.workbook import Workbook



if __name__ == '__main__':
    
    #fileName = "emails.csv"   
    fileName = raw_input (" Enter the path and file name (e.g., path/filename.csv): ")
    
    if not os.path.isfile(fileName):          # check if the file is exist
        print " the file is not "
        sys.exit()                            # terminate the program
        
    workbook  = Workbook(fileName + '.xlsx')
    worksheet = workbook.add_worksheet()    
    
    with open(fileName, 'rb') as f:
        content = csv.reader(f)
        for index_row, data_in_row in enumerate(content):
            for index_col, data_in_cell in enumerate(data_in_row):
                worksheet.write(index_row, index_col, data_in_cell)
    
    workbook.close()
    print " works"