'''
Created on 2015-05-27

@author: legnain
'''
import csv
import os
import sys
import xlsxwriter



if __name__ == '__main__':
    
    fileName = "test"
    #fileName = raw_input (" Enter the file name (e.g., path/filename.csv): ")
    
    #if not os.path.isfile(fileName):          # check if the file is exist
    #    print " the file or directory is not found"
    #    sys.exit()                            # terminate the program
        
    excelFile  = xlsxwriter.Workbook(fileName + '.xlsx')
    worksheet = excelFile.add_worksheet()    
    
    with open(fileName + ".csv", 'rb') as f:
        content = csv.reader(f)
        for index_row, data_in_row in enumerate(content):
            for index_col, data_in_cell in enumerate(data_in_row):
                worksheet.write(index_row, index_col, data_in_cell)
    
    excelFile.close()
    print " === Conversion is done ==="