#!/usr/bin/python

'''
#Created on 2015-05-27
@author: Rajba legnain
To run the scrip, do the following:
1 - copy the script to the  directory that contains the csv files
2 - go to the directory
3- run  the script by writing in linux:  ./main.py  
'''


#import os
import csv
import glob
#import sys
import xlsxwriter



if __name__ == '__main__':
    
    #directory = os.getcwd()
    #fileName = raw_input ("Enter the directory that contains the csv files: \n")        
    #if not os.path.isfile(fileName):              # check if the file is exist
    #    print " the directory you entered is not found."
    #    sys.exit()                                # terminate the program
    
    #listOfFiles = os.listdir(directory)                   #  list of all files in the directory
    listOfFiles = glob.glob("*.csv")                       # list of files with extension of csv
    for index, fileInList in enumerate(listOfFiles):     
        fileName  = fileInList[0:fileInList.find('.')]     # remove the extension (i.e., csv) from the file name
        excelFile = xlsxwriter.Workbook(fileName + '.xlsx')
        worksheet = excelFile.add_worksheet()    
        #with open(fileName + ".csv", 'rb') as f:
        with open(fileInList, 'rb') as f:   
            content = csv.reader(f)
            for index_row, data_in_row in enumerate(content):
                for index_col, data_in_cell in enumerate(data_in_row):
                    worksheet.write(index_row, index_col, data_in_cell)
    
    excelFile.close()
    print " === Conversion is done ==="
