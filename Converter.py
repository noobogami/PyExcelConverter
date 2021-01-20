#https://www.journaldev.com/33306/pandas-read_excel-reading-excel-file-in-python
#https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.dtypes.html

#import excel2json
#excel2json.convert_from_file('Chapter1.xls')

import numpy
import pandas
import os
import json

def GetAllFilesName(directory):
    files = []
    if (not os.path.exists(directory)):
        print ("put your excels in a folder called 'Excels' beside this file and try again")
        return files
    for (dirpath, dirnames, filenames) in os.walk(directory):
        files.extend(filenames)
        break
    return files

def GetValidExcelFiles(files):
    excels = []
    for file in files:
        extension = file.split('.')[-1]
        print("Cheching", file, " with extention:", extension, end = "")
        if ('~' not in file and (extension == "xlsx" or extension == "xls")):
            excels.append(file)
            print (" --------------> ADDED")
        else:
            print()
    return excels

def GetValidSheets(file):
    data = []
    for sheet in file.sheet_names:
        if ('~' not in sheet):
            data.append(sheet)
    return data

def GetValidColumns(dataFrame):
    columns = dataFrame.columns
    validColumns = []
    for column in columns:
        print("Column ", column, end = " : ")
        print(dataFrame[column].dtype, end = "")
        if ('~' not in column):
            validColumns.append(column)
            print (" --------------> ADDED")
        else:
            print()
    return validColumns


def PrintSection(message, items = []):
    print()
    title = " ================== "
    message = title + message
    message += title
    print(message)
    if(items != []):
        print(items)
        print()

def PrintSeperator(mode, amount = 1):
    for i in range (amount):
        if(mode == 0):
            print("█████████████████████████████████████████████████████████")
        if(mode == 1):
            print("<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>")
        if(mode == 2):
            print(".........................................................")


def RemoveExtension(fileName):
    name = ""
    split = fileName.split('.')
    for i in range (len(split) - 2):
        name += split[i] + '.'
    name += split[len(split) - 2]
    return name

def GetJson(filePath, fileName):
    file = filePath + fileName
    if (not os.path.exists(file)):
        print ("json file", fileName, "could not be found in", filePath)
        return False, ""
    with open(file) as f:
        data = json.load(f)
    #print (data)
    return True, data

def CreateFile(directory, name, extension, value):
    CreateDirectory(".", directory)

    path = directory + '/' + name + '.' + extension
    file = open(path,"w+")
    file.write(value)
    file.close()
    print("File Created ", path)

def CreateDirectory(directory, name):
    path = directory + '/' + name
    if (os.path.exists(path)):
        print ("Directory '%s' Exists" % path)
        return
    else:
        try:
            os.makedirs(path)
        except OSError:
            print ("Creation of the directory '%s' failed" % path)
        else:
            print ("Successfully created the directory '%s' " % path)

def ExportExcelsWithoutModel (excelPath, excels):
    for excel in excels:
        ExportExcelWithoutModel(excelPath, excel)

def ExportExcelWithoutModel(excelPath, excel):
    excelFile = excelPath + excel
    file = pandas.ExcelFile(excelFile)
    sheets = GetValidSheets(file)

    PrintSeperator(0)
    PrintSection("Available Sheet for: " + excel, sheets)
    PrintSeperator(1)

    for sheet in sheets:
        PrintSection("Sheet " + sheet)
        df = pandas.read_excel(excelFile, sheet)
        columns = GetValidColumns(df)

        for column in df.columns:
            if (column not in columns):
                df.drop(column, inplace = True, axis=1)
            else:
                if (df[column].dtype == "float64" or df[column].dtype == "int64" ):
                    df[column].fillna(0, inplace = True)
                else:
                    df[column].fillna("", inplace = True)

        #df.fillna("", inplace = True)
        #df = df.astype(str)

        json = df.to_json(double_precision = 0, orient = "records", indent = 3)
        CreateFile("Jsons/" + RemoveExtension(excel), sheet, "json", json)

        PrintSeperator(2)

def ExportExcelsWithModel (excelPath, excels):
    completeSuccess = True
    for excel in excels:
        succeeded = ExportExcelWithModel (excelPath, excel)
        if (not succeeded):
            print("Converting Next File")
            completeSuccess = False
    return completeSuccess
        
def ExportExcelWithModel (excelPath, excel):
    excelFile = excelPath + excel
    fileExist, model = GetJson(excelPath, RemoveExtension(excel) + ".json")
    if (not fileExist):
        print ("json File Could'n Found Aborting Mission! RETREATING TROOPS")
        return False
    PrintSection("Model: ", model)
    file = pandas.ExcelFile(excelFile)
    sheets = GetValidSheets(file)

    PrintSeperator(0)
    PrintSection("Available Sheet for: " + excel, sheets)
    PrintSeperator(1)

    for sheet in sheets:
        PrintSection("Sheet " + sheet)
        df = pandas.read_excel(excelFile, sheet)
        columns = GetValidColumns(df)

        #print (df.dtypes)
        #df = df.astype(model)
        #print (df.dtypes)
        #return

        for column in df.columns:
            if (column not in columns):
                df.drop(column, inplace = True, axis=1)
            else:
                if (model[column] == "float64" or model[column] == "int64" ):
                    df[column].fillna(0, inplace = True)
                else:
                    df[column].fillna("", inplace = True)
                df[column] = df[column].astype(model[column])

        #df.fillna("", inplace = True)
        #df = df.astype(str)

        json = df.to_json(double_precision = 0, orient = "records", indent = 3)
        CreateFile("Jsons/" + RemoveExtension(excel), sheet, "json", json)

        PrintSeperator(2)
    return True

def YesOrNoQuestion(message):
    answer = input(message + " (y/n): ")
    while (answer != "y" and answer != "n"):
        print("Invalid Command Input 'y' or 'n'")
        answer = input(message + " (y/n): ")
    return answer == 'y'

def AskForEachFile(excelPath, excels):
    for excel in excels:
        if (YesOrNoQuestion("ConvertFile '" + excel + "'?")):
            if (YesOrNoQuestion("Is model file exist in folder?")):
                succeded = ExportExcelWithModel (excelPath, excel)
                if (not succeded):
                    PrintSeperator(0, 5)
                    print ("WTF LIAR! it's not here! converting without model")                    
                    PrintSeperator(0, 5)
                    ExportExcelWithoutModel (excelPath, excel)
            else:
                ExportExcelWithoutModel (excelPath, excel)

excelPath = "Excels/"

files = GetAllFilesName(excelPath)
if(files == []):
    #raw_input("Press any key to close ....")
    os.system('pause')
    os._exit(0)
PrintSection("Files in directory: ", files)

excels = GetValidExcelFiles(files)
PrintSection("Excels: ", excels)

if (YesOrNoQuestion("ConvertAllFiles?")):
    if (YesOrNoQuestion("Are model files exist in folder for all files?")):
        succeeded = ExportExcelsWithModel(excelPath, excels)
        if (not succeeded):
            if (YesOrNoQuestion("Do want convert all files without model? (CAUTION: converted file will be removed)")):
                ExportExcelsWithoutModel(excelPath, excels)
            elif (YesOrNoQuestion ("So do you want to specify each file seperatly?")):
                AskForEachFile(excelPath, excels)
            else:
                print ("FINE! Terminating Program ...")
    else:
        ExportExcelsWithoutModel(excelPath, excels)
else:
    AskForEachFile(excelPath, excels)
