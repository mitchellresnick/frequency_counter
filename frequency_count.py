import csv
import pathlib
import sys
from os import path, walk, execv
from pandas import read_excel, DataFrame, ExcelWriter
import math

def get_user_input(prompt, storage_variable, list: list, display_help: bool):
    i = 0
    print()
    for item in list:
        print(i, ": %s" % item)
        i += 1
    print()
    storage_variable = input(prompt)
    if storage_variable.lower() == 'help' and display_help:
        print("If you do not see the file that you want to analyze, please ensure that it is in the same directory as %s\n\n---------" % path.basename(__file__))
    else:
        try:
            return int(storage_variable)
        except ValueError:
            if display_help:
                print("\n!!!!! \nYou did not enter an allowed value. Please enter a numeral or 'help'.")
            else:
                print("You did not enter an allowed value. Please enter a numeral.")
            return None

def list_columns(workbook: str, worksheet: int):
    df = read_excel(workbook, sheet_name=None)

    for value in df[worksheet]:
        print("Columns:", *value.columns, "\n", sep=", ")

def list_worksheets(workbook: str):
    df = read_excel(workbook, sheet_name=None)

    return df.keys()

def main():
    # Gets the directory of this python file
    dir = pathlib.Path(__file__).parent.absolute()
    print(dir)

    # Gets all .xlsx files in the directory
    f = []
    for (firpath, dirnames, filenames) in walk(dir):
        for name in filenames:
            if name.endswith(".xlsx") and not name.endswith("_result.xlsx"):
                f.append(name)

    # Get user input
    file_index = None
    if len(f) == 1:
        print("Only one .xlsx file has been identified. Automatically starting analysis of %s" % f[0])
        file_index = 0
    else:
        user_prompt = "Please enter the number of the file from the list above that you would like to analyze (Enter 'help' for help): "
        while file_index is None:
            file_index = get_user_input(user_prompt, file_index, f, True)


    # Open file for reading
    workbook = None
    try:
        workbook = open(path.join(dir, f[file_index]), 'r')
    except PermissionError:
        print("\n\n!!!!!!!!!\nUnable to access '%s' file. Perhaps you already have it open? If so, close it before trying again.\n!!!!!!!!!\n\n" % (str(f[file_index])))
        sys.exit()
    # Print file contents
    list_worksheets(str(workbook.name))

    worksheet_column_map = {}
    df = read_excel(str(workbook.name), sheet_name=None)

    column_question = "Which column from the list above should be analyzed in "
    user_input = None
    for (key, value) in df.items():
        while user_input is None:
            user_input = get_user_input(column_question + key + "?: ", user_input, value, False)
        worksheet_column_map[key] = user_input
        user_input = None

    worksheet_frequency_holder = {}
    worksheet_frequency = {}
    wordFrequency = {}
    wordFrequencyDF = DataFrame()

    for (key, value) in worksheet_column_map.items():
        wordFrequency.clear()
        wordFrequencyDF = DataFrame()
        # print(key, "\n--------")
        column_name = df[key].columns[value]
        for row in df[key][column_name].values:
            try:
                for word in row.split(' '):
                    if word == '':
                        break
                    else:
                        try:
                            wordFrequency[word] += 1
                        except KeyError:
                            wordFrequency[word] = 1
            except AttributeError:
                # this happens when there is a row without keywords. To generate this error, add a row to the top of the file and only fill in the first column and run the program
                print("\n\n!!!!!!!!!!!!!!!\nEncountered a row without keywords. If this is not expected, please check the format of your file and correct any errors.\n!!!!!!!!!!!!!!!\n\n")
                pass
        dataframe_append = DataFrame.from_dict(wordFrequency, orient='index')
        wordFrequencyDF = wordFrequencyDF.append(dataframe_append)
        try:
            with ExcelWriter(path.join(dir, f[file_index][:-5] + "_result.xlsx"), engine='openpyxl', mode='a') as writer:
                wordFrequencyDF.to_excel(writer, sheet_name=key, header=False, startrow=0, startcol=0)
        except FileNotFoundError:
            print("A file with the name '%s' was not found. This program is unable to create the file. In order to fix this error, you must manually make a BLANK new Excel file with the name of '%s'. Once complete, please run the program again.\n" % ((str(f[file_index][:-5]) + "_result.xlsx"), (str(f[file_index][:-5]) + "_result.xlsx")))
        except PermissionError:
            print("\n\n!!!!!!!!!\nUnable to access '%s'. Perhaps you already have it open? If so, close it before trying again.\n!!!!!!!!!\n\n" % (str(f[file_index][:-5]) + "_result.xlsx"))
            sys.exit()
    workbook.close()
    print("Program successfully completed. Please check '%s' for results." % (str(f[file_index][:-5]) + "_result.xlsx"))


if __name__ == "__main__":
    main()