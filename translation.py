import xlrd
import json
import os
os.environ["PYTHONIOENCODING"] = "utf-8"
# excel file path
EXCEL_FILE_PATH = ''

# excel sheet name
SHEET_NAME = ''

# try:

try:
    workbook = xlrd.open_workbook(
        filename=EXCEL_FILE_PATH)
    sheet = workbook.sheet_by_name(SHEET_NAME)

    language_count = sheet.ncols-1
    key_count = sheet.nrows-1

    for lang in range(language_count):
        data = {}
        for key_index in range(key_count):
            # get key from row of index+1 and column 0
            key = sheet.cell(key_index+1, 0).value
            # if key empty, go to next key
            if not key:
                continue
            # set value from language index+1
            value = sheet.cell(key_index+1, lang+1).value
            # set key and value to dict
            data[str(key)] = value
        # get language
        lang_name = sheet.cell(0, lang+1).value
        # open and write dict to json
        with open(lang_name + '.json', 'w', encoding='utf-8') as outFile:
            json.dump(data, outFile, sort_keys=True, ensure_ascii=False)
except:
    print("Something went wrong. Please check your excel file")
