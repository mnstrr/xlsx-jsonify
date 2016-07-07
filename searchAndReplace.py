import openpyxl
import json
import ftfy


class SearchAndReplace:
    def __init__(self, xlsx_doc, json_doc, json_doc_new, sheet_name, blacklist):
        print('init SearchAndReplace')
        self.__xlsx_doc = xlsx_doc
        self.__json_doc = json_doc
        self.__json_doc_new = json_doc_new
        self.__sheet_name = sheet_name
        self.__blacklist = blacklist

        self.__data = ""
        self.__newdata = ""

        self.__replaceJson()
        # self.__main()

    def __main(self):
        f = open(self.__json_doc, 'r')
        self.__data = f.read()
        f.close()

        self.__newdata = self.__data

        self.__crawlexcel()

        print(self.__newdata)

        f = open(self.__json_doc_new, 'w')
        f.write(self.__newdata)
        f.close()

    def __crawlexcel(self):
        print('start search')

        wb = openpyxl.load_workbook(self.__xlsx_doc)
        sheet = wb.get_sheet_by_name(self.__sheet_name)

        for rowNum in range(2, sheet.max_row):  # skip the first row
            text = sheet.cell(row=rowNum, column=5).value
            translation = sheet.cell(row=rowNum, column=6).value
            if text is not None and translation is not None:
                # print(searchTerm)

                print(text, translation)
                self.___naivereplacement(text, translation)

    def __replaceJson(self):
        testStr = "Home Stories"
        testStr2 = ftfy.fix_encoding(testStr)
        replaceStr = "TROLOLO"
        json_data = open(self.__json_doc, encoding="utf8")
        jdata = json.load(json_data)

        self.__checkdict(jdata, testStr2, replaceStr)

    def __checkdict(self, collection, k, v):
        for key in collection.keys():
            self.__checkitem(collection[key], k, v)

    def __checklist(self, collection, k, v):
        for index in range(len(collection)):
            self.__checkitem(collection[index], k, v)

    def __checkitem(self, item, k, v):
        if type(item) is str:
            if ftfy.fix_encoding(item) == k:
              self.__replaceitem(item)
        if type(item) is dict:
            self.__checkdict(item, k, v)
        if type(item) is list:
            self.__checklist(item, k, v)

    def __replaceitem(self, item):
        print("FOUND YOU: ", ftfy.fix_encoding(item))

    def ___naivereplacement(self, text, translation):
        self.__newdata = self.__newdata.replace(text, translation)