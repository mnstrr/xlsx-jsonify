import openpyxl
import json
import ftfy
import pp


class SearchAndReplace:
    def __init__(self, xlsx_doc, json_doc, json_doc_new, sheet_name, blacklist):
        print('init SearchAndReplace')
        self.__xlsx_doc = xlsx_doc
        self.__json_doc = json_doc
        self.__json_doc_new = json_doc_new
        self.__sheet_name = sheet_name
        self.__blacklist = blacklist

        self.__jadata = {}

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

        data2 = json.load(json_data)

        newcollection = self.__checkdict(data2, testStr2, replaceStr)
        pp(newcollection)

        with open('dataNEW.json', 'w') as fp:
            json.dump(newcollection, fp, sort_keys=True, indent=2)

    def __checkdict(self, collection, k, v):
        collectiontemp = collection
        for index in collectiontemp.keys():
            collectiontemp[index] = self.__checkitem(collectiontemp, k, v, index)
        return collectiontemp

    def __checklist(self, collection, k, v):
        collectiontemp = collection
        for index in range(len(collectiontemp)):
            collectiontemp[index] = self.__checkitem(collectiontemp, k, v, index)
        return collectiontemp

    def __checkitem(self, collection, k, v, index):
        collectiontemp = collection
        item = collectiontemp[index]
        if type(item) is str:
            if ftfy.fix_encoding(item) == k:
                collectiontemp2 = collectiontemp
                collectiontemp2[index] == v
                print(collectiontemp2[index])
                return ftfy.fix_encoding(v)
            else:
                return ftfy.fix_encoding(item)
        if type(item) is dict:
            collectiontemp = self.__checkdict(item, k, v)
        if type(item) is list:
            collectiontemp = self.__checklist(item, k, v)
        if type(item) is int or type(item) is float:
            return item

        return collectiontemp

    def __replaceitem(self, item):
        print("FOUND YOU: ", ftfy.fix_encoding(item))

    def ___naivereplacement(self, text, translation):
        self.__newdata = self.__newdata.replace(text, translation)
