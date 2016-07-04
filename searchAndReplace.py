import openpyxl
import json
import pp


class SearchAndReplace:
    def __init__(self, xlsx_doc, json_doc, sheet_name, blacklist):
        print('init SearchAndReplace')
        self.__xlsx_doc = xlsx_doc
        self.__json_doc = json_doc
        self.__sheet_name = sheet_name
        self.__blacklist = blacklist

        self.__replaceJson()
        # self.__crawlExcel()

    def __crawlExcel(self):
        print('start search')

        wb = openpyxl.load_workbook(self.__xlsx_doc)
        sheet = wb.get_sheet_by_name(self.__sheet_name)

        for rowNum in range(2, sheet.max_row):  # skip the first row
            searchTerm = sheet.cell(row=rowNum, column=5).value
            transTerm = sheet.cell(row=rowNum, column=6).value
            if searchTerm is not None and transTerm is not None:
                # print(searchTerm)

                print(transTerm)

    def __replaceJson(self):
        testStr = "Continue with this browser"
        replaceStr = "TROLOLO"
        json_data = open('data.json')
        jdata = json.load(json_data)
        # self.__fixup(jdata,testStr,replaceStr)
        # print(type(jdata))

        # print(jdata.keys())
        #
        # for key in jdata.keys():
        #     print(type(jdata[key]))
        # # pp(jdata)

        self.__fixup2(jdata, testStr, replaceStr)

        print("REPLACED")

    def __fixup2(self, dict, k, v):
        for key in dict.keys():
            if dict[key] == v:
                print("FOUND YOU", key)
                break
            elif (type(dict[key]) is dict):
                self.__fixup2(dict[key], k, v)
            elif type(dict[key] is list):
                self.__fixupList(dict[key],k,v)


    def __fixupList(self,list,k,v):
        for item in list:
            print(item)

    def __fixup(self, adict, k, v):
        for key in adict.keys():
            print(key)
            if adict[key] == v:
                adict[key] = v
                break
            elif type(adict[key]) is dict:
                self.__fixup(adict[key], k, v)
