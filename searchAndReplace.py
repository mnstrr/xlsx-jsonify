import openpyxl
import json

class SearchAndReplace:

	def __init__(self, xlsx_doc, json_doc, sheet_name, blacklist):
		print('init SearchAndReplace')
		self.__xlsx_doc = xlsx_doc
		self.__json_doc = json_doc
		self.__sheet_name = sheet_name
		self.__blacklist = blacklist
		
		self.__main()


	def __main(self):
		print('start search')

		wb = openpyxl.load_workbook(self.__xlsx_doc)
		sheet = wb.get_sheet_by_name(self.__sheet_name)
		column = sheet.columns[1]
		print(column)
		print(len(column))