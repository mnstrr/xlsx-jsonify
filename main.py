import time
from searchAndReplace import SearchAndReplace as Sar


class Main:
    xlxs_doc = 'translation2.xlsx'
    json_doc = 'data.json'
    json_doc_new = 'data_new.json'
    sheet_name = 'sheet1'
    searchcolumn = 5
    replacecolumn = 6

    blacklist = ['']

    def __init__(self):
        sar = Sar(self.xlxs_doc, self.json_doc, self.json_doc_new, self.sheet_name, self.searchcolumn,
                  self.replacecolumn, self.blacklist)


start_time = time.time()
Main()
print('--------------------')
print("Execution time is %s seconds" % "%0.2f" % (time.time() - start_time))
