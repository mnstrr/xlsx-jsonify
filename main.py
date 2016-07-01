import time
from searchAndReplace import SearchAndReplace as Sar

class Main:    

    xlxs_doc = 'translation2.xlsx'
    json_doc = 'translation.json'
    sheet_name = 'sheet1'
    blacklist = ['']

    def __init__(self):
        sar = Sar(self.xlxs_doc, self.json_doc, self.sheet_name,self.blacklist)

  
start_time = time.time()
Main()
print('--------------------')
print("Execution time is %s seconds" % "%0.2f" % (time.time() - start_time))