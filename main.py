import time
from searchAndReplace import SearchAndReplace as Sar

class Main:    

    def __init__(self):
        sar = Sar()

  
start_time = time.time()
Main()
print('--------------------')
print("Execution time is %s seconds" % "%0.2f" % (time.time() - start_time))