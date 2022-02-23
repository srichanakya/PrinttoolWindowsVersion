from GUI import *
from FileLogic import *
from ExcelLogic import *
from TextExtract import *
if __name__=="__main__":
    el = excel()
    tl = C_TExtract()
    fl = logic(el,tl)
    guInterface= GraphicalInterfaceClass(fl)



