from datetime import datetime
import os
import glob
import PyPDF2
from tkinter import messagebox
import re
from TextExtract import *

class logic:

    def __init__(self, excelLogic,TextLogic):
        self.eLogic = excelLogic
        self.tLogic = TextLogic
        self.Log = []


    def FileExecution(self, source, destination, fName, cName):
        # print(len(source))
        if len(source) != 0 and len(destination) != 0:
            try:
                FileDirectory = source + "\\*.pdf"
                all_Files = glob.glob(FileDirectory)
                # print(all_Files)
                if len(all_Files) > 0:
                    for files in all_Files:
                        #print(files)
                        matches = self.tLogic.FileName(files)
                        #print(f"matches here : {matches} {len(matches)}")
                        if matches != "EMPTY":
                            renamefile = source + "\\" + matches + ".pdf"
                            #print(os.path.exists(renamefile))
                            #print(os.path.exists(files))
                            if os.path.exists(renamefile) == False and os.path.exists(files) == True:
                                os.rename(files, renamefile)
                                DESTINATION = destination + "\\" + matches + ".pdf"
                                initial = renamefile
                                final = DESTINATION
                                #print(os.path.exists(initial))
                                #print(os.path.exists(final))
                                if os.path.exists(final) == False and os.path.exists(initial) == True:
                                    os.renames(initial, final)
                                    self.eLogic.GetExcelFile(final,matches)
                                elif os.path.exists(final) == True and os.path.exists(initial) == True:
                                    self.Log.append([files,matches, "File exists in destination folder"])
                                    # print(self.Log)
                                else:
                                    self.Log.append([files,matches, "couldn't copy file to destination"])
                        else:
                            self.Log.append([files,"Unable to extract file name ", "couldn't copy file to destination"])
                   #End of List

                    messagebox.showinfo(title="Print tool",message="Successful")

                    #print(self.Log)
                    #print(datetime.today())
                else:
                    messagebox.showerror(title="Error!!!", message="End of List")
                # messagebox.showinfo(title="Success", message="File copied from source to destination")
            except:
                messagebox.showerror(title="Error!!!", message="File Missing")
            self.eLogic.infoLog(self.Log)
            self.Log=[]
            self.eLogic.openExcel()
        else:
            messagebox.showerror(title="Error!!!", message="File paths missing. Please enter valid paths")
