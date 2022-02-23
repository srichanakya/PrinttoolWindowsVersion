import re
import PyPDF2


class C_TExtract:
    def FileName(self, sourceFilename):

        # print(file)
        matches = []
        pdfFileObj = open(sourceFilename, mode="rb")
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        # print(pdfReader.numPages)
        pageObj = pdfReader.getPage(0)
        data = pageObj.extractText()
        #print(data)
        pattern = f"{'invoice'}|{'purchase order'}|{'agreement of purchase'}|{'delivery receipt'}"
        matches = re.findall(pattern, data.lower())
        #print(len(matches))
        pdfFileObj.close()
        #print(bool(matches))
        if matches:
            return matches[0]
        else:
            return "EMPTY"
