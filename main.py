import ExcelHandler as eh
import WordHandler as wh

if __name__ == '__main__':
    dataList = eh.getExcelOneRowData("excel.xlsx",0,2)
    print(dataList)
    replaceDict = {"xxx":1,"yyy":2,"zzz":3}
    wh.generateWordDocuments("wordFile.docx",replaceDict,dataList,1)