#This program is used to transform the .xls exported JIRA file into Azure readable .xlsx format where mapping from JIRA to 
#Azure is performed and nnecessary contents eg. logo, blank rows, unwanted statements are removed from JIRA exported .xls file.


import pandas as pd
import numpy as np
import xlrd as xl
import xlwt
import glob
import openpyxl
import os
import easygui
from tkinter import Tk 
from tkinter.filedialog import askopenfilename
import win32com.client
from win32com.client import constants

filePath = easygui.fileopenbox("JIRA to AZURE migrator: Enter JIRA Exported file path with .xls extention for conversion : ")
excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(filePath)

wb.SaveAs(filePath+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()

filePath = easygui.fileopenbox("JIRA to AZURE migrator: Enter Converted JIRA export file path with .xlsx extention for mapping into Azure : ")
filePath1=filePath[15:]
filePath1=filePath1[:-5]

f = filePath
DELETE_THIS = ""    

excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(filePath)

#exc = win32com.client.gencache.EnsureDispatch("Excel.Application")
#exc.Visible = 1
#exc.Workbooks.Open(Filename=f)

totRows = excel.Range("A1048576").End(constants.xlUp).Row
#print(totRows)

excel.Rows("%d:%d" % (totRows, totRows)).Select()
excel.Selection.Delete(Shift=constants.xlUp)

ind = 1
while True:
    excel.Rows("%d:%d" % (ind, ind)).Select()
    excel.Selection.Delete(Shift=constants.xlUp)
    if ind == 2:
        break
    else:
        ind += 1

excel.Rows("%d:%d" % (1, 1)).Select()
excel.Selection.Delete(Shift=constants.xlUp)

wb.SaveAs(filePath, FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                              #FileFormat = 56 is for .xls extension
excel.Application.Quit()

f = pd.read_excel(filePath)
df=pd.DataFrame(f)

#print(df)
rows=len(df.index)
print("Total number of WorkItems to be migrated into ADO : "+str(rows-8))


# function definition 
def highlight_cols(x):
    # copy df to new - original data is not changed 
    df = x.copy() 
    # select all values to yellow color 
    df.loc[:, :] = 'background-color: yellow'
    # return color df 
    return df  

df.style.apply(highlight_cols(df), axis = None)  

df.insert(4, 'Issue Type_Converted', np.nan)
df.insert(6, 'Status_Converted', np.nan)
df.insert(8, 'Priority_Coverted', np.nan)
df.insert(11, 'Assignee_Converted', np.nan)

df.to_excel(r'C:/JiraToAzure/Transformed_Sheet.xlsx')  #temporary file created which will be deleted later

for wsrWB in glob.glob("C:\JiraToAzure\*.xlsx"):
        print("\n\n Excel workbook to be processed :" + wsrWB)
        wb = xl.open_workbook(wsrWB)                    #opening & reading the excel file
        s1 = wb.sheet_by_index(0)                     #extracting the worksheet
        s1.cell_value(0,0)                          #initializing cell from the excel file mentioned through the cell position
        
wsrWB = openpyxl.load_workbook(wsrWB)
wsrSheetList = wsrWB.sheetnames
 
wsrSheet1 = wsrSheetList[0]
sheet_to_focus = wsrSheet1

wsrSheetList[0] = wsrWB.active
wsrSheet1 = wsrWB[wsrSheet1]

#Logic to apply farmula in cell for mapping JIRA items to Azure

for i in range(rows+2):
    if i>1:
        wsrSheet1['F'+ str(i)] = '=''IF(OR(E'+str(i)+'= "Task", E'+str(i)+' = "Sub-Task"),"Task",(IF(E'+str(i)+'="Epic","Epic",IF(E'+str(i)+'="Bug","Bug",IF(E'+str(i)+'="Feature","Feature","User Story")))))'
        wsrSheet1['H'+ str(i)] = '=''IF(OR(G'+str(i)+'= "Removed",G'+str(i)+' = "Cancelled"),"Removed",(IF(G'+str(i)+'="Done","Closed",IF(OR(G'+str(i)+'="In Progress", G'+str(i)+'="Testing",G'+str(i)+'="Deployment",G'+str(i)+'="Awaiting Feedback"),"Active","New"))))'
        wsrSheet1['J'+ str(i)] = '=''IF(OR(I'+str(i)+'= "High",I'+str(i)+' = "Critical",I'+str(i)+'= "Major"),1,(IF(I'+str(i)+'="Medium",2,IF(I'+str(i)+'="Low",3,4))))'
        wsrSheet1['M'+ str(i)] = '=''IF((ISNUMBER(SEARCH("[X]",L'+str(i)+'))),"",(IF(L'+str(i)+'="Unassigned","",(CONCAT((LEFT(L'+str(i)+',FIND(" ",L'+str(i)+',1)-1)),".",(RIGHT(L'+str(i)+',LEN(L'+str(i)+')-FIND(" ",L'+str(i)+'))),".con@derivco.com")))))'.replace("=@","=")

wsrWB.save("C:\JiraToAzure\JIRA_Transformed_"+str(2)+".xlsx") 
wsrWB.close
    
os.remove(r'C:/JiraToAzure/Transformed_Sheet.xlsx') #temporary file deleted


f = r"C:\JiraToAzure\JIRA_Transformed_"+str(2)+".xlsx"
DELETE_THIS = ""

excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
wb = excel.Workbooks.Open(f)

#logic to delete the blank rows from the excel file and shifting it up
row = 1
bow=row+1
while True:
    excel.Range("B%d" % row).Select()
    data = excel.ActiveCell.FormulaR1C1
    excel.Range("A%d" % bow).Select()
    condition = excel.ActiveCell.FormulaR1C1

    if data == '' and condition =='':
        break
    elif data =='':
        excel.Rows("%d:%d" % (row, row)).Select()
        excel.Selection.Delete(Shift=constants.xlUp)
    else:
        row += 1
        bow += 1


#deleting the last row which is unnecessary
totRows = excel.Range("A1048576").End(constants.xlUp).Row
print("Total rows : "+str(totRows))
excel.Rows("%d:%d" % (totRows, totRows)).Select()
excel.Selection.Delete(Shift=constants.xlUp)


for i in range(totRows):
    if i>1:
        excel.Range("M%d" % i).Select()
        excel.ActiveCell.Replace(What="=@", Replacement="=", LookAt=constants.xlPart, SearchOrder=constants.xlByRows, MatchCase=False, SearchFormat=False, ReplaceFormat=False, FormulaVersion=constants.xlReplaceFormula2)
        #wsrSheet1['M'+ str(i)] = str(wsrSheet1['M'+ str(i)].value).replace("=@","=") 

wb.SaveAs(f, FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                       #FileFormat = 56 is for .xls extension
excel.Application.Quit()

#exc.ActiveWorkbook.Save
#exc.Workbooks.Close()
