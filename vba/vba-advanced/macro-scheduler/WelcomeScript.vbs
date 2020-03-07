
Set xlsxApp = CreateObject("Excel.Application")
    xlsxApp.Visible = True

Set xlsxWorkbook = xlsxApp.Workbooks.Open("C:\Users\Alex\Desktop\myWelcomeFile.xlsm")
    xlsxApp.Run("MyWelcomeMessage")