Attribute VB_Name = "Module1"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("D10").Select
    Application.CommandBars("Queries and Connections").Visible = False
    Application.CommandBars("Queries and Connections").Visible = False
    Application.CutCopyMode = False
    Range("C5").Select
End Sub

Sub PowerPivot()
Dim WrkBookConnections As Connections
Dim TblConnection As WorkbookConnection


'ThisWorkbook.Connections.Add2 Name:="MyBakeryDataSet", _
'                              Description:="This contains all the data from my bakery dataset.", _
'                              ConnectionString:="WORKSHEET;https://d.docs.live.net/8bc640c57cda25b6/Growth - Tutorial Videos/Lessons - VBA/Power Pivot/PowerPivot - Model Object.xlsm", _
'                              CommandText:="PowerPivot - Model Object.xlsm!Table1", _
'                              lCmdtype:=7, _
'                              CreateModelConnection:=True, _
'                              ImportRelationships:=True
                              

ThisWorkbook.Connections.Add Name:="MyBakeryDataSet", _
                             Description:="This contains all the data from my bakery dataset.", _
                             ConnectionString:="WORKSHEET;https://d.docs.live.net/8bc640c57cda25b6/Growth - Tutorial Videos/Lessons - VBA/Power Pivot/PowerPivot - Model Object.xlsm", _
                             CommandText:="PowerPivot - Model Object.xlsm!Table1"
                             'lCmdtype:=7

                              

End Sub

