Attribute VB_Name = "Module2"
Private Sub Worksheet_Change(ByVal Target As Range)



    'Set the Variables to be used
    Dim pt As PivotTable
    Dim Field As PivotField
    Dim new_category As String


    If Intersect(Target, Range("D1")) Is Nothing Then Exit Sub

    'Here you amend to filter your data
    Set pt = Worksheets("Sheet19").PivotTables("Fiancial Pivot Table")
    Set Field = pt.PivotFields("Year")
    new_category = ActiveSheet.Range("D1").Value
    xstr = Target.Text
    
    'This updates and refreshes the PIVOT table

        Field.ClearAllFilters
        Field.CurrentPage = xstr
        pt.RefreshTable


End If

End Sub

