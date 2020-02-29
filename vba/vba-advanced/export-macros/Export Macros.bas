Sub ExportModule()

'Declare the variables
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
Dim SelcItem As Variant
Dim ActWrkBook As Workbook

'Turn off screen updates
Application.ScreenUpdating = False
 
'Select the files for export
With Application.FileDialog(msoFileDialogOpen)
    .AllowMultiSelect = True
    .Show

   'Loop through the items we selected in the dialog box.
    For Each SelcItem In .SelectedItems
        
    'Set the active workbook to the one we just opened.
    Set ActWrkBook = Workbooks.Open(Filename:=SelcItem)
        WBName = Replace(ActWrkBook.Name, ".xlsm", "")
    
    'Get the Visual Basic Project for that workbook.
    Set VBProj = ActWrkBook.VBProject

    For Each VBComp In VBProj.VBComponents

        'Check the type of the object make sure its a code module
        If VBComp.Type = vbext_ct_StdModule Then
            
            'Let the user know it's continuing
            Debug.Print VBComp.Name
            Debug.Print VBComp.Type
            
            'Create the file name to export the file
            Filename = "C:\Users\Alex\Desktop\Files\" + WBName + "_" + VBComp.Name + ".bas"
            
            'Export the file.
            VBComp.Export Filename:=Filename
            
        End If
    Next
    
    'Close the active workbook
    ActWrkBook.Close
    Next
    
End With

'Turn off screen updates
Application.ScreenUpdating = True

End Sub
