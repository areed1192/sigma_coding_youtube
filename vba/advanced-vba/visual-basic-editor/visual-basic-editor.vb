Sub VBIDE()

Dim VBAEditor As VBIDE.VBE
Dim VBProj As VBIDE.VBProject
Dim VBProjs As VBIDE.VBProjects
Dim VBComp As VBIDE.VBComponent
Dim CodeMod As VBIDE.CodeModule

'Create a reference to the editor itself
Set VBAEditor = Application.VBE
    Debug.Print VBAEditor.Version
    
'Create a reference to the projects in the Visual Basic Editor
Set VBProjs = VBAEditor.VBProjects

    
'Loop through each project, get the file name and the regular name.
For Each VBProj In VBAEditor.VBProjects
    Debug.Print VBProj.Filename
    Debug.Print VBProj.Name
Next

'Reference a single project
Set VBProj = VBAEditor.VBProjects.Item(3)

'Reference a single project and print some information about that module.
Set VBComp = VBProj.VBComponents.Item("Module1")
    Debug.Print VBComp.CodeModule
    Debug.Print VBComp.Name
    Debug.Print VBComp.Type
    Debug.Print VBComp.CodeModule
    
'Reference the code in that module
Set CodeMod = VBComp.CodeModule
    Debug.Print CodeMod.CountOfLines

'Add a new line of code.
CodeMod.InsertLines Line:=1, String:="Hi there"

End Sub

Sub VbeAddReference()
Dim VbRefs As VBIDE.References
Dim VbRef As VBIDE.Reference
    
'Get the references Collections
Set VbRefs = ThisWorkbook.VBProject.References

'Get some information related to the VBA References.
For Each VbRef In VbRefs
    Debug.Print VbRef.Name
    Debug.Print VbRef.GUID
    Debug.Print VbRef.Description
    Debug.Print VbRef.FullPath
Next

'EXAMPLE
'Excel
'{00020813-0000-0000-C000-000000000046}
'Microsoft Excel 16.0 Object Library
'C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE

'Add a new reference progmatically, this is for the Word Object Library.
VbRefs.AddFromFile ("C:\Program Files\Microsoft Office\root\Office16\MSWORD.OLB")

'Add a new reference progmatically, this is for PowerPoint Object Library
VbRefs.AddFromGuid "{91493440-5A91-11CF-8700-00AA0060263B}", 1, 0

'Count the number of references.
Debug.Print VbRefs.Count
    
End Sub

'' or
'Set VBProj = Application.Workbooks("VBEBook.xlsm").VBProject
'
''''''''''''''''''''''''''''''''''''''''''''
'Set VBComp = ActiveWorkbook.VBProject.VBComponents("Module1")
'' or
'Set VBComp = VBProj.VBComponents("Module1")
''''''''''''''''''''''''''''''''''''''''''''
'
'Set CodeMod = ActiveWorkbook.VBProject.VBComponents("Module1").CodeModule
'' or
'Set CodeMod = VBComp.CodeModule
'Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\AppID\{A34B1EEC-57B6-420D-981E-2D2DB744A174}
