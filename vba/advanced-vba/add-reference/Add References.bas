Option Explicit

Sub WorkingWithReferences()

'Declare our variables
Dim vbProj As VBIDE.VBProject
Dim vbRefs As VBIDE.References
Dim vbRef As VBIDE.Reference

'Get the workbook VBA Project
Set vbProj = ThisWorkbook.VBProject

'Get the references that belong to the VB Project
Set vbRefs = vbProj.References

'Loop through each reference in the reference collection and print some details
For Each vbRef In vbRefs
    Debug.Print "---------------------------------"
    Debug.Print vbRef.Name
    Debug.Print vbRef.Description
    Debug.Print vbRef.GUID
    Debug.Print vbRef.Major
    Debug.Print vbRef.Minor
    Debug.Print vbRef.FullPath
    Debug.Print vbRef.BuiltIn
    Debug.Print vbRef.Type
Next

'PowerPoint
'Microsoft PowerPoint 16.0 Object Library
'{91493440-5A91-11CF-8700-00AA0060263B}
'2
'12
'C:\Program Files\Microsoft Office\Root\Office16\MSPPT.OLB
'False
'0

'Set a reference to a single library
Set vbRef = vbRefs.Item("PowerPoint")
    vbRefs.Remove Reference:=vbRef

'Lets add back the reference we removed
vbRefs.AddFromGuid GUID:="{91493440-5A91-11CF-8700-00AA0060263B}", Major:=2, Minor:=12
vbRefs.AddFromFile Filename:="C:\Program Files\Microsoft Office\Root\Office16\MSPPT.OLB"

End Sub
