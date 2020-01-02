Attribute VB_Name = "WorkingWithTemplates"
Sub WorkingWithTemplates()

'Declare Variables
Dim WrdTemplates As Templates
Dim WrdTemplate As Template
Dim WrdDoc As Document

'Grab the current document.
Set WrdDoc = ActiveDocument

'The Word Application contains all the templates, so if we want the collection we go there.
'It's important to note that templates will only appear if they are either currently opened or
'have been added as a global path.
Set WrdTemplates = Application.Templates

'Loop through each template in the collection.
For Each WrdTemplate In WrdTemplates
    
    'Print some details
    Debug.Print "Template Name: " + WrdTemplate.Name
    Debug.Print "Template Path: " + WrdTemplate.Path
    Debug.Print "Template Full Name: " + WrdTemplate.FullName
    Debug.Print "--------------------"
    
Next

'What if I want to programatically add a Template to my Global Path?
'Well Word considers a Template an AddIn, so if we go to the AddIns collection
'and add it, then it will always be loaded everytime word is started.

'Declare the variables.
Dim WrdAddIns As AddIns
Dim WrdAddIn As AddIn

'Grab the AddIns Collection.
Set WrdAddIns = Application.AddIns

'Let's see which AddIns are currently Installed.
For Each WrdAddIn In WrdAddIns

    Debug.Print "Add In Name: " + WrdAddIn.Name
    Debug.Print "Add In Path: " + WrdAddIn.Path
    Debug.Print "Add In Installed Flag: " + CStr(WrdAddIn.Installed)
    Debug.Print "--------------------"
    
Next

'Lets add a template to our Global Collection by adding it as an AddIn. Again, it will now be available anytime
'you open word.
WrdAddIns.Add FileName:="C:\Users\Alex\AppData\Roaming\Microsoft\Templates\Blue spheres appointment calendar.dotx", Install:=True

'A Word Document can have an attached template, most of the time it's a just a normal word document. Access
'it using the "AttachedTemplate" property.
Set WrdTemplate = WrdDoc.AttachedTemplate
    
    'Now our template is just a normal word document, so the name should just be 'Normal.dotm'
    Debug.Print "Attached Template Name is: " + WrdTemplate.Name

End Sub

Sub ChangeAttachTemplate()

'Declare Variables
Dim WrdTemplates As Templates
Dim WrdTemplate As Template
Dim WrdDoc As Document

'Grab the current document.
Set WrdDoc = ActiveDocument

'Change the attached template.
WrdDoc.AttachedTemplate = "C:\Users\Alex\AppData\Roaming\Microsoft\Templates\Report .dotx"

'Now that we've attached the template, we need to update the styles of the
'document so that they reflect the new template.
WrdDoc.UpdateStyles

End Sub

Sub WorkingWithATemplate()

'Declare Variables
Dim WrdTemplates As Templates
Dim WrdTemplate As Template
Dim WrdDoc As Document

'Grab the current document.
Set WrdDoc = ActiveDocument

'Then the attached template.
Set WrdTemplate = WrdDoc.AttachedTemplate

    'Print some details.
    Debug.Print "Word Template Type: " + CStr(WrdTemplate.Type)
    
    'wdAttachedTemplate  2   An attached template.
    'wdGlobalTemplate    1   A global template.
    'wdNormalTemplate    0   The normal default template.
    
    'See if spell checking is on. 0 - True, 1 - False
    Debug.Print "Spell Checking Flag: " + CStr(WrdTemplate.NoProofing)
    
    'It also has a VB Project associated with it.
    Debug.Print "Word Template VB Project Name: " + CStr(WrdTemplate.VBProject.FileName)
    
    'You can also open it as a document, this returns a normal document object like we've used in the past.
    'WrdTemplate.OpenAsDocument

End Sub

Sub WorkingWithBuildingBlocks()

Dim WrdTemplates As Templates
Dim WrdTemplate As Template

Dim DocBlockEntries As BuildingBlockEntries
Dim DocBlocks As BuildingBlocks
Dim DocBlock As BuildingBlock

Dim WrdDoc As Document

Set WrdDoc = ThisDocument

Set WrdTemplates = Application.Templates
Set WrdTemplate = WrdDoc.AttachedTemplate

    
    Debug.Print WrdDoc.Name
    



End Sub
