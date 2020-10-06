Option Explicit

Sub WorkingWithContentControls()

'Content controls are bounded and potentially labeled regions in a document
'that serve as containers for specific types of content. Individual content
'controls can contain content such as dates, lists, or paragraphs of formatted
'text. In some cases, content controls might remind you of forms. However,
'they are much more powerful, flexible, and useful because they enable you
'to create rich, structured blocks of content. Content controls enable you
'to author templates that insert well-defined blocks into your documents.

Dim WrdContentControls As ContentControls
Dim WrdContentControl As ContentControl

Dim WrdDoc As Document
Dim WrdRng As Range

Dim CurrentDate As Date

'Grab the Current Document.
Set WrdDoc = ThisDocument

'Grab a Range, in this case the bookmark Referencing the Date Picker.
Set WrdRng = WrdDoc.Bookmarks("PlaceholderDatePicker").Range

'Add a Word Content Control.
Set WrdContentControl = WrdDoc.ContentControls.Add(Type:=wdContentControlDate, Range:=WrdRng)
    
    'Set the Appearance
    WrdContentControl.Appearance = wdContentControlTags

    'Give my Content Control a Title.
    WrdContentControl.Title = "Today's Date"

    'Grab Today's Date.
    CurrentDate = Date
    
    'Set the Range Text to the Current Date.
    WrdContentControl.Range.Text = CurrentDate
    
    'Set the Calendar Type.
    WrdContentControl.DateCalendarType = wdCalendarWestern
    
    'Set the Display Format of the Date.
    WrdContentControl.DateDisplayFormat = "MM-dd-yyyy"

End Sub

Sub AddDropDownContentControl()

Dim WrdContentControls As ContentControls
Dim WrdContentControl As ContentControl

Dim WrdDoc As Document
Dim WrdRng As Range

'Grab the Current Document.
Set WrdDoc = ThisDocument

'Grab a Range, in this case the bookmark Referencing the Drop-Down List.
Set WrdRng = WrdDoc.Bookmarks("PlaceholderDropDown").Range

'Add a Word Content Control.
Set WrdContentControl = WrdDoc.ContentControls.Add(Type:=wdContentControlDropdownList, Range:=WrdRng)
    
    'Set the Appearance
    WrdContentControl.Appearance = wdContentControlBoundingBox

    'Give my Content Control a Title.
    WrdContentControl.Title = "Programming Languages"
    
    'Set the Placeholder Text.
    WrdContentControl.SetPlaceholderText Text:="Please select your favorite programming language..."
    
    'Add some items to the dropdown list.
    WrdContentControl.DropdownListEntries.Add Text:="VBA", Index:=0
    WrdContentControl.DropdownListEntries.Add Text:="Python", Index:=1
    WrdContentControl.DropdownListEntries.Add Text:="JavaScript", Index:=2
    WrdContentControl.DropdownListEntries.Add Text:="TypeScript", Index:=3
    WrdContentControl.DropdownListEntries.Add Text:="C++", Index:=4
    WrdContentControl.DropdownListEntries.Add Text:="C", Index:=5
    WrdContentControl.DropdownListEntries.Add Text:="Java", Index:=6

End Sub

Sub AddComboBoxContentControl()

Dim WrdContentControls As ContentControls
Dim WrdContentControl As ContentControl

Dim WrdDoc As Document
Dim WrdRng As Range

'Grab the Current Document.
Set WrdDoc = ThisDocument

'Grab a Range, in this case the bookmark Referencing the Combo-Box.
Set WrdRng = WrdDoc.Bookmarks("PlaceholderComboBox").Range

'Add a Word Content Control.
Set WrdContentControl = WrdDoc.ContentControls.Add(Type:=wdContentControlComboBox, Range:=WrdRng)
    
    'Set the Appearance
    WrdContentControl.Appearance = wdContentControlBoundingBox

    'Give my Content Control a Title.
    WrdContentControl.Title = "Favorite Pokemon"
    
    'Set the Placeholder Text.
    WrdContentControl.SetPlaceholderText Text:="Please select your favorite pokemon..."
    
    'Add some items to the ComboBox, note how we are using "DropdownListEntries" again.
    WrdContentControl.DropdownListEntries.Add Text:="Charizard", Index:=0
    WrdContentControl.DropdownListEntries.Add Text:="Venasaur", Index:=1
    WrdContentControl.DropdownListEntries.Add Text:="Blastoise", Index:=2

End Sub

Sub AddCheckBoxContentControl()

Dim WrdContentControls As ContentControls
Dim WrdContentControl As ContentControl

Dim WrdDoc As Document
Dim WrdRng As Range

'Grab the Current Document.
Set WrdDoc = ThisDocument

'Grab a Range, in this case the bookmark Referencing the Combo-Box.
Set WrdRng = WrdDoc.Bookmarks("PlaceholderCheckBox").Range

'Add a Word Content Control.
Set WrdContentControl = WrdDoc.ContentControls.Add(Type:=wdContentControlCheckBox, Range:=WrdRng)
    
    'Set the Appearance
    WrdContentControl.Appearance = wdContentControlBoundingBox

    'Give my Content Control a Title.
    WrdContentControl.Title = "Do You Like Star Wars?"
    
    'Set the Unchecked Symbol.
    'WrdContentControl.SetUncheckedSymbol CharacterNumber:=&H2610, Font:="Segoe UI"
    
    'Set the Checked Symbol.
    'WrdContentControl.SetCheckedSymbol CharacterNumber:=&H2610, Font:="Segoe UI"
    
    'Set the State to Check.
    WrdContentControl.Checked = True

End Sub
