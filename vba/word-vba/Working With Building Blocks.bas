Attribute VB_Name = "WorkingWithBuildingBlocks"
Sub WorkingWithBuildingBlocks()



' Introduction:
' -------------
' In this module we will go over building blocks in a Word Document. Building blocks are an extensive topic, so
' I wanted to add some notes for you so you can understand how they are organized. All these notes come directly
' from the Microsoft VBA Documentation. If you'd like to read it yourself.
'
' Feel free to go to the following link:
' https://docs.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-building-blocks



' What are building blocks?
' ------------------------
' A building block is pre-built content, similar to autotext,
' that may contain text, images, and formatting.



' How are they organized in the Object model?
' -------------------------------------------
' The building blocks object model includes three objects and four collections.
' These enable you to create an organizational structure that works for your specific
' needs and to modify the structure for a specific solution. Here is a list of all the objects
' and collections.
'
' Here is a table that lists all the objects and collections.
'
' Name                    Description
' --------                --------
' BuildingBlock           A specific building block entry.
' BuildingBlocks          A collection of building block entries in a template that are of the same type and category.
' BuildingBlockEntries    A collection of all the building blocks in a template.
' BuildingBlockType       A building block type.
' BuildingBlockTypes      A collection of building block types.
' Category                A building block category.
' Categories              A collection of building block categories.



' What are the difference between Categories and Types?
' ----------------------------------------------------
'
' TYPES:
' Building blocks are organized by type and category. Building block types are composed of a
' limited number of WdBuildingBlockTypes constants. Although there are a limited number of these constants,
' that number is not small. There are 35 different WdBuildingBlockTypes constants. These types help you to
' define and organize your building blocks and, although you cannot create additional building block types,
' you can create an unlimited number of categories for each type.
'
' CATEGORIES:
' Categories are composed of an unlimited number of strings that you can define to organize your custom building blocks.
' Building blocks are stored in templates. By default, the templates that are included with Word have building block categories like
' "General" and "Built-In". However, you are not limited to just the categories that are included in these templates. A category can
' be any string that you define. Types and categories are explained later in this topic.


' Additional Links:
' -----------------
'
' WdBuildingBlockTypes - Enumeration:
' https://docs.microsoft.com/en-us/office/vba/api/word.wdbuildingblocktypes

' CODE STARTS HERE


'Declare Variables
Dim WrdTemplate As Template

Dim DocBlockEntries As BuildingBlockEntries

Dim DocBlockTypes As BuildingBlockTypes
Dim DocBlockType As BuildingBlockType

Dim DocBlockCategories As Categories
Dim DocBlockCategory As Category

Dim DocBlocks As BuildingBlocks
Dim DocBlock As BuildingBlock

Dim WrdRng As Range

Dim IntCount As Integer

'This is the challenging part, technically Built-In building blocks are their
'own template, so to access them you need to grab that template.
Set WrdTemplate = Application.Templates("C:\Users\Alex\AppData\Roaming\Microsoft\Document Building Blocks\1033\16\Built-In Building Blocks.dotx")

    'Print the Attached Template Name.
    Debug.Print "Template Name is " + WrdTemplate.Name
    Debug.Print "-----------------------"
    
'A template object has a Building Blocks Collection. Remember if we want ALL OF THE BUILDING BLOCKS and not a filtered collection
'then we need to use the "BuildingBlockEntries" property.
Set DocBlockEntries = WrdTemplate.BuildingBlockEntries

'Let's see some details about the BuildingBlockEntries collection.
Debug.Print "There are " + CStr(DocBlockEntries.Count) + " in the `BuildingBlockEntries` collection for the " + WrdTemplate.Name + " template."
Debug.Print "-----------------------"

'NOTES ON USING BUILDING BLOCK ENTRIES
'-----------------------------------
'This collection is a little strange because you can use a For Each loop to iterate over it. It will return an error.
'You need to access the items individually. On top of that I seem to run into issues when it comes to printing or accessing
'details about an individual building block.
'
'Print the name, no issue.
'Debug.Print DocBlockEntries.Item(2).Name
'
'Print the category, issue.
'Debug.Print DocBlockEntries.Item(2).Category


'BUILDING BLOCK TYPES
'--------------------


'Lets maybe try by type instead and see if we have any better luck. First let's see how many Types there are
Debug.Print "There are " + CStr(WrdTemplate.BuildingBlockTypes.Count) + " building block types."
Debug.Print "-----------------------"

'Next store the BuildingBlockTypes collection in an object variable.
Set DocBlockTypes = WrdTemplate.BuildingBlockTypes

'Looping this collection is challenging, you can't use a standard for each loop. Instead we have to use a For loop.
'We want to use the count property of the "BuildingBlockTypes" collection to determine when to stop.
For IntCount = 1 To DocBlockTypes.Count
 
    'As we loop, let's store each BuildingBlockType in an object variable.
    Set DocBlockType = DocBlockTypes(IntCount)
    
    'Print some details.
    Debug.Print "Building Block Type has a name of: " + DocBlockType.Name
    Debug.Print "Building Block Type has a index of: " + CStr(DocBlockType.Index)
    Debug.Print "Building Block Type has " + CStr(DocBlockType.Categories.Count) + " categories in it."
    Debug.Print "-----------------------"

Next

'Now up above, we looped through all of the BuildingBlockTypes. However, we might just want one, so let's access a single type.
'In this case we will grab the Building Block Table Type.

'Method One - Enumeration:
Set DocBlockType = DocBlockTypes(21)

'Method Two - Constant:
Set DocBlockType = DocBlockTypes(wdTypeTables)
    
    Debug.Print "This is the " + CStr(DocBlockType.Name) + " building blocks."
    Debug.Print "It has an index of " + CStr(DocBlockType.Index)
    Debug.Print "-----------------------"

'With a BuildingBlockType we can have multiple categories. The category, as mentioned above, is just another
'way of organizing the content. Let's see if our Table Type has an categories.
If DocBlockType.Categories.Count > 0 Then
    Debug.Print "This Type (" + DocBlockType.Name + ") has " + CStr(DocBlockType.Categories.Count) + " categories."
Else
    Debug.Print "This Type has no categories."
End If


'BUILDING BLOCK CATEGORIES
'-------------------------


'Let's store the Collection in an object variable
Set DocBlockCategories = DocBlockType.Categories

'Just like the BuildingBlockTypes collection, iterating over it requires a tradtional For loop and NOT A FOR EACH LOOP.
For IntCount = 1 To DocBlockCategories.Count
    
    'Assign it to a new object variable.
    Set DocBlockCategory = DocBlockCategories.Item(IntCount)
    
    'Print some details.
    Debug.Print "Category has a name of: " + DocBlockCategory.Name
    Debug.Print "Category has a index of: " + CStr(DocBlockCategory.Index)
    Debug.Print "-----------------------"
    
Next

'Method One - Index:
Set DocBlockCategory = DocBlockCategories.Item(1)

'Method Two - Key:
Set DocBlockCategory = DocBlockCategories.Item("Built-In")


'BUILDING BLOCKS
'---------------


'With our category now in hand, let's grab the building blocks that belong to it.
Set DocBlocks = DocBlockCategory.BuildingBlocks

'Print the count of Building Blocks in that category.
Debug.Print "There are " + CStr(DocBlocks.Count) + " in the " + DocBlockCategory.Name + " category that" _
; " belongs to the " + DocBlockType.Name + " type."
Debug.Print "-----------------------"

'You guessed it, to iterate over them we again have to use a For loop.
For IntCount = 1 To DocBlocks.Count
    
    'Assign it to a new object variable.
    Set DocBlock = DocBlocks.Item(IntCount)
    
    'Print some details.
    Debug.Print "Building Block has a name of: " + DocBlock.Name
    Debug.Print "Building Block has a index of: " + CStr(DocBlock.Index)
    Debug.Print "-----------------------"

Next

'Let's grab a single Building Block, in this case a "With Subheads 1".

'Method One - Index:
Set DocBlock = DocBlocks.Item(4)

'Method Two - Key:
Set DocBlock = DocBlocks.Item("With subheads 1")

'Print some details about our building block.
Debug.Print "Building Block Name: " + DocBlock.Name
Debug.Print "Building Block Description: " + DocBlock.Description
Debug.Print "Building Block Category Name: " + DocBlock.Category.Name
Debug.Print "Building Block ID: " + DocBlock.ID
Debug.Print "Building Block Value: " + DocBlock.Value 'This seems to return the actual content itself.
Debug.Print "-----------------------"

'Lets insert the Building Block, but we need to specify where we want it. Let's do the end of paragraph 1.
'First let's define the end of the paragraph.
EndRng = ActiveDocument.Paragraphs(1).Range.End

'Now the End property returns a Long value, so we need to define our "Range" object.
'With the range object we specify a starting and ending "CHARACTER POSITION". In this case we start at character 0 and
'end at character 1.
Set WrdRng = ActiveDocument.Range(Start:=0, End:=EndRng)
    
    'Use the insert method to insert it. I also want Rich Formatting Text.
    DocBlock.Insert Where:=WrdRng, RichText:=True
    
    'WARNING!
    'WARNING!
    'WARNING!
    
    '------------------
    
    'We could delete it, but keep in mind IT WILL BE PERMANENT! That means if you delete a built-in
    'one it's gone FOREVER! I learned this fact the hard way.
    
    '------------------
    
    'WARNING!
    'WARNING!
    'WARNING!
    
    'DocBlock.Delete

End Sub


