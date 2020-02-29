Sub FileSys()

Dim FilSysObj As Scripting.FileSystemObject
Dim Drv As Drive
Dim Fldr As Folder
Dim SubFld As Folder
Dim Fil As File
Dim Fils As Files
Dim FilPath As String

'Create a new file system object
Set FilSysObj = New Scripting.FileSystemObject

    ' Let's see all the drives we have and display some info about them.
    For Each Drv In FilSysObj.Drives:
        Debug.Print "The path to this drive is: " + Drv.Path
        Debug.Print "The type of drive is a: " + CStr(Drv.DriveType)
    Next
    
' Define a folder we want to work with
Set Fldr = FilSysObj.GetFolder("C:\Users\Alex\Desktop\YouTube Tutorials")
    
    ' Loop through the sub folders, these are also folder objects.
    For Each SubFld In Fldr.SubFolders
        Debug.Print SubFld.Path
        Debug.Print SubFld.Name
        Debug.Print SubFld.ParentFolder
        Debug.Print SubFld.Size
    Next

' Define a group of files
Set Fils = Fldr.Files

    For Each Fil In Fils
        Debug.Print Fil.Name
        Debug.Print Fil.Path
        Debug.Print Fil.Type
        Debug.Print Fil.ShortPath
    Next

'Define a file
FilPath = "C:\Users\Alex\Desktop\YouTube Tutorials\Python VBA - API Library\PythonAPILibrary.py"
    
' Print some details about the file
Debug.Print "The File Name is: " + FilSysObj.GetBaseName(FilPath)
Debug.Print "The File Absolute Path Name is: " + FilSysObj.GetAbsolutePathName(FilPath)
Debug.Print "The File Extension is: " + FilSysObj.GetExtensionName(FilPath)
Debug.Print "The File Name with it's extension is: " + FilSysObj.GetFileName(FilPath)
Debug.Print "The File is in Drive: " + FilSysObj.GetDriveName(FilPath)
Debug.Print "The File is in Folder: " + FilSysObj.GetParentFolderName(FilPath)
    
' Build a new path
Debug.Print FilSysObj.BuildPath(Path:="C:\Users\Alex\Desktop\YouTube Tutorials\Python VBA - API Library\", Name:="PythonAPILibrary.py")

' Grab a folder and move it or copy it.
Set Fldr = FilSysObj.GetFolder("C:\Users\Alex\Desktop\YouTube Tutorials\Python VBA - API Library")
    Fldr.Move ("C:\Users\Alex\Desktop\")
    Fldr.Copy ("C:\Users\Alex\Desktop\")
End Sub
