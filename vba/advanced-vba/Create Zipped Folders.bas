
Sub ZipFiles()

Dim ShellObj As Shell32.Shell
Dim TargetFolder As Shell32.Folder
Dim DestinationFolder As Shell32.Folder
Dim DesktopName As Shell32.Folder

'Define the desktop
DesktopPath = "C:\Users\305197\Desktop"

'Define the folder that contains the objects you want to zip
TargetPath = "C:\Users\305197\Desktop\Target"

'Define the Destination folder path, this is the zipped file
DestinationPath = "C:\Users\305197\Desktop\Destination"

'Declare a new shell object
Set ShellObj = New Shell32.Shell

'Create a new folder on the desktop, to contain all the items we want to compress
Set DesktopName = ShellObj.Namespace(DesktopPath)

'Define the folder we are copying the items from.
Set TargetFolder = ShellObj.Namespace(TargetPath)

    'Make the new folder.
    DesktopName.NewFolder bName:=DestinationPath
    
    'Copy each item to the destination folder.
    For Each fldrItem In TargetFolder.Items
        ShellObj.Namespace(DestinationPath).CopyHere vItem:=fldrItem
    Next
    
    'If we use power shell we can use command line arguments
    args = "Compress-Archive -Path " + DestinationPath + " -DestinationPath " + DestinationPath + ".zip"
    
    'Call PowerShell and pass through the arguments.
    ShellObj.ShellExecute File:="Powershell", vArgs:=args

End Sub
