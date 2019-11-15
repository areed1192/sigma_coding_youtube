
Sub ShellControls()

Dim ShellObj As Shell32.Shell
Dim ShellObjFldr As Shell32.Folder
Dim FldrItems As Shell32.FolderItems
Dim FldrItem As Shell32.FolderItem
Dim PythonExe, PythonScript, PythonArgs As String

'HTML OBJECT LIBRARY
'https://docs.microsoft.com/en-us/windows/desktop/shell/shell

'Create a new shell object
Set ShellObj = New Shell32.Shell

'Provide file path to Python.exe
'USE TRIPLE QUOTES WHEN FILE PATH CONTAINS SPACES.
PythonScript = "C:\Users\Alex\Desktop\pythonArgs.py"
PythonArgs = " arg1 arg2"

'Open the notepade application
ShellObj.ShellExecute file:=PythonScript, vArgs:=PythonArgs

'Grabbing System information
'Full List Found Here: https://docs.microsoft.com/en-us/windows/desktop/shell/shell-getsysteminformation
memory = ShellObj.GetSystemInformation("PhysicalMemoryInstalled")
dblclick = ShellObj.GetSystemInformation("DoubleClickTime")
ProcLevel = ShellObj.GetSystemInformation("ProcessorLevel")
procspeed = ShellObj.GetSystemInformation("ProcessorSpeed")

'Print out the details
Debug.Print "----------------------"
Debug.Print memory
Debug.Print dblclick
Debug.Print ProcLevel
Debug.Print procspeed
Debug.Print "----------------------"

'Print the status of a setting in this case, what is the state of hidden folders.
'False means it not set, True means it's set
'Full List Found Here: https://docs.microsoft.com/en-us/windows/desktop/shell/shell-getsetting
Debug.Print ShellObj.GetSetting(SSF_SHOWALLOBJECTS)
Debug.Print "----------------------"


'Define the name space, here I use a Special Folder Name Constant
'Full List Found Here: https://docs.microsoft.com/en-us/windows/desktop/api/Shldisp/ne-shldisp-shellspecialfolderconstants
Set ShellObjFldr = ShellObj.Namespace(ssfRECENT)

    'If we have a folder then continue
    If (Not ShellObjFldr Is Nothing) Then
        
        'Print out the folder title
        Debug.Print ShellObjFldr.Title
        
        'Print out the Parent Folder
        Debug.Print ShellObjFldr.ParentFolder
        
        'Grab the Items in the folder
        Set FldrItems = ShellObjFldr.Items()
        
        'Print the Item count
        Debug.Print FldrItems.Count
        
        'Loop through each item in the folder
        For Each FldrItem In FldrItems

            Debug.Print "----------------------"

            'Print the path.
            Debug.Print FldrItem.Path
            
            'Print last modified
            Debug.Print FldrItem.ModifyDate

            'Print Item Type
            Debug.Print FldrItem.Type

            'Print Item Size
            Debug.Print FldrItem.Size
        Next
        
        'Grab Details of the item, doesn't seem to work.
        Debug.Print ShellObjFldr.GetDetailsOf(vItem:=FldrItem, iColumn:=2)
        
    End If

'Declare some more variables
Dim AllWindows As Object
Dim IEObject As InternetExplorer

'Grab all the open windows
Set AllWindows = ShellObj.Windows()

    'Count all the windows
    Debug.Print AllWindows.Count
    
    'Grab the window
    Set IEObject = AllWindows.Item
        Debug.Print IEObject.Path

End Sub

Sub IShellDispatch()

Dim ShellDispatch As Shell32.Shell
Set ShellDispatch = New Shell32.Shell

    'Display the file run
    ShellDispatch.FileRun
    
    'Display the folder of my choice, either a special folder or a folder I specified
    ShellDispatch.Explore vDir:=ssfFONTS
    ShellDispatch.Explore vDir:="C:\Users\Alex\OneDrive\Growth - Tutorial Videos"
    
    'Open a control panel item, very limited
    ShellDispatch.ControlPanelItem bstrDir:="desk.cpl"
  
    'Kinda works
    ShellDispatch.Help
    
    'Cascade the windows.
    ShellDispatch.CascadeWindows
    
    ShellDispatch.Open ("C:\Program Files\IronPython 2.7\ipy.exe")
        
End Sub
