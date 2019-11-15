Sub RunPythonScript()

'Declare Variables
Dim objShell As Object
Dim PythonExe, PythonScript As String

'Create a new Object shell.
Set objShell = VBA.CreateObject("Wscript.Shell")

'Provide file path to Python.exe
'USE TRIPLE QUOTES WHEN FILE PATH CONTAINS SPACES.
PythonExe = """C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\python.exe"""
PythonScript = "C:\Users\Alex\Desktop\ExcelToPowerPoint.py"

'Run the Python Script
objShell.Run PythonExe & PythonScript

End Sub
