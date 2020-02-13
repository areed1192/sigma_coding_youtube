import win32com.client

# Grab the Active Instance of Excel.
ExcelApp = win32com.client.GetActiveObject("Excel.Application")

# Grab a workbook called "HasMyMacros". Technically you don't have to do this, but for demonstration purposes it helps out.
xlWorkbook = ExcelApp.Workbooks("HasMyMacros.xlsm")

# Execute the Macro "PopulateTwoCellsWithArguments", pass through the arguments ["Bob","Smith"]
ExcelApp.Run("PopulateTwoCellsWithArguments", "Bob", "Smith")

# Execute the Macro "PopulateTwoCells", which doesn't require any arguments.
ExcelApp.Run("PopulateTwoCells")

# ---------------------------------------------------------------------------------
# VBA CODE - PopulateTwoCellsWithArguments
# ---------------------------------------------------------------------------------

# Sub PopulateTwoCellsWithArguments(FirstName As String, LastName As String)

# Dim Cell1 As Range
# Dim Cell2 As Range

# Set Cell1 = Sheet1.Range("C5")
# Set Cell2 = Sheet1.Range("C6")

#     Cell1.Value = FirstName
#     Cell2.Value = LastName

# End Sub

# ---------------------------------------------------------------------------------
# VBA CODE - PopulateTwoCells
# ---------------------------------------------------------------------------------

# Sub PopulateTwoCells()

# Dim Cell1 As Range
# Dim Cell2 As Range

# Set Cell1 = Sheet1.Range("C2")
# Set Cell2 = Sheet1.Range("C3")

#     Cell1.Value = 3000
#     Cell2.Value = 4000

# End Sub
