import win32com.client as win32

# Grab the Active Instance of Word
WrdApp = win32.GetActiveObject("Word.Application")

# Grab the current document.
WrdDoc = WrdApp.ActiveDocument

# Reference the Table in it.
WrdTable = WrdDoc.Tables.Item(1)

# Grab all the columns
SaleColumn = WrdTable.Columns(1)
CostColumn = WrdTable.Columns(2)
ProfitColumn = WrdTable.Columns(3)


# Loop through each cell in the Sales Column.
for SaleCell in list(SaleColumn.Cells)[1:]:
    
    # Grab the Text
    SaleCellText = SaleCell.Range.Text
    
    # Clear out the old text
    SaleCell.Range.Text = ""

    # Create a Formula String
    formula_string = "={my_number}\#""$#,##0.00;($#,##0.00)""".format(my_number = SaleCellText)

    # Create the Range
    SaleCell.Range.Select()

    # Collapse the Range
    WrdApp.Selection.Collapse(Direction=1)

    # Define the new Selection Range
    SelecRng = WrdApp.Selection.Range

    # Set the Formula
    SelecRng.Fields.Add(Range=SelecRng, Type=-1, Text=formula_string, PreserveFormatting=True)


# Loop through each cell in the Cost Column.
for CostCell in list(CostColumn.Cells)[1:]:
    
    # Grab the Text
    CostCellText = CostCell.Range.Text

    # Clear the Original Text
    CostCell.Range.Text = ""

    # Create a Formula String
    formula_string = "={my_number}\#""$#,##0.00;($#,##0.00)""".format(my_number = SaleCellText)

    # Create the Range
    CostCell.Range.Select()

    # Collapse the Range
    WrdApp.Selection.Collapse(Direction=1)

    # Define the new Selection Range
    SelecRng = WrdApp.Selection.Range

    # Set the Formula
    SelecRng.Fields.Add(Range=SelecRng, Type=-1, Text=formula_string, PreserveFormatting=True)


# Loop through each cell in the Profit Column.
for ProfitCell in list(ProfitColumn.Cells)[1:]:

    # Clear the Original Text
    ProfitCell.Range.Text = ""

    # Create a Formula String
    formula_string = "=R{row_number}C1 - R{row_number}C2 \#""$#,##0.00;($#,##0.00)""".format(row_number = ProfitCell.Row.Index)

    # Create the Range
    ProfitCell.Range.Select()

    # Collapse the Range
    WrdApp.Selection.Collapse(Direction=1)

    # Define the new Selection Range
    SelecRng = WrdApp.Selection.Range

    # Set the Formula
    SelecRng.Fields.Add(Range=SelecRng, Type=-1, Text=formula_string, PreserveFormatting=True)
