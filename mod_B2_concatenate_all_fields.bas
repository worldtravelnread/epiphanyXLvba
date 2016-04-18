Attribute VB_Name = "mod_B2_concatenate_all_fields"
Sub B2_concatenate_all_fields()

    'This is part of the sequence of procedures
    'to identify duplicate records and put the
    'duplicate records in a new sheet.
    
    'This procedure creates a new column.
    'The cells in the column have a concatenate formula
    'that strings together all the values of all the fields
    'in the spreadsheet.
    
    Dim lngColCt As Long
    Dim strCellAddr As String
    Dim strConcatFormula As String
    Dim intCounter As Integer
    Dim rngFormula As Range
    Dim intFormLeng As Integer
    Dim strFormula As String
    Dim lngRowCt As Long
    Dim rngTarget As Range
    
    'activate the first cell
    Cells(1, 1).Activate

    'Count the number of columns in the current region
    lngColCt = ActiveCell.CurrentRegion.Columns.Count
    
    'Count the number of rows in the current region
    lngRowCt = ActiveCell.CurrentRegion.Rows.Count

    'Select the first cell in row 2
    'Set the intCounter to 1 to start with the 1st column
    'Set the open parentheses for the concatenate formula
    Cells(2, 1).Select
    intCounter = 1
    strConcatFormula = "("
    
    'loop through each field in row 2 to record the cell address
    'this will build the expression for the concatenate formula
    For intCounter = 1 To lngColCt
        strCellAddress = Cells(2, intCounter).Address(False)
        strConcatFormula = strConcatFormula & strCellAddress & ","
    Next intCounter
    
        
    'Select the first blank cell after the last record in the 1st row of data
    Set rngFormula = Cells(2, lngColCt + 1)
    rngFormula.Select
    
    'get the length of strConcatFormula
    intFormLeng = Len(strConcatFormula)
    
    'set strConCatFormula so that the last comma is not included
    'and add the close parentheses
    strConcatFormula = Left(strConcatFormula, intFormLeng - 1) & ")"
    
    'create the concatenate function
    strFormula = "=Concatenate" & strConcatFormula
      
    'Set the cell formula
    rngFormula.Formula = strFormula
    
    'set the range to be autofilled
    Set rngTarget = Range(rngFormula.Offset, rngFormula.Offset(lngRowCt - 2, 0))
    
    'autofill the formula down to the last row
    Selection.AutoFill Destination:=rngTarget
    
    'Select the first cell in the column
    'And label it Concatenated
    rngFormula.Offset(-1, 0).Select
    Selection.Value = "Concatenated"
          

End Sub

