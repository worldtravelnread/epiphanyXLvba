Attribute VB_Name = "mod_B7_ID_1st_dupes"

Sub B7_ID_1st_dupes()

    'This is part of the sequence of procedures
    'to identify duplicate records and put the
    'duplicate records in a new sheet.
    
    'This procedure creates a new column to
    'identify the 1st duplicates in the list
    'of duplicates that is created with the
    'prior procedures.
    'This is used for troubleshooting
    'the query statement.
    
    Dim lngColCt As Long            'the number of columns in the current region
    Dim lngRowCt As Long            'the number of rows in the current region
    Dim intColTarget As Integer     'the number of the 1st blank column
    Dim strStartCellAddr As String  'the address of the starting cell to format & copy
    Dim varDupForm As Variant       'the formula for duplicate conditional format
    Dim varUniqForm As Variant      'the formula for the unique conditional format
    Dim rngStartCell As Range       'the starting cell for the formula
    Dim strEndRowAddr As String     'the address of the last row in the column
    Dim strDestRng As String        'the range destination to autofill
    Dim strCountIfForm As String    'the formula to count the number of duplicates
    Dim lngDupeCt As Long           'the number of duplicates in the sheet
    
    
    'explicitly set the active cell to A1
    Cells(1, 1).Activate

    'calculate the number of columns in the current region
    lngColCt = ActiveCell.CurrentRegion.Columns.Count
    
    'calculate the number of rows in the current region
    lngRowCt = ActiveCell.CurrentRegion.Rows.Count
    
    'set the column number of the 1st blank column after the end of the current region
    intColTarget = lngColCt + 1
    
    'Activate the top cell in the 1st blank column
    Cells(1, intColTarget).Activate
    'set the title of the new column
    ActiveCell.Value = "1st Duplicate?"
    
    'set the range of the 1st cell to enter the formula
    'this should be the 2nd row of the new column
    'Activate the starting cell
    Set rngStartCell = ActiveCell.Offset(1, 0)
    rngStartCell.Activate
    
    'Get the address of the starting cell
    strStartCellAddr = ActiveCell.Address(rowabsolute:=False)
    
    'get the address of the last row in the new column
    strEndRowAddr = Cells(lngRowCt, intColTarget).Address(rowabsolute:=False)
    
    'set the variable for the autofill destination range
    strDestRng = strStartCellAddr & ":" & strEndRowAddr
    
    'Calculate whether the value of concatenated for the current record equals
    'the value of the concatenated record above
    'uses relative references in the formula
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]<>R[-1]C[-2],""1st duplicate"",""repeat"")"
    
    
    'Select the active cell and add the conditional format
    'set the conditional format formula if the value is duplicate
    ActiveCell.Select
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$q2=""repeat"""
    
    'add the conditional format
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    'if duplicate, the font is black
    With Selection.FormatConditions(1).Font
        .Color = RGB(0, 0, 0)            'black
        .TintAndShade = 0
    End With
    'if duplicate, the fill is yellow
    With Selection.FormatConditions(1).Interior
        .Color = RGB(255, 255, 0)               'yellow
        .TintAndShade = 0
    End With
    'do not stop evaluating condition if true
    Selection.FormatConditions(1).StopIfTrue = False
    
    
    'set the conditional format formula if the value is unique
    Selection.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=$q2=""1st duplicate"""
    
    'add the conditional format
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    'if unique, the font is black
    With Selection.FormatConditions(1).Font
        .Color = RGB(0, 0, 0)            'black
        .TintAndShade = 0
    End With
    'if unique, the fill is teal
    With Selection.FormatConditions(1).Interior
        .Color = RGB(0, 255, 200)                  'teal
        .TintAndShade = 0
    End With
    'do not stop evaluating condition if true
    Selection.FormatConditions(1).StopIfTrue = False
    
    'activate the starting cell
    rngStartCell.Activate
    ActiveCell.Select
    
    'Autofill the formula and conditional formats to the rest of the rows
    Selection.AutoFill Destination:=Range(strDestRng)
    Range(strDestRng).Select
    
    'activate the starting cell
    rngStartCell.Activate
    
    'format the column to fit width
    ActiveCell.Columns.EntireColumn.AutoFit
    
    'get the number of duplicates in the sheet
    lngDupeCt = ActiveCell.Application.WorksheetFunction.CountIf(Range(strDestRng), "1st duplicate")


End Sub
