Attribute VB_Name = "mod_C_document_0_len_strings"
Sub C_Document_0_len_strings()

    'This procedure checks the data for
    'fields with zero-length strings: ""
    'It creates new columns for each existing
    'column and creates formulas for each column
    'to check for zero-length strings.
    
    'A dynamic array is created to put the new
    'zero-length string column titles
    'and count of zero-length string records
    'in the notes sheet


    Dim intTargetCol As Integer
    Dim intNextTargetCol As Integer
    Dim intLoopCounter As Integer
    Dim strZeroLenStr As String
    Dim intSourceCol As Integer
    Dim intNextSourceCol As Integer
    Dim strSourceText As String
    Dim strSourceAddr As String
    Dim strZeroLenStrCkFormula As String
    Dim intLastZeroLenStrCol As Integer
    Dim strLastZeroLenStrColAddr As String
    Dim strFirstZeroLenStrColAddr As String
    Dim lngLastRow As Long
    Dim lngLastCol As Long
    Dim rng1stTargetCol As Range
    Dim strTargetColNm As String
    Dim str1stTargetColTitleAddr As String
    Dim strLastTargetColTitleAddr As String
    Dim strZeroTitleRngNm As String
    Dim lngZeroColCt As Long
    Dim intColCounter As Integer
    Dim lngRowCounter As Variant
    Dim strLastRowAddr As String
    Dim strZeroDataRngNm As String
    Dim lngZeroDataCt As Long
    Dim varZeroArr() As Variant                     'dynamic array
    Dim lngNotesRowCt As Long                       '# of rows in current region of notes sheet
    Dim lngNotesRowTarget As Long                    'row number to start entering data in notes sheet
    
    
    'activate the 1st cell
    Cells(1, 1).Activate
    
    'set the beginning of the title of the
    'new columns that will record the
    'zero-length strings for each field.
    strZeroLenStr = "Zero-Length String "
    
    'count the number of columns and rows in the region
    lngLastCol = ActiveCell.CurrentRegion.Columns.Count
    lngLastRow = ActiveCell.CurrentRegion.Rows.Count
        
    'get the name of the 1st column to be checked
    'for zero-length strings
    intSourceCol = ActiveCell.Column
    strSourceText = ActiveCell.Value
    
    'set the number of the 1st column to be added
    'it is the column immediately after the last
    'original column
    intTargetCol = lngLastCol + 1
    'define a range for the 1st target column
    Set rng1stTargetCol = Cells(1, intTargetCol)
     
    'loop through each column heading
    'and set the name for a new column that has
    'Zero-Length String and the original column name
    For intLoopCounter = intSourceCol To lngLastCol
        strSourceText = Cells(1, intLoopCounter).Value
        strTargetValue = strZeroLenStr & strSourceText
        
        'get the address of the 1st data cell in the source column
        strSourceAddr = Cells(2, intLoopCounter).Address(rowabsolute:=False, columnabsolute:=False)
        
        'set the formula to check if the 1st data cell in the source
        'column is a zero-length string
        strZeroLenStrCkFormula = "=if(" & strSourceAddr & "="""",""zero-length string"",""ok"")"
        '=(IF(A2="","empty string","ok"))
    
        'activate the 1st cell in the 1st blank column
        'and set the column title
        Cells(1, intTargetCol).Activate
        ActiveCell.Value = strTargetValue
        'activate the 1st data cell in the new column
        'and put in the formula to check for a zero-length string
        Cells(2, intTargetCol).Activate
        ActiveCell.Formula = strZeroLenStrCkFormula
                    
        'increment the target column number to
        'the next blank column
        intTargetCol = intTargetCol + 1
    
    'increment to the next source column
    Next intLoopCounter
    
    'set the last column of the added columns
    'this is the last column overall
    intLastZeroLenStrCol = ActiveCell.Column
    'get the address of the 1st data cell (row 2)
    'in the last added column
    strLastZeroLenStrColAddr = ActiveCell.Address(rowabsolute:=False, _
        columnabsolute:=False)
    
    'activate the 1st target column
    rng1stTargetCol.Activate
    'activate the 1st data cell (row 2)
    ActiveCell.Offset(1, 0).Activate
    'get the address of this data cell
    strFirstZeroLenStrColAddr = ActiveCell.Address(rowabsolute:=False, _
        columnabsolute:=False)
       
    'set the range of cells to be copied
    Range(strFirstZeroLenStrColAddr, strLastZeroLenStrColAddr).Select
    'autofill the formulas down to the last record
    Selection.AutoFill Destination:=Range(strFirstZeroLenStrColAddr, _
        Cells(lngLastRow, intLastZeroLenStrCol))
    
    'select the current region
    Cells(1, 1).CurrentRegion.Select
    
    'format text alignment
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
        
    'activate the 1st target column
    rng1stTargetCol.Activate
    'get the address of the title cell of the 1st
    'target column
    str1stTargetColTitleAddr = ActiveCell.Address(rowabsolute:=False, _
        columnabsolute:=False)
    
    'get the number of columns in the current range
    'this should include the newly added columns
    lngLastCol = ActiveCell.CurrentRegion.Columns.Count
    'activate the title cell of the last target column
    Cells(1, lngLastCol).Activate
    'get the address of the title cell of the last
    'target column
    strLastTargetColTitleAddr = ActiveCell.Address(rowabsolute:=False, _
        columnabsolute:=False)
    
    'set the name of the range of target cells with
    'column titles/headings
    strZeroTitleRngNm = str1stTargetColTitleAddr & ":" & _
        strLastTargetColTitleAddr
    
    'get the number of column headings that start with Zero
    'using the COUNTIF function
    '=COUNTIF(M1:X1,"Zero*")
    lngZeroColCt = ActiveCell.Application.WorksheetFunction.CountIf(Range(strZeroTitleRngNm), _
        "zero*")
        
    'reset the range to the title cell of the
    '1st target column
    rng1stTargetCol.Activate
   
    'set the intLoopCounter to 0
    lngRowCounter = 0
    'set the intTargetCol back to the 1st target column
    intTargetCol = ActiveCell.Column
   
    'Create an array of all the new column titles
    '(they start with Zero) in the 1st dimension
    'with a second dimension to COUNTIF any zero-length strings
    ReDim varZeroArr(0 To lngZeroColCt - 1, 0 To 1) As Variant
    
    For lngRowCounter = 0 To lngZeroColCt - 1
        'set the first dimension of the array to the
        'column title name
        varZeroArr(lngRowCounter, 0) = ActiveCell.Offset(0, lngRowCounter).Value
        'get the address of the 1st data cell (row 2) of this column
        strFirstZeroLenStrColAddr = ActiveCell.Offset(1, lngRowCounter).Address(rowabsolute:=False, _
            columnabsolute:=False)
        'get the address of the last data cell of this column
        strLastRowAddr = ActiveCell.Offset(lngLastRow, lngRowCounter).Address(rowabsolute:=False, _
            columnabsolute:=False)
        'set the name of the range of data cells in this column
        strZeroDataRngNm = strFirstZeroLenStrColAddr & ":" & strLastRowAddr
        'get the number of cells that indicate a zero-length string
        'using the COUNTIF function
        lngZeroDataCt = ActiveCell.Application.WorksheetFunction.CountIf(Range(strZeroDataRngNm), _
            "zero*")
        'set the second dimension of the array to
        'the number of cells with zero-length strings
        varZeroArr(lngRowCounter, 1) = lngZeroDataCt
              
    Next lngRowCounter
    
        
    'activate the notes sheet
    'activate the 1st cell
    Sheets("notes").Activate
    Cells(1, 1).Activate
    'get the number of rows in the current region
    'set the number of the target row - 1st blank row
    lngNotesRowCt = ActiveCell.CurrentRegion.Rows.Count
    lngNotesRowTarget = lngNotesRowCt + 1
    
    'activate the 1st target cell
    Cells(lngNotesRowTarget, 1).Activate
    
    'loop through the array and enter the column title in Column A
    'and the number of zero-length strings in Column B
    
    For lngRowCounter = 0 To lngZeroColCt - 1
        ActiveCell.Offset(lngRowCounter, 0).Value = varZeroArr(lngRowCounter, 0)
        ActiveCell.Offset(lngRowCounter, 1).Value = varZeroArr(lngRowCounter, 1)
    Next lngRowCounter
    
    'select the first two columns and autofit contents
    Columns("A:B").EntireColumn.AutoFit
    
   
   'show message that the procedure is complete
   MsgBox ("Finished identifying and counting zero-length strings.")


End Sub
