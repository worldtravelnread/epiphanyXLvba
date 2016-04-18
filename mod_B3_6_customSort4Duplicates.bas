Attribute VB_Name = "mod_B3_6_customSort4Duplicates"
Sub B3_6_customSortForDuplicates()
Attribute B3_6_customSortForDuplicates.VB_ProcData.VB_Invoke_Func = " \n14"

    'This is part of the sequence of procedures
    'to identify duplicate records and put the
    'duplicate records in a new sheet.
    
    'This procedure creates a custom sort made up of all the
    'columns in the spreadsheet, except the last one.
    'This procedures ensures the data is sorted correctly so that
    'the maximum number of duplicates is identified.
    'This procedure is the equivalent of manually creating a custom
    'sort using the option from the Home ribbon for all the fields
    'in the spreadsheet.


    Dim lngColCt As Long
    Dim intSortCol As Integer
    Dim rngStartCell As Range
    Dim intStartCol As Integer
    Dim intColCounter As Integer
    Dim strColNm As String
    Dim strShtSortStartAddr As String
    Dim strShtSortEndAddr As String
    Dim lngRowCt As Long
    Dim strShtSortAddr As String    'text string to enter as  the worksheet sort range
    Dim strColSortStartAddr As String
    Dim strColSortEndAddr As String
    Dim strColSortAddr As String
    Dim lngSortFieldsCt As Long        'number of sort fields in the custom sort
    
    'Set the active cell to cell A1
    Set rngStartCell = Cells(1, 1)
    rngStartCell.Activate
    
    'get the column number of the starting column
    intStartCol = ActiveCell.Column
    
    'Count the number of columns in the current region
    lngColCt = ActiveCell.CurrentRegion.Columns.Count
    
    'Count the number of rows in the current region
    lngRowCt = ActiveCell.CurrentRegion.Rows.Count
    
    'set the beginning cell of the sort range to the 1st cell in the 1st row,
    'the sort method below will explicitly state there is a header row
    strShtSortStartAddr = Cells(1, 1).Address
    
    'set the end cell of the sort range to the last cell in the current region
    strShtSortEndAddr = Cells(lngRowCt, lngColCt - 1).Address
    
    'set the string for the worksheet's sort range
    strShtSortAddr = strShtSortStartAddr & ":" & strShtSortEndAddr
    
    'Clear any existing sort fields
    ActiveSheet.Sort.SortFields.Clear
        
    'Loop through the column titles in the current region,
    'Except the concatenated column
    For intColCounter = intStartCol To lngColCt - 1
        'Get the value of each column title/header
        strColNm = Cells(1, intColCounter).Value
        
        'Set the address of the 1st cell in the column's sort range
        'This is row 2 of the column, the first data row
        strColSortStartAddr = Cells(2, intColCounter).Address
        
        'Set the address of the last cell in the column's sort range
        strColSortEndAddr = Cells(lngRowCt, intColCounter).Address
        
        'Set the A1C1 address of the column sort range
        strColSortAddr = strColSortStartAddr & ":" & strColSortEndAddr
        
        
        'Add a sort field for the current column in the loop
        ActiveSheet.Sort.SortFields.Add Key:=Range( _
            strColSortAddr), SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortTextAsNumbers
           
        'Count the number of sort fields
        'should be the same number as intColCounter, this is a check step
        lngSortFieldsCt = ActiveSheet.Sort.SortFields.Count
           
    Next intColCounter
    
    'Apply the custom sort to the entire worksheet
    With ActiveSheet.Sort
        .SetRange Range(strShtSortAddr)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With


End Sub
