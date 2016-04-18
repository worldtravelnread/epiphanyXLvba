Attribute VB_Name = "mod_X_add_named_range"

Sub X_addNamedRange()

    'This module creates a named range for the data
    'in a sheet.
    'This is useful for vlookup formulas.
    
    Dim strSourceRngNm As String
    Dim rngSourceRng As Range
    Dim strSourceWorksheetNm As String
    Dim lngLastCol As Long
    Dim lngLastRow As Long
    Dim strLastRow As String
    Dim strLastCol As String
    Dim strRefersTo As String
    Dim strMsg As String
    
    'set the range name to [activesheet name]_all
    strSourceWorksheetNm = ActiveSheet.Name
    strSourceRngNm = strSourceWorksheetNm & "_all"
    
    'get the last row number & the last column number
    lngLastRow = ActiveCell.SpecialCells(xlCellTypeLastCell).Row
    lngLastCol = ActiveCell.SpecialCells(xlCellTypeLastCell).Column
    
    'set the address of the last row and column
    strLastRow = "R" & lngLastRow
    strLastCol = "C" & lngLastCol
    
    'set the name of the range that the name refers to
    strRefersTo = strSourceWorksheetNm & "!R1C1:" & strLastRow & strLastCol
    
    Selection.CurrentRegion.Select
    
    'add the named range
    ActiveWorkbook.Names.Add Name:=strSourceRngNm, RefersToR1C1:="=" & strRefersTo

    'show a message box that the named range has been added
    strMsg = strSourceRngNm & " added as named range."
    MsgBox (strMsg)



End Sub

