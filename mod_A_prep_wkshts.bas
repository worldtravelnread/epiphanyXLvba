Attribute VB_Name = "mod_A_prep_wkshts"

Sub A_prep_wkshts()
Attribute A_prep_wkshts.VB_ProcData.VB_Invoke_Func = " \n14"
    
    'This procedure should be run 1st
    'It renames the sheet with the raw data
    'to "raw_data" and counts the number of rows and columns
    'It creates a new sheet named notes
    'and adds the current date and time, then copies
    'and pastes the date/time value as values so it
    'is a static value and doesn't automatically update
    'It adds information on the number of records
    'and fields in the "raw_data" sheet
    'It adds a sheet named "source_info" where the user should
    'enter information about the source of the data (e.g., SQL statement,
    'screenshots of the fields used in a report builder form)
    
    Dim strRawNm As String
    Dim lngRowCt As Long
    Dim lngColCt As Long
    Dim strNotesNm As String
    Dim rngTarget As Range
    Dim strSourceNm As String
    Dim strMsg As String
    
    
    'Select the active worksheet and rename it to "raw_data"
    ActiveSheet.Select
    strRawNm = "raw_data"
    ActiveSheet.Name = strRawNm
    
    'get the number of rows and columns
    Cells(1, 1).Activate
    lngRowCt = ActiveCell.CurrentRegion.Rows.Count
    lngColCt = ActiveCell.CurrentRegion.Columns.Count
    
    
    'Add a new sheet and name it notes
    Sheets.Add after:=Sheets(strRawNm)
    ActiveSheet.Select
    strNotesNm = "notes"
    ActiveSheet.Name = strNotesNm
    
    'add the current date & time to the notes sheet
    Cells(1, 1).Activate
    ActiveCell.Value = "Date / Time"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "=NOW()"
    'copy and paste the value of the date time
    'so it does not update when saved
    Set rngTarget = ActiveCell
    rngTarget.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, _
        operation:=xlNone, skipblanks:=False, _
        Transpose:=False
    
    'add the number of records (lngRowCt - 1) in the raw_data sheet
    'subtract 1 to account for the header row
    'apply number format with thousands separator (comma)
    Cells(2, 1).Activate
    ActiveCell.Value = "# of records in " & strRawNm & " sheet"
    With ActiveCell.Offset(0, 1)
        .Value = lngRowCt - 1
        .NumberFormat = "#,##0"
    End With
    
    'add the number of fields (lngColCt) in the orig sheet
    Cells(3, 1).Activate
    ActiveCell.Value = "# of fields in " & strRawNm & " sheet"
    ActiveCell.Offset(0, 1).Value = lngColCt
    
    'select the first two columns and autofit contents
    Columns("A:B").EntireColumn.AutoFit
            
    'Add a new sheet and name it "source_info"
    Sheets.Add after:=Sheets(Sheets.Count)
    ActiveSheet.Select
    strSourceNm = "source_info"
    ActiveSheet.Name = strSourceNm
    
    strMsg = "Finished preparing worksheets. Copy and paste information about the data source into the 'source_info' worksheet."
    MsgBox (strMsg)
    
End Sub
