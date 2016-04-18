Attribute VB_Name = "mod_B5_create_pvt_list_dupes"
Sub B5_create_pvt_list_dupes()
Attribute B5_create_pvt_list_dupes.VB_ProcData.VB_Invoke_Func = " \n14"
    
    'This is part of the sequence of procedures
    'to identify duplicate records and put the
    'duplicate records in a new sheet.
    
    'This procedure creates a pivot table to
    'identify the count of duplicates in the working
    'worksheet.
    'It also creates the detail sheet of all the duplicate values.
    
    'Check steps:
    'Make another copy of the working worksheet,
    'and delete the Duplicate? column.
    'Then use the Remove Duplicates menu item on the Data ribbon.
    'Compare the number of duplicates displayed by Excel to
    'the number shown on the pivot table.
    'If the Excel number is higher, make sure you have
    'run the customSortForDuplicates procedure.
    
    
    Dim rngStartCell As Range               'the starting cell for the pivot table range
    Dim lngRowCt As Long                    'the number of rows in the current region
    Dim lngColCt As Long                    'the number of columns in the current region
    Dim varPvtShtNm As Variant              'the name of the pivot table sheet
    Dim strDupePvtFld As String             'the name of the field to be pivoted
    Dim strSourceShtNm As String            'name of the source worksheet to be pivoted
    Dim varSourceDataRngNm As Variant       'the name of the SourceData range for the pivot table
    Dim strStartCellAddr As String          'the address of the starting cell
    Dim strEndCellAddr As String            'the address of the ending cell
    Dim rngTargetCell As Range              'the range of the target pivot table cell
    Dim varDestinNm As Variant              'the name of the destination range
    Dim strTargetCellAddr As String         'the address of the target cell
    Dim strPvtPgFldColNm As String          'the name of the page field to be pivoted
    Dim strPvtValFldColNm As String         'the name of the field for which to count values
    Dim strPvtValFieldCountNm As String     'the name of the count field in the pivot table
    Dim rngPvtValue As Range                'the range of the pivot table value/results
    Dim strDetailShtNm As String            'the name of the detail sheet created from the pivot table
    
    
    'Explicitly set the active cell to A1
    'set A1 as the starting cell for the pivot table range
    Cells(1, 1).Activate
    Set rngStartCell = ActiveCell
    
    'put the Column 1 header text in
    'the pivot table value name variable
    strPvtValFldColNm = ActiveCell.Value
    
    'get the address of the starting cell
    strStartCellAddr = ActiveCell.Address(rowabsolute:=False, _
        columnabsolute:=False, ReferenceStyle:=xlA1)
    
    'Calculate the number of rows in the current region
    lngRowCt = ActiveCell.CurrentRegion.Rows.Count
    
    'Calculate the number of columns in the current region
    lngColCt = ActiveCell.CurrentRegion.Columns.Count
    
    'put the Duplicates? column heading in
    'the pivot page field name variable
    strPvtPgFldColNm = Cells(1, lngColCt).Value
    
    
    'get the address of the ending cell
    strEndCellAddr = Cells(lngRowCt, lngColCt).Address(rowabsolute:=False, _
        columnabsolute:=False, ReferenceStyle:=xlA1)
    
    'set the name of the source worksheet
    strSourceShtNm = ActiveSheet.Name
    
    'define the name of the SourceData
    varSourceDataRngNm = strSourceShtNm & "!" & strStartCellAddr _
        & ":" & strEndCellAddr
    
    
    'add the new sheet for the pivot table
    Sheets.Add before:=Sheets(strSourceShtNm)
    'set the name of the new sheet
    varPvtShtNm = "pvt_" & strSourceShtNm
    ActiveSheet.Name = varPvtShtNm
    
    
    'select A3 as the starting cell for the pivot table & set the range
    Cells(3, 1).Activate
    Set rngTargetCell = ActiveCell
    'set the name of the target cell
    strTargetCellAddr = ActiveCell.Address(rowabsolute:=False, _
        columnabsolute:=False, ReferenceStyle:=xlA1)
    
    'set the name of the TableDestination
    varDestinNm = varPvtShtNm & "!" & strTargetCellAddr
    
    'create the pivot table
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        varSourceDataRngNm, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:=rngTargetCell, TableName:=varPvtShtNm, DefaultVersion _
        :=xlPivotTableVersion14
                 
    'set the pivot page field to Duplicates?
    With ActiveSheet.PivotTables(varPvtShtNm).PivotFields(strPvtPgFldColNm)
        .Orientation = xlPageField
        .Position = 1
    End With
    
    'set the name of the count field
    strPvtValFieldCountNm = "Count of " & strPvtValFldColNm
    
    'add the items to be counted by the pivot table
    ActiveSheet.PivotTables(varPvtShtNm).AddDataField ActiveSheet.PivotTables( _
        varPvtShtNm).PivotFields(strPvtValFldColNm), strPvtValFieldCountNm, _
        xlCount
        
    'format the count value
    ActiveSheet.PivotTables(varPvtShtNm).PivotFields(strPvtValFieldCountNm) _
        .NumberFormat = "#,##0"
    
    'clear the page field filter
    ActiveSheet.PivotTables(varPvtShtNm).PivotFields(strPvtPgFldColNm). _
        ClearAllFilters
        
    'set the page field to show duplicates
    ActiveSheet.PivotTables(varPvtShtNm).PivotFields(strPvtPgFldColNm). _
        CurrentPage = "duplicate"
        
    'activate the target cell
    rngTargetCell.Activate
    
    'activate & select the cell below where the pivot table value/result is
    ActiveCell.Offset(1, 0).Activate
    Set rngPvtValue = ActiveCell
    rngPvtValue.Select
    
    
    'show the details behind the pivot table value
    'this is the same as double-clicking on the value
    'it will create a new sheet with all the records with the value
    Selection.ShowDetail = True
    
    'set the name of the detail sheet
    strDetailShtNm = "duplicates_only"
    ActiveSheet.Name = strDetailShtNm
    
    
End Sub
