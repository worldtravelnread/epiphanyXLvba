Attribute VB_Name = "mod_B0_ID_dupes_master_proc"


Sub B0_ID_dupes_master_proc()

    'This procedure calls all the individual procedures
    'needed to identify duplicate records and place
    'the duplicate records in a new sheet.
    
    Dim strID_dupes_sht_nm As String
    Dim lngDuplicateCt As Long
    Dim strMsg As String
    
    Call B1_prep_ID_Duplicates_sht
    
    Call B2_concatenate_all_fields
    
    Call B3_6_customSortForDuplicates
    
    Call B4_ID_duplicates
    
    'get the name of the active sheet
    strID_dupes_sht_nm = ActiveSheet.Name
    
    'Go to the notes sheet and check if there are
    'duplicates
    Sheets("notes").Activate
    Cells(1.1).Select
    Selection.CurrentRegion.Select
    Selection.Find(what:="ID_duplicates", after:=ActiveCell, _
        LookIn:=xlValues, lookat:=xlPart, searchorder:=xlByRows, _
        searchdirection:=xlNext, MatchCase:=False, searchformat:=False).Activate
    'select the cell with the number of duplicates
    ActiveCell.Offset(0, 1).Activate
    lngDuplicateCt = ActiveCell.Value
    
    'if there are no duplicates, then exit the procedure
    If lngDuplicateCt = 0 Then
    
        MsgBox ("No duplicates identified.")
        Exit Sub
    Else
           
        'activate the ID duplicates sheet
        Sheets(strID_dupes_sht_nm).Activate
        
        Call B5_create_pvt_list_dupes
    
        Call B3_6_customSortForDuplicates
    
        Call B7_ID_1st_dupes

        strMsg = "Finished identifying duplicate records."
        MsgBox (strMsg)
    End If

End Sub
