Attribute VB_Name = "mod_B1_prep_ID_Duplicates_sht"

Sub B1_prep_ID_Duplicates_sht()

    'This is part of the sequence of procedures
    'to identify duplicate records and put the
    'duplicate records in a new sheet.
    
    'This procedure makes a copy of the sheet with the
    'source data and names it "ID_duplicates_[source name]"
       
    Dim strSourceSht As String
    Dim strTargetSht As String
    Dim strIndxNm As String
    Dim intIndxCol As Integer
    Dim strMsg As String
    
    'activate the 1st cell in the sheet
    Cells(1, 1).Activate
    
    'Get the name of the worksheet to be copied and
    'Copy the sheet and place the copy before the original sheet
    strSourceSh = ActiveSheet.Name
    ActiveSheet.Copy before:=Worksheets(strSourceSh)
    
    'Name the newly copied sheet
    strTargetSht = "ID_duplicates_" & strSourceSh
    ActiveSheet.Name = strTargetSht
    
    
    'Format the 1st row: Horizontal & vertical are centered, text is wrapped
    Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    'Set the column width of all columns in the range to 15
    Selection.CurrentRegion.Select
    Selection.ColumnWidth = 15
    
    'show the form with the list box that shows the column titles
    'so the user can pick which column should be the first one
    
    Call B1a_show_frmSelectIndex
    
    'Search for the column title selected in the form
    'activate the cell with the selected column title
    strIndxNm = strSelectedCol
    Selection.Find(what:=strIndxNm, after:=ActiveCell, LookIn:= _
        xlValues, lookat:=xlPart, searchorder:=xlByRows, searchdirection:=xlNext _
        , MatchCase:=False, searchformat:=False).Activate
    
    'get the number of the selected column
    intIndxCol = ActiveCell.Column
    
    'check to see if the selected column is already the
    'first column
    'if not, select the column then cut and
    'insert it as the first column
    'then align the values: horizontal & vertical to center
    
    If intIndxCol <> 1 Then
        Columns(intIndxCol).Select
        Selection.Cut
        Columns("A:A").Select
        Selection.Insert shift:=xlToRight
        With Selection
           .HorizontalAlignment = xlCenter
           .VerticalAlignment = xlCenter
        End With
        Selection.CurrentRegion.Select
    End If
    

End Sub
