VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectIndex 
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmSelectIndex.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



    

Private Sub cmdSelect_Click()

    'hide the form
    frmSelectIndex.Hide


End Sub

Private Sub lstSelectColumn_AfterUpdate()

    
    'get the value of the selecttion
    strSelectedCol = lstSelectColumn.Value

    

End Sub

Private Sub UserForm_Initialize()

    Dim lngColCt As Long
    Dim lngColCounter As Long
    Dim varColTitleArr As Variant
    Dim varListIndex As Variant
    Dim lngColTemp As Long
    Dim strColTempAddr As String
    Dim strSelectedCol As String
    
    
    'clear the listbox
    lstSelectColumn.Clear
    
    'activate 1st cell & count # columns in current region
    Cells(1, 1).Activate
    lngColCt = ActiveCell.CurrentRegion.Columns.Count
    
    'set the column number of the temporary cell to
    'the first blank cell after the titles
    'lngColTemp = lngColCt + 1
    
    'set the address of the temporary cell
    'strColTempAddr = Cells(1, lngColTemp).Address(rowabsolute:=False, _
        columnabsolute:=False)
    
    'set the lngColCounter for the For...Next loop to 0
    lngColCounter = 0
    
    'create an array with the names of all the
    'column titles
    
    ReDim varColTitleArr(0 To lngColCt - 1) As Variant
    
    For lngColCounter = 0 To lngColCt - 1
        'set the value of the array to the column title
        varColTitleArr(lngColCounter) = ActiveCell.Offset(0, lngColCounter).Value
    Next lngColCounter
    
    'clear the listbox
    lstSelectColumn.Clear
    'populate the listbox with the array
    lstSelectColumn.List() = varColTitleArr
    
    

End Sub
