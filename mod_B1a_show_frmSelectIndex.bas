Attribute VB_Name = "mod_B1a_show_frmSelectIndex"
Public strSelectedCol As String
    
Sub B1a_show_frmSelectIndex()


    
    'show the form to select the column to index
    frmSelectIndex.Show
       
    'get the value of the selection
    strSelectedCol = frmSelectIndex.lstSelectColumn.Value

End Sub
