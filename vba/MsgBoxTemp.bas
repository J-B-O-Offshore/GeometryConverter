Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("Please fully populate the NEW Meta table to create a new DB entry or clear it of all data to overwrite the loaded structure.", 64, "Message")
    End Function
    
