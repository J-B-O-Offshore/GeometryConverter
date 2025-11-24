Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("PY data file could not be read, make sure it is the right format and it is reachable.", 64, "Message")
    End Function
    
