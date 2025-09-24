Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("PY data file could not be read or FLS_(Reloading_BE) not part of the file, make shure it is the right format and it is reachable.", 64, "Message")
    End Function
    
