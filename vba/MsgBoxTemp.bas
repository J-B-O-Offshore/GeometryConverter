Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("Height references are the same or not defined. (MP: None, TP: LAT, TOWER: None).", 64, "Message")
    End Function
    
