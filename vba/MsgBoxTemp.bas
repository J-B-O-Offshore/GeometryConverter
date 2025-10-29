Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("The MP and the TP are overlapping by 15.0m. Combine stiffness etc as grouted connection (yes) or add as skirt (no)?", 4, "Message")
    End Function
    
