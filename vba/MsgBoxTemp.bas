Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("RNA table contains duplicate identifiers. Please privide a unique identifier for each row. Aborting.", 64, "Message")
    End Function
    
