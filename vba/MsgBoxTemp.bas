Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("Please define a value for 'deflection TP'. If no deflection is required, please set it to 0. Aborting.", 64, "Message")
    End Function
    
