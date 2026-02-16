Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("Please provide a numeric input for the 'Height difference to MSL' in the Global Parameters. Aborting", 64, "Message")
    End Function
    
