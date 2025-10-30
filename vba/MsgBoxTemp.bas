Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("Please provide a seabed level in the Global Parameters. Aborting", 64, "Message")
    End Function
    
