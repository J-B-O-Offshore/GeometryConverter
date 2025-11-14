Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("Seabed level has to be below the structure top and above or at the structure bottom. Aborting", 64, "Message")
    End Function
    
