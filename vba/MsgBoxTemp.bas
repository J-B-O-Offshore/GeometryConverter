Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("Failed to create or write to table ''." & vbNewLine & "Python Error: Empty table or column name specified", 64, "Message")
    End Function
    
