Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("Sesam file created:" & vbNewLine & "C:\temp\tutorial\output\Heike_.js" & vbNewLine & "" & vbNewLine & "Rows (tbl_Export_Sesam): 117 " & vbNewLine & "Rows (tbl_Export_Sesam_Mass): 73" & vbNewLine & "Preface lines (tbl_Export_Text): 7", 64, "Message")
    End Function
    
