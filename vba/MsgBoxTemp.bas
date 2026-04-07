Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("Cannot save 'C:\Users\aaron.lange\Desktop\Projekte\Geometrie_Converter\JBOOSTReloaded_test\inclination_results.xlsx': Permission denied.", 64, "Message")
    End Function
    
