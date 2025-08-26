Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("Error reading C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/MPTP_DP-B2_L0_G3_S2.xlsm TP and MP. Please make sure, the path leads to a valid MP_tool xlsm file and has the TP data on the 'Geometry' sheet 3 rows under the first 'Section' header, empty rows allowed. Error thrown by Python: (-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None).", 64, "Message")
    End Function
    
