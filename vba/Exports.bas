Attribute VB_Name = "Exports"
'==============================================================================
' JBOOST Path Selection Subroutines
'==============================================================================

Sub select_JBOOST_out()
    Dim prevValue As Variant
    
    PickFolderDialog "JBOOST_Path"
    
End Sub

Sub select_Bladad_py_curves()
    Dim prevValue As Variant
    
    PickFolderDialog "Bladed_py_path"
    
End Sub

Sub select_Bladad_py_curves_output()
    Dim prevValue As Variant
    
    PickFolderDialog "Bladed_py_export_path"
    
End Sub


Sub load_Bladed_dropdown()
    Dim lua_path As Variant
    Dim run As Boolean
    Dim args As New Collection
    
    path = Range("Bladed_py_path").Value
    success = FileExists(path)

    If Not success Then
        MsgBox "The Bladed file does not exist or is not reachable: " & path, vbExclamation, "Error"
        Exit Sub
    End If
    
    args.Add path

    RunPythonWrapper "export", "fill_bladed_py_dropdown", path
End Sub

Sub plot_Bladed_py_curves()
    Dim db_path As String
    Dim config_name As String
    Dim args As New Collection
    
    db_path = Range("Bladed_py_path").Value
    config_name = get_dropdown_value("ExportStructure", "Dropdown_Bladed_py_loadcase")
    
    success = FileExists(db_path)

    If Not success Then
        MsgBox "The Bladed file does not exist or is not reachable: " & path, vbExclamation, "Error"
        Exit Sub
    End If
    
    args.Add db_path
    args.Add config_name
    
    RunPythonWrapper "export", "plot_bladed_py", args
End Sub

Sub apply_py_curves()
    Dim Bladed_py_path As String
    Dim Bladed_py_export_path As String
    Dim config_name As String
    Dim args As New Collection
    
    Bladed_py_path = Range("Bladed_py_path").Value
    Bladed_py_export_path = Range("Bladed_py_export_path").Value
    config_name = get_dropdown_value("ExportStructure", "Dropdown_Bladed_py_loadcase")
    
    success = FileExists(Bladed_py_path)
    If Not success Then
        MsgBox "The Bladed file does not exist or is not reachable: " & path, vbExclamation, "Error"
        Exit Sub
    End If
    
    success = FolderExists(Bladed_py_export_path)
    If Not success Then
        MsgBox "The PJ output folder does not exist or is not reachable: " & path, vbExclamation, "Error"
        Exit Sub
    End If
    
    args.Add Bladed_py_path
    args.Add Bladed_py_export_path
    args.Add config_name
    
    RunPythonWrapper "export", "apply_bladed_py_curves", args
End Sub

Sub export_JBOOST()
    Dim lua_path As Variant
    Dim run As Boolean
    Dim args As New Collection
    
    jboost_path = Range("JBOOST_Path").Value
    
    args.Add jboost_path

    RunPythonWrapper "export", "export_JBOOST", args
End Sub


Sub export_run_JBOOST()
    Dim jboost_path As Variant
    Dim args As New Collection
    
    jboost_path = Range("JBOOST_Path").Value
    
    
    RunPythonWrapper "export", "run_JBOOST_excel", jboost_path
End Sub


Sub fill_JBOOST_auto_values()
    RunPythonWrapper "export", "fill_JBOOST_auto_excel"
End Sub


Sub run_JBOOST()
    Dim lua_path As Variant
    Dim run As Boolean
    Dim args As New Collection
    
    jboost_path = ""
    
    RunPythonWrapper "export", "run_JBOOST_excel", jboost_path
End Sub


Sub select_WLGen_out()
    Dim prevValue As Variant
    PickFolderDialog "WLGen_Path"
End Sub

Sub export_WLGen()
    Dim lua_path As Variant
    lua_path = Range("WLGen_Path").Value
    RunPythonWrapper "export", "export_WLGen", lua_path
End Sub

Sub fill_WLGenMasses()
    ClearTableContents "ExportStructure", "APPURTANCES"
    RunPythonWrapper "export", "fill_WLGenMasses"
End Sub

Sub fill_Bladed_table()
    ClearTableContents "ExportStructure", "Bladed_Nodes"
    ClearTableContents "ExportStructure", "Bladed_Elements"
    RunPythonWrapper "export", "fill_Bladed_table"
End Sub

Sub show_WLGen_section()
    ShowOnlySelectedColumns "E:BW", "E:O"
End Sub

Sub show_Bladed_section()
    ShowOnlySelectedColumns "E:BW", "S:AP"
End Sub

Sub show_JBOOST_section()
    ShowOnlySelectedColumns "E:BW", "AQ:BW"
End Sub

Sub open_PY_csv()
    OpenFileDialog "Bladed_py_path", "select PY cureve csv file", "csv files", "*.csv"
End Sub


