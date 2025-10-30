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

Sub select_Bladad_py_curves_insert()
    Dim prevValue As Variant
    
    OpenFileDialog "Bladed_py_insert_path"
    
End Sub

Sub select_Bladad_py_curves_fig_insert()
    Dim prevValue As Variant
    
    PickFolderDialog "Bladed_py_insert_fig_path"
    
End Sub

Sub load_Bladed_dropdown()
    Dim path As String
    Dim run As Boolean
    Dim args As New Collection
    
    path = Range("Bladed_py_path").Value
    success = FileExists(path)
    
    ClearFormDropDown "ExportStructure", "Dropdown_Bladed_py_loadcase"
    DeleteFigure "ExportStructure", "Fig_FIG_PY_CURVES"
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


Sub load_JBOOST_soil_stiffness()
    Dim path As String
    Dim run As Boolean
    Dim args As New Collection
    
    path = Range("JBOOST_soil_path").Value
    
    ClearTableContents "ExportStructure", "JBOOST_soil_stiffness"
    
    success = FileExists(path)
    
    If Not success Then
        MsgBox "The soil stiffness matrix csv file does not exist or is not reachable: " & path, vbExclamation, "Error"
        Exit Sub
    End If
    
    args.Add path

    RunPythonWrapper "export", "load_JBOOST_soil_file", path
End Sub


Sub load_Bladed_soil_stiffness_mat()
    Dim pj_path As String
    Dim args As New Collection

    pj_path = Range("Bladed_soil_mat_path").Value
    
    ClearTableContents "ExportStructure", "Bladed_soil_stiffness_mat"
    
    success = FileExists(pj_path)
    
    If Not success Then
        MsgBox "The soil stiffness matrix csv file does not exist or is not reachable: " & pj_path, vbExclamation, "Error"
        Exit Sub
    End If

    
    RunPythonWrapper "export", "load_Bladed_soil_file_mat", pj_path
End Sub

Sub apply_soil_stiff_Bladed()
    Dim Bladed_stiff_path As String
    Dim Bladed_pj_export_path As String
    Dim config_name As String
    Dim args As New Collection
    
    Bladed_stiff_path = Range("Bladed_soil_mat_path").Value
    Bladed_pj_export_path = Range("Bladed_pj_file_stiff_mat_path").Value
    config_name = get_dropdown_value("ExportStructure", "Dropdown_Bladed_stiff_mat")
    
    success = FileExists(Bladed_stiff_path)
    If Not success Then
        MsgBox "The Bladed file does not exist or is not reachable: " & Bladed_stiff_path, vbExclamation, "Error"
        Exit Sub
    End If
    
    success = FileExists(Bladed_pj_export_path)
    If Not success Then
        MsgBox "The PJ output folder does not exist or is not reachable: " & Bladed_pj_export_path, vbExclamation, "Error"
        Exit Sub
    End If
    
        
    ClearTableContents "ExportStructure", "Bladed_Nodes"
    ClearTableContents "ExportStructure", "Bladed_Elements"
    RunPythonWrapper "export", "fill_Bladed_table"

    args.Add Bladed_stiff_path
    args.Add Bladed_pj_export_path
    args.Add config_name
    
    RunPythonWrapper "export", "apply_bladed_stiff_mat", args
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

Sub apply_py_curves_insert_PJ()
    Dim Bladed_py_path As String
    Dim Bladed_py_insert_path As String
    Dim Bladed_py_insert_fig_path As String
    Dim config_name As String
    Dim args As New Collection
    Dim success As Boolean
    Dim fso As Object
    
    Bladed_py_path = Range("Bladed_py_path").Value
    Bladed_py_insert_path = Range("Bladed_py_insert_path").Value
    Bladed_py_insert_fig_path = Range("Bladed_py_insert_fig_path").Value
    
    ' ? If Bladed_py_insert_fig_path is empty, use the folder of the insert file
    If Trim(Bladed_py_insert_fig_path) = "" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Bladed_py_insert_fig_path = fso.GetParentFolderName(Bladed_py_insert_path)
    End If
    
    config_name = get_dropdown_value("ExportStructure", "Dropdown_Bladed_py_loadcase")
    
    ' --- File existence checks ---
    success = FileExists(Bladed_py_path)
    If Not success Then
        MsgBox "The Bladed file does not exist or is not reachable: " & Bladed_py_path, vbExclamation, "Error"
        Exit Sub
    End If
    
    success = FileExists(Bladed_py_insert_path)
    If Not success Then
        MsgBox "The PJ output file does not exist or is not reachable: " & Bladed_py_insert_path, vbExclamation, "Error"
        Exit Sub
    End If
    
    ' ?? No FolderExists() check here — path can be a file or folder
    ' The Python function will handle both
    
    args.Add Bladed_py_path
    args.Add Bladed_py_insert_path
    args.Add config_name
    args.Add "True"
    args.Add Bladed_py_insert_fig_path
    args.Add "False"
    
    
    RunPythonWrapper "export", "apply_bladed_py_curves", args
End Sub

Sub export_JBOOST()
    Dim jboost_path As String
    Dim run As Boolean
    Dim args As New Collection
    
    jboost_path = Range("JBOOST_Path").Value
    
    
        ' --- File existence checks ---
    success = FolderExists(jboost_path)
    If Not success Then
        MsgBox "JBOOST folder does not exist or is not reachable: " & jboost_path, vbExclamation, "Error"
        Exit Sub
    End If
    
    
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

Sub fill_JBOOST_soil_configs()
    RunPythonWrapper "export", "create_JBOOST_soil_configs"
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
    Dim path As String
    path = Range("WLGen_Path").Value
    
    ' --- File existence checks ---
    success = FolderExists(path)
    If Not success Then
        MsgBox "WLgen folder does not exist or is not reachable: " & path, vbExclamation, "Error"
        Exit Sub
    End If
    
    RunPythonWrapper "export", "export_WLGen", path
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

Sub fill_Bladed_table_py()
    Dim args As New Collection
    Dim Bladed_py_path As String
    Dim config_name As String
    
    Bladed_py_path = Range("Bladed_py_path").Value
    config_name = get_dropdown_value("ExportStructure", "Dropdown_Bladed_py_loadcase")
    
    success = FileExists(Bladed_py_path)
    If Not success Then
        MsgBox "The Bladed file does not exist or is not reachable: " & Bladed_py_path, vbExclamation, "Error"
        Exit Sub
    End If
    
    ClearTableContents "ExportStructure", "Bladed_Nodes"
    ClearTableContents "ExportStructure", "Bladed_Elements"
    
    args.Add "True"
    args.Add config_name
    args.Add Bladed_py_path
    
    
    RunPythonWrapper "export", "fill_Bladed_table", args
End Sub


Sub show_WLGen_section()
    ShowOnlySelectedColumns "E:BW", "E:S"
End Sub

Sub show_Bladed_section()
    ShowOnlySelectedColumns "E:BW", "T:AQ"
End Sub

Sub show_JBOOST_section()
    ShowOnlySelectedColumns "E:BW", "AS:BW"
End Sub

Sub open_PY_csv()
    
    OpenFileDialog "Bladed_py_path", "select PY curve csv file", "csv files", "*.csv"
End Sub


Sub open_JBOOST_soil_csv()
    OpenFileDialog "JBOOST_soil_path", "select so csv file", "csv files", "*.csv"
End Sub

Sub open_BLADED_pj_file_stiff_mat()
    OpenFileDialog "Bladed_pj_file_stiff_mat_path", "Select %pj or prj file"
End Sub


Sub open_Bladed_soil_mat_csv()
    OpenFileDialog "Bladed_soil_mat_path", "select so csv file", "csv files", "*.csv"
End Sub
