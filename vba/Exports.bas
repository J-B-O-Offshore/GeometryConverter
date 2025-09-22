Attribute VB_Name = "Exports"
'==============================================================================
' JBOOST Path Selection Subroutines
'==============================================================================

Sub select_JBOOST_out()
    Dim prevValue As Variant
    
    PickFolderDialog "JBOOST_Path"
    
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
    ShowOnlySelectedColumns "E:BW", "S:AK"
End Sub

Sub show_JBOOST_section()
    ShowOnlySelectedColumns "E:BW", "AO:BW"
End Sub
