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
    lua_path = Range("JBOOST_Path").Value
    RunPythonWrapper "export", "export_JBOOST", lua_path
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
    RunPythonWrapper "export", "fill_WLGenMasses"
End Sub



Sub fill_Bladed_table()
    RunPythonWrapper "export", "fill_Bladed_table"
End Sub
