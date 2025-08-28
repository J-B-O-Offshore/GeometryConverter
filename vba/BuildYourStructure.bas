Attribute VB_Name = "BuildYourStructure"

'==============================================================================
' Database Path Selection Subroutines
'==============================================================================

Sub load_MP_dialog()
    OpenFileDialog "TextBox_MP_db_path", "Select a MP database file", "sql lite database", "*.db"
    load_MP_DB
End Sub

Sub load_TP_dialog()
    OpenFileDialog "TextBox_TP_db_path", "Select a TP database file", "sql lite database", "*.db"
    load_TP_DB
End Sub

Sub load_TOWER_dialog()
    OpenFileDialog "TextBox_TOWER_db_path", "Select a TOWER database file", "sql lite database", "*.db"
    load_TOWER_DB
End Sub


Sub load_RNA_dialog()
    OpenFileDialog "TextBox_RNA_db_path", "Select a RNA database file", "sql lite database", "*.db"
    load_RNA_DB
    save_RNA_Data
End Sub

Sub load_MP_DB()
    Dim prevValue As Variant
    Dim db_path As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BuildYourStructure")
    db_path = ws.Range("TextBox_MP_db_path").Value

    
    ClearTableContents "BuildYourStructure", "MP_DATA_TRUE"
    ClearTableContents "BuildYourStructure", "MP_DATA"
    ClearTableContents "BuildYourStructure", "MP_META_TRUE"
    ClearTableContents "BuildYourStructure", "MP_META"
    ClearTableContents "BuildYourStructure", "MP_META_FULL"
    ClearTableContents "BuildYourStructure", "MP_META_NEW", 1, 6
    ClearTableContents "BuildYourStructure", "MP_MASSES_TRUE"
    ClearTableContents "BuildYourStructure", "MP_MASSES"
    ClearFormDropDown "BuildYourStructure", "Dropdown_MP_Structures2"

    If Not CheckPath(db_path, "db") Then
        Exit Sub
    End If
        
    load_MP_META

End Sub


Sub load_TP_DB()

    Dim db_path As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BuildYourStructure")
    db_path = ws.Range("TextBox_TP_db_path").Value

        
    ClearTableContents "BuildYourStructure", "TP_DATA_TRUE"
    ClearTableContents "BuildYourStructure", "TP_DATA"
    ClearTableContents "BuildYourStructure", "TP_META_TRUE"
    ClearTableContents "BuildYourStructure", "TP_META"
    ClearTableContents "BuildYourStructure", "TP_META_FULL"
    ClearTableContents "BuildYourStructure", "TP_META_NEW", 1, 6
    ClearTableContents "BuildYourStructure", "TP_MASSES_TRUE"
    ClearTableContents "BuildYourStructure", "TP_MASSES"
    ClearFormDropDown "BuildYourStructure", "Dropdown_TP_Structures2"
        
    If Not CheckPath(db_path, "db") Then
        Exit Sub
    End If
    
    load_TP_META
    

End Sub

Sub load_RNA_DB()
    Dim db_path As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BuildYourStructure")
    db_path = ws.Range("TextBox_RNA_db_path").Value

    ClearTableContents "BuildYourStructure", "RNA_DATA_TRUE"
    ClearTableContents "BuildYourStructure", "RNA_DATA"
    ClearFormDropDown "BuildYourStructure", "Dropdown_RNA_Structures"
    
    If Not CheckPath(db_path, "db") Then
        Exit Sub
    End If

    load_RNA_DATA

End Sub


Sub load_TOWER_DB()

    Dim db_path As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BuildYourStructure")
    db_path = ws.Range("TextBox_TOWER_db_path").Value

        
    ClearTableContents "BuildYourStructure", "TOWER_DATA_TRUE"
    ClearTableContents "BuildYourStructure", "TOWER_DATA"
    ClearTableContents "BuildYourStructure", "TOWER_META_TRUE"
    ClearTableContents "BuildYourStructure", "TOWER_META"
    ClearTableContents "BuildYourStructure", "TOWER_META_FULL"
    ClearTableContents "BuildYourStructure", "TOWER_META_NEW", 1, 6
    ClearTableContents "BuildYourStructure", "TOWER_MASSES_TRUE"
    ClearTableContents "BuildYourStructure", "TOWER_MASSES"
    ClearFormDropDown "BuildYourStructure", "Dropdown_TOWER_Structures2"
    
    If Not CheckPath(db_path, "db") Then
        Exit Sub
    End If
    
    load_TOWER_META
    

End Sub

'==============================================================================
' Database Metadata Loading Subroutines
' Purpose: Load metadata from different databases (MP, TP, TOWER) by reading
'          the database path from a named range and calling corresponding
'          Python functions through RunPythonWrapper.
'==============================================================================


Sub load_MP_META()
    Dim db_path As Variant
    db_path = Range("TextBox_MP_db_path").Value
    RunPythonWrapper "db_handling", "load_MP_META", db_path
End Sub

    
Sub load_TP_META()
    Dim db_path As Variant
    db_path = Range("TextBox_TP_db_path").Value
    RunPythonWrapper "db_handling", "load_TP_META", db_path
End Sub

Sub load_TOWER_META()
    Dim db_path As Variant
    db_path = Range("TextBox_TOWER_db_path").Value
    RunPythonWrapper "db_handling", "load_TOWER_META", db_path
End Sub

'==============================================================================
' Database Data Loading Subroutines
' Purpose: Load specific data from MP, TP, and TOWER databases by reading
'          the structure name from dropdowns and the database path, then
'          calling Python functions with both arguments.
'==============================================================================

Sub load_MP_Data()
    Dim Structure_names As String, db_path As String
    Dim args As New Collection
    
    db_path = Range("TextBox_MP_db_path").Value
    Structure_names = get_dropdown_value("BuildYourStructure", "Dropdown_MP_Structures2")
    
    args.Add Structure_names
    args.Add db_path
    
    RunPythonWrapper "db_handling", "load_MP_DATA", args
End Sub

Sub load_TP_Data()
    Dim Structure_names As String, db_path As String
    Dim args As New Collection
    
    db_path = Range("TextBox_TP_db_path").Value
    Structure_names = get_dropdown_value("BuildYourStructure", "Dropdown_TP_Structures2")
    
    args.Add Structure_names
    args.Add db_path
    
    RunPythonWrapper "db_handling", "load_TP_DATA", args
End Sub

Sub load_TOWER_Data()
    Dim Structure_names As String, db_path As String
    Dim args As New Collection
    
    db_path = Range("TextBox_TOWER_db_path").Value
    Structure_names = get_dropdown_value("BuildYourStructure", "Dropdown_TOWER_Structures2")
    
    args.Add Structure_names
    args.Add db_path
    
    RunPythonWrapper "db_handling", "load_TOWER_DATA", args
End Sub

Sub load_RNA_DATA()
    Dim db_path As Variant
    db_path = Range("TextBox_RNA_db_path").Value
    RunPythonWrapper "db_handling", "load_RNA_DATA", db_path
End Sub
'==============================================================================
' Database Data Saving Subroutines
' Purpose: Save data for selected structures in MP, TP, and TOWER databases by
'          passing database path and structure name to Python save functions.
'          Updates dropdowns and clears related tables as needed.
'==============================================================================

' Save data for selected MP structure and update UI accordingly
Sub save_MP_Data()
    Dim args As New Collection, db_path As String, selected_structure As String, structure_name As String
    
    db_path = Range("TextBox_MP_db_path").Value
    selected_structure = get_dropdown_value("BuildYourStructure", "Dropdown_MP_Structures2")
    
    args.Add db_path
    args.Add selected_structure
    RunPythonWrapper "db_handling", "save_MP_data", args
    
End Sub

Sub save_TP_Data()
    Dim args As New Collection, db_path As String, selected_structure As String, structure_name As String
    
    db_path = Range("TextBox_TP_db_path").Value
    selected_structure = get_dropdown_value("BuildYourStructure", "Dropdown_TP_Structures2")
    
    args.Add db_path
    args.Add selected_structure
    RunPythonWrapper "db_handling", "save_TP_data", args
    
End Sub

Sub save_TOWER_Data()
    Dim args As New Collection, db_path As String, selected_structure As String, structure_name As String
    
    db_path = Range("TextBox_TOWER_db_path").Value
    selected_structure = get_dropdown_value("BuildYourStructure", "Dropdown_TOWER_Structures2")
    
    args.Add db_path
    args.Add selected_structure
    RunPythonWrapper "db_handling", "save_TOWER_data", args
    
End Sub


Sub save_RNA_Data()
    Dim args As New Collection, db_path As String, selected_structure As String, structure_name As String
    
    db_path = Range("TextBox_RNA_db_path").Value
    selected_structure = get_dropdown_value("BuildYourStructure", "Dropdown_RNA_Structures")
    
    args.Add db_path
    args.Add selected_structure
    RunPythonWrapper "db_handling", "save_RNA_data", args

End Sub


'==============================================================================
' Database Data Deletion Subroutines
' Purpose: Delete selected structures from MP, TP, and TOWER databases by
'          passing database path and structure name to Python delete functions.
'==============================================================================

Sub delete_MP_Data()
    Dim args As New Collection, db_path As String, selected_structure As String
    
    db_path = Range("TextBox_MP_db_path").Value
    selected_structure = get_dropdown_value("BuildYourStructure", "Dropdown_MP_Structures2")
    
    args.Add db_path
    args.Add selected_structure
    RunPythonWrapper "db_handling", "delete_MP_data", args
End Sub

Sub delete_TP_Data()
    Dim args As New Collection, db_path As String, selected_structure As String
    
    db_path = Range("TextBox_TP_db_path").Value
    selected_structure = get_dropdown_value("BuildYourStructure", "Dropdown_TP_Structures2")
    
    args.Add db_path
    args.Add selected_structure
    RunPythonWrapper "db_handling", "delete_TP_data", args
End Sub

Sub delete_TOWER_Data()
    Dim args As New Collection, db_path As String, selected_structure As String
    
    db_path = Range("TextBox_TOWER_db_path").Value
    selected_structure = get_dropdown_value("BuildYourStructure", "Dropdown_TOWER_Structures2")
    
    args.Add db_path
    args.Add selected_structure
    RunPythonWrapper "db_handling", "delete_TOWER_data", args
End Sub


Sub delete_RNA_Data()
    Dim args As New Collection, db_path As String, selected_structure As String
    
    db_path = Range("TextBox_RNA_db_path").Value
    selected_structure = get_dropdown_value("BuildYourStructure", "Dropdown_RNA_Structures")
    
    args.Add selected_structure
    RunPythonWrapper "db_handling", "delete_RNA_data", args
End Sub

'==============================================================================
' Utility Subroutines
'==============================================================================


Sub UpdateIdentifierColumn(wsName As String, TableName As String)
    Dim ws As Worksheet, tbl As ListObject, row As ListRow
    Dim idColIndex As Long, projectID As String, phase As String, structureID As String
    Dim col As ListColumn, colFound As Boolean
    
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Worksheets(wsName)
    Set tbl = ws.ListObjects(TableName)
    
    colFound = False
    For Each col In tbl.ListColumns
        If Trim(col.name) = "Identifier" Then
            idColIndex = col.Index
            colFound = True
            Exit For
        End If
    Next col
    
    If Not colFound Then
        MsgBox "Column 'Identifier' not found!", vbExclamation
        Exit Sub
    End If
    
    For Each row In tbl.ListRows
        With row.Range
            projectID = Trim(.Columns(tbl.ListColumns("Project ID").Index).Value)
            phase = Trim(.Columns(tbl.ListColumns("Phase").Index).Value)
            structureID = Trim(.Columns(tbl.ListColumns("Structure ID").Index).Value)
            
            If projectID <> "" And phase <> "" And structureID <> "" Then
                .Cells(1, idColIndex).Value = projectID & "_" & phase & "_" & structureID
            Else
                .Cells(1, idColIndex).Value = ""
            End If
        End With
    Next row
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub import_MP_from_MPTool()
    
    Dim path As String

    OpenFileDialog "TextBox_Import_MP_path", "select MP Tool file", "Excel files", "*.xlsm"
    path = ActiveSheet.Range("TextBox_Import_MP_path").Value
    If path = "" Then
        Exit Sub
    End If
    
    If Range("USE_MP").Value Then
        
        ClearTableContents "BuildYourStructure", "MP_DATA_TRUE"
        ClearTableContents "BuildYourStructure", "MP_DATA"
        ClearTableContents "BuildYourStructure", "MP_META_TRUE"
        ClearTableContents "BuildYourStructure", "MP_META"
        ClearTableContents "BuildYourStructure", "MP_META_NEW", 1, 6
        ClearTableContents "BuildYourStructure", "MP_MASSES_TRUE"
        ClearTableContents "BuildYourStructure", "MP_MASSES"
        set_dropdown_value "BuildYourStructure", "Dropdown_MP_Structures2", ""
        
        show_MP_section
        Application.Wait (Now + TimeValue("0:00:01"))
        RunPythonWrapper "db_handling", "load_MP_from_MPTool", path
        
    End If
    
    If Range("USE_TP").Value Then
        
        ClearTableContents "BuildYourStructure", "TP_DATA_TRUE"
        ClearTableContents "BuildYourStructure", "TP_DATA"
        ClearTableContents "BuildYourStructure", "TP_META_TRUE"
        ClearTableContents "BuildYourStructure", "TP_META"
        ClearTableContents "BuildYourStructure", "TP_META_NEW", 1, 6
        ClearTableContents "BuildYourStructure", "TP_MASSES_TRUE"
        ClearTableContents "BuildYourStructure", "TP_MASSES"
        set_dropdown_value "BuildYourStructure", "Dropdown_TP_Structures2", ""
        
        show_TP_section
        Application.Wait (Now + TimeValue("0:00:01"))
        RunPythonWrapper "db_handling", "load_TP_from_MPTool", path
        
    End If
    

    If Not Range("USE_MP").Value And Not Range("USE_TP") Then

        MsgBox "Please check Use MP or Use TP to load data from path", vbInformation, "Response"
        
    End If
    
End Sub

Sub import_Masses_from_GConverter()
    
    Dim path As String
    
    OpenFileDialog "Import_GeomConv_path", "select a file", "Excel files", "*.xlsm"
    path = ActiveSheet.Range("Import_GeomConv_path").Value
    If path = "" Then
        Exit Sub
    End If
    
    If Range("UseMPGeomConv").Value Then
        ClearTableContents "BuildYourStructure", "MP_MASSES_TRUE"
        ClearTableContents "BuildYourStructure", "MP_MASSES"
        show_MP_section
        Application.Wait (Now + TimeValue("0:00:01"))
        RunPythonWrapper "db_handling", "load_MPMasses_from_GeomConv", path
        
    End If
    
    If Range("UseTPGeomConv").Value Then
        ClearTableContents "BuildYourStructure", "TP_MASSES_TRUE"
        ClearTableContents "BuildYourStructure", "TP_MASSES"
        show_TP_section
        Application.Wait (Now + TimeValue("0:00:01"))
        RunPythonWrapper "db_handling", "load_TPMasses_from_GeomConv", path
        
    End If
    
    If Range("UseTOWERGeomConv").Value Then
        ClearTableContents "BuildYourStructure", "TOWER_MASSES_TRUE"
        ClearTableContents "BuildYourStructure", "TOWER_MASSES"
        show_TOWER_section
        Application.Wait (Now + TimeValue("0:00:01"))
        RunPythonWrapper "db_handling", "load_TOWERMasses_from_GeomConv", path
        
    End If
    

    If Not Range("UseMPGeomConv").Value And Not Range("UseTPGeomConv") And Not Range("UseTOWERGeomConv") Then

        MsgBox "Please check Use MP, TP and/or TOWRER to load data from path", vbInformation, "Response"
        
    End If
    
End Sub

Sub move_structure_MP()
    Dim displacement As String
    Dim args As New Collection
    displacement = ActiveSheet.Range("DISPL_MP").Value
    args.Add displacement
    
    RunPythonWrapper "misc", "move_structure_MP", args
    ActiveSheet.Range("DISPL_MP").Value = ""
End Sub

Sub move_structure_TP()
    Dim displacement As String
    Dim args As New Collection
    displacement = ActiveSheet.Range("DISPL_TP").Value
    args.Add displacement
    
    RunPythonWrapper "misc", "move_structure_TP", args
    ActiveSheet.Range("DISPL_TP").Value = ""
    
    
    
End Sub


Sub assamble_structure()
    Dim args As New Collection
    Dim RNA_config As String
    ClearTableContents "StructureOverview", "WHOLE_STRUCTURE"
    ClearTableContents "StructureOverview", "ALL_ADDED_MASSES"
    ClearTableContents "StructureOverview", "SKIRT_POINTMASS"
    ClearTableContents "StructureOverview", "SKIRT"
    ClearTableContents "StructureOverview", "RNA"
    
    args.Add 7860
    RNA_config = get_dropdown_value("BuildYourStructure", "Dropdown_RNA_Structures")
    args.Add RNA_config
    
    RunPythonWrapper "misc", "assemble_structure_excel", args
    
    Worksheets("StructureOverview").Activate
End Sub

Sub show_MP_section()
    ShowOnlySelectedColumns "E:EN", "E:X"
End Sub


Sub show_TP_section()
    ShowOnlySelectedColumns "E:EN", "AP:BI"
End Sub

Sub show_TOWER_section()
    ShowOnlySelectedColumns "E:EN", "CB:CU"
End Sub
    
Sub show_RNA_section()
    ShowOnlySelectedColumns "E:EN", "DN:DW"
End Sub


Sub plot_MP()
    RunPythonWrapper "plot", "plot_MP"
End Sub


Sub plot_TP()
    RunPythonWrapper "plot", "plot_TP"
End Sub

Sub plot_TOWER()
    RunPythonWrapper "plot", "plot_TOWER"
End Sub


Sub plot_Assambly()
    RunPythonWrapper "plot", "plot_Assambly_Build"
End Sub



Public Sub BuildYourStructureChange(ByVal Target As Range)

    On Error GoTo CleanExit
    'Application.EnableEvents = False
    'Application.ScreenUpdating = False

    Dim section As Variant
    Dim tblName As String
    Dim watchRng As Range
    Dim idCol As Range
    Dim sections As Variant
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("BuildYourStructure")
    
    ' --- Special case for RNA path ---
    Set watchRng = RangeFromNameOrTable(ws, "TextBox_RNA_db_path")
    If Not watchRng Is Nothing Then
        If Not Intersect(Target, watchRng) Is Nothing Then
            ' Add code to handle RNA path changes (similar to MP/TP/TOWER)
            load_RNA_DB
        End If
    End If
    
    ' --- Special case for RNA data ---
    Set watchRng = RangeFromNameOrTable(ws, "RNA_DATA")
    If Not watchRng Is Nothing Then
        If Not Intersect(Target, watchRng) Is Nothing Then
            CompareTablesAndHighlightDifferences "BuildYourStructure", "RNA_DATA_TRUE", "BuildYourStructure", "RNA_DATA", , RGB(255, 199, 206)
        End If
    End If
    
    sections = Array("MP", "TP", "TOWER")
    
    ' Loop through all sections
    For Each section In sections
        
        ' --- Path ---
        Set watchRng = RangeFromNameOrTable(ws, "TextBox_" & section & "_db_path")
        If Not watchRng Is Nothing Then
            If Not Intersect(Target, watchRng) Is Nothing Then
                 Select Case section
                    Case "MP": load_MP_DB
                    Case "TP": load_TP_DB
                    Case "TOWER": load_TOWER_DB
                 End Select
            End If
        End If

        ' --- Data ---
        tblName = section & "_DATA"
        Set watchRng = RangeFromNameOrTable(ws, tblName)
        If Not watchRng Is Nothing Then
            If Not Intersect(Target, watchRng) Is Nothing Then
                ResizeTableToData tblName
                CompareTablesAndHighlightDifferences "BuildYourStructure", tblName & "_TRUE", "BuildYourStructure", tblName, , RGB(255, 199, 206)

            End If
        End If

        ' --- Masses ---
        tblName = section & "_MASSES"
        Set watchRng = RangeFromNameOrTable(ws, tblName)
        If Not watchRng Is Nothing Then
            If Not Intersect(Target, watchRng) Is Nothing Then
                ResizeTableToData tblName
                CompareTablesAndHighlightDifferences "BuildYourStructure", tblName & "_TRUE", "BuildYourStructure", tblName, , RGB(255, 199, 206)

            End If
        End If

        ' --- Meta ---
        tblName = section & "_META"
        Set watchRng = RangeFromNameOrTable(ws, tblName)
        If Not watchRng Is Nothing Then
            If Not Intersect(Target, watchRng) Is Nothing Then
                Set idCol = ws.ListObjects(tblName).ListColumns("Identifier").DataBodyRange
                If Intersect(Target, idCol) Is Nothing Or ForceUpdate Then
                    UpdateIdentifierColumn "BuildYourStructure", tblName
                End If
                CompareTablesAndHighlightDifferences "BuildYourStructure", tblName & "_TRUE", "BuildYourStructure", tblName, , RGB(255, 199, 206)
            End If
        End If

        '--- Meta new ---
        tblName = section & "_META_NEW"
        Set watchRng = RangeFromNameOrTable(ws, tblName)
        If Not watchRng Is Nothing Then
            If Not Intersect(Target, watchRng) Is Nothing Then
                Set idCol = ws.ListObjects(tblName).ListColumns("Identifier").DataBodyRange
                If Intersect(Target, idCol) Is Nothing Or ForceUpdate Then
                    UpdateIdentifierColumn "BuildYourStructure", tblName
                End If
            End If
        End If

    Next section

CleanExit:
    'Application.EnableEvents = True
    'Application.ScreenUpdating = True

End Sub


Sub Import_GeomConv_useMP_Klicken()
    If Range("UseMPGeomConv").Value = True Then
        Range("UseTPGeomConv").Value = False
        Range("UseTOWERGeomConv").Value = False
    End If
End Sub

Sub Import_GeomConv_useTP_Klicken()
    If Range("UseTPGeomConv").Value = True Then
        Range("UseMPGeomConv").Value = False
        Range("UseTOWERGeomConv").Value = False
    End If
End Sub

Sub Import_GeomConv_useTOWER_Klicken()
    If Range("UseTOWERGeomConv").Value = True Then
        Range("UseMPGeomConv").Value = False
        Range("UseTPGeomConv").Value = False
    End If
End Sub


Sub Import_MPTool_useMP_Klicken()
    If Range("Use_MP").Value = True Then
        Range("Use_TP").Value = False
    End If
End Sub

Sub Import_MPTool_useTP_Klicken()
    If Range("Use_TP").Value = True Then
        Range("Use_MP").Value = False
    End If
End Sub
