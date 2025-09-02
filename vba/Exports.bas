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
    ClearTableContents "ExportStructure", "APPURTANCES"
    RunPythonWrapper "export", "fill_WLGenMasses"
End Sub



Sub fill_Bladed_table()
    ClearTableContents "ExportStructure", "Bladed_Nodes"
    ClearTableContents "ExportStructure", "Bladed_Elements"
    RunPythonWrapper "export", "fill_Bladed_table"
End Sub


Sub copy_BladedNode_Output_To_Clipboard()
    Const H_ELEV As String = "Elevation [m]"
    Const H_LX   As String = "Local x [m]"
    Const H_LY   As String = "Local y [m]"
    Const H_MASS As String = "Point mass [m]"

    Dim lo As ListObject
    Dim colElev As Long, colLx As Long, colLy As Long, colMass As Long
    Dim r As ListRow, txt As String, lineCount As Long

    ' Get table by name
    Set lo = ActiveSheet.ListObjects("Bladed_Nodes")

    ' Find column indices by header names
    colElev = lo.ListColumns(H_ELEV).Index
    colLx = lo.ListColumns(H_LX).Index
    colLy = lo.ListColumns(H_LY).Index
    colMass = lo.ListColumns(H_MASS).Index

    ' Build tab-separated text from each row
    For Each r In lo.ListRows
        txt = txt & r.Range.Cells(1, colElev).Text & vbTab & _
                    r.Range.Cells(1, colLx).Text & vbTab & _
                    r.Range.Cells(1, colLy).Text & vbTab & _
                    r.Range.Cells(1, colMass).Text & vbCrLf
        lineCount = lineCount + 1
    Next r

    If Len(txt) = 0 Then
        MsgBox "Nothing to copy.", vbExclamation
        Exit Sub
    End If

    If ClipboardSetTextUnicode(txt) Then
        MsgBox "Copied " & lineCount & " rows to clipboard." & vbCrLf & _
               "Columns:" & vbCrLf & H_ELEV & ", " & H_LX & ", " & H_LY & ", " & H_MASS, _
               vbInformation, "Bladed Export"
    Else
        MsgBox "Clipboard error.", vbExclamation
    End If
End Sub

Sub copy_BladedMember_Output_To_Clipboard()

    Const H_NODE As String = "Node [-]"
    Const H_DIAMETER As String = "Diameter [m]"
    Const H_WALL_THICKNESS As String = "Wall thickness [mm]"

    Dim lo As ListObject
    Dim colNode As Long, colDiameter As Long, colWallTh As Long
    Dim r As ListRow, txt As String, lineCount As Long

    ' Get table by name
    Set lo = ActiveSheet.ListObjects("Bladed_Elements")

    ' Find column indices by header names
    colNode = lo.ListColumns(H_NODE).Index
    colDiameter = lo.ListColumns(H_DIAMETER).Index
    colWallTh = lo.ListColumns(H_WALL_THICKNESS).Index

    ' Build tab-separated text from each row
    For Each r In lo.ListRows
        txt = txt & r.Range.Cells(1, colNode).Text & vbTab & _
                    r.Range.Cells(1, colDiameter).Text & vbTab & _
                    r.Range.Cells(1, colWallTh).Text & vbCrLf
        lineCount = lineCount + 1
    Next r

    If Len(txt) = 0 Then
        MsgBox "Nothing to copy.", vbExclamation
        Exit Sub
    End If

    If ClipboardSetTextUnicode(txt) Then
        MsgBox "Copied " & lineCount & " rows to clipboard." & vbCrLf & _
               "Columns:" & vbCrLf & H_NODE & ", " & H_DIAMETER & ", " & H_WALL_THICKNESS, _
               vbInformation, "Bladed Export"
    Else
        MsgBox "Clipboard error.", vbExclamation
    End If

End Sub


'''''''''''''
''' Show'''
'''''''''''''

Sub show_JBOOST_section()
    ShowOnlySelectedColumns "E:BX", "E:Q"
End Sub

Sub show_WLGen_section()
    ShowOnlySelectedColumns "E:BX", "R:AE"
End Sub

Sub show_Bladed_section()
    ShowOnlySelectedColumns "E:BX", "AF:AX"
End Sub
Sub show_Sesam_section()
    ShowOnlySelectedColumns "E:BX", "BB:BX"
End Sub
'''''''''''''
''' Sesam '''
'''''''''''''
Sub fill_Sesam_table()
    RunPythonWrapper "export", "fill_Sesam_table"
    Update_Sesam_Export_Yellow
End Sub

Sub select_Sesam_out()
    Dim prevValue As Variant
    
    PickFolderDialog "Sesam_Path"
    
End Sub
Sub export_Sesam()
    Dim lua_path As Variant
    lua_path = Range("Sesam_Path").Value
    RunPythonWrapper "export", "export_Sesam", lua_path
End Sub
Sub Update_Sesam_Export_Yellow()
    Const TBL_STRUCT_SRC     As String = "tbl_ExportStructure_Structure"
    Const TBL_MASS_SRC       As String = "tbl_ExportStructure_Mass"
    Const TBL_YELLOW_STRUCT  As String = "tbl_Export_Sesam"
    Const TBL_YELLOW_MASS    As String = "tbl_Export_Sesam_Mass"

    Dim ws As Worksheet
    Dim loStructSrc As ListObject, loMassSrc As ListObject
    Dim loYellowStruct As ListObject, loYellowMass As ListObject

    Set ws = ActiveSheet
    Set loStructSrc = ws.ListObjects(TBL_STRUCT_SRC)
    Set loMassSrc = ws.ListObjects(TBL_MASS_SRC)
    Set loYellowStruct = ws.ListObjects(TBL_YELLOW_STRUCT)
    Set loYellowMass = ws.ListObjects(TBL_YELLOW_MASS)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1) Structure-driven yellow table
    ResizeAndFillTable loYellowStruct, loStructSrc.ListRows.Count

    ' 2) Mass-driven yellow table
    ResizeAndFillTable loYellowMass, loMassSrc.ListRows.Count

    ' 3) Force the Mass table to be EXACTLY one column (header col) and real used rows
    Fix_Mass_Table_Range ws, TBL_YELLOW_MASS

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Yellow tables updated successfully:" & vbCrLf & vbCrLf & _
           "Source (Structure): " & loStructSrc.ListRows.Count & " rows" & vbCrLf & _
           "Export table (Structure): " & loYellowStruct.ListRows.Count & " rows" & vbCrLf & vbCrLf & _
           "Source (Mass): " & loMassSrc.ListRows.Count & " rows" & vbCrLf & _
           "Export table (Mass): " & loYellowMass.ListRows.Count & " rows", vbInformation


End Sub

Private Sub ResizeAndFillTable(lo As ListObject, ByVal nRows As Long)
    Const BASE_WIDTH As Double = 255

    With lo
        ' --- resize (keep at least 1 data row for formulas) ---
        If .DataBodyRange Is Nothing Then
            .Resize .HeaderRowRange.Resize(2, .ListColumns.Count)
        End If
        If nRows <= 0 Then
            .Resize .Range.Resize(2, .ListColumns.Count)
        Else
            .Resize .Range.Resize(nRows + 1, .ListColumns.Count)
        End If

        ' --- autofill from first data row ---
        If Not .DataBodyRange Is Nothing Then
            Dim r1 As Range
            Set r1 = .DataBodyRange.Rows(1)
            r1.AutoFill Destination:=.DataBodyRange.Resize(Application.Max(1, nRows), r1.Columns.Count), _
                        Type:=xlFillDefault
        End If

        ' --- layout: no wrap, wide then shrink, equal row height ---
        .Range.WrapText = False
        .Range.Columns.ColumnWidth = BASE_WIDTH
        .Range.Columns.AutoFit

        Dim maxH As Double, r As Range
        maxH = .HeaderRowRange.RowHeight
        For Each r In .DataBodyRange.Rows
            If r.RowHeight > maxH Then maxH = r.RowHeight
        Next r
        .HeaderRowRange.RowHeight = maxH
        .DataBodyRange.Rows.RowHeight = maxH
    End With
End Sub

Private Sub Fix_Mass_Table_Range(ws As Worksheet, ByVal massTableName As String)
    ' Makes the Mass table exactly ONE column wide (its header col)
    ' and extends down to the last used cell in that column.
    Dim lo As ListObject
    Dim hdr As Range, firstHdrCell As Range
    Dim lastRow As Long
    Dim newRange As Range

    Set lo = ws.ListObjects(massTableName)
    Set hdr = lo.HeaderRowRange
    Set firstHdrCell = hdr.Cells(1)  ' first (and only) column header

    ' last used row in that column (ensure at least 1 data row exists)
    lastRow = ws.Cells(ws.Rows.Count, firstHdrCell.Column).End(xlUp).row
    If lastRow < firstHdrCell.row + 1 Then lastRow = firstHdrCell.row + 1

    ' new single-column range
    Set newRange = ws.Range(firstHdrCell, ws.Cells(lastRow, firstHdrCell.Column))
    lo.Resize newRange

    ' tidy
    lo.Range.WrapText = False
    lo.Range.Columns.AutoFit
End Sub

