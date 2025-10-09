Attribute VB_Name = "Helpers"

'------------------------------------------------------------------------------
' OpenFileDialog
'
' Opens a file dialog to allow the user to select a file and writes the selected
' file path into a specified target cell.
'
' Features:
' - Displays a file picker dialog with a customizable title.
' - Lets you define a file filter (e.g., only .txt or .xlsx files).
' - Writes the selected file path into the specified target cell.
' - If the user cancels, the target cell is cleared.
'
' Parameters:
'   TargetCellAddress (String) - Address of the cell where the file path should be written.
'   DialogTitle (String)       - The title shown in the file picker dialog.
'   FilterName (String)        - Description of the filter (e.g., "Text Files").
'   FilterPattern (String)     - Pattern for the filter (e.g., "*.txt").
'
' Example:
'   OpenFileDialog "B2", "Select a text file", "Text Files", "*.txt"
'   -> Opens a file dialog showing only .txt files, writes chosen file path into cell B2.
'
Sub OpenFileDialog(TargetCellAddress As String, _
                   Optional DialogTitle As String = "Select a file", _
                   Optional FilterName As String = "All Files", _
                   Optional FilterPattern As String = "*.*")

    Dim FileDialog As FileDialog
    Dim filePath As String
    Dim TargetCell As Range

    Set TargetCell = ActiveSheet.Range(TargetCellAddress)
    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)

    With FileDialog
        .Title = DialogTitle
        .InitialFileName = "C:\"   ' Starting folder
        .Filters.Clear
        .Filters.Add FilterName, FilterPattern
        
        If .Show <> -1 Then
            TargetCell.Value = ""   ' User cancelled -> clear target cell
            Exit Sub
        End If
        
        filePath = .SelectedItems(1)
    End With
    TargetCell.Value = filePath
End Sub

'------------------------------------------------------------------------------
' PickFolderDialog
'
' Opens a folder selection dialog and writes the selected folder path
' into the specified worksheet cell.
'
' Parameters:
'   Target_cell (String) - The address of the cell to write the folder path (e.g., "B2")
'
' Example:
'   Call PickFolderDialog("B2")
'------------------------------------------------------------------------------
Sub PickFolderDialog(Target_cell As String)
    Dim FolderDialog As FileDialog
    Dim SelectedFolder As String
    Dim TargetCell As Range

    ' Open folder picker
    Set FolderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With FolderDialog
        .Title = "Select Folder to Save File"
        .InitialFileName = "C:\"
        If .Show <> -1 Then Exit Sub ' User cancelled
        SelectedFolder = .SelectedItems(1)
    End With

    ' Write folder path to target cell
    On Error GoTo NoCell
    Set TargetCell = ActiveSheet.Range(Target_cell)
    TargetCell.Value = SelectedFolder
    Exit Sub

NoCell:
    MsgBox "Target cell not found: " & Target_cell, vbExclamation
End Sub

'*******************************************************************************
' RunPythonWrapper (Shell version, no xlwings)
'
' Description:
'   Runs a Python function from a script in the "python_scripts" folder
'   by launching Python via the Shell.
'
' Parameters:
'   module_name (String)     - Python module name (without .py)
'   Optional function_name   - Function to run inside module (default = "main")
'   Optional args            - Either a string (path) or a Range/Collection of strings
'
' Example Usage:
'   RunPythonWrapper "load_MP_META", "main", "C:\my\path\file.db"
'   RunPythonWrapper "my_script", "main", Range("A1:A5")
'*******************************************************************************
Sub RunPythonWrapper(module_name As String, Optional function_name As String = "main", Optional args As Variant)
    Dim scriptPath As String
    Dim pythonExe As String
    Dim checkPythonExe As String
    Dim checkScriptPath As String
    Dim argsString As String
    Dim excelFileName As String
    Dim item As Variant
    Dim cmd As String
    Dim showShell As Boolean
    Dim wsh As Object
    Dim retCode As Long
    
    ' Turn off Excel updates for speed & stability
    Application.ScreenUpdating = False
    'Application.EnableEvents = False
    'Application.Calculation = xlCalculationManual
    
    On Error GoTo CleanFail

    ' Check debug mode
    On Error Resume Next
    showShell = Sheets("GlobalConfig").Range("debug_mode").Value
    On Error GoTo 0

    ' Get Python executable from config
    On Error GoTo PythonPathError
    pythonExe = Sheets("GlobalConfig").Range("python_path").Value
    On Error GoTo 0

    ' Quote path if it contains spaces
    If InStr(pythonExe, " ") > 0 And Left(pythonExe, 1) <> """" Then
        pythonExe = """" & pythonExe & """"
    End If

    ' Get the script path from the named range
    On Error GoTo ScriptPathError
    scriptPath = Range("python_script_path").Value
    On Error GoTo 0

    ' Remove trailing slash if present
    If Right(scriptPath, 1) = "\" Or Right(scriptPath, 1) = "/" Then
        scriptPath = Left(scriptPath, Len(scriptPath) - 1)
    End If

    If Not FileExists(pythonExe) Then
        MsgBox "Python executable not found: " & pythonExe, vbCritical
        GoTo CleanExit
    End If

    If Not FolderExists(scriptPath) Then
        MsgBox "Python script path not found: " & scriptPath, vbCritical
        GoTo CleanExit
    End If
    ' Always include Excel filename first
    excelFileName = ThisWorkbook.FullName
    argsString = ProcessArgument(excelFileName)

    ' Append other arguments
    If Not IsMissing(args) Then
        If TypeName(args) = "Collection" Then
            For Each item In args
                argsString = argsString & ", " & ProcessArgument(item)
            Next item
        Else
            argsString = argsString & ", " & ProcessArgument(args)
        End If
    End If

    ' Build full command line
    cmd = pythonExe & " -c " & _
          """import sys; sys.path.append(r'" & scriptPath & "'); " & _
          "import " & module_name & "; " & _
          module_name & "." & function_name & "(" & argsString & ")"""

    ' Run Python synchronously
    Set wsh = CreateObject("WScript.Shell")
    If showShell Then
        ' Keep shell open in debug mode
        retCode = wsh.run("cmd /k " & cmd, 1, True)
    Else
        ' Run hidden and wait until finished
        retCode = wsh.run("cmd /c " & cmd, 0, True)
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
        
CleanExit:
    Exit Sub

PythonPathError:
    MsgBox "Named range 'python_path' not found on sheet 'GlobalConfig'.", vbCritical
    Resume CleanExit

ScriptPathError:
    MsgBox "Named range 'python_script_path' not found or invalid.", vbCritical
    Resume CleanExit

CleanFail:
    MsgBox "Error in RunPythonWrapper: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Public Function FileExists(path As String) As Boolean
    Dim fso As Object
    Dim p As String
    
    If Len(path) = 0 Then
        FileExists = False
        Exit Function
    End If
    
    ' Remove quotes
    p = path
    If Left(p, 1) = """" And Right(p, 1) = """" Then
        p = Mid(p, 2, Len(p) - 2)
    End If
    p = Trim(p)
    
    ' Use FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(p)
End Function


Public Function FolderExists(path As String) As Boolean
    Dim fso As Object
    Dim p As String
    
    If Len(path) = 0 Then
        FolderExists = False
        Exit Function
    End If
    
    ' Remove quotes
    p = path
    If Left(p, 1) = """" And Right(p, 1) = """" Then
        p = Mid(p, 2, Len(p) - 2)
    End If
    p = Trim(p)
    
    ' Remove trailing slashes
    Do While Right(p, 1) = "\" Or Right(p, 1) = "/"
        p = Left(p, Len(p) - 1)
    Loop
    
    ' Use FileSystemObject for reliability
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(p)
End Function

'------------------------------------------------------------------------------
' Helper: turns VBA args into proper Python literal strings
'------------------------------------------------------------------------------
Function ProcessArgument(item As Variant) As String
    Dim result As String
    Dim i As Long
    
    If TypeName(item) = "Range" Then
        result = "'" & Replace(item.Value, "'", "''") & "'"
    ElseIf TypeName(item) = "String" Then
        result = "r'" & Replace(item, "\", "\\") & "'"
    ElseIf IsArray(item) Then
        result = "["
        For i = LBound(item) To UBound(item)
            result = result & "r'" & Replace(item(i), "\", "\\") & "', "
        Next i
        If Right(result, 2) = ", " Then
            result = Left(result, Len(result) - 2)
        End If
        result = result & "]"
    Else
        result = CStr(item)
    End If
    
    ProcessArgument = result
End Function

'------------------------------------------------------------------------------
' Function: get_dropdown_value
'
' Purpose:
'   Retrieves the currently selected value from a dropdown (form control)
'   located on a specified worksheet.
'
' Arguments:
'   sheet_name (String): Name of the worksheet containing the dropdown.
'   dropdown_name (String): Name of the dropdown shape (form control) to read.
'
' Returns:
'   String: The text of the currently selected item in the dropdown list.
'------------------------------------------------------------------------------

Function get_dropdown_value(sheet_name As String, dropdown_name As String) As String
    Dim dropdown_value As String
    Dim dropdown_value_id As Integer

    ' Get the selected item index
    dropdown_value_id = ThisWorkbook.Sheets(sheet_name).Shapes(dropdown_name).ControlFormat.Value

    ' If nothing is selected, return an empty string
    If dropdown_value_id = 0 Then
        get_dropdown_value = ""
        Exit Function
    End If

    ' Get the text value of the selected item based on the index
    dropdown_value = ThisWorkbook.Sheets(sheet_name).Shapes(dropdown_name).ControlFormat.List(dropdown_value_id)

    ' Return the value
    get_dropdown_value = dropdown_value
End Function

Sub ClearTableContents(worksheetName As String, TableName As String, Optional fromCol As Long = -1, Optional toCol As Long = -1)
    ' Clears the contents of the specified table, keeping its structure.
    ' If fromCol and toCol are specified, only those columns within the table are cleared.
    '
    ' Arguments:
    ' worksheetName (String) : The name of the worksheet containing the table.
    ' tableName (String)     : The name of the table to clear.
    ' fromCol (Long)         : Optional start column index (1-based, relative to the table).
    ' toCol (Long)           : Optional end column index (1-based, relative to the table).
    '
    ' Example:
    ' ClearTableContents "Sheet1", "Table1"            ' clears all columns
    ' ClearTableContents "Sheet1", "Table1", 2, 4      ' clears columns 2 to 4 in the table

    Dim ws As Worksheet
    Dim table As ListObject
    Dim i As Long

    ' Set the worksheet and table
    Set ws = ThisWorkbook.Sheets(worksheetName)
    Set table = ws.ListObjects(TableName)

    If table.DataBodyRange Is Nothing Then Exit Sub ' exit if table is empty

    Dim startCol As Long
    Dim endCol As Long

    ' Determine column range to clear
    If fromCol = -1 Or toCol = -1 Then
        ' Default: clear all columns
        table.DataBodyRange.ClearContents
    Else
        ' Validate and adjust column range
        startCol = Application.Max(1, fromCol)
        endCol = Application.Min(table.ListColumns.Count, toCol)

        If startCol <= endCol Then
            For i = startCol To endCol
                table.ListColumns(i).DataBodyRange.ClearContents
            Next i
        End If
    End If
End Sub

Function set_dropdown_value(sheet_name As String, dropdown_name As String, new_value As String)
    Dim dropdown_list As Variant
    Dim i As Integer
    Dim dropdown_found As Boolean
    Dim control As Object

    Set control = ThisWorkbook.Sheets(sheet_name).Shapes(dropdown_name).ControlFormat

    With control
        ' If dropdown is empty and trying to set "", then add "" and select it
        If .ListCount = 0 And new_value = "" Then
            .AddItem ""  ' Add an empty selection option
            .Value = 1   ' Select the first item (which is "")
            Exit Function
        End If

        ' Handle unset case ("" means "no selection")
        If new_value = "" Then
            ' Try to find "" in the list
            dropdown_found = False
            For i = 1 To .ListCount
                If .List(i) = "" Then
                    .Value = i
                    dropdown_found = True
                    Exit For
                End If
            Next i
            ' If not found, add it
            If Not dropdown_found Then
                .AddItem ""
                .Value = .ListCount ' Select the last added item
            End If
            Exit Function
        End If

        ' Exit if the dropdown has no items (and new_value is not "")
        If .ListCount = 0 Then
            MsgBox "Dropdown '" & dropdown_name & "' is empty.", vbExclamation
            Exit Function
        End If

        ' Get the dropdown list items
        dropdown_list = .List

        ' Loop through the list to find the matching value
        dropdown_found = False
        For i = 1 To .ListCount
            If .List(i) = new_value Then
                .Value = i ' 1-based index
                dropdown_found = True
                Exit For
            End If
        Next i

        ' Optional: notify if value not found
        If Not dropdown_found Then
            MsgBox "Value '" & new_value & "' not found in dropdown '" & dropdown_name & "'.", vbExclamation
        End If
    End With
End Function

Sub CompareTablesAndHighlightDifferences(sheet1Name As String, table1Name As String, _
                                         sheet2Name As String, table2Name As String, _
                                         Optional matchColor As Variant, _
                                         Optional mismatchColor As Variant)

    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim table1 As ListObject, table2 As ListObject
    Dim r As Long, c As Long
    Dim cell1 As Range, cell2 As Range
    Dim tolerance As Double
    tolerance = 0.00001

    ' Apply default colors if not provided
    If IsMissing(matchColor) Then matchColor = xlNone
    If IsMissing(mismatchColor) Then mismatchColor = RGB(255, 199, 206)

    Set ws1 = ThisWorkbook.Sheets(sheet1Name)
    Set ws2 = ThisWorkbook.Sheets(sheet2Name)
    Set table1 = ws1.ListObjects(table1Name)
    Set table2 = ws2.ListObjects(table2Name)

    ' Ensure the tables have the same number of columns
    If table1.ListColumns.Count <> table2.ListColumns.Count Then
        MsgBox "Tables '" & table1Name & "' and '" & table2Name & "' must have the same number of columns!", vbExclamation
        Exit Sub
    End If

    ' Clear previous formatting if DataBodyRange exists
    If Not table2.DataBodyRange Is Nothing Then
        table2.DataBodyRange.Interior.ColorIndex = xlNone
    End If
    If Not table1.DataBodyRange Is Nothing Then
        table1.DataBodyRange.Interior.ColorIndex = xlNone
    End If

    ' If either table is empty, exit or highlight new rows
    If table1.DataBodyRange Is Nothing And table2.DataBodyRange Is Nothing Then
        ' Both tables empty, nothing to compare
        Exit Sub
    End If

    ' Get number of rows (0 if empty)
    Dim rows1 As Long, rows2 As Long

    If table1.DataBodyRange Is Nothing Then
        rows1 = 0
    Else
         rows1 = table1.DataBodyRange.Rows.Count
    End If

    If table2.DataBodyRange Is Nothing Then
        rows2 = 0
    Else
        rows2 = table2.DataBodyRange.Rows.Count
    End If
    ' Compare cell by cell for overlapping rows
    For r = 1 To Application.Min(rows1, rows2)
        For c = 1 To table1.ListColumns.Count
            Set cell1 = table1.DataBodyRange.Cells(r, c)
            Set cell2 = table2.DataBodyRange.Cells(r, c)

            If IsNumeric(cell1.Value) And IsNumeric(cell2.Value) Then
                If Abs(cell1.Value - cell2.Value) > tolerance Then
                    cell2.Interior.Color = mismatchColor
                ElseIf matchColor <> xlNone Then
                    cell2.Interior.Color = matchColor
                End If
            Else
                If cell1.Value <> cell2.Value Then
                    cell2.Interior.Color = mismatchColor
                ElseIf matchColor <> xlNone Then
                    cell2.Interior.Color = matchColor
                End If
            End If
        Next c
    Next r

    ' Highlight extra rows in table1 (new rows not in table2)
    If rows1 > rows2 Then
        For r = rows2 + 1 To rows1
            For c = 1 To table1.ListColumns.Count
                Set cell1 = table1.DataBodyRange.Cells(r, c)
                If cell1.Value <> "" Then
                    cell1.Interior.Color = mismatchColor
                End If
            Next c
        Next r
    End If

    ' Highlight extra rows in table2 (new rows not in table1)
    If rows2 > rows1 Then
        For r = rows1 + 1 To rows2
            For c = 1 To table2.ListColumns.Count
                Set cell2 = table2.DataBodyRange.Cells(r, c)
                If cell2.Value <> "" Then
                    cell2.Interior.Color = mismatchColor
                End If
            Next c
        Next r
    End If
End Sub

'===============================================================================
' Sub ToggleComparisonGeneric
'===============================================================================
' Toggles table comparison highlighting on or off for a specific pair of tables.
'
' This macro is designed to work with Form Control buttons. It uses metadata
' encoded in the button’s `OnAction` property to determine which tables and sheet
' to work with. The format must be:
'
'     table1_name|table2_name|sheet_name
'
' When the button is clicked:
'   - If comparison is off, it turns on: differences are highlighted (red) and
'     matches can be optionally highlighted (green).
'   - If comparison is on, it turns off: formatting is cleared.
'
' The toggle state is stored in a shared dictionary `comparisonEnabledDict` keyed
' by the combination of table1, table2, and sheet names.
'
' Notes:
' - Supports multiple buttons/table pairs independently.
' - You must assign this macro to each button and set the OnAction string properly.
'
' Example usage:
'   Button's assigned macro: ToggleComparisonGeneric
'   Button's OnAction: MP_DATA_TRUE|MP_DATA|BuildYourStructure
'
' Dependencies:
'   - Requires CompareTablesAndHighlightDifferences and ResetNormalColoring subs.
'   - Requires `comparisonEnabledDict` defined as `Public` in a standard module.
'
'===============================================================================

Sub ToggleComparisonGeneric()
    Dim btnName As String
    Dim ws As Worksheet
    Dim btn As Button
    Dim tableInfo() As String
    Dim key As String

    btnName = Application.Caller

    ' Find the worksheet and button object
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set btn = ws.Buttons(btnName)
        On Error GoTo 0
        If Not btn Is Nothing Then Exit For
    Next ws

    If btn Is Nothing Then
        MsgBox "Button '" & btnName & "' not found.", vbExclamation
        Exit Sub
    End If

    ' First time use: create the dictionary
    If comparisonEnabledDict Is Nothing Then Set comparisonEnabledDict = CreateObject("Scripting.Dictionary")

    ' Parse metadata from the button’s OnAction string
    tableInfo = Split(btn.OnAction, "|")
    If UBound(tableInfo) < 2 Then
        MsgBox "Button OnAction must be in format: table1|table2|sheetName", vbExclamation
        Exit Sub
    End If

    Dim table1Name As String: table1Name = tableInfo(0)
    Dim table2Name As String: table2Name = tableInfo(1)
    Dim sheetName As String: sheetName = tableInfo(2)

    key = table1Name & "|" & table2Name & "|" & sheetName

    ' Toggle state
    If Not comparisonEnabledDict.Exists(key) Then
        comparisonEnabledDict.Add key, False
    End If
    comparisonEnabledDict(key) = Not comparisonEnabledDict(key)

    If comparisonEnabledDict(key) Then
        btn.Text = "Stop Comparison"
        CompareTablesAndHighlightDifferences sheetName, table1Name, sheetName, table2Name, RGB(198, 239, 206), RGB(255, 199, 206)
    Else
        btn.Text = "Start Comparison"
        ResetNormalColoring sheetName, table1Name
    End If
End Sub

Sub ResetNormalColoring(sheetName As String, TableName As String)
    ' Resets the background color (fill) of the data rows in the specified table.
    
    Dim ws As Worksheet
    Dim table As ListObject

    ' Get the worksheet and table
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set table = ws.ListObjects(TableName)

    On Error Resume Next
    ' Clear manual fill color from the table's data range
    table.DataBodyRange.Interior.ColorIndex = xlColorIndexNone
    On Error GoTo 0
End Sub

Sub ShowOnlyColumns_old(columnRange As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BuildYourStructure") ' Adjust sheet name if needed

    On Error GoTo InvalidRange
    ws.Columns.Hidden = False ' Show all columns
    
    Dim col As Range
    For Each col In ws.UsedRange.Columns
        If Intersect(col, ws.Range(columnRange)) Is Nothing Then
            col.EntireColumn.Hidden = True
        End If
    Next col

    Exit Sub

InvalidRange:
    MsgBox "Invalid column range: " & columnRange, vbCritical
End Sub

Sub ShowOnlySelectedColumns(rngAllCols As String, rngVisibleCols As String)
    Dim ws As Worksheet
    Dim col As Range
    Dim colNum As Long
    Dim visibleDict As Object
    Dim shp As Shape
    Dim topLeftCol As Long

    Set ws = ActiveSheet
    Set visibleDict = CreateObject("Scripting.Dictionary")

    ' -- Speed up execution
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Always show column A
    ws.Columns("A").Hidden = False

    ' Build dictionary of visible columns
    For Each col In ws.Range(rngVisibleCols).Columns
        visibleDict(col.Column) = True
    Next col

    ' Hide/show columns
    For Each col In ws.Range(rngAllCols).Columns
        colNum = col.Column
        col.EntireColumn.Hidden = Not visibleDict.Exists(colNum)
    Next col

    ' Hide dropdowns in hidden columns
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlDropDown Then
                On Error Resume Next ' Prevents crash if TopLeftCell is off-sheet
                topLeftCol = shp.TopLeftCell.Column
                On Error GoTo 0
                If ws.Columns(topLeftCol).Hidden = True And topLeftCol <> 1 Then
                    shp.Visible = msoFalse
                Else
                    shp.Visible = msoTrue
                End If
            End If
        End If
    Next shp

    ' Hide pictures in hidden columns (PNG, SVG, etc.)
    For Each shp In ws.Shapes
        Select Case shp.Type
            Case msoPicture, msoLinkedPicture, msoLinkedGraphic, msoGraphic
                On Error Resume Next ' Prevents crash if TopLeftCell is off-sheet
                topLeftCol = shp.TopLeftCell.Column
                On Error GoTo 0
                If ws.Columns(topLeftCol).Hidden = True And topLeftCol <> 1 Then
                    shp.Visible = msoFalse
                Else
                    shp.Visible = msoTrue
                End If
        End Select
    Next shp

    ' -- Restore application settings
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub toggle_folding()
    Dim groupRow As Integer
    groupRow = 28 ' ?? Change to the row ABOVE the group

    With Rows(groupRow)
        If .OutlineLevel > 1 Then
            .ShowDetail = Not .ShowDetail ' ? Toggle (fold/unfold)
        Else
            MsgBox "Row " & groupRow & " is not grouped.", vbExclamation
        End If
    End With
End Sub

Sub DeleteNamedRange_AllScopes()
    Dim nm As name
    For Each nm In ThisWorkbook.Names
        If InStr(1, nm.name, "Dropdown_TP_Structures", vbTextCompare) > 0 Then
            Debug.Print "Deleting: " & nm.name
            nm.Delete
        End If
    Next nm
End Sub

Public Sub ResizeTableToData(ByVal TableName As String, _
                             Optional ByVal CountFormulaBlanksAsData As Boolean = False)
    Dim ws As Worksheet, lo As ListObject
    Dim headerRow As Long, firstCol As Long, lastCol As Long, lastRow As Long
    Dim col As Long, rng As Range, searchRange As Range
    Dim lookIn As XlFindLookIn, hadTotals As Boolean

    ' Find the table by name across all sheets
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(TableName)
        On Error GoTo 0
        If Not lo Is Nothing Then Exit For
    Next ws
    If lo Is Nothing Then
        MsgBox "Table '" & TableName & "' not found.", vbExclamation
        Exit Sub
    End If

    headerRow = lo.HeaderRowRange.row
    firstCol = lo.Range.Column
    lastCol = lo.HeaderRowRange.Cells(lo.HeaderRowRange.Columns.Count).Column

    ' Choose whether "" from formulas counts as data
    lookIn = IIf(CountFormulaBlanksAsData, xlFormulas, xlValues)

    ' Scan every table column from the header+1 down to the bottom of the sheet
    lastRow = headerRow
    For col = firstCol To lastCol
        Set searchRange = ws.Range(ws.Cells(headerRow + 1, col), ws.Cells(ws.Rows.Count, col))
        Set rng = searchRange.Find(What:="*", lookIn:=lookIn, LookAt:=xlPart, _
                                   SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
        If Not rng Is Nothing Then
            If rng.row > lastRow Then lastRow = rng.row
        End If
    Next col

    ' Keep at least one data row to avoid ListObject resize errors
    If lastRow < headerRow + 1 Then lastRow = headerRow + 1

    ' Respect Totals row if present
    hadTotals = lo.ShowTotals
    If hadTotals Then lo.ShowTotals = False

    lo.Resize ws.Range(ws.Cells(headerRow, firstCol), ws.Cells(lastRow, lastCol))

    If hadTotals Then lo.ShowTotals = True
End Sub

Function CheckPath(filePath As String, Optional extension As String = "") As Boolean
    ' Default return value
    CheckPath = False
    
    ' Check if empty
    If Trim(filePath) = "" Then
        MsgBox "Error: File path is empty!", vbCritical, "Invalid Path"
        Exit Function
    End If
    
    ' Check extension if provided
    If extension <> "" Then
        If LCase(Right$(filePath, Len(extension) + 1)) <> "." & LCase(extension) Then
            MsgBox "Error: File must have the extension ." & extension & "!", vbCritical, "Invalid Path"
            Exit Function
        End If
    End If
    
    ' Check if file exists
    If Dir(filePath) = "" Then
        MsgBox "Error: File not found at given path!" & vbCrLf & filePath, vbCritical, "Invalid Path"
        Exit Function
    End If
    
    ' If all checks pass
    CheckPath = True
End Function

Sub ClearFormDropDown(sheetName As String, dropName As String)
    With Worksheets(sheetName).DropDowns(dropName)
        .RemoveAllItems
    End With
End Sub

Public Function TableExists(ws As Worksheet, tblName As String) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tblName)
    TableExists = Not lo Is Nothing
    On Error GoTo 0
End Function

Public Function RangeExists(ws As Worksheet, rngName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.Range(rngName) 'works for named ranges
    RangeExists = (Err.Number = 0 And Not rng Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

Public Function RangeFromNameOrTable(ws As Worksheet, name As String) As Range
    On Error Resume Next
    If TableExists(ws, name) Then
        Set RangeFromNameOrTable = ws.ListObjects(name).Range
    Else
        Set RangeFromNameOrTable = ws.Range(name)
    End If
    On Error GoTo 0
End Function

Sub InstallPythonRequirements()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim pyPath As String
    Dim reqCell As Range
    Dim reqList As String
    Dim cmd As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("GlobalConfig")
    
    ' Get Python path from named range
    pyPath = ws.Range("python_path").Value
    If pyPath = "" Then
        MsgBox "Python path is not defined in named range 'python_path'.", vbCritical
        Exit Sub
    End If
    
    ' Get the table
    On Error Resume Next
    Set tbl = ws.ListObjects("requirements")
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "Table 'requirements' not found on sheet 'GlobalConfig'.", vbCritical
        Exit Sub
    End If
    
    ' Build space-separated list of packages
    reqList = ""
    For Each reqCell In tbl.ListColumns(1).DataBodyRange
        If Trim(reqCell.Value) <> "" Then
            reqList = reqList & " " & Trim(reqCell.Value)
        End If
    Next reqCell
    
    If reqList = "" Then
        MsgBox "No requirements found in table.", vbExclamation
        Exit Sub
    End If
    
    ' Build command to run in cmd and leave window open
    ' /K keeps the cmd window open after execution
    cmd = "cmd /K """ & pyPath & " -m pip install" & reqList & """"
    
    ' Run command
    shell cmd, vbNormalFocus
End Sub

Option Explicit

Public Sub MapNetworkDrive()
    On Error Resume Next  ' Suppress all VBA runtime errors
    
    Dim batFile As String
    Dim shell As Object
    Dim tempPath As String
    Dim fnum As Integer
    
    ' Create temporary batch file
    tempPath = Environ$("TEMP") & "\map_drive.bat"
    fnum = FreeFile
    Open tempPath For Output As #fnum
    Print #fnum, "@ECHO OFF"
    Print #fnum, "NET USE O: \\HH-DC02.jbo.local\daten >nul 2>&1"
    Close #fnum
    
    ' Run batch file silently
    Set shell = CreateObject("WScript.Shell")
    shell.run """" & tempPath & """", 0, True  ' 0 = hidden window, True = wait until done
    
    ' Cleanup silently
    Kill tempPath
End Sub
