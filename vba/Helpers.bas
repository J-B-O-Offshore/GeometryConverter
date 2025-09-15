Attribute VB_Name = "Helpers"
' PRIVATE FUNCTIONS:

' -----------------------------------------------------------------------------
' Function: FileExists
'
' Description:
'   Checks whether a specified file exists on disk.
'   Handles paths with optional quotes and trims extra spaces.
'
' Parameters:
'   path (String) [Required]
'       - The full path to the file, optionally quoted.
'
' Returns:
'   Boolean
'       - True if the file exists, False otherwise.
'
' Notes:
'   - Uses the FileSystemObject for reliable file existence checking.
'   - Returns False if the input path is empty.
' -----------------------------------------------------------------------------
Private Function FileExists(path As String) As Boolean
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

' -----------------------------------------------------------------------------
' Function: FolderExists
'
' Description:
'   Checks whether a specified folder exists on disk.
'   Handles paths with optional quotes, trims extra spaces, and removes
'   trailing slashes.
'
' Parameters:
'   path (String) [Required]
'       - The full path to the folder, optionally quoted.
'
' Returns:
'   Boolean
'       - True if the folder exists, False otherwise.
'
' Notes:
'   - Uses the FileSystemObject for reliable folder existence checking.
'   - Returns False if the input path is empty.
' -----------------------------------------------------------------------------

Private Function FolderExists(path As String) As Boolean
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


' -----------------------------------------------------------------------------
' Subroutine: OpenFileDialog
'
' Description:
'   Opens a file picker dialog in Excel and writes the selected file path into
'   a specified target cell. Provides options to customize the dialog title
'   and file filters.
'
' Parameters:
'   TargetCellAddress (String) [Required]
'       - The address of the cell (e.g., "B2") where the selected file path
'         will be written.
'
'   DialogTitle (String) [Optional, default: "Select a file"]
'       - The title displayed on the file dialog window.
'
'   FilterName (String) [Optional, default: "All Files"]
'       - The display name of the file filter (e.g., "Excel Files").
'
'   FilterPattern (String) [Optional, default: "*.*"]
'       - The file pattern for filtering files (e.g., "*.xlsx;*.xls").
'
' Behavior:
'   1. Uses the workbook's folder as the starting location for the dialog.
'      If the workbook is unsaved, defaults to "C:\".
'   2. Clears any existing filters and adds the specified filter.
'   3. Opens the file picker dialog and waits for user selection.
'   4. Writes the selected file path into the target cell. If the user cancels,
'      the target cell is cleared.
'
' Example Usage:
'   OpenFileDialog "B2"
'   OpenFileDialog "C5", "Select an Excel file", "Excel Files", "*.xlsx;*.xls"
'
' Notes:
'   - Only the first selected file is used if multiple selection is allowed.
' -----------------------------------------------------------------------------
Sub OpenFileDialog(TargetCellAddress As String, _
                   Optional DialogTitle As String = "Select a file", _
                   Optional FilterName As String = "All Files", _
                   Optional FilterPattern As String = "*.*")

    Dim FileDialog As FileDialog
    Dim filePath As String
    Dim TargetCell As Range
    Dim startFolder As String

    Set TargetCell = ActiveSheet.Range(TargetCellAddress)
    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)

    ' Use workbook location as starting folder (fallback to C:\ if unsaved)
    If ThisWorkbook.path <> "" Then
        startFolder = ThisWorkbook.path & "\"
    Else
        startFolder = "C:\"
    End If

    With FileDialog
        .Title = DialogTitle
        .InitialFileName = startFolder
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

' -----------------------------------------------------------------------------
' Subroutine: RunPythonWrapper
'
' Description:
'   Executes a specified Python module and function from within Excel using VBA.
'   This wrapper constructs a command line call to Python, passes the current
'   workbook filename and any additional arguments, and optionally displays
'   the command prompt for debugging.
'
' Parameters:
'   module_name (String) [Required]
'       - The name of the Python module to import and execute.
'
'   function_name (String) [Optional, default: "main"]
'       - The function within the module to call.
'
'   args (Variant) [Optional]
'       - Additional arguments to pass to the Python function. Can be a single
'         value or a Collection of values.
'
' Behavior:
'   1. Reads the Python executable path and script folder from the "GlobalConfig"
'      sheet's named ranges: 'python_path' and 'python_script_path'.
'   2. Verifies that the Python executable and script folder exist.
'   3. Constructs a Python command line that appends the script path to sys.path,
'      imports the module, and calls the specified function with the workbook
'      filename and additional arguments.
'   4. Executes the command synchronously using WScript.Shell.
'   5. In debug mode (controlled via 'debug_mode' on GlobalConfig), the command
'      prompt remains open after execution; otherwise, it runs hidden.
'   6. Temporarily disables Excel screen updating, events, and calculation for
'      performance and stability, restoring them afterward.
'
' Error Handling:
'   - Displays a message box if the Python executable or script path is missing or invalid.
'   - Catches other runtime errors and reports them via message box.
'
' Example Usage:
'   RunPythonWrapper "my_module"
'   RunPythonWrapper "my_module", "my_function", Array("arg1", 123)
'
' Notes:
'   - The workbook filename is always passed as the first argument to the Python function.
'   - Spaces in paths are automatically quoted.
' -----------------------------------------------------------------------------
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
        retCode = wsh.Run("cmd /k " & cmd, 1, True)
    Else
        ' Run hidden and wait until finished
        retCode = wsh.Run("cmd /c " & cmd, 0, True)
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


' -----------------------------------------------------------------------------
' Function: ProcessArgument
'
' Description:
'   Converts a VBA variable into a string formatted as a Python literal.
'   Handles ranges, strings, arrays, and other types appropriately so that
'   they can be safely passed as arguments to Python functions via a command line.
'
' Parameters:
'   item (Variant) [Required]
'       - The value to convert. Can be a single value, a Range, or an Array.
'
' Returns:
'   String
'       - A string representing the Python literal version of the input.
'
' Behavior:
'   - Range: returns the cell value as a single-quoted Python string, escaping
'     single quotes inside the value.
'   - String: returns a raw Python string (r'...') and escapes backslashes.
'   - Array: returns a Python list with each element converted as a raw string.
'   - Other types (numbers, Boolean, etc.): converted to string using CStr.
'
' Example Usage:
'   ProcessArgument("C:\path\file.txt")  ->  r'C:\\path\\file.txt'
'   ProcessArgument(Range("A1"))        ->  'cell_value'
'   ProcessArgument(Array("a", "b"))    ->  [r'a', r'b']
' -----------------------------------------------------------------------------
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


' -----------------------------------------------------------------------------
' Function: set_dropdown_value
'
' Description:
'   Sets the selected value of a Form Control dropdown (ComboBox) on a
'   specified worksheet. Handles empty values, adds missing items if needed,
'   and provides optional warnings if the value is not found.
'
' Parameters:
'   sheet_name (String) [Required]
'       - The name of the worksheet containing the dropdown.
'
'   dropdown_name (String) [Required]
'       - The name of the Form Control dropdown shape.
'
'   new_value (String) [Required]
'       - The value to set in the dropdown. Use "" to clear or set an empty
'         selection.
'
' Returns:
'   None explicitly (Function used for its side-effect of setting the dropdown).
'
' Behavior:
'   1. If the dropdown is empty and new_value is "", adds an empty item and selects it.
'   2. If new_value is "", selects an existing empty item or adds one if missing.
'   3. If new_value exists in the dropdown list, selects it.
'   4. If new_value is not found in the list, optionally shows a message box.
'   5. Warns if attempting to set a non-empty value when the dropdown has no items.
'
' Example Usage:
'   set_dropdown_value "Sheet1", "DropDown1", "Option A"
'   set_dropdown_value "Sheet1", "DropDown1", ""      ' clear selection
'
' Notes:
'   - Works only with Form Control dropdowns (not ActiveX controls).
'   - Dropdown items are 1-based indexed in VBA.
' -----------------------------------------------------------------------------
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

    ' Hide pictures in hidden columns
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Then
            On Error Resume Next ' Prevents crash if TopLeftCell is off-sheet
            topLeftCol = shp.TopLeftCell.Column
            On Error GoTo 0
            If ws.Columns(topLeftCol).Hidden = True And topLeftCol <> 1 Then
                shp.Visible = msoFalse
            Else
                shp.Visible = msoTrue
            End If
        End If
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
    Shell cmd, vbNormalFocus
End Sub
