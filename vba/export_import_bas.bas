Attribute VB_Name = "export_import_bas"
'========================  Module: export_import_bas  =========================
' Ex- & Import aller VBA-Module (bas/cls/frm) für Git-Diffs & Re-Import
' - ExportAllModules [outFolder]: schreibt alle Module als Textdateien
' - ImportAllModules [inFolder, replaceExisting]: importiert Dateien zurück
'     replaceExisting:=True   -> gleichnamige Module werden vorher entfernt
' Hinweis: Dokument-Module (DieseArbeitsmappe/Blatt-Module) werden nie exportiert/gelöscht.
'==========================================================================

Option Explicit

' -------- Public API --------

Public Sub ExportAllModules(Optional ByVal outFolder As String = "")
    Dim fso As Object, vbcomp As Object, ext As String, path As String
    If outFolder = "" Then outFolder = ThisWorkbook.path & "\vba"
    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolder fso, outFolder
    ' Alte Exporte aufräumen (nur Code-Dateien)
    DeleteCodeFiles fso, outFolder

    For Each vbcomp In ThisWorkbook.VBProject.VBComponents
        Select Case vbcomp.Type
            Case 1: ext = "bas"  ' std module
            Case 2: ext = "cls"  ' class module
            Case 3: ext = "frm"  ' userform (Excel legt zusätzlich .frx an)
            Case Else: ext = ""
        End Select
        If ext <> "" Then
            path = outFolder & "\" & vbcomp.name & "." & ext
            On Error Resume Next
            vbcomp.Export path
            On Error GoTo 0
        End If
    Next vbcomp
End Sub
Public Sub ImportAllModules(Optional ByVal inFolder As String = "", Optional ByVal replaceExisting As Boolean = True)
    Dim fso As Object, fol As Object, fil As Object
    Dim ext As String, internalName As String
    Dim newComp As Object, tmpName As String
    Dim attempts As Long
    
    If inFolder = "" Then inFolder = ThisWorkbook.path & "\vba"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(inFolder) Then
        MsgBox "Folder not found: " & inFolder, vbCritical
        Exit Sub
    End If
    
    Set fol = fso.GetFolder(inFolder)
    For Each fil In fol.Files
        ext = LCase$(fso.GetExtensionName(fil.name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            internalName = GetInternalVBName(fil.path)
            If internalName = "" Then internalName = fso.GetBaseName(fil.name)
            
            If replaceExisting Then SafeRemoveComponent internalName
            
            ' --- import with temporary unique name
            tmpName = "__TMP_IMPORT__" & Format(Now, "hhmmss") & "_" & Int(Rnd() * 1000)
            On Error Resume Next
            Set newComp = ThisWorkbook.VBProject.VBComponents.Import(fil.path)
            If Err.Number <> 0 Then
                MsgBox "Import failed: " & fil.path & vbCrLf & Err.Description, vbExclamation
                Err.Clear
            Else
                ' --- retry renaming until successful
                attempts = 0
                Do
                    On Error Resume Next
                    newComp.name = internalName
                    On Error GoTo 0
                    attempts = attempts + 1
                    If newComp.name = internalName Then Exit Do
                    DoEvents
                    Application.Wait Now + TimeValue("0:00:01") ' wait 1 second
                Loop While attempts < 5
                Debug.Print "Imported: " & fil.name & " as " & internalName
            End If
            On Error GoTo 0
        End If
    Next fil
End Sub

' -------- Helpers (private) --------

' Reads the "Attribute VB_Name" line from a code file
Private Function GetInternalVBName(ByVal filePath As String) As String
    Dim f As Integer, line As String
    On Error GoTo CleanFail
    
    f = FreeFile
    Open filePath For Input As #f
    Do Until EOF(f)
        Line Input #f, line
        If LCase$(Left$(Trim$(line), 18)) = "attribute vb_name" Then
            GetInternalVBName = Replace(Trim$(Split(line, "=")(1)), """", "")
            Exit Do
        End If
    Loop
    Close #f
    Exit Function
    
CleanFail:
    On Error Resume Next
    If f > 0 Then Close #f
    GetInternalVBName = ""
End Function

' Removes only normal modules/classes/forms – never document modules
Private Sub SafeRemoveComponent(ByVal compName As String)
    Dim vbcomp As Object
    On Error Resume Next
    Set vbcomp = ThisWorkbook.VBProject.VBComponents(compName)
    On Error GoTo 0
    
    If Not vbcomp Is Nothing Then
        Select Case vbcomp.Type
            Case 1, 2, 3  ' std module, class, userform
                On Error Resume Next
                ThisWorkbook.VBProject.VBComponents.Remove vbcomp
                On Error GoTo 0
            Case Else
                ' document modules untouched
        End Select
    End If
End Sub

' -------- Helpers (private) --------

Private Sub EnsureFolder(ByVal fso As Object, ByVal folder As String)
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
End Sub

Private Sub DeleteCodeFiles(ByVal fso As Object, ByVal folder As String)
    Dim f As Object, ext As String
    For Each f In fso.GetFolder(folder).Files
        ext = LCase$(fso.GetExtensionName(f.name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Or ext = "frx" Then
            On Error Resume Next
            f.Delete True
            On Error GoTo 0
        End If
    Next f
End Sub


