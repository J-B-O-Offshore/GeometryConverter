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


