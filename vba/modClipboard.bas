Attribute VB_Name = "modClipboard"
Option Explicit

'=== Clipboard helper (ohne MSForms) =================================
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As LongPtr, ByVal lpSrc As LongPtr, ByVal cb As LongPtr)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal cb As Long)
#End If

Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const CF_UNICODETEXT As Long = 13

Public Function ClipboardSetTextUnicode(ByVal s As String) As Boolean
    Dim cb As LongPtr, hMem As LongPtr, pMem As LongPtr, ok As Boolean
    cb = (Len(s) + 1) * 2 ' bytes inkl. Nullterminator

    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, cb)
    If hMem = 0 Then Exit Function
    pMem = GlobalLock(hMem)
    If pMem = 0 Then GlobalFree hMem: Exit Function

    CopyMemory pMem, StrPtr(s), cb
    GlobalUnlock hMem

    If OpenClipboard(0) <> 0 Then
        EmptyClipboard
        ok = (SetClipboardData(CF_UNICODETEXT, hMem) <> 0)
        CloseClipboard
        If ok Then
            ClipboardSetTextUnicode = True
            Exit Function ' Ownership geht ans Clipboard
        End If
    End If
    GlobalFree hMem ' nur bei Fehler
End Function


