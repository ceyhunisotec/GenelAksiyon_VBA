Attribute VB_Name = "Mod_Lock"
Option Explicit

Private Const LOCK_FILE_NAME As String = "GenelAksiyon_Schedule.lock"

Public Function AcquireGlobalLock() As Boolean
    Dim fso As Object
    Dim lockFile As String
    Dim ts As Object

    lockFile = Environ$("TEMP") & "\" & LOCK_FILE_NAME
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(lockFile) Then
        AcquireGlobalLock = False
        Exit Function
    End If

    Set ts = fso.CreateTextFile(lockFile, True)
    ts.Write Now
    ts.Close

    AcquireGlobalLock = True
End Function

Public Sub ReleaseGlobalLock()
    Dim fso As Object
    Dim lockFile As String

    lockFile = Environ$("TEMP") & "\" & LOCK_FILE_NAME
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(lockFile) Then
        fso.DeleteFile lockFile, True
    End If
End Sub

