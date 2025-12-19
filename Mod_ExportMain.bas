Attribute VB_Name = "Mod_ExportMain"
Option Explicit

Public Sub ExportCommitPush_Smart()
    Dim errLog As String
    Dim changedFiles As String
    Dim sh As Object, execObj As Object
    Dim cmd As String
    MsgBox "Makro baþladý", vbInformation

    If Not AcquireGlobalLock Then Exit Sub
    On Error GoTo CLEANUP

    ' 1?? Her zaman export
    ExportAllModules_Stable errLog

    If errLog <> "" Then
        MsgBox "Export sýrasýnda hata oluþtu:" & vbCrLf & errLog, vbCritical
        GoTo CLEANUP
    End If

    ' 2?? Git deðiþiklik kontrolü
    changedFiles = GetGitChangedFiles()

    If changedFiles = "" Then
        MsgBox "Export yapýldý. Git farký yok.", vbInformation
        GoTo CLEANUP
    End If

    ' 3?? Commit + Push
    cmd = "cmd.exe /c cd /d " & GIT_PATH & _
          " && git add ." & _
          " && git commit -m ""Auto export: " & changedFiles & """" & _
          " && git push"

    Set sh = CreateObject("WScript.Shell")
    Set execObj = sh.Exec(cmd)

    Do While execObj.status = 0
        DoEvents
    Loop

    MsgBox "Export + Commit + Push tamamlandý:" & vbCrLf & changedFiles, vbInformation

CLEANUP:
    ReleaseGlobalLock
End Sub


