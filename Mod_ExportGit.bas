Attribute VB_Name = "Mod_ExportGit"
Option Explicit

Private Const GIT_EXPORT_PATH As String = "C:\Git\GenelAksiyon_VBA\"

Public Sub ExportAllModules_Safe(ByRef ErrorLog As String)
    Dim comp As Object
    Dim fso As Object
    Dim exportFile As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    ErrorLog = ""

    For Each comp In ThisWorkbook.VBProject.VBComponents
        On Error Resume Next

        Select Case comp.Type
            Case 1, 2, 3
                exportFile = GIT_PATH & comp.Name & GetExt(comp.Type)

                If fso.FileExists(exportFile) Then
                    fso.DeleteFile exportFile, True
                End If

                comp.Export exportFile
        End Select

        If Err.Number <> 0 Then
            ErrorLog = ErrorLog & vbCrLf & "• " & comp.Name & " › " & Err.Description
            Err.Clear
        End If

        On Error GoTo 0
    Next comp
End Sub

Private Function GetExt(t As Long) As String
    Select Case t
        Case 1: GetExt = ".bas"
        Case 2, 3: GetExt = ".cls"
    End Select
End Function

Public Sub ExportCommitPush_OneClick()
    Dim errLog As String
    Dim sh As Object
    Dim cmd As String

    ExportAllModules_Safe errLog

    If errLog <> "" Then
        MsgBox "Export sýrasýnda hata oluþtu:" & vbCrLf & errLog, vbCritical
        Exit Sub
    End If

    Set sh = CreateObject("WScript.Shell")

    cmd = "cmd.exe /c cd /d " & GIT_PATH & _
          " && git status --porcelain"

    If sh.Exec(cmd).StdOut.ReadAll = "" Then
        MsgBox "Export tamamlandý. Git farký yok.", vbInformation
        Exit Sub
    End If

    cmd = "cmd.exe /c cd /d " & GIT_PATH & _
          " && git add . && git commit -m ""Auto VBA export"" && git push"

    sh.Run cmd, 0, True

    MsgBox "Export + Commit + Push baþarýyla tamamlandý.", vbInformation
End Sub


Public Sub ExportAllModulesToGit()
    Dim comp As Object
    Dim fso As Object
    Dim exportFile As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1, 2, 3 ' StdModule, Class, ThisWorkbook
                exportFile = GIT_EXPORT_PATH & comp.Name & GetExtension(comp.Type)

                ' Eski dosyayý sil (diff temiz olsun)
                If fso.FileExists(exportFile) Then
                    fso.DeleteFile exportFile, True
                End If

                comp.Export exportFile
        End Select
    Next comp

    MsgBox "Tüm VBA modülleri Git klasörüne export edildi.", vbInformation
End Sub

Private Function GetExtension(vbType As Long) As String
    Select Case vbType
        Case 1: GetExtension = ".bas"
        Case 2: GetExtension = ".cls"
        Case 3: GetExtension = ".cls"
    End Select
End Function

Public Sub GitDiffAfterExport()
    Dim sh As Object
    Dim cmd As String

    cmd = "cmd.exe /k cd /d C:\Git\GenelAksiyon_VBA && git diff"

    Set sh = CreateObject("WScript.Shell")
    sh.Run cmd, 1, False
End Sub

Public Sub ExportAndCheckGitDiff()
    On Error GoTo SAFE_EXIT

    ExportAllModulesToGit
    Application.Wait Now + TimeValue("00:00:01")
    GitDiffAfterExport

SAFE_EXIT:
    ' Excel'in kapanmasýný engellemek için bilinçli boþ exit
End Sub

