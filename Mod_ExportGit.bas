Attribute VB_Name = "Mod_ExportGit"
Option Explicit

Private Const GIT_EXPORT_PATH As String = "C:\Git\GenelAksiyon_VBA\"

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

    cmd = "cmd /c cd /d C:\Git\GenelAksiyon_VBA && git diff"

    Set sh = CreateObject("WScript.Shell")
    sh.Run cmd, 1, False
End Sub

Public Sub ExportAndCheckGitDiff()
    ExportAllModulesToGit
    Application.Wait Now + TimeValue("00:00:01")
    GitDiffAfterExport
End Sub

