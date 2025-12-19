Attribute VB_Name = "Mod_SmartExport"
Option Explicit


Public Function ExportAllModules_Stable(ByRef ErrorLog As String) As Boolean
    Dim comp As Object
    Dim fso As Object
    Dim exportFile As String
    MsgBox "Export baþlýyor", vbInformation

    Set fso = CreateObject("Scripting.FileSystemObject")
    ErrorLog = ""

    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 1 Or comp.Type = 2 Or comp.Type = 3 Then
            On Error Resume Next
            
            MsgBox "Git commit/push aþamasýna geçiliyor", vbInformation
            exportFile = GIT_PATH & comp.Name & GetExt(comp.Type)
            If fso.FileExists(exportFile) Then fso.DeleteFile exportFile, True
            comp.Export exportFile
            Debug.Print "Export edildi: " & comp.Name

            If Err.Number <> 0 Then
                ErrorLog = ErrorLog & vbCrLf & "• " & comp.Name & " › " & Err.Description
                Err.Clear
            End If

            On Error GoTo 0
        End If
    Next comp

    ExportAllModules_Stable = True
    MsgBox "Export bitti", vbInformation
End Function

Public Function GetGitChangedFiles() As String
    Dim sh As Object, execObj As Object
    Dim cmd As String, output As String

    cmd = "cmd.exe /c cd /d " & GIT_PATH & " && git status --porcelain"

    Set sh = CreateObject("WScript.Shell")
    Set execObj = sh.Exec(cmd)

    Do While execObj.status = 0
        DoEvents
    Loop

    output = execObj.StdOut.ReadAll

    If Trim(output) = "" Then Exit Function

    Dim lines() As String, i As Long
    lines = Split(output, vbCrLf)

    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            GetGitChangedFiles = GetGitChangedFiles & Mid(lines(i), 4) & ", "
        End If
    Next i

    If GetGitChangedFiles <> "" Then
        GetGitChangedFiles = left(GetGitChangedFiles, Len(GetGitChangedFiles) - 2)
    End If
End Function


Private Function GitDiffFile(filePath As String) As String
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    GitDiffFile = sh.Exec( _
        "cmd.exe /c cd /d " & GIT_PATH & _
        " && git diff -- """ & filePath & """").StdOut.ReadAll
End Function

Private Function GetExt(t As Long) As String
    If t = 1 Then GetExt = ".bas" Else GetExt = ".cls"
End Function


