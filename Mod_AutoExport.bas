Attribute VB_Name = "Mod_AutoExport"
Option Explicit
Public gNextAutoExport As Date

Private Function IsWorkingHours() As Boolean
    Dim h As Integer
    h = Hour(Now)
    IsWorkingHours = (Weekday(Now, vbMonday) <= 5) And (h >= 8 And h < 18)
End Function

Public Sub StartAutoExport()
    ScheduleNext
End Sub

Private Sub ScheduleNext()
    gNextAutoExport = Now + TimeSerial(0, 30, 0)
    Application.onTime gNextAutoExport, "AutoExportRunner"
End Sub

Public Sub AutoExportRunner()
    If IsWorkingHours Then
        ExportCommitPush_Smart
    End If
    ScheduleNext
End Sub

Public Sub StopAutoExport()
    On Error Resume Next
    Application.onTime gNextAutoExport, "AutoExportRunner", , False
    On Error GoTo 0
End Sub

