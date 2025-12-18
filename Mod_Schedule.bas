Attribute VB_Name = "Mod_Schedule"

'== Mod_Schedule.bas ==
Option Explicit

' 17:00 "Geciken Görevler" zamanlayýcý için global randevu
Public gNextRun As Date

' 08:15 "Bugün Ýþ Planý" zamanlayýcý için global randevu
Public gNextRunPlan As Date

' ---------- 17:00 GECÝKEN GÖREV RAPORLARI ----------
Public Sub ScheduleDailyOverdueReports()
    Dim runTime As Date, wd As Integer
    wd = Weekday(Date, vbMonday)
    runTime = Date + TimeValue("17:00:00")
    If wd <= 5 Then
        If Now > runTime Then runTime = Date + 1 + TimeValue("17:00:00")
    Else
        runTime = Date + (8 - wd) + TimeValue("17:00:00")
    End If
    gNextRun = runTime
    Application.onTime EarliestTime:=gNextRun, Procedure:="RunDailyOverdueReports", Schedule:=True
End Sub

Public Sub RunDailyOverdueReports()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsMeetingSheet(ws) Then
            SendOverdueReportForSheet ws
        End If
    Next ws
    ScheduleDailyOverdueReports  ' bir sonraki randevuyu kur
End Sub

Public Sub RunOverdueReportsNow()
    RunDailyOverdueReports
End Sub

Public Sub CancelDailySchedule()
    On Error Resume Next
    If gNextRun <> 0 Then
        Application.onTime EarliestTime:=gNextRun, Procedure:="RunDailyOverdueReports", Schedule:=False
        gNextRun = 0
    End If
End Sub

' ---------- 08:15 BUGÜN ÝÞ PLANI ----------
Public Sub ScheduleDailyTodayPlans()
    Dim runTime As Date, wd As Integer
    wd = Weekday(Date, vbMonday)
    runTime = Date + TimeValue("08:15:00")
    If wd <= 5 Then
        If Now > runTime Then runTime = Date + 1 + TimeValue("08:15:00")
    Else
        runTime = Date + (8 - wd) + TimeValue("08:15:00")
    End If
    gNextRunPlan = runTime
    Application.onTime EarliestTime:=gNextRunPlan, Procedure:="RunDailyTodayPlans", Schedule:=True
End Sub

Public Sub RunDailyTodayPlans()
    SendTodayPlansForAll
    ScheduleDailyTodayPlans  ' bir sonraki randevuyu kur
End Sub

Public Sub CancelDailyTodayPlans()
       On Error Resume Next
    If gNextRunPlan <> 0 Then
        Application.onTime EarliestTime:=gNextRunPlan, Procedure:="RunDailyTodayPlans", Schedule:=False
        gNextRunPlan = 0
    End If
End Sub
