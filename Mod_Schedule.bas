Attribute VB_Name = "Mod_Schedule"

'== Mod_Schedule.bas ==
Option Explicit

'==========================
'  G L O B A L   T I M E R S
'==========================
' 09:00 ve 17:00 geciken rapor slotlarý
Public gNextRun09 As Date
Public gNextRun17 As Date

' 08:15 "Bugün Ýþ Planý" zamanlayýcý
Public gNextRunPlan As Date

'==========================
'  H E L P E R
'==========================
' Haftaiçi ise ayný gün/sýradaki gün; hafta sonu ise Pazartesi planlar
Private Function NextWeekdayDateTime(ByVal timeStr As String) As Date
    Dim wd As Integer
    Dim t As Date
    wd = Weekday(Date, vbMonday)          ' 1=Mon ... 7=Sun
    t = Date + TimeValue(timeStr)

    If wd <= 5 Then
        If Now > t Then
            NextWeekdayDateTime = Date + 1 + TimeValue(timeStr)
        Else
            NextWeekdayDateTime = t
        End If
    Else
        ' Cumartesi(6) -> +2, Pazar(7) -> +1
        NextWeekdayDateTime = Date + (8 - wd) + TimeValue(timeStr)
    End If
End Function

Private Sub ScheduleOverdueSlot(ByVal slotTime As String, ByVal slotName As String, ByRef nextRun As Date)
    nextRun = NextWeekdayDateTime(slotTime)

    If slotName = "09" Then
        Application.onTime EarliestTime:=nextRun, Procedure:="RunDailyOverdueReports_09", Schedule:=True
    Else
        Application.onTime EarliestTime:=nextRun, Procedure:="RunDailyOverdueReports_17", Schedule:=True
    End If
End Sub

'==========================
'  O V E R D U E   (09:00 / 17:00)
'==========================
' Workbook_Open içinden çaðrýlýr
Public Sub ScheduleDailyOverdueReports()
    ' Çift planlamayý engelle (önce iptal, sonra kur)
    CancelDailySchedule

    ScheduleOverdueSlot "09:00:00", "09", gNextRun09
    ScheduleOverdueSlot "17:00:00", "17", gNextRun17
End Sub

Public Sub RunDailyOverdueReports_09()
    ' NOTE: Bu prosedür Mod_Report içinde olmalý:
    '       RunDailyOverdueReports_WithSlot(slot As String)
    RunDailyOverdueReports_WithSlot "09"

    ' Bir sonraki 09:00'u kur
    ScheduleOverdueSlot "09:00:00", "09", gNextRun09
End Sub

Public Sub RunDailyOverdueReports_17()
    RunDailyOverdueReports_WithSlot "17"

    ' Bir sonraki 17:00'yi kur
    ScheduleOverdueSlot "17:00:00", "17", gNextRun17
End Sub

' Dashboard butonu için tek entrypoint (istersen iki turu da gönderir)
Public Sub RunOverdueReportsNow()
    ' Ýstersen sadece akþam turu:
    'RunDailyOverdueReports_WithSlot "17"

    ' Ýstersen iki turu da manuel gönder:
    RunDailyOverdueReports_WithSlot "09"
    RunDailyOverdueReports_WithSlot "17"
End Sub

Public Sub CancelDailySchedule()
    On Error Resume Next
    If gNextRun09 <> 0 Then
        Application.onTime EarliestTime:=gNextRun09, Procedure:="RunDailyOverdueReports_09", Schedule:=False
    End If
    If gNextRun17 <> 0 Then
        Application.onTime EarliestTime:=gNextRun17, Procedure:="RunDailyOverdueReports_17", Schedule:=False
    End If
    gNextRun09 = 0
    gNextRun17 = 0
End Sub

'==========================
'  T O D A Y   P L A N  (08:15)
'==========================
Public Sub ScheduleDailyTodayPlans()
    ' Önce iptal (çift planlama engeli)
    CancelDailyTodayPlans

    gNextRunPlan = NextWeekdayDateTime("08:15:00")
    Application.onTime EarliestTime:=gNextRunPlan, Procedure:="RunDailyTodayPlans", Schedule:=True
End Sub

Public Sub RunDailyTodayPlans()
    SendTodayPlansForAll
    ScheduleDailyTodayPlans
End Sub

Public Sub CancelDailyTodayPlans()
    On Error Resume Next
    If gNextRunPlan <> 0 Then
        Application.onTime EarliestTime:=gNextRunPlan, Procedure:="RunDailyTodayPlans", Schedule:=False
       End If
    gNextRunPlan = 0
End Sub
