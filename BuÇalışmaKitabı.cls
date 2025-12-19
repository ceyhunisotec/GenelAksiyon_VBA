VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BuÇalýþmaKitabý"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'== ThisWorkbook ==
Private Sub Workbook_Open()
    Set AppEv = New CAppEvents
    Set AppEv.App = Application

    On Error Resume Next
    Call RebuildDashboard
    On Error GoTo 0

    ' Önce eski schedule varsa temizle (çift planlamayý engeller)
    Call CancelDailySchedule
    Call CancelDailyTodayPlans

    ' Sonra yeniden kur
    Call ScheduleDailyOverdueReports   ' 09:00 + 17:00 artýk burada kurulacak
    Call ScheduleDailyTodayPlans       ' 08:15 günlük iþ planý
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Set AppEv.App = Nothing
    Set AppEv = Nothing

    ' Zamanlamayý iptal et (güvenli)
    Call CancelDailySchedule
    Call CancelDailyTodayPlans
End Sub
