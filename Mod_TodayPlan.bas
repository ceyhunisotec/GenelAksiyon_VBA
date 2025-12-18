Attribute VB_Name = "Mod_TodayPlan"

'== Mod_TodayPlan.bas ==
Option Explicit

' --- Yardýmcý: Bugün planlý mý? ---
Private Function IsDueToday(ByVal v As Variant) As Boolean
    If IsDate(v) Then
        IsDueToday = (CLng(CDate(v)) = CLng(Date))
    End If
End Function

' --- Kiþiye göre bugün plan listesi topla (tüm toplantý sayfalarý) ---
Private Function CollectTodayTasksForPerson(ByVal personKey As String) As Collection
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim col As New Collection
    Dim vJ As Variant
    For Each ws In ThisWorkbook.Worksheets
        If IsMeetingSheet(ws) Then ' Koordinasyon, Sipariþ, Þikayet, Atýl_Stok, Kalite
            lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
            For r = 5 To lastRow  ' veri 5. satýrdan baþlýyor
                If Trim$(ws.Cells(r, "F").text) <> "" Then
                    If StrComp(Trim$(ws.Cells(r, "F").text), Trim$(personKey), vbTextCompare) = 0 Then
                        vJ = ws.Cells(r, "J").Value2
                        If IsNumeric(vJ) Then
                            If vJ < 0.99 And IsDueToday(ws.Cells(r, "H").value) Then
                                ' ws, r sakla
                                col.Add Array(ws.Name, r)
                            End If
                        End If
                    End If
                End If
            Next r
        End If
    Next ws
    Set CollectTodayTasksForPerson = col
End Function

' --- Bugün iþ planý: tek mail (TO: kiþi, CC: CB) ---
Private Function SendTodayPlanMail(ByVal personKey As String, ByVal items As Collection) As Boolean
    Dim toAddr As String, fullName As String
    Dim OutApp As Object, OutMail As Object, rec As Object
    Dim html As String, i As Long, it, ws As Worksheet, r As Long
    Dim subjectText As String

    On Error GoTo ErrH
    SendTodayPlanMail = False

    toAddr = ResolveRecipient(personKey) ' kod/ad/email -> email
    If Len(toAddr) = 0 Then Exit Function
    fullName = ResolveFullName(personKey) ' kod/ad/email -> Ad Soyad

    subjectText = "AAA Ýþ Planý - " & fullName & ", tarih " & Format(Date, "dd.MM.yyyy")

    ' --- HTML gövde ---
    html = ""
    html = html & "<p>Sayýn " & fullName & ",</p>"
    html = html & "<p>Bugün için planlanan görevleriniz aþaðýdadýr. Ýyi çalýþmalar.</p>"
    html = html & "<table border='1' cellspacing='0' cellpadding='6' "
    html = html & "style='border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt;'>"
    html = html & "<tr style='background:#f2f2f2;'>"
    html = html & "<th align='right'>No</th>"
    html = html & "<th align='left'>Toplantý</th>"
    html = html & "<th align='left'>Görev</th>"
    html = html & "<th align='left'>Termin</th>"
    html = html & "<th align='left'>Sorumlu</th>"
    html = html & "<th align='left'>Durum</th>"
    html = html & "</tr>"

    For i = 1 To items.Count
        it = items(i): Set ws = ThisWorkbook.Worksheets(CStr(it(0))): r = CLng(it(1))
        html = html & "<tr>"
        html = html & "<td align='right'>" & i & "</td>"
        html = html & "<td>" & ws.Name & "</td>"
        html = html & "<td>" & ws.Cells(r, "E").text & "</td>"
        html = html & "<td>" & Format(ws.Cells(r, "H").value, "dd.MM.yyyy") & "</td>"
        html = html & "<td>" & ws.Cells(r, "F").text & "</td>"
        html = html & "<td>" & Format(ws.Cells(r, "J").value, "0%") & "</td>"
        html = html & "</tr>"
    Next i
    html = html & "</table>"
    html = html & "<p>Dokümana hýzlý eriþim: <a href='" & ONEDRIVE_LINK & "'>buraya týklayýn</a>.</p>"
    html = html & "<p>Saygýlarýmla</p>"

    ' --- Gönderim ---
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        Set rec = .recipients.Add(toAddr): rec.Resolve
        ' Sadece CB CC olacak (talebiniz gereði)
        With .recipients.Add("ceyhun@isotec.com.tr"): .Type = 2: .Resolve: End With
        .Subject = subjectText
        .BodyFormat = 2
        .htmlBody = html
        .Save
        .Send
    End With

    SendTodayPlanMail = True

CleanExit:
    On Error Resume Next
    Set rec = Nothing: Set OutMail = Nothing: Set OutApp = Nothing
    Exit Function
ErrH:
    Resume CleanExit
End Function

' --- Tüm kiþilere gönder: H=today olan görevlerden benzersiz sorumlularý bul ---
Public Sub SendTodayPlansForAll()
    Dim dict As Object, ws As Worksheet, lastRow As Long, r As Long
    Dim person As String, vJ As Variant
    Set dict = CreateObject("Scripting.Dictionary")

    For Each ws In ThisWorkbook.Worksheets
        If IsMeetingSheet(ws) Then
            lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
            For r = 5 To lastRow
                If IsDueToday(ws.Cells(r, "H").value) Then
                    vJ = ws.Cells(r, "J").Value2
                    If IsNumeric(vJ) And vJ < 0.99 Then
                        person = Trim$(ws.Cells(r, "F").text)
                        If Len(person) > 0 Then dict(person) = True
                    End If
                End If
            Next r
        End If
    Next ws

    ' Kiþi kiþi gönder
    Dim k As Variant, items As Collection
    For Each k In dict.keys
        Set items = CollectTodayTasksForPerson(CStr(k))
        If items.Count > 0 Then Call SendTodayPlanMail(CStr(k), items)
    Next k
End Sub
