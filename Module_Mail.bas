Attribute VB_Name = "Module_Mail"

'== Module_Mail.bas ==
Option Explicit

' === Yardýmcý: Data hücrelerinden çýkan e-postayý temiz ve tek adres haline getir ===
Private Function CleanAddress(ByVal s As String) As String
    Dim startB As Long, endB As Long

    s = Trim$(s)
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")

    If InStr(1, s, "mailto:", vbTextCompare) > 0 Then
        s = Replace$(s, "mailto:", "", , , vbTextCompare)
    End If

    startB = InStr(s, "["): endB = InStr(s, "]")
    If startB > 0 And endB > startB Then
        s = Mid$(s, startB + 1, endB - startB - 1)
    End If

    If InStr(s, "(") > 0 Then s = left$(s, InStr(s, "(") - 1)

    s = Replace$(s, "<", "")
    s = Replace$(s, ">", "")

    If InStr(s, " ") > 0 Then s = left$(s, InStr(s, " ") - 1)

    CleanAddress = Trim$(s)
End Function

' === Basit e-posta doðrulamasý: "x@y.z" desenini ve illegal karakterleri kontrol et ===
Private Function IsValidEmail(ByVal addr As String) As Boolean
    Dim atPos As Long, dotPos As Long

    addr = Trim$(addr)
    If Len(addr) = 0 Then Exit Function
    If InStr(addr, " ") > 0 Then Exit Function
    If InStr(addr, vbCr) > 0 Or InStr(addr, vbLf) > 0 Or InStr(addr, vbTab) > 0 Then Exit Function

    atPos = InStr(addr, "@"): If atPos = 0 Then Exit Function
    dotPos = InStr(atPos + 1, addr, "."): If dotPos = 0 Then Exit Function
    If left$(addr, 1) = "." Or right$(addr, 1) = "." Then Exit Function

    IsValidEmail = True
End Function

' === F sütunundaki Sorumlu girdisini e-postaya çevir: kod / ad / direkt e-posta desteklenir ===
Public Function ResolveRecipient(ByVal responsibleInput As String) As String
    Dim ws As Worksheet, f As Range, s As String, candidate As String

    s = Trim$(responsibleInput)
    If Len(s) = 0 Then Exit Function

    If InStr(1, s, "@") > 0 Then
        candidate = CleanAddress(s)
        If IsValidEmail(candidate) Then ResolveRecipient = candidate: Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets("Data")

    Set f = ws.Columns("A").Find(What:=s, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then
        candidate = CleanAddress(CStr(f.Offset(0, 3).value))  ' D sütunu: Mail
        If IsValidEmail(candidate) Then ResolveRecipient = candidate: Exit Function
    End If

    Set f = ws.Columns("B").Find(What:=s, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not f Is Nothing Then
        candidate = CleanAddress(CStr(f.Offset(0, 2).value))  ' D sütunu: Mail
        If IsValidEmail(candidate) Then ResolveRecipient = candidate: Exit Function
    End If

    ResolveRecipient = ""
End Function

' === GÖNDERÝM: Recipients.Add ile güvenli ekleme; TRUE dönerse gerçekten gönderilmiþtir ===
Public Function SendTaskMail(ByVal ws As Worksheet, ByVal r As Long, Optional ByVal oneDriveLink As String = "") As Boolean
    Dim OutApp As Object, OutMail As Object, recp As Object
    Dim subjectText As String, toAddr As String
    Dim html As String, ccSplit As Variant, i As Long

    On Error GoTo ErrHandler
    SendTaskMail = False

    If Len(oneDriveLink) = 0 Then oneDriveLink = ONEDRIVE_LINK

    toAddr = ResolveRecipient(CStr(ws.Cells(r, "F").value))
    If Len(toAddr) = 0 Then
        MsgBox "Mail gönderilemedi: F hücresindeki deðer için geçerli e-posta bulunamadý.", vbExclamation, "Alýcý yok"
        Exit Function
    End If

    subjectText = "Yeni Görev Eklendi / " & ws.Name & " Toplantýsý"
    
   
    ' --- TABLOYU GÜNCELLE ---
    html = ""
    html = html & "<p>Sayýn Yetkili,</p>"
    html = html & "<p>Tarafýnýza görev atanmýþtýr, lütfen iþ planýnýza alýnýz.</p>"
    html = html & "<table border='1' cellspacing='0' cellpadding='6' style='border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt;'>"
    html = html & "<tr style='background:#f2f2f2;'>"
    html = html & "<th align='left'>Konu/Madde (D)</th>"
    html = html & "<th align='left'>Görev/Aksiyon (E)</th>"
    html = html & "<th align='left'>Planlanan Tarih (H)</th>"
    html = html & "<th align='left'>Sorumlu (F)</th>"
    html = html & "</tr><tr>"
    html = html & "<td>" & ws.Cells(r, "D").text & "</td>"
    html = html & "<td>" & ws.Cells(r, "E").text & "</td>"
    html = html & "<td>" & ws.Cells(r, "H").text & "</td>"
    html = html & "<td>" & ws.Cells(r, "F").text & "</td>"
    html = html & "</tr></table>"
    html = html & "<p>Dokümana hýzlý eriþim: <a href='" & oneDriveLink & "'>buraya týklayýn</a>.</p>"
    html = html & "<p>Saygýlarýmýzla</p>"
    

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        ' Alýcý
        Set recp = .recipients.Add(toAddr)
        recp.Resolve

        ' CC'ler
        ccSplit = Split(CC_LIST, ";")
        For i = LBound(ccSplit) To UBound(ccSplit)
            If Len(Trim$(ccSplit(i))) > 0 Then
                Set recp = .recipients.Add(Trim$(ccSplit(i)))
                recp.Type = 2   ' olCC
                recp.Resolve
            End If
        Next i

        .Subject = subjectText
        .BodyFormat = 2        ' olFormatHTML (late binding)
        .htmlBody = html

        '.Display              ' TEST için açýn
        .Save
        .Send                  ' canlý gönderim
    End With

    SendTaskMail = True

CleanUp:
    On Error Resume Next
    Set recp = Nothing
    Set OutMail = Nothing
    Set OutApp = Nothing
    Exit Function

ErrHandler:
    MsgBox "Mail gönderimi sýrasýnda hata: " & Err.Description, vbCritical, "Outlook hatasý"
    Resume CleanUp
End Function

' === Tamamlanan görev bildirimi: CB & HAO'ya gönder ===
'== Module_Mail.bas -- SendCompletionMail güncel ==
Public Function SendCompletionMail(ByVal ws As Worksheet, ByVal r As Long, Optional ByVal oneDriveLink As String = "") As Boolean
    Dim OutApp As Object, OutMail As Object
    Dim subjectText As String, html As String, kanitText As String
    Dim toAddr As String

    On Error GoTo ErrHandler
    SendCompletionMail = False
    If Len(oneDriveLink) = 0 Then oneDriveLink = ONEDRIVE_LINK

    toAddr = ResolveRecipient(CStr(ws.Cells(r, "F").value)) ' F=Sorumlu -> e-posta
    If Len(toAddr) = 0 Then Exit Function

    subjectText = "Görev Tamamlandý / " & ws.Name & " Toplantýsý"
    kanitText = Trim$(ws.Cells(r, "K").text): If Len(kanitText) = 0 Then kanitText = "—"

    html = ""
    html = html & "<p>Sayýn Yetkili,</p>"
    html = html & "<p>Aþaðýdaki görev <b>tamamlanmýþtýr</b>:</p>"
    html = html & "<table border='1' cellspacing='0' cellpadding='6' style='border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt;'>"
    html = html & "<tr style='background:#f2f2f2;'><th align='left'>Konu/Madde (D)</th><th align='left'>Görev/Aksiyon (E)</th><th align='left'>Planlanan (H)</th><th align='left'>Sorumlu (F)</th><th align='left'>Bitiþ (I)</th></tr>"
    html = html & "<tr><td>" & ws.Cells(r, "D").text & "</td><td>" & ws.Cells(r, "E").text & "</td><td>" & ws.Cells(r, "H").text & "</td><td>" & ws.Cells(r, "F").text & "</td><td>" & ws.Cells(r, "I").text & "</td></tr>"
    html = html & "</table>"
    html = html & "<p><b>Kanýt</b>: " & kanitText & "</p>"
    html = html & "<p>Dokümana hýzlý eriþim: <a href='" & oneDriveLink & "'>buraya týklayýn</a>.</p>"
    html = html & "<p>Bilginize sunarým.</p>"

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        .recipients.Add toAddr ' TO: Sorumlu
        ' CC: CB ve HAO
        With .recipients.Add("ceyhun@isotec.com.tr"): .Type = 2: .Resolve: End With
        With .recipients.Add("halisahmet.orhan@isotec.com.tr"): .Type = 2: .Resolve: End With
        .Subject = subjectText
        .BodyFormat = 2
        .htmlBody = html
        .Save
        .Send
    End With
    SendCompletionMail = True

CleanExit:
    On Error Resume Next
    Set OutMail = Nothing
    Set OutApp = Nothing
    Exit Function
    
ErrHandler:
    MsgBox "Tamamlanan görev mailinde hata: " & Err.Description, vbCritical, "Outlook hatasý"
    Resume CleanExit
    
End Function

' -- Rapor maillerinde CC kuralýný uygula --
Public Sub AddCC_ForReport(ByVal mailItem As Object, ByVal sheetName As String)
    Dim addr As String

    ' CB ve HAO her zaman CC
    addr = ResolveRecipient("CB"): If Len(addr) = 0 Then addr = "ceyhun@isotec.com.tr"
    With mailItem.recipients.Add(addr): .Type = 2: .Resolve: End With

    addr = ResolveRecipient("HAO"): If Len(addr) = 0 Then addr = "halisahmet.orhan@isotec.com.tr"
    With mailItem.recipients.Add(addr): .Type = 2: .Resolve: End With
    
    ' Koordinasyon sayfasý ise EÖ ve ZÖ de CC
    If LCase$(sheetName) = LCase$("Koordinasyon") Or LCase$(sheetName) = LCase$("Koordinasyon") Then
        addr = ResolveRecipient("EÖ"): If Len(addr) = 0 Then addr = "erkan@isotec.com.tr"
        With mailItem.recipients.Add(addr): .Type = 2: .Resolve: End With

        addr = ResolveRecipient("ZÖ"): If Len(addr) = 0 Then addr = "zahide@isotec.com.tr"
        With mailItem.recipients.Add(addr): .Type = 2: .Resolve: End With
    End If

    ' Sipariþ sayfasý ise EÖ, ZÖ, EG ve HB de CC
    If LCase$(sheetName) = LCase$("Sipariþ") Or LCase$(sheetName) = LCase$("Siparis") Then
        addr = ResolveRecipient("EÖ"): If Len(addr) = 0 Then addr = "erkan@isotec.com.tr"
        With mailItem.recipients.Add(addr): .Type = 2: .Resolve: End With

        addr = ResolveRecipient("ZÖ"): If Len(addr) = 0 Then addr = "zahide@isotec.com.tr"
        With mailItem.recipients.Add(addr): .Type = 2: .Resolve: End With
        
        addr = ResolveRecipient("EG"): If Len(addr) = 0 Then addr = "eray.guzel@isotec.com.tr"
        With mailItem.recipients.Add(addr): .Type = 2: .Resolve: End With

        addr = ResolveRecipient("HB"): If Len(addr) = 0 Then addr = "hakan.bildirici@isotec.com.tr"
        With mailItem.recipients.Add(addr): .Type = 2: .Resolve: End With
    End If
End Sub

