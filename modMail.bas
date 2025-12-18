Attribute VB_Name = "modMail"

'=== Modül: modMail ===
Option Explicit

Sub AskAI_SendMail()
    Dim ans As String: ans = Range("rngAI_Answer").value
    If Len(Trim(ans)) = 0 Then
        MsgBox "Gönderilecek bir AI cevabý yok.", vbExclamation: Exit Sub
    End If
    
    Dim toAddr As String
    toAddr = InputBox("TO e-posta adresi (ör: gokhan.baskoy@isotec.com.tr):", "Mail Gönder")
    If Len(Trim(toAddr)) = 0 Then Exit Sub
    
    Dim ccAddr As String: ccAddr = GetEmailByCode("CB") ' Data sayfasýndan CB emailini al
    Dim subj As String: subj = "AI Özeti - " & Format(Date, "dd.MM.yyyy")
    Dim htmlBody As String
    
    htmlBody = "<div style='font-family:Segoe UI;'>" & _
               "<h3 style='color:#0078D4;margin-bottom:4px;'>Toplantý/Ýþ Planý AI Özeti</h3>" & _
               "<p><strong>Tarih:</strong> " & Format(Date, "dd.MM.yyyy") & "</p>" & _
               "<hr style='border:0;border-top:1px solid #e1e1e1;'/>" & _
               "<pre style='white-space:pre-wrap; font-size:14px;'>" & HtmlEncode(ans) & "</pre>" & _
               "</div>"
    
    SendViaOutlook toAddr, ccAddr, subj, htmlBody
    
    LogToSysLog toAddr, "Gündem", "AI_Mail"
    MsgBox "E-posta gönderildi.", vbInformation
End Sub

Private Function GetEmailByCode(code As String) As String
    ' Data sayfasý: Kod, Adý Soyadý, Görevi, Mail Adresi...
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Data")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).value) = code Then
            GetEmailByCode = ExtractEmail(ws.Cells(i, 4).value)
            Exit For
        End If
    Next i
End Function

Private Function ExtractEmail(textVal As String) As String
    ' Hücrede [name@domain] formatý varsa e-posta adresini ayýklar
    Dim s As String: s = textVal
    s = Replace(s, "mailto:", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    ExtractEmail = s
End Function

Private Sub SendViaOutlook(toAddr As String, ccAddr As String, subj As String, htmlBody As String)
    Dim olApp As Object, mail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set mail = olApp.CreateItem(0)
    With mail
        .To = toAddr
        If Len(Trim(ccAddr)) > 0 Then .CC = ccAddr
        .Subject = subj
        .htmlBody = htmlBody
        .Send
    End With
End Sub

Private Function HtmlEncode(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    HtmlEncode = s
End Function


