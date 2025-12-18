Attribute VB_Name = "modAI"

'=== Cloudflare Workers AI (Run Model) ===
Option Explicit
Private Const CF_MODEL As String = "@cf/meta/llama-3-8b-instruct"  ' sizin curl örneðinizle ayný

Public Function AskAI_Cloudflare(promptText As String) As String
    On Error GoTo ErrHandler

    Dim token As String: token = GetCFToken()
    If Len(token) = 0 Then
        AskAI_Cloudflare = "Token boþ: CF_API_TOKEN okunamadý."
        Exit Function
    End If

    Dim accountId As String: accountId = GetCFAccountId()
    If Len(accountId) <> 32 Then
        AskAI_Cloudflare = "Geçersiz CF_ACCOUNT_ID: " & accountId
        Exit Function
    End If

    Dim url As String
    url = "https://api.cloudflare.com/client/v4/accounts/" & accountId & "/ai/run/" & CF_MODEL

    ' ? Sizin curl örneðiniz gibi "messages" ile gönderelim:
    Dim payload As String
    payload = "{""messages"":[{""role"":""system"",""content"":""Sen bir aksiyon takip asistanýsýn. Türkçe ve kýsa cevap ver.""}," & _
              "{""role"":""user"",""content"":""" & EscapeJson(promptText) & """}]}"

    Dim status As Long
    Dim resp As String
    resp = HttpPostJson_Bearer(url, token, payload, status)

    ' Status 200 deðilse bile, JSON'da success:true ise baþarýlý kabul et
    If status <> 200 Then
        If InStr(1, resp, """success"":true", vbTextCompare) > 0 Then
            AskAI_Cloudflare = ParseCloudflareResult(resp)
            Exit Function
        End If

        AskAI_Cloudflare = "Cloudflare HTTP " & status & vbCrLf & resp
        Exit Function
    End If

    AskAI_Cloudflare = ParseCloudflareResult(resp)
    Exit Function
    
    ' Parse sonucu boþsa ham yanýtý göster (debug için)
    If Len(Trim$(AskAI_Cloudflare)) = 0 Then
        AskAI_Cloudflare = "AI yanýtý boþ döndü. Ham cevap:" & vbCrLf & left$(resp, 1500)
    End If

ErrHandler:
       AskAI_Cloudflare = "VBA Error: " & Err.Number & " - " & Err.Description

End Function

Private Function GetCFAccountId() As String
    On Error Resume Next
    
    Dim s As String: s = ""
    
    ' 1) Önce named range’den dene
    s = CStr(ThisWorkbook.Names("CF_ACCOUNT_ID").RefersToRange.value)
    
    ' 2) Yoksa Config!B2’den oku
    If Len(Trim$(s)) = 0 Then
        s = CStr(ThisWorkbook.Worksheets("Config").Range("B2").value)
    End If
    
    ' Temizlik
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, "<", "")
    s = Replace(s, ">", "")
    s = Trim$(s)
    
    GetCFAccountId = s
End Function

Private Function IsValidAccountId(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) <> 32 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr("0123456789abcdefABCDEF", ch) = 0 Then Exit Function
    Next i
       IsValidAccountId = True
End Function


Private Function GetCFToken() As String
    On Error Resume Next
    Dim t As String
    t = ThisWorkbook.Names("CF_API_TOKEN").RefersToRange.value

    ' boþluk ve satýr sonlarýný temizle
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Trim$(t)

    GetCFToken = t
End Function

Private Function HttpPostJson_Bearer(url As String, token As String, jsonBody As String, ByRef statusCode As Long) As String
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Accept", "application/json"
    http.Send jsonBody

    ' Yanýtýn gerçekten geldiðinden emin ol
    http.WaitForResponse 60

    On Error Resume Next
    statusCode = CLng(http.status)
    On Error GoTo 0

    ' ? ResponseText yerine ResponseBody (byte) alýp UTF-8 decode et
    HttpPostJson_Bearer = Utf8BytesToString(http.ResponseBody)
End Function

Private Function Utf8BytesToString(ByVal bytes As Variant) As String
    ' bytes: WinHTTP ResponseBody -> Byte()
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 1 ' adTypeBinary
    stm.Open
    stm.Write bytes
    stm.Position = 0

    stm.Type = 2 ' adTypeText
    stm.Charset = "utf-8"
    Utf8BytesToString = stm.ReadText

    stm.Close

End Function


Private Function ParseCloudflareResult(respJson As String) As String
    Dim p As Long, q As Long
    p = InStr(1, respJson, """response"":")
    If p = 0 Then
        ParseCloudflareResult = "Cevap okunamadý (response alaný yok): " & left$(respJson, 500)
        Exit Function
    End If

    p = InStr(p + 11, respJson, """") + 1
    q = InStr(p, respJson, """")

    ParseCloudflareResult = Replace(Mid$(respJson, p, q - p), "\n", vbCrLf)
End Function

Private Function EscapeJson(ByVal text As String) As String
    text = Replace(text, "\", "\\")
    text = Replace(text, """", "\""")
    text = Replace(text, vbCrLf, "\n")
    text = Replace(text, vbCr, "\n")
    text = Replace(text, vbLf, "\n")
    EscapeJson = text
End Function


