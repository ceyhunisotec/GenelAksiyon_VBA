Attribute VB_Name = "modFlow"

'=== Module: modFlow ===

Option Explicit

Public Sub RepairAgendaNamedRanges()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Gündem")

    Application.ScreenUpdating = False

    ' 1) Eski isimleri sil (varsa)
    On Error Resume Next
    ThisWorkbook.Names("rngAI_Question").Delete
    ThisWorkbook.Names("rngAI_Answer").Delete
    On Error GoTo 0

    ' 2) UI'daki hücreleri garanti merge yap
    ws.Range("B4:H5").UnMerge
    ws.Range("B4:H5").Merge
    ws.Range("B4:H5").WrapText = True

    ws.Range("B9:H30").UnMerge
    ws.Range("B9:H30").Merge
    ws.Range("B9:H30").WrapText = True
    ws.Range("B9:H30").VerticalAlignment = xlVAlignTop

    ' 3) Ýsimleri yeniden oluþtur (workbook-scope)
    ' Ýsimleri merge alanýn sol-üst hücresine baðlamak en saðlýklýsýdýr.
    ThisWorkbook.Names.Add Name:="rngAI_Question", RefersTo:=ws.Range("B4")
    ThisWorkbook.Names.Add Name:="rngAI_Answer", RefersTo:=ws.Range("B9")

    Application.ScreenUpdating = True

    MsgBox "rngAI_Question ve rngAI_Answer yeniden oluþturuldu." & vbCrLf & _
           "Þimdi Gündem'de Soruyu Çalýþtýr'ý tekrar deneyin.", vbInformation
End Sub


Sub AskAI_Run()
    Dim ws As Worksheet, wsc As Worksheet
    Dim lastStep As String
    
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Worksheets("Gündem")
    Set wsc = ThisWorkbook.Worksheets("Config")
    
    lastStep = "1) Baþladý"
    wsc.Range("D9").value = lastStep
    
    ' 1) Soru oku
    lastStep = "2) Soru okunuyor (rngAI_Question)"
    wsc.Range("D9").value = lastStep
    
    Dim q As String
    q = CStr(ws.Range("rngAI_Question").MergeArea.Cells(1, 1).value)
    
    If Len(Trim$(q)) = 0 Then
        MsgBox "Lütfen bir soru yazýn.", vbInformation
        Exit Sub
    End If
    
    ' 2) Baðlam üret
    lastStep = "3) BuildContextSummary çalýþýyor"
    wsc.Range("D9").value = lastStep
    
    Dim ctx As String
    ctx = BuildContextSummaryDetailed()
    
    ' 3) Prompt hazýrla
    lastStep = "4) BuildSmartPrompt çalýþýyor"
    wsc.Range("D9").value = lastStep
    
    Dim promptAll As String
    promptAll = BuildSmartPrompt(q, ctx)
    
    ' 4) AI çaðrýsý
    lastStep = "5) AskAI_Cloudflare çaðrýsý"
    wsc.Range("D9").value = lastStep
    
    Application.StatusBar = "AI yanýt üretiyor..."
    Dim ans As String
    ans = AskAI_Cloudflare(promptAll)
    Application.StatusBar = False
    
    ' 5) Cevabý yaz
    lastStep = "6) Cevap yazýlýyor (rngAI_Answer)"
    wsc.Range("D9").value = lastStep
    
    SafeWriteMerged ws.Range("rngAI_Answer"), ans
    
    lastStep = "7) Log yazýlýyor"
    wsc.Range("D9").value = lastStep
    
    LogToSysLog Environ$("username"), "Gündem", "AI_Chat_CF"
    
    lastStep = "8) Bitti"
    wsc.Range("D9").value = lastStep
    
    Exit Sub
    
    
    If Len(Trim$(ans)) = 0 Then ans = "(Boþ yanýt alýndý) Lütfen tekrar deneyin."


ErrHandler:
    Application.StatusBar = False
    
    ' D9: son adým, D10: hata metni
    On Error Resume Next
    wsc.Range("D10").value = "HATA: " & Err.Number & " - " & Err.Description
    On Error GoTo 0
    
    MsgBox "AskAI_Run hata: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Son adým (Config!D9): " & wsc.Range("D9").value, vbExclamation
End Sub

Private Sub SafeWriteMerged(ByVal rng As Range, ByVal txt As String)
    If rng.MergeCells Then
        rng.MergeArea.Cells(1, 1).value = vbNullString
        rng.MergeArea.Cells(1, 1).value = txt
        rng.MergeArea.WrapText = True
    Else
        rng.value = txt
        rng.WrapText = True
    End If
End Sub

''API testi
Public Sub Test_Cloudflare()
    Dim ans As String
    ans = AskAI_Cloudflare("Sadece OK yaz.")
    MsgBox ans, vbInformation, "Cloudflare Test"
End Sub

Public Sub Check_CF_Token_Length()
    Dim t As String
    t = GetCFTokenSafe()
    
    If Len(t) = 0 Then
        MsgBox "CF_API_TOKEN okunamadý. Name Manager'da 'Workbook scope' ile tanýmlý mý kontrol edin." & vbCrLf & _
               "Öneri: Fix_CF_Names makrosunu çalýþtýrýn.", vbExclamation
        Exit Sub
    End If
    
    MsgBox "Token karakter sayýsý: " & Len(t), vbInformation
End Sub

Private Function GetCFTokenSafe() As String
    On Error Resume Next
    
    Dim t As String: t = ""
    
    ' 1) Workbook-level name
    t = CStr(ThisWorkbook.Names("CF_API_TOKEN").RefersToRange.value)
    
    ' 2) Hâlâ boþsa Config!B1 fallback
    If Len(Trim$(t)) = 0 Then
        t = CStr(ThisWorkbook.Worksheets("Config").Range("B1").value)
    End If
    
    ' Temizlik
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Trim$(t)
    
    GetCFTokenSafe = t
End Function


Public Sub Fix_CF_Names()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Config")
    
    ' Token hücrenizi burada net belirtin:
    Dim tokenCell As Range
    Set tokenCell = ws.Range("B1")
    
    On Error Resume Next
    ' Varsa silip yeniden ekler (scope/baðlantý sorunlarýný temizler)
    ThisWorkbook.Names("CF_API_TOKEN").Delete
    On Error GoTo 0
    
    ThisWorkbook.Names.Add Name:="CF_API_TOKEN", RefersTo:=tokenCell
    
    MsgBox "CF_API_TOKEN isimli aralýk oluþturuldu: " & tokenCell.Address(External:=True), vbInformation
End Sub


Public Sub Test_Cloudflare_Diag_ToCell()
    Dim ans As String
    ans = AskAI_Cloudflare("Sadece OK yaz.")
    
    ' Config sayfasýna yaz
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Config")
    
    ws.Range("A5").value = "Cloudflare Test Sonucu"
    ws.Range("B5").value = ans
    ws.Range("B5").WrapText = True
    ws.Rows(5).RowHeight = 120
    
    MsgBox "Test sonucu Config!B5 hücresine yazýldý. Lütfen oradaki ilk satýrdaki HTTP kodunu gönderin.", vbInformation
End Sub


Public Sub Fix_CF_AccountId_Name()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Config")
    
    On Error Resume Next
    ThisWorkbook.Names("CF_ACCOUNT_ID").Delete
    On Error GoTo 0
    
    ThisWorkbook.Names.Add Name:="CF_ACCOUNT_ID", RefersTo:=ws.Range("B2")
    MsgBox "CF_ACCOUNT_ID isimli aralýk oluþturuldu: " & ws.Range("B2").Address(External:=True), vbInformation
End Sub


Public Sub Test_Context()
    Dim ctx As String
    ctx = BuildContextSummary()
    ThisWorkbook.Worksheets("Config").Range("B6").value = ctx
    MsgBox "Baðlam Config!B6'ya yazýldý. Boþ mu kontrol edin.", vbInformation
End Sub

Private Function GetMergedText(rng As Range) As String
    If rng.MergeCells Then
        GetMergedText = CStr(rng.MergeArea.Cells(1, 1).value)
    Else
        GetMergedText = CStr(rng.value)
    End If
End Function

Private Sub SetMergedText(rng As Range, txt As String)
    If rng.MergeCells Then
        rng.MergeArea.ClearContents
        rng.MergeArea.Cells(1, 1).value = txt
        rng.MergeArea.WrapText = True
    Else
        rng.ClearContents
        rng.value = txt
        rng.WrapText = True
    End If
End Sub
