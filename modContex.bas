Attribute VB_Name = "modContex"

'=== Modül: modContext ===
Option Explicit

Public Function BuildContextSummary() As String
    Dim sb As String
    sb = "Veri Özeti (bugün: " & Format(Date, "dd.MM.yyyy") & "):" & vbCrLf
    sb = sb & ContextForSheet("Koordinasyon") & vbCrLf
    sb = sb & ContextForSheet("Sipariþ") & vbCrLf
    sb = sb & ContextForSheet("Þikayet") & vbCrLf
    sb = sb & ContextForSheet("Atýl_Stok") & vbCrLf
    sb = sb & ContextForSheet("Kalite") & vbCrLf
    BuildContextSummary = sb
End Function

Public Function BuildContextSummaryDetailed() As String
    Dim sb As String
    sb = "Bugün: " & Format(Date, "dd.MM.yyyy") & vbCrLf & _
         "Aþaðýdaki özet; sayfa bazýnda KPI + kritik maddeler içerir." & vbCrLf & vbCrLf
    
    sb = sb & SummaryWithTop("Koordinasyon") & vbCrLf
    sb = sb & SummaryWithTop("Sipariþ") & vbCrLf
    sb = sb & SummaryWithTop("Þikayet") & vbCrLf
    sb = sb & SummaryWithTop("Atýl_Stok") & vbCrLf
    sb = sb & SummaryWithTop("Kalite") & vbCrLf
    
    BuildContextSummaryDetailed = sb
End Function

Private Function SummaryWithTop(sheetName As String) As String
    On Error GoTo Fail
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim firstDataRow As Long: firstDataRow = FindFirstDataRow(ws, lastRow)
    
    Dim totalCnt As Long, openCnt As Long, overdueCnt As Long, todayCnt As Long
    Dim i As Long
    Dim pDate As Variant, done As Variant
    
    ' Top listeler
    Dim topOverdue As String: topOverdue = ""
    Dim topToday As String: topToday = ""
    
    For i = firstDataRow To lastRow
        If IsNumeric(ws.Cells(i, 1).value) And Len(Trim$(ws.Cells(i, 1).value)) > 0 Then
            totalCnt = totalCnt + 1
            
            pDate = ws.Cells(i, 7).value  ' Planlanan Tarih
            done = ws.Cells(i, 9).value   ' % Tamamlanma
            
            If Val(done) < 1 Then
                openCnt = openCnt + 1
                
                If IsDate(pDate) Then
                    If CDate(pDate) < Date Then
                        overdueCnt = overdueCnt + 1
                        If CountLines(topOverdue) < 5 Then
                            topOverdue = topOverdue & "- " & ItemLine(ws, i) & vbCrLf
                        End If
                    ElseIf CDate(pDate) = Date Then
                        todayCnt = todayCnt + 1
                        If CountLines(topToday) < 5 Then
                            topToday = topToday & "- " & ItemLine(ws, i) & vbCrLf
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    Dim s As String
    s = sheetName & " | Toplam:" & totalCnt & " Açýk:" & openCnt & _
        " Gecikmiþ:" & overdueCnt & " Bugün:" & todayCnt & vbCrLf
    
    If Len(topOverdue) > 0 Then s = s & "Gecikmiþ Ýlk 5:" & vbCrLf & topOverdue
    If Len(topToday) > 0 Then s = s & "Bugün Planlý Ýlk 5:" & vbCrLf & topToday
    
    SummaryWithTop = s
    Exit Function

Fail:
    SummaryWithTop = sheetName & " | Özet alýnamadý: " & Err.Description
End Function

Private Function FindFirstDataRow(ws As Worksheet, lastRow As Long) As Long
    Dim i As Long
    FindFirstDataRow = 1
    For i = 1 To Application.Min(25, lastRow)
        If UCase$(Trim$(ws.Cells(i, 1).value)) Like "*SIRA*" Then
            FindFirstDataRow = i + 1
            Exit Function
        End If
    Next i
End Function

Private Function ItemLine(ws As Worksheet, r As Long) As String
    ' Kolonlar: Konu/Madde (C), Aksiyon (D), Sorumlu (E), Planlanan (G), % (I)
    ItemLine = _
        "No:" & ws.Cells(r, 1).value & _
        " | Konu:" & left$(CStr(ws.Cells(r, 3).value), 60) & _
        " | Aksiyon:" & left$(CStr(ws.Cells(r, 4).value), 80) & _
        " | Sorumlu:" & CStr(ws.Cells(r, 5).value) & _
        " | Plan:" & DateText(ws.Cells(r, 7).value) & _
        " | %:" & CStr(ws.Cells(r, 9).value)
End Function

Private Function DateText(v As Variant) As String
    If IsDate(v) Then
        DateText = Format(CDate(v), "dd.MM.yyyy")
    Else
        DateText = CStr(v)
    End If
End Function

Private Function CountLines(s As String) As Long
    If Len(Trim$(s)) = 0 Then
        CountLines = 0
    Else
        CountLines = UBound(Split(Trim$(s), vbCrLf)) + 1
    End If
End Function

Private Function ContextForSheet(sheetName As String) As String
    On Error GoTo SafeExit
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long, openCnt As Long, overdueCnt As Long, totalCnt As Long, todayCnt As Long
    Dim pDate As Variant, done As Variant
    
    ' Baþlýk satýrlarýný atlamak için basit heuristik (ilk veri satýrýný arar)
    Dim firstDataRow As Long: firstDataRow = 1
    For i = 1 To Application.Min(20, lastRow)
        If UCase(ws.Cells(i, 1).value) Like "*SIRA*" Then
            firstDataRow = i + 1
            Exit For
        End If
    Next i
    
    For i = firstDataRow To lastRow
        If Len(Trim(ws.Cells(i, 1).value)) > 0 And IsNumeric(ws.Cells(i, 1).value) Then
            totalCnt = totalCnt + 1
            pDate = ws.Cells(i, 7).value         ' Planlanan Tarih sütunu
            done = ws.Cells(i, 9).value          ' % Tamamlanma sütunu
            If Val(done) < 1 Then
                openCnt = openCnt + 1
                If IsDate(pDate) Then
                    If CDate(pDate) < Date Then overdueCnt = overdueCnt + 1
                    If CDate(pDate) = Date Then todayCnt = todayCnt + 1
                End If
            End If
        End If
    Next i
    
    ContextForSheet = sheetName & " — Toplam:" & totalCnt & _
                      ", Açýk:" & openCnt & ", Gecikmiþ:" & overdueCnt & _
                      ", Bugün Planlý:" & todayCnt
SafeExit:

End Function
