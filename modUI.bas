Attribute VB_Name = "modUI"

'=== Modül: modUI ===
Option Explicit

Sub SetupAgendaUI()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Gündem")
    With ws
        .Range("A1").value = "AI Sohbet"
        .Range("A3").value = "Soru:"
        .Range("A6").value = "Cevap:"
        .Range("B3").ClearContents
        .Range("B6:B20").ClearContents
        .Range("B3").Name = "rngAI_Question"
        .Range("B6:B20").Name = "rngAI_Answer"
        .Range("B6:B20").WrapText = True
    End With
    
    ' Buton: Soruyu Çalýþtýr
    AddSheetButton "Gündem", "Soruyu Çalýþtýr", "AskAI_Run", 100, 30, "D3"
    ' Buton: Temizle
    AddSheetButton "Gündem", "Temizle", "AskAI_Clear", 100, 30, "E3"
    ' Buton: Mail Gönder
    AddSheetButton "Gündem", "Mail Gönder", "AskAI_SendMail", 120, 30, "F3"
End Sub

Private Sub AddSheetButton(sheetName As String, caption As String, macroName As String, w As Double, h As Double, topLeftCell As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range(topLeftCell).left, ws.Range(topLeftCell).Top, w, h)
    With shp
        .TextFrame.Characters.text = caption
        .OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
        .Fill.ForeColor.RGB = RGB(0, 120, 215)
        .Line.Visible = msoFalse
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.Characters.Font.Color = vbWhite
        .TextFrame.Characters.Font.Bold = True
    End With
End Sub


Public Sub SetupAgendaUI_Pro()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Gündem")
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    ' Eski shape'leri temizle (butonlar vs.)
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    
    ' Sayfayý düzenle
    ws.Cells.Clear
    
    ' Kolon geniþlikleri (okunaklý panel)
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 18
    
    ' Baþlýk
    With ws.Range("A1:H1")
        .Merge
        .value = "AI Sohbet (Aksiyon Listesi Asistaný)"
        .Font.Size = 16
        .Font.Bold = True
        .Interior.Color = RGB(0, 120, 215)
        .Font.Color = vbWhite
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignCenter
        .RowHeight = 34
        .IndentLevel = 1
    End With
    
    ' Açýklama satýrý
    With ws.Range("A2:H2")
        .Merge
        .value = "Soru yazýn › Soruyu Çalýþtýr › Cevabý görüntüleyin / Tek tuþ mail gönderin."
        .Font.Size = 10
        .Font.Color = RGB(60, 60, 60)
        .HorizontalAlignment = xlHAlignLeft
        .RowHeight = 18
        .IndentLevel = 1
    End With
    
    ' Soru etiketi + soru kutusu
    ws.Range("A4").value = "Soru:"
    ws.Range("A4").Font.Bold = True
    With ws.Range("B4:H5")
        .Merge
        .Name = "rngAI_Question"
        .value = ""
        .WrapText = True
        .Font.Size = 11
        .Interior.Color = RGB(245, 245, 245)
        .Borders.LineStyle = xlContinuous
        .RowHeight = 24
    End With
    
    ' Butonlar alaný (A7:H7)
    ws.Rows(7).RowHeight = 28
    
    AddSheetButtonPro ws, "Soruyu Çalýþtýr", "AskAI_Run", "B7", 140, 28, RGB(0, 120, 215)
    AddSheetButtonPro ws, "Temizle", "AskAI_Clear", "D7", 100, 28, RGB(90, 90, 90)
    AddSheetButtonPro ws, "Mail Gönder", "AskAI_SendMail", "E7", 120, 28, RGB(0, 153, 51)
    
    ' Cevap etiketi + cevap paneli
    ws.Range("A9").value = "Cevap:"
    ws.Range("A9").Font.Bold = True
    
    With ws.Range("B9:H30")
        .Merge
        .Name = "rngAI_Answer"
        .value = ""
        .WrapText = True
        .Font.Size = 11
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlVAlignTop
        .HorizontalAlignment = xlHAlignLeft
    End With
    
    ws.Range("B9:H30").IndentLevel = 1
    
    Application.ScreenUpdating = True
End Sub

Private Sub AddSheetButtonPro(ws As Worksheet, caption As String, macroName As String, topLeftCell As String, w As Double, h As Double, fillRGB As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range(topLeftCell).left, ws.Range(topLeftCell).Top, w, h)
    With shp
        .TextFrame.Characters.text = caption
        .OnAction = "'" & ThisWorkbook.Name & "'!" & macroName  ' <-- çok önemli
        .Fill.ForeColor.RGB = fillRGB
        .Line.Visible = msoFalse
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.Characters.Font.Color = vbWhite
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
    End With
End Sub


Sub AskAI_Clear()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Gündem")
       ws.Range("rngAI_Question").MergeArea.ClearContents
    ws.Range("rngAI_Answer").MergeArea.ClearContents
End Sub

