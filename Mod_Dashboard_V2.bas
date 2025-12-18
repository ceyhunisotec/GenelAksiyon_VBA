Attribute VB_Name = "Mod_Dashboard_V2"

'== Mod_Dashboard_V2.bas ==
Option Explicit

' -------------------------
' YERLEÞÝM SABÝTLERÝ (HÜCRE-ANKORLU)
' -------------------------
' Bu hücreleri kendi görüntünüze göre kolayca ayarlayabilirsiniz.
Private Const ROW_KPI_TOP     As Long = 6    ' KPI barlarýnýn üst satýrý (logo hizasý sað üst)
Private Const COL_KPI_LEFT    As String = "M" ' KPI sol üst hücresi (ör. M6)
Private Const ROW_BUTTONS     As Long = 9    ' Butonlarýn üst satýrý (baþlýða yakýn)
Private Const COL_BUTTONS_CTR As String = "H" ' Butonlarý bu sütunun etrafýna ortalayacaðýz (H)
Private Const ROW_HEADER      As Long = 10   ' Baþlýk bandýnýn baþladýðý satýr
Private Const ROW_GRID_START  As Long = 12   ' Kart ýzgarasýnýn baþladýðý satýr (headerdan en az 4 satýr alt)
Private Const ROW_TREND_TOP   As Long = 26   ' Trend grafiðinin üst satýrý (kartlardan sonra)
Private Const TILE_W          As Single = 240
Private Const TILE_H          As Single = 92
Private Const GAP_X           As Single = 16
Private Const GAP_Y           As Single = 16
Private Const COLS            As Long = 4    ' satýr baþýna 4 kart

' (Opsiyonel) harici logo yolu; boþsa Assets/CompanyLogo kullanýlacak
Private Const LOGO_PATH As String = ""

' -------------------------
' RENK / LOGO / BAÞLIK
' -------------------------
Private Function Dash_HexToRGB(ByVal hexColor As String) As Long
    Dim h As String, r As Long, g As Long, b As Long
    h = Replace$(Trim$(hexColor), "#", "")
    If Len(h) <> 6 Then Dash_HexToRGB = RGB(47, 85, 151): Exit Function
    r = CLng("&H" & Mid$(h, 1, 2))
    g = CLng("&H" & Mid$(h, 3, 2))
    b = CLng("&H" & Mid$(h, 5, 2))
    Dash_HexToRGB = RGB(r, g, b)
End Function

Private Function Dash_EnsureSingleLogo(ByVal ws As Worksheet, Optional ByVal desiredWidth As Single = 220) As Shape
    Dim s As Shape, base As Shape
    On Error Resume Next
    For Each s In ws.Shapes
        If (s.Type = msoPicture Or s.Type = msoLinkedPicture) Then
            If LCase$(s.Name) Like "*report_logo*" Or LCase$(s.AlternativeText) Like "*report_logo*" Then
                If base Is Nothing Then
                    Set base = s
                    base.Name = "Report_Logo": base.AlternativeText = "Report_Logo"
                Else
                    s.Delete
                End If
            End If
        End If
    Next s
    If base Is Nothing Then
        If Len(Dir(LOGO_PATH)) > 0 Then
            Set base = ws.Shapes.AddPicture(LOGO_PATH, False, True, 0, 0, desiredWidth, desiredWidth * 0.45)
        Else
            ThisWorkbook.Worksheets("Assets").Shapes("CompanyLogo").Copy
            ws.Paste
            If Err.Number = 0 Then Set base = ws.Shapes(ws.Shapes.Count)
            Err.Clear
        End If
    End If
    If Not base Is Nothing Then
        base.LockAspectRatio = msoTrue
        base.width = desiredWidth
        base.left = ws.Range("A1").left + 12
        base.Top = ws.Range("A1").Top + 6
        base.ZOrder msoBringToFront
        base.Name = "Report_Logo": base.AlternativeText = "Report_Logo"
    End If
    Set Dash_EnsureSingleLogo = base
End Function

' Baþlýk bandý (koyu gri), HÜCRE-ANKOR: ROW_HEADER
Private Sub Dash_TitleBand(ByVal ws As Worksheet, ByVal titleText As String)
    Dim left As Single, width As Single, lastCol As Long, right As Single
    Dim bandRGB As Long: bandRGB = RGB(64, 64, 64)
    Dim topOffset As Single: topOffset = ws.Range("A" & ROW_HEADER).Top

    ' Eski band/baþlýklarý temizle
    Dim shp As Shape
    On Error Resume Next
    For Each shp In ws.Shapes
        If LCase$(shp.Name) Like "*dash_header*" Or LCase$(shp.Name) Like "*dash_band*" Then shp.Delete
    Next shp
    On Error GoTo 0

    ' Geniþlik
    left = ws.Range("A1").left + 5
    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
    right = ws.Cells(1, lastCol).left + ws.Cells(1, lastCol).width
    width = right - left - 5

    ' Bant
    Dim band As Shape
    Set band = ws.Shapes.AddShape(msoShapeRectangle, left, topOffset, width, 44)
    With band
        .Name = "Dash_Band"
        .Fill.ForeColor.RGB = bandRGB
        .Line.Visible = msoFalse
        .ZOrder msoSendToBack
    End With

    ' Baþlýk metni
    Dim hdr As Shape
    Set hdr = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left, topOffset, width, 44)
    With hdr
        .Name = "Dash_Header"
        .TextFrame.Characters.text = titleText
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        With .TextFrame.Characters.Font
            .Name = "Segoe UI": .Size = 20: .Bold = True: .Color = RGB(255, 255, 255)
        End With
        .Line.Visible = msoFalse: .Fill.Visible = msoFalse
    End With
End Sub

' Sayfa sað sýnýrý
Private Function Dash_RightBound(ByVal ws As Worksheet) As Single
    Dim lastCol As Long
    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
    Dash_RightBound = ws.Cells(1, lastCol).left + ws.Cells(1, lastCol).width
End Function

' Ýzgarayý yatayda merkezle
Private Function Dash_CenterGridLeft(ByVal ws As Worksheet, ByVal totalWidth As Single) As Single
    Dim leftBound As Single: leftBound = ws.Range("A1").left
    Dim rightBound As Single: rightBound = Dash_RightBound(ws)
    Dash_CenterGridLeft = leftBound + (rightBound - leftBound - totalWidth) / 2
    If Dash_CenterGridLeft < leftBound + 12 Then Dash_CenterGridLeft = leftBound + 12
End Function

' Gridlines kapatma (tüm pencerelerde + aktif pencere)
Private Sub Dash_HideGridlinesAll()
    Dim wnd As Window
    For Each wnd In Application.Windows
        On Error Resume Next
        wnd.DisplayGridlines = False
        On Error GoTo 0
    Next wnd
    On Error Resume Next
    Application.ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub

' -------------------------
' KART STÝLÝ (baþlýk 12pt, sayý 28pt)
' -------------------------
Private Sub Dash_DrawTile(ByVal ws As Worksheet, ByVal x As Single, ByVal y As Single, _
                          ByVal w As Single, ByVal h As Single, ByVal title As String, _
                          ByVal value As Long, ByVal hex As String, _
                          Optional ByVal clickTag As String = "")
    Dim bg As Shape, hdr As Shape, valBox As Shape, nmBase As String
    Dim titleH As Single, valH As Single

    nmBase = "TL_" & Replace(title, " ", "_") & "_" & Format(Timer, "0")

    Set bg = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    With bg
        .Fill.ForeColor.RGB = Dash_HexToRGB(hex)
        .Line.Visible = msoFalse
        .Name = nmBase & "_bg"
        .Adjustments.Item(1) = 0.12
        If Len(clickTag) > 0 Then .OnAction = "Dash_TileClick": .AlternativeText = clickTag
        .Shadow.Visible = msoTrue: .Shadow.Blur = 8
        .Shadow.OffsetX = 3: .Shadow.OffsetY = 3
        .Shadow.ForeColor.RGB = RGB(180, 180, 180)
        .ZOrder msoBringToFront
    End With

    titleH = h * 0.45: valH = h - titleH

    Set hdr = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, titleH)
    With hdr
        .Name = nmBase & "_hdr"
        If Len(clickTag) > 0 Then .OnAction = "Dash_TileClick": .AlternativeText = clickTag
        .TextFrame.Characters.text = title
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        With .TextFrame.Characters.Font
            .Name = "Segoe UI": .Size = 12: .Bold = True: .Color = RGB(255, 255, 255)
        End With
        .Line.Visible = msoFalse: .Fill.Visible = msoFalse
        .ZOrder msoBringToFront
    End With

    Set valBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y + titleH, w, valH)
    With valBox
        .Name = nmBase & "_val"
        If Len(clickTag) > 0 Then .OnAction = "Dash_TileClick": .AlternativeText = clickTag
        .TextFrame.Characters.text = CStr(value)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        With .TextFrame.Characters.Font
            .Name = "Segoe UI": .Size = 28: .Bold = True: .Color = RGB(255, 255, 255)
        End With
        .Line.Visible = msoFalse: .Fill.Visible = msoFalse
        .ZOrder msoBringToFront
    End With
End Sub

' -------------------------
' KART TIKLAMA › FÝLTRE
' -------------------------
Public Sub Dash_TileClick()
    Dim shp As Shape, tag As String
    On Error Resume Next
    Set shp = ThisWorkbook.Worksheets("Dashboard").Shapes(Application.Caller)
    If shp Is Nothing Then Exit Sub
    tag = shp.AlternativeText
    If Len(tag) = 0 Then Exit Sub
    Dash_ExecuteTileAction tag
End Sub

Private Sub Dash_ExecuteTileAction(ByVal tag As String)
    Dim parts() As String, i As Long, kv As Variant
    Dim targetSheet As String, targetType As String
    parts = Split(tag, ";")
    For i = LBound(parts) To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) = 1 Then
            If UCase$(Trim$(kv(0))) = "SHEET" Then targetSheet = Trim$(kv(1))
            If UCase$(Trim$(kv(0))) = "TYPE" Then targetType = UCase$(Trim$(kv(1)))
        End If
    Next i
    If Len(targetSheet) = 0 Then Exit Sub
    If targetType = "OVERDUE" Then Dash_FilterSheetForOverdue targetSheet
End Sub

Public Sub Dash_FilterSheetForOverdue(ByVal sheetName As String)
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, headerRow As Long, rng As Range, i As Long
    Dim dec As String, critJ As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    headerRow = 4
    For i = 1 To 10
        If UCase$(Trim$(ws.Cells(i, 1).text)) = "SIRA" Then headerRow = i: Exit For
    Next i

    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    Set rng = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    rng.AutoFilter

    dec = Application.DecimalSeparator
    critJ = "<0" & IIf(dec = ",", ",", ".") & "99"

    rng.AutoFilter Field:=8, Criteria1:="<" & CDbl(Date), Operator:=xlAnd      ' H < bugün
    rng.AutoFilter Field:=10, Criteria1:=critJ, Operator:=xlAnd                ' J < %99

    ws.Activate
    ws.Cells(headerRow + 1, 1).Select
End Sub

' -------------------------
' DASHBOARD YÜZEY TEMÝZLÝÐÝ
' -------------------------
Private Sub Dash_ClearSurface(ByVal ws As Worksheet)
    Dim shp As Shape, keep As Boolean
    On Error Resume Next
    For Each shp In ws.Shapes
        keep = (LCase$(shp.Name) Like "*report_logo*")
        If Not keep Then shp.Delete
    Next shp
    ws.Cells.ClearContents
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
End Sub

' -------------------------
' GÜVENLÝ DÖNÜÞTÜRÜCÜLER (Tarih / Yüzde)
' -------------------------
Private Function Dash_ParsePercent(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then Dash_ParsePercent = CDbl(v): Exit Function
    If VarType(v) = vbString Then
        Dim s As String: s = Trim$(v)
        If s = "" Then Exit Function
        Dim hasPct As Boolean: hasPct = (InStr(s, "%") > 0)
        s = Replace$(s, "%", ""): s = Replace$(s, " ", ""): s = Replace$(s, ",", ".")
        On Error Resume Next
        Dim d As Double: d = CDbl(s)
        If Err.Number = 0 Then
            If hasPct And d > 1 Then d = d / 100
            Dash_ParsePercent = d
        End If
        Err.Clear: On Error GoTo 0
    End If
End Function

Private Function Dash_TryParseDate(ByVal v As Variant) As Variant
    If IsError(v) Then Exit Function
    If IsDate(v) Then Dash_TryParseDate = CDate(v): Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then Dash_TryParseDate = CDate(v)
        Exit Function
    End If
    If VarType(v) = vbString Then
        Dim s As String: s = Trim$(v)
        If s = "" Or s = "-" Or s = "—" Then Exit Function
        s = Replace$(s, ".", "/")
        On Error Resume Next
        Dim d As Date: d = DateValue(s)
        If Err.Number = 0 Then Dash_TryParseDate = d
        Err.Clear: On Error GoTo 0
    End If
End Function

' -------------------------
' KÝÞÝ BAZLI OPEN/OVERDUE
' -------------------------
Private Sub Dash_CollectOverduePerPerson(ByRef dict As Object)
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim vJ As Double, dPlan As Variant, key As String, fullName As String
    Dim arr As Variant
    Set dict = CreateObject("Scripting.Dictionary")

    For Each ws In ThisWorkbook.Worksheets
        If IsMeetingSheet(ws) Then
            lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
            For r = 5 To lastRow
                key = Trim$(ws.Cells(r, "F").text)
                If key <> "" Then
                    vJ = Dash_ParsePercent(ws.Cells(r, "J").value)        ' J=%Tamamlanma
                    dPlan = Dash_TryParseDate(ws.Cells(r, "H").value)      ' H=Planlanan
                    fullName = ResolveFullName(key)

                    If Not dict.Exists(key) Then dict.Add key, Array(fullName, 0&, 0&)
                    arr = dict(key)

                    If vJ < 0.99 Then
                        arr(1) = arr(1) + 1
                        If Not IsEmpty(dPlan) Then
                            If CLng(CDate(dPlan)) < CLng(Date) Then arr(2) = arr(2) + 1
                        End If
                    End If

                    dict(key) = arr
                End If
            Next r
        End If
    Next ws
End Sub

' -------------------------
' SHEET ÝSTATÝSTÝKLERÝ (UDT: TStats) – güvenli
' -------------------------
Public Sub Dash_GetSheetStats(ByVal ws As Worksheet, ByRef st As TStats)
    Dim lastRow As Long, r As Long
    Dim vJ As Double, dPlan As Variant
    Dim hasE As Boolean, hasF As Boolean

    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    For r = 5 To lastRow
        hasE = (Len(Trim$(ws.Cells(r, "E").text)) > 0)
        hasF = (Len(Trim$(ws.Cells(r, "F").text)) > 0)
        If hasE And hasF Then
            st.total = st.total + 1

            vJ = Dash_ParsePercent(ws.Cells(r, "J").value)
            dPlan = Dash_TryParseDate(ws.Cells(r, "H").value)

            If vJ < 0.99 Then
                st.Open = st.Open + 1
                If Not IsEmpty(dPlan) Then
                    Dim dNum As Long, tNum As Long
                    dNum = CLng(CDate(dPlan))
                    tNum = CLng(Date)

                    If dNum < tNum Then
                        st.Overdue = st.Overdue + 1
                    ElseIf dNum = tNum Then
                        st.DueToday = st.DueToday + 1
                    End If
                End If
            End If
        End If
    Next r
End Sub

' -------------------------
' SYSLOG TREND + MÝNÝ GRAFÝK (merkezlenmiþ, kartlardan sonra)
' -------------------------
Private Sub Dash_BuildSysLogTrends(ByVal ws As Worksheet, ByVal leftCell As Range)
    Dim logWs As Worksheet, lastRow As Long, r As Long
    Dim d As Date, key As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim startDate As Date: startDate = Date - 29

    On Error Resume Next
    Set logWs = ThisWorkbook.Worksheets("SysLog")
    On Error GoTo 0
    If logWs Is Nothing Then Exit Sub

    lastRow = logWs.Cells(logWs.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        If IsDate(logWs.Cells(r, "A").value) Then
            d = CDate(logWs.Cells(r, "A").value)
            If d >= startDate And d <= Date Then
                key = Format(d, "yyyy-mm-dd")
                If Not dict.Exists(key) Then dict.Add key, 0
                dict(key) = dict(key) + 1
            End If
        End If
    Next r

    Dim outTop As Range: Set outTop = leftCell
    outTop.Offset(0, 0).value = "Tarih"
    outTop.Offset(0, 1).value = "Adet"

    Dim i As Long, dayCount As Long: dayCount = 0
    For i = 0 To 29
        d = Date - (29 - i)
        outTop.Offset(i + 1, 0).value = d
        outTop.Offset(i + 1, 1).value = IIf(dict.Exists(Format(d, "yyyy-mm-dd")), dict(Format(d, "yyyy-mm-dd")), 0)
        dayCount = dayCount + outTop.Offset(i + 1, 1).value
    Next i

    ' Eski grafiði sil ve yenisini ekle
    Dim co As ChartObject, dataRng As Range
    Set dataRng = ws.Range(outTop.Offset(1, 0), outTop.Offset(30, 1))
    For Each co In ws.ChartObjects
        If co.Name = "ch_Trends30" Then co.Delete
    Next co

    ' Kart ýzgarasýnýn geniþliði: 4 kart + boþluklar
    Dim gridWidth As Single, gridLeft As Single
    gridWidth = (TILE_W * COLS) + (GAP_X * (COLS - 1))
    gridLeft = Dash_CenterGridLeft(ws, gridWidth)

    Dim chartTop As Single: chartTop = ws.Range("A" & ROW_TREND_TOP).Top
    Set co = ws.ChartObjects.Add(left:=gridLeft, Top:=chartTop, width:=gridWidth, Height:=170)
    co.Name = "ch_Trends30"
    With co.Chart
        .ChartType = xlLine
        .SetSourceData Source:=dataRng
        .HasTitle = True: .ChartTitle.text = "Son 30 Gün Gönderim Trendi"
        .Axes(xlCategory).TickLabels.NumberFormat = "dd.MM"
        .Legend.Delete
        ' Ýç grid çizgilerini kapat (arka çizgi algýsýný azaltýr)
        .Axes(xlValue).HasMajorGridlines = False
        .Axes(xlCategory).HasMajorGridlines = False
    End With
End Sub

' -------------------------
' KPI BARLARI – SAÐ ÜST (logo hizasý)
' -------------------------
Private Sub Dash_DrawProgressBarAt(ByVal ws As Worksheet, ByVal leftPx As Single, ByVal topPx As Single, _
                                   ByVal w As Single, ByVal h As Single, ByVal ratio As Double, _
                                   ByVal backRGB As Long, ByVal fillRGB As Long, ByVal caption As String)
    Dim bg As Shape, fg As Shape, tx As Shape
    If ratio < 0 Then ratio = 0
    If ratio > 1 Then ratio = 1
    Set bg = ws.Shapes.AddShape(msoShapeRectangle, leftPx, topPx, w, h)
    With bg: .Fill.ForeColor.RGB = backRGB: .Line.Visible = msoFalse: End With
    Set fg = ws.Shapes.AddShape(msoShapeRectangle, leftPx, topPx, w * ratio, h)
    With fg: .Fill.ForeColor.RGB = fillRGB: .Line.Visible = msoFalse: End With
    Set tx = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPx, topPx - 20, w, 20)
    With tx.TextFrame
        .Characters.text = caption & " - " & Format(ratio, "0%")
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    With tx.TextFrame.Characters.Font
        .Name = "Segoe UI": .Size = 11
    End With
    tx.Line.Visible = msoFalse: tx.Fill.Visible = msoFalse
End Sub

Private Sub Dash_AddKpisTopRight(ByVal ws As Worksheet)
    ' Sað üst KPI’lar için hücre ankörü: COL_KPI_LEFT & ROW_KPI_TOP (ör. M6)
    Dim anchor As Range: Set anchor = ws.Range(COL_KPI_LEFT & ROW_KPI_TOP)
    Dim leftPx As Single: leftPx = anchor.left
    Dim topPx  As Single: topPx = anchor.Top
    Dim w As Single: w = 300
    Dim h As Single: h = 18

    ' Hesaplar
    Dim total As Long, done As Long, onTime As Long
    Dim wsM As Worksheet, r As Long, lastRow As Long
    For Each wsM In ThisWorkbook.Worksheets
        If IsMeetingSheet(wsM) Then
            lastRow = wsM.Cells(wsM.Rows.Count, "F").End(xlUp).Row
            For r = 5 To lastRow
                If Len(Trim$(wsM.Cells(r, "E").text)) > 0 And Len(Trim$(wsM.Cells(r, "F").text)) > 0 Then
                    total = total + 1
                    If Dash_ParsePercent(wsM.Cells(r, "J").value) >= 0.99 Then
                        done = done + 1
                        If IsDate(wsM.Cells(r, "I").value) And IsDate(wsM.Cells(r, "H").value) Then
                            If CDate(wsM.Cells(r, "I").value) <= CDate(wsM.Cells(r, "H").value) Then
                                onTime = onTime + 1
                            End If
                        End If
                    End If
                End If
            Next r
        End If
    Next wsM

    ' Çizim (sað üst – logo hizasý civarý)
    Dash_DrawProgressBarAt ws, leftPx, topPx, w, h, IIf(total > 0, done / total, 0), _
                           RGB(230, 230, 230), RGB(0, 120, 212), "Genel Tamamlanma Oraný"
    Dash_DrawProgressBarAt ws, leftPx, topPx + 40, w, h, IIf(total > 0, onTime / total, 0), _
                           RGB(230, 230, 230), RGB(16, 124, 16), "Zamanýnda Tamamlama Oraný"
End Sub

' -------------------------
' ANA: DASHBOARD
' -------------------------
Public Sub RebuildDashboard()
    Dim ws As Worksheet, shpLogo As Shape
    Dim gridLeft As Single, gridTop As Single
    Dim x As Single, y As Single, w As Single, h As Single

    Dim sh As Worksheet, hex As String
    Dim st As TStats

    Dim dict As Object, keys As Variant
    Dim a As Long, b As Long, r As Long, topN As Long
    Dim tmp As Variant, tblTop As Range
    Dim btn As Shape

    ' 1) Sayfa hazýrlýðý
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = "Dashboard"
    End If

    Application.ScreenUpdating = False
    Dash_ClearSurface ws
    Dash_HideGridlinesAll

    ' 2) Logo
    Set shpLogo = Dash_EnsureSingleLogo(ws, 200)

    ' 3) Baþlýk bandý (koyu gri) – hücreye sabit ROW_HEADER
    Dash_TitleBand ws, "Toplantýlar Dashboard – " & Format(Date, "dd.MM.yyyy")

    ' 4) KPI barlarý – sað üst (logo hizasý civarý, M6)
    Dash_AddKpisTopRight ws

    ' 5) Butonlar – kart ýzgarasýna göre ortalý (baþlýðýn hemen altýnda, ROW_BUTTONS)
    Dim btnW As Single, btnH As Single, gapBtn As Single
    btnW = 180: btnH = 28: gapBtn = 12
    Dim centerCell As Range: Set centerCell = ws.Range(COL_BUTTONS_CTR & ROW_BUTTONS)
    Dim btnTop As Single:  btnTop = centerCell.Top
    Dim gridWidth As Single: gridWidth = (TILE_W * COLS) + (GAP_X * (COLS - 1))
    Dim btnLeft As Single: btnLeft = Dash_CenterGridLeft(ws, gridWidth) + (gridWidth - (btnW * 2 + gapBtn)) / 2

    Set btn = ws.Shapes.AddShape(msoShapeActionButtonForwardorNext, btnLeft, btnTop, btnW, btnH)
    With btn
        .OnAction = "RunOverdueReportsNow"
        .Fill.ForeColor.RGB = RGB(0, 120, 212)
        .Line.Visible = msoFalse
        .TextFrame.Characters.text = "Raporlarý Þimdi Gönder"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        With .TextFrame.Characters.Font
            .Name = "Segoe UI": .Size = 11: .Bold = True: .Color = RGB(255, 255, 255)
        End With
    End With

    Set btn = ws.Shapes.AddShape(msoShapeActionButtonCustom, btnLeft + btnW + gapBtn, btnTop, 160, btnH)
    With btn
        .OnAction = "RebuildDashboard"
        .Fill.ForeColor.RGB = RGB(16, 124, 16)
        .Line.Visible = msoFalse
        .TextFrame.Characters.text = "Dashboard'u Yenile"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        With .TextFrame.Characters.Font
            .Name = "Segoe UI": .Size = 11: .Bold = True: .Color = RGB(255, 255, 255)
        End With
    End With

    ' 6) Kart ýzgarasý – hücreye sabit ROW_GRID_START
    w = TILE_W: h = TILE_H
    gridTop = ws.Range("A" & ROW_GRID_START).Top
    gridLeft = Dash_CenterGridLeft(ws, (w * COLS) + (GAP_X * (COLS - 1)))
    y = gridTop

    For Each sh In ThisWorkbook.Worksheets
        If IsMeetingSheet(sh) Then
            st.total = 0: st.Open = 0: st.Overdue = 0: st.DueToday = 0
            Dash_GetSheetStats sh, st
            hex = GetThemeColorHex(sh.Name)

            x = gridLeft
            Dash_DrawTile ws, x, y, w, h, sh.Name & " – Toplam", st.total, hex
            Dash_DrawTile ws, x + (w + GAP_X), y, w, h, "Açýk", st.Open, hex
            Dash_DrawTile ws, x + 2 * (w + GAP_X), y, w, h, "Gecikmiþ", st.Overdue, hex, _
                           "SHEET=" & sh.Name & ";TYPE=OVERDUE"
            Dash_DrawTile ws, x + 3 * (w + GAP_X), y, w, h, "Bugün Planlý", st.DueToday, hex

            y = y + h + GAP_Y
        End If
    Next sh

    ' 7) Trend grafiði – hücreye sabit ROW_TREND_TOP, merkezlenmiþ
    Dash_BuildSysLogTrends ws, ws.Range("A1000")

    Application.ScreenUpdating = True
End Sub


