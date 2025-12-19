Attribute VB_Name = "Mod_Report"

Option Explicit
'==========================
' R A P O R   M O D Ü L Ü
'==========================

'-- Basit URL kontrolü (elde var, koruyoruz) --
Private Function IsUrl(ByVal s As String) As Boolean
    s = LCase$(Trim$(s))
    IsUrl = (left$(s, 7) = "http://" Or left$(s, 8) = "https://")
End Function

'==========================
'  GÖRSEL / BAÞLIK YARDIMCILARI
'==========================

'== Hex (#RRGGBB) -> RGB dönüþtürücü (tema rengi için) ==
'== Hex (#RRGGBB) -> RGB (ÇAKIÞMASIZ SÜRÜM) ==
Private Function HexColorToRGB(ByVal hexColor As String) As Long
    Dim h As String, r As Long, g As Long, b As Long
    h = Replace$(Trim$(hexColor), "#", "")
    If Len(h) <> 6 Then
        HexColorToRGB = RGB(0, 120, 212) ' Varsayýlan mavi
        Exit Function
    End If
    r = CLng("&H" & Mid$(h, 1, 2))
    g = CLng("&H" & Mid$(h, 3, 2))
    b = CLng("&H" & Mid$(h, 5, 2))
    HexColorToRGB = RGB(r, g, b)
End Function

'== Sayfada TEK logo garantisi (ROBUST): varsa birini býrakýr, yoksa dosyadan ekler.
'   Copy/Paste yerine AddPicture kullanýr (PDF export sýrasýnda logo kaybolmasýný engeller)
Private Function EnsureSingleLogo(ByVal ws As Worksheet, Optional ByVal desiredWidth As Single = 230) As Shape
    On Error Resume Next

    Dim s As Shape, base As Shape
    ' 1) Varsa mevcut Report_Logo'yu kullan (fazlalarý sil)
    For Each s In ws.Shapes
        If (s.Type = msoPicture Or s.Type = msoLinkedPicture) Then
            If LCase$(s.Name) Like "*report_logo*" Or LCase$(s.AlternativeText) Like "*report_logo*" Or _
               InStr(1, LCase$(s.Name), "logo") > 0 Or InStr(1, LCase$(s.AlternativeText), "logo") > 0 Then
                If base Is Nothing Then
                    Set base = s
                    base.Name = "Report_Logo"
                    base.AlternativeText = "Report_Logo"
                Else
                    s.Delete
                End If
            End If
        End If
    Next s

    ' 2) Logo yoksa: LOGO_PATH varsa onu ekle; yoksa Assets/CompanyLogo'yu TEMP'e export edip ekle
    If base Is Nothing Then
        Dim logoFile As String
        logoFile = GetLogoFilePathSafe()   ' LOGO_PATH veya exported temp png

        If Len(logoFile) > 0 And Len(Dir(logoFile)) > 0 Then
            Set base = ws.Shapes.AddPicture( _
                        Filename:=logoFile, _
                        LinkToFile:=msoFalse, _
                        SaveWithDocument:=msoTrue, _
                        left:=0, Top:=0, width:=desiredWidth, Height:=desiredWidth * 0.45)
        End If
    End If

    ' 3) Konum/ölçek/print garantisi
    If Not base Is Nothing Then
        base.LockAspectRatio = msoTrue
        base.width = desiredWidth
        base.left = ws.Range("A1").left + 12
        base.Top = ws.Range("A1").Top + 10
        base.ZOrder msoBringToFront

        ' PDF'de kaybolmamasý için kritik ayarlar
        base.Visible = msoTrue
        base.PrintObject = True
        base.Placement = xlMoveAndSize

        base.Name = "Report_Logo"
        base.AlternativeText = "Report_Logo"
    End If

    On Error GoTo 0
    Set EnsureSingleLogo = base
End Function


' LOGO_PATH çalýþýyorsa onu döndürür, çalýþmýyorsa Assets/CompanyLogo'yu TEMP'e export eder ve path döndürür
Private Function GetLogoFilePathSafe() As String
    On Error Resume Next

    ' 1) Mod_Settings içindeki LOGO_PATH
    If Len(Trim$(LOGO_PATH)) > 0 Then
        If Len(Dir(LOGO_PATH)) > 0 Then
            GetLogoFilePathSafe = LOGO_PATH
            Exit Function
        End If
    End If

    ' 2) Assets sayfasýndaki CompanyLogo'yu export et
    GetLogoFilePathSafe = ExportAssetsLogoToTempPng()
End Function


' Assets!CompanyLogo þeklinin PNG çýktýsýný TEMP'e alýr
Private Function ExportAssetsLogoToTempPng() As String
    On Error GoTo Fail

    Dim tmpPng As String
    tmpPng = Environ$("TEMP") & "\isotec_company_logo.png"

    Dim wsA As Worksheet
    Set wsA = ThisWorkbook.Worksheets("Assets")

    Dim shp As Shape
    Set shp = wsA.Shapes("CompanyLogo")

    ' Export baþarýlýysa dosya oluþur
    shp.Export Filename:=tmpPng, FilterName:="PNG"

    If Len(Dir(tmpPng)) > 0 Then
        ExportAssetsLogoToTempPng = tmpPng
    Else
        ExportAssetsLogoToTempPng = ExportAssetsLogoToTempPng = ""
    End If
    Exit Function

Fail:
    ExportAssetsLogoToTempPng = ""
End Function

'== Logonun SAÐINA tema bantlý BÜYÜK baþlýk ==

'== Hex (#RRGGBB) -> RGB dönüþtürücü (tema rengi için) ==
Private Function HexToRGB(ByVal hexColor As String) As Long
    Dim h As String, r As Long, g As Long, b As Long
    h = Replace$(Trim$(hexColor), "#", "")
    If Len(h) <> 6 Then
        HexToRGB = RGB(0, 120, 212) ' Varsayýlan: mavi
        Exit Function
    End If
    r = CLng("&H" & Mid$(h, 1, 2))
    g = CLng("&H" & Mid$(h, 3, 2))
    b = CLng("&H" & Mid$(h, 5, 2))
    HexToRGB = RGB(r, g, b)
End Function

'== Rengi koyulaþtýr (0..1 faktör; 0=siyah, 1=ayný renk) ==
Private Function DarkenColor(ByVal c As Long, Optional ByVal factor As Double = 0.7) As Long
    Dim r As Long, g As Long, b As Long
    r = (c And &HFF)
    g = (c \ &H100) And &HFF
    b = (c \ &H10000) And &HFF
    If factor < 0 Then factor = 0
    If factor > 1 Then factor = 1
    DarkenColor = RGB(r * factor, g * factor, b * factor)
End Function

'== Logonun ALTINDA, sayfa geniþliðinde KOYU bant + ORTALANMIÞ beyaz baþlýk ==
Private Sub StampHeaderBand(ByVal ws As Worksheet, ByVal titleText As String, ByVal themeHex As String)
    On Error Resume Next
    ' Eski baþlýk/bant varsa temizle
    Dim shp As Shape, i As Long
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If LCase$(shp.Name) Like "*report_header*" Or LCase$(shp.Name) Like "*report_header_band*" Then shp.Delete
    Next i

    ' Logo: yükseklik hizasý için
    Dim logo As Shape, topPos As Single
    On Error Resume Next: Set logo = ws.Shapes("Report_Logo"): On Error GoTo 0
    If Not logo Is Nothing Then
        topPos = logo.Top + logo.Height + 10   ' << LOGONUN ALTINA
    Else
        topPos = ws.Range("A1").Top + 60       ' logo yoksa güvenli offset
    End If

    ' Sayfa sol/sað sýnýr (UsedRange)
    Dim firstLeft As Single, lastCol As Long, rightBound As Single, widthPx As Single
    firstLeft = ws.Range("A1").left + 5
    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
    rightBound = ws.Cells(1, lastCol).left + ws.Cells(1, lastCol).width
    widthPx = rightBound - firstLeft - 5

    ' Sekme rengi (yoksa tema) + koyulaþtýrma
    Dim tabCol As Long, bandCol As Long
    tabCol = ws.Tab.Color
    If tabCol = 0 Then
        bandCol = HexColorToRGB(themeHex)      ' Mod_Report’taki HexColorToRGB kullanýlmalý
    Else
        bandCol = tabCol
    End If
    bandCol = DarkenColor(bandCol, 0.7)

    ' Bant
    Dim band As Shape
    Set band = ws.Shapes.AddShape(msoShapeRectangle, firstLeft, topPos, widthPx, 50)
    With band
        .Name = "Report_Header_Band"
        .Line.Visible = msoFalse
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = bandCol
        .ZOrder msoSendToBack
    End With

    ' Metin kutusu: bant geniþliði kadar ve ORTALANMIÞ
    Dim hdr As Shape
    Set hdr = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, firstLeft, topPos, widthPx, 50)
    With hdr
        .Name = "Report_Header"
        .TextFrame.Characters.text = titleText
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        With .TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Size = 24
            .Bold = True
            .Color = RGB(255, 255, 255)
        End With
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .ZOrder msoBringToFront
    End With
End Sub


'==========================
'  SYSLOG / ALICI YARDIMCILARI  (mevcut iþlevler korunur)
'==========================

Private Function EnsureSysLogSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("SysLog")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "SysLog"
        ws.Range("A1:D1").value = Array("Tarih", "Email", "Sheet", "Note")
        ws.Columns("A:D").HorizontalAlignment = xlLeft
        ws.Visible = xlSheetVeryHidden
    Else
        If Trim$(CStr(ws.Cells(1, "A").value)) <> "Tarih" Then
            ws.Range("A1:D1").value = Array("Tarih", "Email", "Sheet", "Note")
        End If
    End If
    Set EnsureSysLogSheet = ws
End Function

' Mod_Report.bas içindeki sürümün yerine
Private Function HasDailyMailSent(ByVal email As String, _
                                  Optional ByVal scope As String = "global", _
                                  Optional ByVal sheetName As String = "", _
                                  Optional ByVal slot As String = "") As Boolean
    Dim ws As Worksheet, lastRow As Long, i As Long, d0 As Long, di As Long
    Dim e As String, sh As String, note As String

    Set ws = EnsureSysLogSheet()
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    d0 = CLng(Date)

    For i = 2 To lastRow
        If Not IsEmpty(ws.Cells(i, "A").value) Then
            On Error Resume Next
            di = CLng(CDate(ws.Cells(i, "A").value))
            On Error GoTo 0

            If di = d0 Then
                e = LCase$(Trim$(CStr(ws.Cells(i, "B").value)))
                sh = CStr(ws.Cells(i, "C").value)
                note = CStr(ws.Cells(i, "D").value)

                ' Sadece gecikmiþ rapor notlarýnda slot kontrolü uygula
                If InStr(1, note, "OverdueReport", vbTextCompare) > 0 Then
                    If Len(slot) > 0 Then
                        If InStr(1, note, "-" & slot, vbTextCompare) = 0 Then GoTo ContinueLoop
                    End If
                End If

                If LCase$(Trim$(email)) = e Then
                    If LCase$(scope) = "sheet" Then
                        If sh = sheetName Then HasDailyMailSent = True: Exit Function
                    Else
                        HasDailyMailSent = True: Exit Function
                    End If
                End If
            End If
        End If
ContinueLoop:
    Next i
End Function

Private Sub MarkDailyMailSent(ByVal email As String, ByVal sheetName As String, Optional ByVal slot As String = "")
    Dim ws As Worksheet, nextRow As Long, noteVal As String
    Set ws = EnsureSysLogSheet()
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    noteVal = "OverdueReport" & IIf(Len(slot) > 0, "-" & slot, "")
    ws.Cells(nextRow, "A").value = Date
    ws.Cells(nextRow, "B").value = email
    ws.Cells(nextRow, "C").value = sheetName
    ws.Cells(nextRow, "D").value = noteVal
End Sub


Private Function BuildOverdueRecipients(ByVal ws As Worksheet, Optional ByVal slot As String = "") As Collection
    Dim col As New Collection
    Dim lastRow As Long, r As Long
    Dim vJ As Variant, planDate As Variant, addr As String
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    For r = 5 To lastRow
        vJ = ws.Cells(r, "J").Value2
        planDate = ws.Cells(r, "H").value
        If IsNumeric(vJ) Then
            If vJ < 0.99 Then
                If IsDate(planDate) Then
                    If CDate(planDate) < Date Then
                        addr = ResolveRecipient(CStr(ws.Cells(r, "F").value))
                        If Len(addr) > 0 Then
                            If Not HasDailyMailSent(addr, "sheet", ws.Name, slot) Then
                                On Error Resume Next
                                col.Add addr, key:=LCase$(addr)
                                On Error GoTo 0
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r
    Set BuildOverdueRecipients = col
End Function

'==========================
'  RAPOR GÖNDERÝMÝ (mevcut akýþ korunur)
'==========================

Private Function ExportSheetPDF(ByVal ws As Worksheet) As String
    Dim path As String
    path = Environ$("TEMP") & "\" & ws.Name & "_" & Format(Date, "yyyymmdd") & ".pdf"
    On Error Resume Next
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=path, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    If Err.Number <> 0 Then path = ""
    On Error GoTo 0
    ExportSheetPDF = path
End Function



Public Sub SendOverdueReportForSheet(ByVal ws As Worksheet, Optional ByVal slot As String = "")
    Dim recipients As Collection, i As Long
    Dim OutApp As Object, OutMail As Object
    Dim subjectText As String, html As String, pdfPath As String
    Dim lastRow As Long, r As Long
    Dim vJ As Variant, planDate As Variant
    Dim totalTasks As Long, openTasks As Long, overdueNow As Long
    Dim closeSum As Double, closeCnt As Long, avgClose As Double
    Dim themeHex As String

    ' -- Kiþi bazlý istatistik sözlükleri --
    Dim dTotal As Object, dOpen As Object, dOver As Object
    Dim dCloseSum As Object, dCloseCnt As Object
    Set dTotal = CreateObject("Scripting.Dictionary")
    Set dOpen = CreateObject("Scripting.Dictionary")
    Set dOver = CreateObject("Scripting.Dictionary")
    Set dCloseSum = CreateObject("Scripting.Dictionary")
    Set dCloseCnt = CreateObject("Scripting.Dictionary")

    ' 1) Alýcýlar (ayný kiþi günde 1 kez)
    Set recipients = BuildOverdueRecipients(ws, slot)
    If recipients Is Nothing Then Exit Sub
    If recipients.Count = 0 Then Exit Sub

    ' 2) Metrikler
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    For r = 5 To lastRow
        Dim hasE As Boolean, hasF As Boolean, personKey As String
        hasE = (Len(Trim$(ws.Cells(r, "E").text)) > 0)
        hasF = (Len(Trim$(ws.Cells(r, "F").text)) > 0)
        vJ = ws.Cells(r, "J").Value2
        planDate = ws.Cells(r, "H").value

        personKey = Trim$(ws.Cells(r, "F").text)
        If personKey = "" Then personKey = "(Boþ)"

        ' Toplam görev (E & F dolu)
        If hasE And hasF Then
            totalTasks = totalTasks + 1
            If Not dTotal.Exists(personKey) Then dTotal.Add personKey, 0
            dTotal(personKey) = dTotal(personKey) + 1
        End If

        ' Açýk görev (J<%99)
        If hasE And hasF And IsNumeric(vJ) Then
            If vJ < 0.99 Then
                openTasks = openTasks + 1
                If Not dOpen.Exists(personKey) Then dOpen.Add personKey, 0
                dOpen(personKey) = dOpen(personKey) + 1
            End If
        End If

        ' Gecikmiþ görev (J<%99 ve H<Bugün)
        If hasE And hasF And IsNumeric(vJ) And IsDate(planDate) Then
            If vJ < 0.99 And CDate(planDate) < Date Then
                overdueNow = overdueNow + 1
                If Not dOver.Exists(personKey) Then dOver.Add personKey, 0
                dOver(personKey) = dOver(personKey) + 1
            End If
        End If

        ' Ortalama kapanma süresi (tamamlananlar) -> hem genel hem kiþi bazlý
        If IsNumeric(vJ) And vJ >= 0.99 Then
            If IsDate(ws.Cells(r, "I").value) And IsDate(ws.Cells(r, "G").value) Then
                Dim diff As Long
                diff = DateDiff("d", CDate(ws.Cells(r, "G").value), CDate(ws.Cells(r, "I").value))
                If diff < 0 Then diff = 0
                ' Genel
                closeSum = closeSum + diff
                closeCnt = closeCnt + 1
                ' Kiþi bazlý
                If Not dCloseSum.Exists(personKey) Then dCloseSum.Add personKey, 0
                If Not dCloseCnt.Exists(personKey) Then dCloseCnt.Add personKey, 0
                dCloseSum(personKey) = dCloseSum(personKey) + diff
                dCloseCnt(personKey) = dCloseCnt(personKey) + 1
            End If
        End If
    Next r
    If closeCnt > 0 Then avgClose = closeSum / closeCnt Else avgClose = 0

    ' 3) Tema rengi (Mod_Settings) -> hex
    themeHex = GetThemeColorHex(ws.Name)

    ' 4) Mail gövdesi (üst bant + iki özet + gecikmiþler tablosu)
    html = ""
    html = html & "<div style='background:" & themeHex & ";color:#ffffff;padding:10px 14px;"
    html = html & "font-family:Segoe UI,Arial,sans-serif;font-size:12pt;font-weight:600;'>"
    html = html & Format(Date, "dd.MM.yyyy") & " - Geciken Görevler Raporu / " & ws.Name & " Toplantýsý</div>"

    html = html & "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt;'>"
    html = html & "<p>Sayýn Yetkili,</p>"
    html = html & "<p>Bugün itibarýyla geciken görevlerin raporu ekte (PDF) ve aþaðýda özet/tablo halinde paylaþýlmýþtýr.</p>"

    ' --- Kýsa Özet - GENEL ---
    html = html & "<div style='border-left:4px solid " & themeHex & ";padding:8px 12px;background:#f7f9fc;'>"
    html = html & "<b>Kýsa Özet - Genel</b><br>"
    html = html & "• Toplam Görev Sayýsý: " & CStr(totalTasks) & "<br>"
    html = html & "• Açýk Görev Sayýsý: " & CStr(openTasks) & "<br>"
    html = html & "• Anlýk Geciken Görev Sayýsý: " & CStr(overdueNow) & "<br>"
    html = html & "• Ortalama Görev Kapanma Süresi: " & Format(avgClose, "0.0") & " gün"
    html = html & "</div><br>"

    ' --- Kýsa Özet - PERSONEL (Top 4) -> Açýk görev sayýsýna göre sýralama ---
    Dim top4 As Variant, idx As Long, key As String
    top4 = Top4ByOpen(dOpen, dOver, dTotal)

    html = html & "<div style='border-left:4px solid " & themeHex & ";padding:8px 12px;background:#f7f9fc;'>"
    html = html & "<b>Kýsa Özet - Personel (Top 4)</b><br>"

    html = html & "<table border='1' cellspacing='0' cellpadding='6' style='border-collapse:collapse;"
    html = html & "font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt;'>"
    html = html & "<tr style='background:" & themeHex & ";color:#ffffff;'>"
    html = html & "<th align='left'>Sorumlu</th>"
    html = html & "<th align='right'>Toplam</th>"
    html = html & "<th align='right'>Açýk</th>"
    html = html & "<th align='right'>Gecikmiþ</th>"
    html = html & "<th align='right'>Ort. Kapanma</th></tr>"

    If IsArray(top4) Then
        For idx = LBound(top4) To UBound(top4)
            key = CStr(top4(idx))
            Dim nmDisp As String, t As Long, o As Long, ov As Long, av As String
            nmDisp = ResolveFullName(key) ' Kod yerine Ad Soyadý

            t = IIf(dTotal.Exists(key), dTotal(key), 0)
            o = IIf(dOpen.Exists(key), dOpen(key), 0)
            ov = IIf(dOver.Exists(key), dOver(key), 0)

            If dCloseCnt.Exists(key) And dCloseCnt(key) > 0 Then
                av = Format(dCloseSum(key) / dCloseCnt(key), "0.0") & " gün"
            Else
                av = "—"
            End If

            html = html & "<tr><td>" & nmDisp & "</td>"
            html = html & "<td align='right'>" & t & "</td>"
            html = html & "<td align='right'>" & o & "</td>"
            html = html & "<td align='right'>" & ov & "</td>"
            html = html & "<td align='right'>" & av & "</td></tr>"
        Next idx
    Else
        html = html & "<tr><td colspan='5' align='center'>Veri bulunamadý</td></tr>"
    End If

    html = html & "</table></div><br>"

    ' --- Gecikmiþler detay tablosu (Notlar L ve Sorumlu = Ad Soyad) ---
    html = html & "<table border='1' cellspacing='0' cellpadding='6' style='border-collapse:collapse;"
    html = html & "font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt;'>"
    html = html & "<tr style='background:" & themeHex & ";color:#ffffff;'>"
    html = html & "<th align='left'>Konu/Madde (D)</th>"
    html = html & "<th align='left'>Görev/Aksiyon (E)</th>"
    html = html & "<th align='left'>Planlanan Tarih (H)</th>"
    html = html & "<th align='left'>Sorumlu (F)</th>"
    html = html & "<th align='left'>Notlar (L)</th>"
    html = html & "<th align='left'>% Tamamlanma (J)</th>"
    html = html & "</tr>"

    For r = 5 To lastRow
        vJ = ws.Cells(r, "J").Value2
        planDate = ws.Cells(r, "H").value
        If IsNumeric(vJ) And IsDate(planDate) Then
            If vJ < 0.99 And CDate(planDate) < Date Then
                Dim respDisp As String
                respDisp = ResolveFullName(ws.Cells(r, "F").text)

                html = html & "<tr>"
                html = html & "<td>" & ws.Cells(r, "D").text & "</td>"
                html = html & "<td>" & ws.Cells(r, "E").text & "</td>"
                html = html & "<td>" & Format(ws.Cells(r, "H").value, "dd.MM.yyyy") & "</td>"
                html = html & "<td>" & respDisp & "</td>"
                html = html & "<td>" & ws.Cells(r, "L").text & "</td>"
                html = html & "<td>" & Format(ws.Cells(r, "J").value, "0%") & "</td>"
                html = html & "</tr>"
            End If
        End If
    Next r
    html = html & "</table>"
    html = html & "<p>Dokümana hýzlý eriþim: <a href='" & ONEDRIVE_LINK & "'>buraya týklayýn</a>.</p>"
    html = html & "<p>Ýyi çalýþmalar.</p></div>"

    ' 5) Konu ve PDF
    subjectText = Format(Date, "dd.MM.yyyy") & " - Geciken Görevler Raporu / " & ws.Name & " Toplantýsý"
    pdfPath = ExportOverduePDF_FilteredCopy(ws)

    ' 6) Outlook gönderimi
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        For i = 1 To recipients.Count
            .recipients.Add recipients(i)
        Next i
        Call AddCC_ForReport(OutMail, ws.Name)
        .Subject = subjectText
        .BodyFormat = 2
        .htmlBody = html
        If Len(pdfPath) > 0 Then .Attachments.Add pdfPath
        .Save
        .Send
    End With

    For i = 1 To recipients.Count
        Call MarkDailyMailSent(recipients(i), ws.Name, slot)
       Next i

CleanExit:
    On Error Resume Next
    Set OutMail = Nothing
    Set OutApp = Nothing
    Exit Sub
    
End Sub


'================== Yardýmcýlar: Top 4 (Açýk göreve göre) + Ad Soyadý çözümleme ==================

' Top4: Önce Açýk (desc), sonra Gecikmiþ (desc), sonra Toplam (desc)
Private Function Top4ByOpen(ByVal dOpen As Object, ByVal dOver As Object, ByVal dTotal As Object) As Variant
    Dim unionDict As Object, k As Variant, keys As Variant
    Set unionDict = CreateObject("Scripting.Dictionary")

    ' Tüm kiþi anahtarlarýnýn birleþimi
    For Each k In dTotal.keys: unionDict(k) = True: Next k
    For Each k In dOpen.keys:  unionDict(k) = True: Next k
    For Each k In dOver.keys:  unionDict(k) = True: Next k

    If unionDict.Count = 0 Then
        Top4ByOpen = Array()
        Exit Function
    End If

    keys = unionDict.keys

    ' Küçük veri setleri için basit sýralama: desc (Açýk -> Gecikmiþ -> Toplam)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CompareByOpen(keys(i), keys(j), dOpen, dOver, dTotal) < 0 Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i

    ' Ýlk 4 kiþiyi döndür
    Dim n As Long: n = Application.Min(4, UBound(keys) - LBound(keys) + 1)
    ReDim Preserve keys(0 To n - 1)
    Top4ByOpen = keys
End Function

' Sýralama kriteri: Açýk v, eþitse Gecikmiþ v, yine eþitse Toplam v
Private Function CompareByOpen(ByVal a As String, ByVal b As String, _
                               ByVal dOpen As Object, ByVal dOver As Object, ByVal dTotal As Object) As Long
    Dim ao As Long, bo As Long, ag As Long, bg As Long, at As Long, bt As Long
    ao = IIf(dOpen.Exists(a), dOpen(a), 0)
    bo = IIf(dOpen.Exists(b), dOpen(b), 0)
    If ao <> bo Then CompareByOpen = Sgn(ao - bo): Exit Function

    ag = IIf(dOver.Exists(a), dOver(a), 0)
    bg = IIf(dOver.Exists(b), dOver(b), 0)
    If ag <> bg Then CompareByOpen = Sgn(ag - bg): Exit Function

    at = IIf(dTotal.Exists(a), dTotal(a), 0)
    bt = IIf(dTotal.Exists(b), dTotal(b), 0)
    CompareByOpen = Sgn(at - bt)
End Function

'==========================
'  PDF ÜRETÝMÝ — TEK SAYFA / LANDSCAPE
'==========================

Private Function ExportOverduePDF_FilteredCopy(ByVal ws As Worksheet) As String
    Dim tmp As Worksheet, lastRow As Long, lastCol As Long, r As Long
    Dim vJ As Variant, planDate As Variant
    Dim path As String, newName As String, startRow As Long
    Dim themeHex As String
    Dim printRange As Range

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 1) Kopya üret (stiller korunur)
    ws.Copy After:=ws
    Set tmp = ActiveSheet
    newName = "__tmp_pdf_" & left$(ws.Name, 20) & "_" & Format(Now, "hhmmss")
    On Error Resume Next: tmp.Name = newName: Err.Clear: On Error GoTo 0

    ' 2) Yalnýz geciken satýrlar kalsýn
    startRow = 5
    lastRow = tmp.Cells(tmp.Rows.Count, "F").End(xlUp).Row
    For r = lastRow To startRow Step -1
        vJ = tmp.Cells(r, "J").Value2
        planDate = tmp.Cells(r, "H").value
        If IsNumeric(vJ) And IsDate(planDate) Then
            If Not (vJ < 0.99 And CDate(planDate) < Date) Then
                tmp.Rows(r).Delete
            End If
        Else
            tmp.Rows(r).Delete
        End If
    Next r

    ' 3) Üste daha fazla boþluk (logo/baþlýk için)
    tmp.Rows(1).Resize(8).Insert            ' 8 satýr boþluk
    ' Döngüsüz toplu satýr yüksekliði — daha stabil
    tmp.Rows("1:8").RowHeight = 22

    ' 4) Logo + Baþlýk (tek logo, koyu bantta ortalanmýþ baþlýk)
    Dim logo As Shape
    Set logo = EnsureSingleLogo(tmp, 230)    ' tek logo ve büyük yerleþim (mevcut yordamýnýza uygun) [1](https://isotecsolar.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B7632AC1B-A852-4283-BD26-ABDCD6BFE201%7D&file=Genel_Aksiyon_Listesi.xlsm&action=default&mobileredirect=true)
    themeHex = GetThemeColorHex(ws.Name)     ' tema rengi (Mod_Settings) [2](https://isotecsolar-my.sharepoint.com/personal/ceyhun_bostanci_isotec_com_tr/Documents/Microsoft%20Copilot%20Chat%20Dosyalar%C4%B1/Module_Mail.rtf)
    Dim headerText As String
    headerText = Format(Date, "dd.MM.yyyy") & " - Geciken Görevler Raporu / " & ws.Name & " Toplantýsý"
    Call StampHeaderBand(tmp, headerText, themeHex)  ' ortalanmýþ koyu bant sürümünü kullanýyoruz

    ' 5) Baský ayarlarýný temizle + tek sayfa LANDSCAPE
    With tmp.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .PrintArea = ""
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1        ' tek sayfa zorlamasý
        .CenterHorizontally = True
        .CenterVertically = False
        .LeftMargin = Application.InchesToPoints(0.35)
        .RightMargin = Application.InchesToPoints(0.35)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
        .PrintGridlines = False
    End With

    ' 6) Sayfa sonlarýný sýfýrla
    tmp.ResetAllPageBreaks

    ' 7) Baský alaný (üst boþluk + tüm tablo)
    lastCol = tmp.UsedRange.Columns(tmp.UsedRange.Columns.Count).Column
    lastRow = tmp.Cells(tmp.Rows.Count, "F").End(xlUp).Row
    If lastRow < startRow Then
        ExportOverduePDF_FilteredCopy = ""
        GoTo CleanExit
    End If
    Set printRange = tmp.Range(tmp.Cells(1, 1), tmp.Cells(lastRow, lastCol))
    tmp.PageSetup.PrintArea = printRange.Address

    ' 8) PDF üret
    path = Environ$("TEMP") & "\" & ws.Name & "_Overdue_" & Format(Date, "yyyymmdd") & ".pdf"
    tmp.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False

    ExportOverduePDF_FilteredCopy = path

CleanExit:
    On Error Resume Next
    tmp.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Function

Fail:
    On Error Resume Next
    If Not tmp Is Nothing Then tmp.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ExportOverduePDF_FilteredCopy = ""
End Function

Public Sub RunDailyOverdueReports_WithSlot(ByVal slot As String)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsMeetingSheet(ws) Then
            SendOverdueReportForSheet ws, slot
        End If
    Next ws
End Sub


