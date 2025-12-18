Attribute VB_Name = "modPrompt"

Option Explicit

Public Function BuildSmartPrompt(question As String, ctx As String) As String
    Dim q As String: q = LCase$(Trim$(question))
    Dim styleGuide As String
    Dim outputFormat As String

    styleGuide = _
        "Sen bir toplantý karar takip ve iþ planý asistanýsýn." & vbCrLf & _
        "Cevap Türkçe olacak. Kýsa, net, madde madde yaz." & vbCrLf & _
        "Sadece verilen BAÐLAM'a dayan; uydurma bilgi verme." & vbCrLf & _
        "Mümkünse sayfa adlarýný (Koordinasyon/Sipariþ/Þikayet/Atýl_Stok/Kalite) belirt." & vbCrLf

    outputFormat = _
        "Çýktý formatý:" & vbCrLf & _
        "1) Genel durum (Toplam/Açýk/Gecikmiþ/Bugün planlý)" & vbCrLf & _
        "2) En kritik 3 konu (kýsa gerekçe)" & vbCrLf & _
        "3) Önerilen aksiyonlar (maks 5 madde)" & vbCrLf

    If InStr(q, "gecik") > 0 Or InStr(q, "geç") > 0 Or InStr(q, "overdue") > 0 Then
        outputFormat = _
            "Çýktý formatý (GECÝKENLER):" & vbCrLf & _
            "- Sayfa bazýnda gecikmiþ sayýsý" & vbCrLf & _
            "- Risk/etki (kýsa)" & vbCrLf & _
            "- Ýlk 5 öncelik (neden?)" & vbCrLf & _
            "- Hýzlý aksiyon önerisi (maks 5 madde)" & vbCrLf
    End If

    If InStr(q, "bugün") > 0 Or InStr(q, "today") > 0 Then
        outputFormat = _
            "Çýktý formatý (BUGÜN):" & vbCrLf & _
            "- Bugün planlý maddeler: sayfa bazýnda adet" & vbCrLf & _
            "- Bugün en kritik 3 baþlýk" & vbCrLf & _
            "- Engeller/Riskler" & vbCrLf & _
            "- Gün sonu hedefi" & vbCrLf
    End If

    If InStr(q, "risk") > 0 Or InStr(q, "kritik") > 0 Or InStr(q, "acil") > 0 Then
        outputFormat = _
            "Çýktý formatý (RÝSK/KRÝTÝK):" & vbCrLf & _
            "- En kritik 5 risk (sayfa + kýsa açýklama)" & vbCrLf & _
            "- Etki (Yüksek/Orta/Düþük) + gerekçe" & vbCrLf & _
            "- Önleyici aksiyon önerisi" & vbCrLf
    End If

    If InStr(q, "mail") > 0 Or InStr(q, "e-posta") > 0 Or InStr(q, "gönder") > 0 Then
        outputFormat = _
            "Çýktý formatý (MAÝL UYUMLU):" & vbCrLf & _
            "- Konu önerisi (1 satýr)" & vbCrLf & _
            "- 5 satýrlýk yönetici özeti" & vbCrLf & _
            "- Madde madde aksiyon listesi (maks 7 madde)" & vbCrLf
    End If

    BuildSmartPrompt = _
        styleGuide & vbCrLf & _
        outputFormat & vbCrLf & _
        "BAÐLAM:" & vbCrLf & ctx & vbCrLf & vbCrLf & _
        "SORU:" & vbCrLf & question
        
    
    If InStr(q, "koordinasyon") = 0 And InStr(q, "sipariþ") = 0 And InStr(q, "þikayet") = 0 And InStr(q, "kalite") = 0 And InStr(q, "atýl") = 0 And InStr(q, "aksiyon") = 0 Then
        ' Genel soru: baðlamý minimum tut
        BuildSmartPrompt = styleGuide & vbCrLf & _
                        "Genel bir soru. Dosya baðlamýna ihtiyaç yoksa kullanma." & vbCrLf & _
                        "SORU:" & vbCrLf & question
        Exit Function
    End If
    
End Function
