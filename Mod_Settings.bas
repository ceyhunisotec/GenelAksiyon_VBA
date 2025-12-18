Attribute VB_Name = "Mod_Settings"

'== Mod_Settings.bas ==
Option Explicit

' -- AYARLAR --
Public Const ONEDRIVE_LINK As String = _
    "https://isotecsolar.sharepoint.com/:x:/g/IQAbrDJ2UqiDQr0mq9zWv-IBAbagXQ95upAHjnfof9XWS0U?e=4P2CXX"

Public Const CC_LIST As String = "ceyhun@isotec.com.tr; halisahmet.orhan@isotec.com.tr"

' LOGO (PNG önerilir) - yerel yol hýzlýdýr; yoksa Assets/CompanyLogo kullanýlacak
Public Const LOGO_PATH As String = _
    "C:\Users\ceyhun.bostanci\OneDrive - ISOTEC Enerji A.Þ\Masaüstü\Resim1.png"

' Tema renkleri
Public Function GetThemeColorHex(ByVal sheetName As String) As String
    Select Case LCase$(sheetName)
        Case LCase$("Koordinasyon"): GetThemeColorHex = "#0078D4"
        Case LCase$("Sipariþ"), LCase$("Siparis"): GetThemeColorHex = "#107C10"
        Case LCase$("Þikayet"), LCase$("Sikayet"): GetThemeColorHex = "#C50F1F"
        Case LCase$("Atýl_Stok"), LCase$("Atil_Stok"): GetThemeColorHex = "#8E562E"
        Case LCase$("Kalite"): GetThemeColorHex = "#2F5597"
        Case Else: GetThemeColorHex = "#2F5597"
    End Select
End Function

' Toplantý sayfasý mý?
Public Function IsMeetingSheet(ByVal sh As Object) As Boolean
    Select Case sh.Name
        Case "Koordinasyon", "Sipariþ", "Þikayet", "Atýl_Stok", "Kalite"
            IsMeetingSheet = True
        Case Else
            IsMeetingSheet = False
    End Select
End Function


