Attribute VB_Name = "Mod_Resolver"

'== Mod_Resolver.bas ==
Option Explicit

' Kod/ad/e-posta giriþi -> "Adý Soyadý" (Data sayfasý: A=Kod, B=Ad Soyad, D=Mail)
Public Function ResolveFullName(ByVal responsibleInput As String) As String
    Dim ws As Worksheet, f As Range, s As String, mail As String
    s = Trim$(responsibleInput)
    If Len(s) = 0 Then ResolveFullName = "": Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Data")
    On Error GoTo 0
    If ws Is Nothing Then
        ResolveFullName = responsibleInput
        Exit Function
    End If

    ' 1) Kod ile TAM eþleþme (A sütunu)
    Set f = ws.Columns("A").Find(What:=s, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then
        ResolveFullName = CStr(f.Offset(0, 1).value) ' B: Ad Soyad
        Exit Function
    End If

    ' 2) Eðer giriþ e-posta içeriyorsa D sütununda ara
    mail = LCase$(s)
    mail = Replace$(mail, "mailto:", "")
    mail = Replace$(mail, "<", "")
    mail = Replace$(mail, ">", "")
    If InStr(mail, " ") > 0 Then mail = left$(mail, InStr(mail, " ") - 1)
    If InStr(mail, "@") > 0 Then
        Set f = ws.Columns("D").Find(What:=mail, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not f Is Nothing Then
            ResolveFullName = CStr(f.Offset(0, -2).value) ' D->B = -2
            Exit Function
        End If
    End If

    ' 3) Ad ile KISMÝ arama (B sütunu)
    Set f = ws.Columns("B").Find(What:=s, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not f Is Nothing Then
        ResolveFullName = CStr(f.value)
        Exit Function
    End If

    ' 4) Bulunamazsa olduðu gibi döndür
       ResolveFullName = responsibleInput
       
End Function

