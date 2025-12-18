VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'== CAppEvents ==
Option Explicit
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_SheetChange(ByVal sh As Object, ByVal Target As Range)
    On Error GoTo ExitPoint

    If Not IsMeetingSheet(sh) Then Exit Sub
    If Application.EnableEvents = False Then Exit Sub

    Application.EnableEvents = False

    Dim rngF As Range, rngJ As Range, rngEFH As Range

    ' -- F (Sorumlu) › G (Listeye Giriþ) tarih --
    Set rngF = Intersect(Target, sh.Columns("F"))
    If Not rngF Is Nothing Then
        Dim c As Range
        For Each c In rngF.Cells
            If Len(c.value) > 0 And IsEmpty(sh.Cells(c.Row, "G").value) Then
                sh.Cells(c.Row, "G").value = Date
            End If
        Next c
    End If

    
' -- J (%Tamamlanma) ? %99 › I tarih; K boþsa uyar; TAMAMLAMA MAILI (bir kez) --
Set rngJ = Intersect(Target, sh.Columns("J"))
If Not rngJ Is Nothing Then
    Dim jCell As Range, v As Variant
    For Each jCell In rngJ.Cells
        v = jCell.Value2
        If IsNumeric(v) Then
            If v >= 0.99 Then
                ' Bitiþ tarihi bir kez
                If IsEmpty(sh.Cells(jCell.Row, "I").value) Then
                    sh.Cells(jCell.Row, "I").value = Date
                End If
                ' Kanýt uyarýsý
                If Len(Trim$(sh.Cells(jCell.Row, "K").value)) = 0 Then
                    MsgBox "Kanýt Ekleyiniz (Link)", vbExclamation, "Uyarý"
                End If
                ' Tamamlama maili (N bayraðý yoksa)
                If Trim$(CStr(sh.Cells(jCell.Row, "N").value)) <> "TamMail" Then
                    If SendCompletionMail(sh, jCell.Row, ONEDRIVE_LINK) = True Then
                        sh.Cells(jCell.Row, "N").value = "TamMail"
                    End If
                End If
            End If
        End If
    Next jCell
End If


    ' -- E, F, H ilk kez dolduysa ve J<%99 ise mail gönder (M boþsa) --
    Set rngEFH = Intersect(Target, sh.Range("E:H"))
    If Not rngEFH Is Nothing Then
        Dim r As Range, rowNum As Long, vJ As Variant
        For Each r In rngEFH.Rows
            rowNum = r.Row

            If Len(sh.Cells(rowNum, "E").value) > 0 _
               And Len(sh.Cells(rowNum, "F").value) > 0 _
               And Len(sh.Cells(rowNum, "H").value) > 0 Then

                vJ = sh.Cells(rowNum, "J").Value2
                If Not IsNumeric(vJ) Then vJ = 0

                If Trim$(CStr(sh.Cells(rowNum, "M").value)) <> "Gönderildi" _
                   And vJ < 0.99 Then

                    If SendTaskMail(sh, rowNum, ONEDRIVE_LINK) = True Then
                        sh.Cells(rowNum, "M").value = "Gönderildi"
                    End If
                End If
            End If
        Next r
    End If

ExitPoint:
    Application.EnableEvents = True
End Sub


