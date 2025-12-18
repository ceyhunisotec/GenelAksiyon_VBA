Attribute VB_Name = "modLog"

'=== Module: modLog ===

Option Explicit

Public Sub LogToSysLog(senderEmail As String, sheetName As String, note As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("SysLog")
    Dim r As Long: r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ws.Cells(r, 1).value = Date
    ws.Cells(r, 2).value = senderEmail
    ws.Cells(r, 3).value = sheetName
    ws.Cells(r, 4).value = note
End Sub

