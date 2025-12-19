Attribute VB_Name = "Mod_ManualExport"
Option Explicit

Public Sub ManualExportCommitPush()
    If Not AcquireGlobalLock Then
        MsgBox "Baþka bir iþlem çalýþýyor.", vbExclamation
        Exit Sub
    End If

    On Error GoTo CLEANUP

    ExportCommitPush_Smart

CLEANUP:
    ReleaseGlobalLock
End Sub

