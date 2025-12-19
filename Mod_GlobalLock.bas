Attribute VB_Name = "Mod_GlobalLock"
Option Explicit

Private Function GetGlobalLockFile() As String
    GetGlobalLockFile = Environ$("TEMP") & "\GenelAksiyon_GLOBAL.lock"
End Function

