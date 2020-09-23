Attribute VB_Name = "modINI"
Option Compare Text

Public INIFile As String

Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ReadINI(Section As String, KeyName As String, sFileName As String) As String
    Dim sRet As String
    On Local Error Resume Next
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, getprivateprofilestring(Section, ByVal KeyName$, "", sRet, Len(sRet), sFileName))
End Function

Public Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    On Local Error Resume Next
    r = writeprivateprofilestring(sSection, sKeyName, sNewString, sFileName)
End Function


