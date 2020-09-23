Attribute VB_Name = "modGetCompName"
Option Explicit
Option Compare Text

Private Declare Function w32_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetComputerName() As String
    On Local Error Resume Next
    
    Dim sComputerName As String
    Dim sCompName As String
    Dim lComputerNameLen As Long
    Dim lResult As Long

    lComputerNameLen = 256
    sComputerName = Space$(lComputerNameLen)

    lResult = w32_GetComputerName(sComputerName, lComputerNameLen)
    If lResult = 1 Then
        sCompName = sComputerName
        sCompName = LTrim$(RTrim$(sCompName))
        sCompName = StripNulls(sCompName)
        GetComputerName = sCompName
    Else
        GetComputerName = vbNullString
    End If

End Function

Private Function StripNulls(sTrip As String) As String
    Dim i As Integer
        
    On Local Error Resume Next
    
    StripNulls = sTrip
    If Len(sTrip) Then
        i = InStr(sTrip, Chr$(0))
        If i Then StripNulls = left$(sTrip, i - 1)
    End If

End Function
