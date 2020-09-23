Attribute VB_Name = "modPING"
Private Type QOCINFO
  dwSize As Long
  dwFlags As Long
  dwInSpeed As Long 'in bytes/second
  dwOutSpeed As Long 'in bytes/second
End Type


Private Declare Function IsDestinationReachable Lib "SENSAPI.DLL" Alias "IsDestinationReachableA" (ByVal lpszDestination As String, ByRef lpQOCInfo As QOCINFO) As Long


Public Function Ping(IP As String) As Boolean
    Dim Ret As QOCINFO
  
    On Local Error Resume Next
    
    Ret.dwSize = Len(Ret)
    If IsDestinationReachable(IP, Ret) = 0 Then
        Ping = False
    Else
        Ping = True
        'MsgBox "The destination can be reached!" + vbCrLf + _
        '"The speed of data coming in from the destination is " + Format$(Ret.dwInSpeed / 1048576, "#.0") + " Mb/s," + vbCrLf + _
        '"and the speed of data sent to the destination is " + Format$(Ret.dwOutSpeed / 1048576, "#.0") + " Mb/s."
    End If
End Function



