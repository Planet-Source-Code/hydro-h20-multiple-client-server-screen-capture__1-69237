VERSION 5.00
Begin VB.Form frmCapture 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capture"
   ClientHeight    =   2310
   ClientLeft      =   3465
   ClientTop       =   4050
   ClientWidth     =   3795
   Icon            =   "frmCapture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3795
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5880
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5880
      Top             =   2400
   End
   Begin VB.PictureBox picRSetMode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1560
      Left            =   1080
      ScaleHeight     =   1560
      ScaleWidth      =   1560
      TabIndex        =   1
      Top             =   3000
      Width           =   1560
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   4440
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6120
      Picture         =   "frmCapture.frx":0D4A
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5160
      Picture         =   "frmCapture.frx":1A94
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dataBuffer As String

Private Sub Form_Load()
    On Local Error Resume Next
    Me.Show
    DoEvents
    
    MkDir App.Path & "\Captures"
    
    TrayAdd hwnd, Image1.Picture, "CaptureIt", MouseMove
    Me.WindowState = vbMinimized
    Me.Hide
    
    ' before activating timer2
    ' kill any startcapture/stopcapture/recording files
    ' keep things tidy
    If Len(Dir$(App.Path & "\StartCapture.dat")) <> 0 Then Kill App.Path & "\StartCapture.dat"
    If Len(Dir$(App.Path & "\StopCapture.dat")) <> 0 Then Kill App.Path & "\StopCapture.dat"
    If Len(Dir$(App.Path & "\Recording.dat")) <> 0 Then Kill App.Path & "\Recording.dat"
    
    Timer2.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim cEvent As Single
    
    On Local Error Resume Next
    
    cEvent = x / Screen.TwipsPerPixelX
    Select Case cEvent
        Case MouseMove
            'Debug.Print "MouseMove"
        Case LeftUp
            'Debug.Print "Left Up"
        Case LeftDown
            'Debug.Print "LeftDown"
        Case LeftDbClick
            'Debug.Print "LeftDbClick"
            If Me.WindowState = 1 Then
                WindowState = 0
                Me.Show
            End If
        Case MiddleUp
            'Debug.Print "MiddleUp"
        Case MiddleDown
            'Debug.Print "MiddleDown"
        Case MiddleDbClick
            'Debug.Print "MiddleDbClick"
        Case RightUp
            'Debug.Print "RightUp"
        Case RightDown
            'Debug.Print "RightDown"
        Case RightDbClick
            'Debug.Print "RightDbClick"
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Local Error Resume Next
    TrayDelete
End Sub

Private Sub Timer1_Timer()
    On Local Error Resume Next
    
    If LastTimeCaptured = Format$(Time$, "HH:MM:SS") Then Exit Sub
    
    LastTimeCaptured = Format$(Time$, "HH:MM:SS")

    FrameCounter = FrameCounter + 1
    CaptureDesktop
    
    
    ' the server may go down or be turned off, this could lead to the capture files
    ' filling up the HD
    ' so to keep things tidy when there are more than 300 files or more
    ' then recording will stop automatically which gives us 5mins worth
    Call CheckCaptures
    
End Sub

Private Sub CheckCaptures()
    Dim l As String
    Dim i As Integer
    Dim sFN As Integer
    
    l = Dir$(App.Path & "\Captures\cap*.jpg")
    While Len(l) <> 0
        i = i + 1
        l = Dir$
    Wend
    
    If i >= 300 Then
        sFN = FreeFile
        Open App.Path & "\StopCapture.dat" For Random As #sFN
        Close #sFN
    End If

End Sub

Private Sub Timer2_Timer()
    On Local Error Resume Next
    
    If Len(Dir$(App.Path & "\StartCapture.dat")) <> 0 Then
        TrayModify Tray_Icon, Image2.Picture
        If Timer1.Enabled Then
            ' if already capturing do nothing
            Kill App.Path & "\StartCapture.dat"
            Exit Sub
        End If
        Call Start_New_Capture
    End If
        
    If Len(Dir$(App.Path & "\StopCapture.dat")) <> 0 Then
        TrayModify Tray_Icon, Image1.Picture
        Timer1.Enabled = False
        Kill App.Path & "\StopCapture.dat"
        If Len(Dir$(App.Path & "\Recording.dat")) <> 0 Then Kill App.Path & "\Recording.dat"
    End If
        
End Sub


Private Sub Start_New_Capture()
    Dim s As String
    Dim fn As Integer
    Dim dataCMD As String
    Dim dataVAL As String
    
    On Local Error Resume Next
    
    List1.Clear
    
    fn = FreeFile
    Open App.Path & "\StartCapture.dat" For Input As #fn
    
    While EOF(fn) = False
        Line Input #fn, s
        
        Call Add_To_List(s)
    
        dataCMD = Left$(s, InStr(1, s, "=", vbTextCompare) - 1)
        dataVAL = Right$(s, Len(s) - (Len(dataCMD) + 1))
    
        Select Case dataCMD
            Case "srcHeight"
                CaptureArea.srcHeight = Val(dataVAL)
                Picture1.Height = CaptureArea.srcHeight * Screen.TwipsPerPixelY
            Case "srcWidth"
                CaptureArea.srcWidth = Val(dataVAL)
                Picture1.Width = CaptureArea.srcWidth * Screen.TwipsPerPixelX
            Case "startLeft"
                CaptureArea.startLeft = Val(dataVAL)
            Case "startTop"
                CaptureArea.startTop = Val(dataVAL)
            Case "targetWidth"
                CaptureArea.targetWidth = Val(dataVAL)
                picRSetMode.Width = CaptureArea.targetWidth * Screen.TwipsPerPixelX
            Case "targetHeight"
                CaptureArea.targetHeight = Val(dataVAL)
                picRSetMode.Height = CaptureArea.targetHeight * Screen.TwipsPerPixelY
            Case "rate"
'                Timer1.Interval = Val(dataVAL)
            Case "compression"
                CaptureArea.Compression = Val(dataVAL)
            Case "DTStamp"
                CaptureArea.DTStamp = CBool(dataVAL)
            Case "Quality"
                CaptureArea.qualityH = Val(Left$(dataVAL, 1))
                CaptureArea.qualityV = Val(Right$(dataVAL, 1))
            Case "Convert256"
                CaptureArea.Convert256 = CBool(dataVAL)
            Case "CaptureMouse"
                CaptureArea.CaptureMouse = CBool(dataVAL)
        End Select
    Wend
    
    Close #fn
    Kill App.Path & "\StartCapture.dat"
    
    Open App.Path & "\Recording.dat" For Output As #fn
    Close #fn
    
    FrameCounter = 0
    tmpBMPname = App.Path & "\capCurrentFrame.bmp"
    tmpBMP256name = App.Path & "\capCurrentFrame256.bmp"
    
    LastTimeCaptured = Format$(Time$, "HH:MM:SS")
    Timer1.Enabled = True

End Sub

Private Sub Add_To_List(lMSG As String)
    On Local Error Resume Next
    List1.AddItem lMSG
    List1.ListIndex = List1.ListCount - 1
End Sub
