VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmCapture 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Till Capture System"
   ClientHeight    =   7800
   ClientLeft      =   9540
   ClientTop       =   5160
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   7065
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "Capture From"
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   2535
      Begin VB.Frame Frame3 
         Caption         =   "IP"
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1455
         Begin VB.TextBox txtIP 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Text            =   "192.168.0.6"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Port"
         Height          =   615
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   855
         Begin VB.TextBox txtPORT 
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Text            =   "2001"
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2760
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
      LocalPort       =   1001
   End
   Begin VB.Frame Frame1 
      Caption         =   "Capture Area"
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   25
         Text            =   "1"
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox chkPreview 
         Caption         =   "Show Preview"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   15
         Text            =   "300"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   14
         Text            =   "400"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   13
         Text            =   "600"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   12
         Text            =   "800"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   11
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   10
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "rate"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "targetHeight"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "targetWidth"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "srcHeight"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "srcWidth"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "startTop"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "startLeft"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Preview Window"
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Waiting...."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frameInCounter As Long
Dim dataBuffer As String
Dim tmpCaptureStream As String
Dim capFN As Integer
Dim StartTime As Date
Dim StopTime As Date
Dim DataLastReceived As Date

Private Sub cmdStart_Click()
    Dim cSTART As Date
    Dim i As Integer
    Dim dataSEND As String
    Dim fc As cssHEADER
    Dim sendVAL As String
    
    On Local Error Resume Next
    
    Label1 = "Initialising Stream..."
    DoEvents
    
    Winsock1.RemoteHost = txtIP
    Winsock1.RemotePort = txtPORT
    
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    
    Winsock1.Connect txtIP, txtPORT
    DoEvents
    DoEvents
    cSTART = Now()
    While Winsock1.State <> sckConnected And Abs(DateDiff("s", Now(), cSTART)) < 10
        DoEvents
    Wend
    
    If Winsock1.State <> sckConnected Then
        MsgBox "Connection Failed: Current State " & Winsock1.State
        Exit Sub
    End If
    
    ' send cpature settings
    For i = 0 To 6
        sendVAL = Text1(i).Text
        
        If Label2(i) = "targetWidth" Then
            Image1.Width = Val(Text1(i)) * Screen.TwipsPerPixelX
            Me.Width = Image1.Left + Image1.Width + 200
        ElseIf Label2(i) = "targetHeight" Then
            Image1.Height = Val(Text1(i)) * Screen.TwipsPerPixelY
            If Me.Height < Image1.Top + Image1.Height + 250 Then
                Me.Height = Image1.Top + Image1.Height + 250
            End If
        ElseIf Label2(i) = "rate" Then
            sendVAL = Val(sendVAL) * 1000
        End If
        
        dataSEND = Label2(i) & "=" & sendVAL & "&&"
        List1.AddItem dataSEND
        
        Winsock1.SendData dataSEND
        
        ' wait 1 second
        cSTART = Now()
        While Abs(DateDiff("s", Now(), cSTART)) < 1
            DoEvents
        Wend
    Next i
    
    frameInCounter = 0
    fc.jpgSize = 0
    
    ' stream ready for start command
    tmpCaptureStream = App.Path & "\cap" & Format$(Date, "ddmmyyyy") & Format$(Time$, "HHMMSS") & ".tmp"
    capFN = FreeFile
    Open tmpCaptureStream For Binary Access Read Write As #capFN
    Put #capFN, , fc
    
    Winsock1.SendData "start=0&&"
    List1.AddItem "start=0&&"
    StartTime = Now()
    cmdStart.Enabled = False
    cmdStop.Enabled = True
End Sub

Private Sub cmdStop_Click()
    Dim cSTART As Date
    Dim capFile As String
    Dim fc As cssHEADER
    
    On Local Error Resume Next

    cmdStop.Enabled = False
    Label1 = "Closing Stream..."
    DoEvents
    
    Winsock1.SendData "stop=0&&"
    StopTime = Now()
    
    ' wait for last bit of stream
    cSTART = Now()
    While Abs(DateDiff("s", Now(), DataLastReceived)) < 10
        DoEvents
    Wend
    
    Winsock1.Close
    
    fc.jpgSize = frameInCounter
    Put #capFN, 1, fc
    
    
    Close #capFN
    capFile = App.Path & "\Capture~" & Format$(Date, "ddmmmyyyy") & "~" & Format$(StartTime, "HHMMSS") & "~" & Format$(StopTime, "HHMMSS") & ".cap"
    Name tmpCaptureStream As capFile
    
    cmdStop.Enabled = False
    cmdStart.Enabled = True
    Label1 = "Ready..."
End Sub



Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim dataIN As String
    Dim frameFN As Integer
    Dim frameDATA As String
    Dim frameEND As Long
    Dim cFILE As String
    Dim cssHDR As cssHEADER
    
    On Local Error Resume Next
    
    Winsock1.GetData dataIN, vbString
    dataBuffer = dataBuffer & dataIN
    DataLastReceived = Now()
    
    If Left$(dataBuffer, 6) <> "//fs\\" Then Exit Sub
    If InStr(1, dataBuffer, "//fe\\", vbTextCompare) = 0 Then Exit Sub
    
    frameInCounter = frameInCounter + 1
    Label1 = "Frame Received " & frameInCounter
    Label1.Refresh
    
    ' we have a frame
    frameEND = InStr(1, dataBuffer, "//fe\\", vbTextCompare) + 5
    frameDATA = Left$(dataBuffer, frameEND)
    dataBuffer = Right$(dataBuffer, Len(dataBuffer) - frameEND)
    
    ' strip control characters
    frameDATA = Left$(frameDATA, Len(frameDATA) - 6)
    frameDATA = Right$(frameDATA, Len(frameDATA) - 6)
    
    ' write to a tmp bmp file
    frameFN = FreeFile
    Open App.Path & "\frame.jpg" For Binary Access Write As #frameFN
    Put #frameFN, , frameDATA
    Close #frameFN
    frameDATA = ""
    
    If chkPreview.Value = 1 Then
        Image1.Picture = LoadPicture(App.Path & "\frame.jpg")
    End If
    
    'add JPG to stream
    cssHDR.jpgSize = FileLen(App.Path & "\frame.jpg")
    Put #capFN, , cssHDR
    
    frameDATA = Space$(cssHDR.jpgSize)
    
    ' 1st need to load jpg into a buffer
    Open App.Path & "\frame.jpg" For Binary Access Read As #frameFN
    Get #frameFN, , frameDATA
    Close #frameFN
    
    Put #capFN, , frameDATA
    
    frameDATA = ""
    
End Sub

