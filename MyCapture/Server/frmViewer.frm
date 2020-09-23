VERSION 5.00
Begin VB.Form frmViewer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2895
   ClientLeft      =   4230
   ClientTop       =   4275
   ClientWidth     =   4785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2895
   ScaleWidth      =   4785
   Visible         =   0   'False
   Begin VB.Timer timerAuto 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4920
      Top             =   1560
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   6480
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4920
      Top             =   960
   End
   Begin VB.Label lblMachine 
      Caption         =   "Label2"
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgBLUE 
      Height          =   420
      Left            =   4920
      Picture         =   "frmViewer.frx":000C
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   50
      TabIndex        =   0
      Top             =   50
      Width           =   4215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3015
      Left            =   50
      Picture         =   "frmViewer.frx":0305
      Stretch         =   -1  'True
      Top             =   355
      Width           =   4215
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastImageCount As Integer
Dim InCaptureMode As Boolean
Dim LastTime As String



Private Sub Form_Load()
    On Local Error Resume Next
    ' helps when loading up forms
    ' reduces flicker
    Me.Width = 0
    Me.Height = 0
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    Label1.Width = Me.Width - 130
    Image1.Width = Me.Width - 130
    Image1.Height = Me.Height - Label1.Height - 180
    LastTime = Format$(Time$, "HH:MM")
End Sub

Private Sub Image1_DblClick()
    On Local Error Resume Next
    
    If ViewMode = 1 Then Exit Sub
    
    ViewMode = 1
    ViewerID = Left$(Label1.Caption, 2)
    
    Call Arrange_Capture_Forms(ViewMode, ViewerID)
End Sub

Private Sub Timer1_Timer()
    Dim cFRAME As String
    
    On Local Error Resume Next
    
    
    File1.Refresh
    If File1.ListCount = 0 Then
        LastImageCount = LastImageCount + 1
        If LastImageCount < 20 Then Exit Sub
        
        Image1.Picture = imgBLUE.Picture
        
        ' 10secs past and no image
        LastImageCount = 0
        InCaptureMode = False
        If Len(Dir$(App.Path & "\" & lblMachine.Caption & "\NewCapture.mcp")) = 0 Then Exit Sub
    
        ' call a generic function that will rename capture file off
        Call Close_Capture(App.Path & "\" & lblMachine.Caption & "\NewCapture.mcp", lblMachine.Caption)
        ActiveRecordings = ActiveRecordings - 1
        Exit Sub
    End If
    
    If InCaptureMode = False Then
        InCaptureMode = True
        ActiveRecordings = ActiveRecordings + 1
    End If
    
    LastImageCount = 0
    cFRAME = File1.Path & "\" & File1.List(0)
    
    Image1.Picture = LoadPicture(cFRAME)
    
    ' call generic function that will append frame to a mcp file
    Call Write_Frame(App.Path & "\" & lblMachine.Caption & "\NewCapture.mcp", cFRAME, lblMachine.Caption)
    
    
    Kill cFRAME
    
End Sub

Private Sub timerAuto_Timer()
    Dim cTIME As String
    Dim i As Integer
    Dim schedSTART As String
    Dim schedSTOP As String
    Dim startTIME As Date
    Dim stopTIME As Date
    Dim mName As String
    
    On Local Error Resume Next
    
    cTIME = Format$(Time$, "HH:MM")
    If LastTime = cTIME Then Exit Sub
    LastTime = cTIME
    
    i = Left$(Label1.Caption, 2)
    
    schedSTART = ReadINI("Location" & i, "StartAt", INIFile)
    schedSTOP = ReadINI("Location" & i, "StopAt", INIFile)
    mName = ReadINI("Location" & i, "Name", INIFile)
    
    startTIME = CDate(Format$(Date, "dd/mmm/yyyy") & " " & schedSTART & ":00")
    stopTIME = CDate(Format$(Date, "dd/mmm/yyyy") & " " & schedSTOP & ":00")
    
    ' see if after schedule start but before schedule end
    If Now() >= startTIME And Now() < stopTIME Then
        ' should be recording now
        If ReadINI("Location" & i, "LastSchedStart", INIFile) <> Format$(Date, "dd/mmm/yyyy") Then
            ' not started today
            Call Start_Capture(mName)
            WriteINI "Location" & i, "LastSchedStart", Format$(Date, "dd/mmm/yyyy"), INIFile
        End If
        Exit Sub
    End If
    
    ' see if after stoptime, but prior to starttime
    If Now() >= stopTIME Then
        If ReadINI("Location" & i, "LastSchedStop", INIFile) <> Format$(Date, "dd/mmm/yyyy") Then
            ' not stopped today
            Call Stop_Capture(mName)
            WriteINI "Location" & i, "LastSchedStop", Format$(Date, "dd/mmm/yyyy"), INIFile
        End If
        Exit Sub
    End If
            
End Sub
