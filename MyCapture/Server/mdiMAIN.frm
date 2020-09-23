VERSION 5.00
Begin VB.MDIForm mdiMAIN 
   BackColor       =   &H8000000C&
   Caption         =   "Capture Suite Server"
   ClientHeight    =   9975
   ClientLeft      =   2010
   ClientTop       =   2355
   ClientWidth     =   12570
   Icon            =   "mdiMAIN.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   1080
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   9975
      Left            =   9630
      ScaleHeight     =   9945
      ScaleWidth      =   2910
      TabIndex        =   0
      Top             =   0
      Width           =   2940
      Begin Server.ucGIF cmdConfigure 
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   4200
         Width           =   2655
         _ExtentX        =   1429
         _ExtentY        =   1005
         autoSize        =   0   'False
         Caption         =   "Configure Capture Server"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         gifLoopInfinity =   -1  'True
         ShowFocus       =   0   'False
         gifPosition     =   1
         fileLen         =   6160
         fileData        =   "mdiMAIN.frx":0E42
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recording"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   2655
         Begin VB.CommandButton cmdStopRec 
            Caption         =   "Stop Recording"
            Enabled         =   0   'False
            Height          =   495
            Left            =   1440
            TabIndex        =   6
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton cmdStartRec 
            Caption         =   "Start Recording"
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Machine"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   2415
            Begin VB.ComboBox cboLocation 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   4
               Text            =   "Combo1"
               Top             =   240
               Width           =   2175
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Views"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2655
         Begin VB.Image imgSINGLE 
            Enabled         =   0   'False
            Height          =   480
            Left            =   240
            Picture         =   "mdiMAIN.frx":2673
            Top             =   360
            Width           =   480
         End
         Begin VB.Image imgFour 
            Enabled         =   0   'False
            Height          =   480
            Left            =   1080
            Picture         =   "mdiMAIN.frx":33BD
            Top             =   360
            Width           =   480
         End
         Begin VB.Image imgAll 
            Enabled         =   0   'False
            Height          =   480
            Left            =   1920
            Picture         =   "mdiMAIN.frx":4107
            Top             =   360
            Width           =   480
         End
      End
      Begin Server.ucGIF cmdPlayBack 
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   2655
         _ExtentX        =   1323
         _ExtentY        =   1323
         autoSize        =   0   'False
         Caption         =   "Play Back Facilty"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         gifLoopInfinity =   -1  'True
         ShowFocus       =   0   'False
         gifPosition     =   1
         fileLen         =   3534
         fileData        =   "mdiMAIN.frx":4E51
      End
   End
End
Attribute VB_Name = "mdiMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastTestTime As String
Dim LastTestDate As String

Private Sub cboLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    KeyDown = 0
End Sub

Private Sub cboLocation_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    KeyAscii = 0
End Sub

Private Sub cmdConfigure_Click()
    On Local Error Resume Next
    
    If IsFormVisible("frmConfigure") Then Exit Sub
    
    If IsFormVisible("frmPlayback") Then
        MsgBox "Sorry, But You Cannot Use Configure Your System While The Playback Facility Is Running.", vbOKOnly + vbExclamation, ""
        Exit Sub
    End If
        
    If ActiveRecordings > 0 Then
        MsgBox "You Must Stop All Recordings, Then You You Can Configure Your System.", vbOKOnly + vbExclamation, "System Is Recording"
        Exit Sub
    End If
    
    frmConfigure.Show , Me
End Sub

Private Sub cmdPlayBack_Click()
    On Local Error Resume Next
    
    If IsFormVisible("frmSelectPath") Then Exit Sub
    
    If IsFormVisible("frmConfigure") Then
        MsgBox "Sorry, But You Cannot Use Play Back Facility While You Are Configuring The System.", vbOKOnly + vbExclamation, ""
        Exit Sub
    End If
    
    frmPlayback.Show , Me
End Sub

Private Sub cmdStartRec_Click()
    On Local Error Resume Next
    
    cmdStartRec.Enabled = False
    Call Start_Capture(cboLocation.Text)
    cmdStartRec.Enabled = True
    
End Sub

Private Sub cmdStopRec_Click()
    On Local Error Resume Next
    
    cmdStopRec.Enabled = False
    Call Stop_Capture(cboLocation)
    cmdStopRec.Enabled = True
    
End Sub



Private Sub imgAll_Click()
    On Local Error Resume Next
    
    If ViewMode = 3 Then Exit Sub
    
    ViewMode = 3
    ViewerID = 1
    Call Arrange_Capture_Forms(ViewMode, ViewerID)

End Sub

Private Sub imgFour_Click()
    On Local Error Resume Next
    
    If ViewMode = 2 Then
        ViewerID = ViewerID + 4
        If ViewerID > 16 Then ViewerID = 1
    Else
        ViewMode = 2
        ViewerID = 1
    End If
    
    Call Arrange_Capture_Forms(ViewMode, ViewerID)


End Sub

Private Sub imgSINGLE_Click()
    On Local Error Resume Next
    
    If ViewMode = 1 Then
        ViewerID = ViewerID + 1
        If ViewerID > 16 Then ViewerID = 1
    Else
        ViewMode = 1
    End If
    
    Call Arrange_Capture_Forms(ViewMode, ViewerID)

End Sub

Private Sub MDIForm_Load()
    On Local Error Resume Next
    
    LastTestTime = Format$(Time$, "HH:MM")
    LastTestDate = Format$(Date, "dd/mmm/yyyy")
    
    
End Sub

Private Sub MDIForm_Resize()
    On Local Error Resume Next
    Call Arrange_Capture_Forms(ViewMode, ViewerID)
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    On Local Error Resume Next
    

    If ActiveRecordings > 0 Then
        MsgBox "You Must Stop All Recordings, Then You You Can Configure Your System.", vbOKOnly + vbExclamation, "System Is Recording"
        Cancel = -1
        Exit Sub
    End If
    
    If IsFormVisible("frmFolderDelete") Then
        Cancel = -1
        Exit Sub
    End If
    
    If Len(Dir$(App.Path & "\header.bmp")) <> 0 Then
        Kill App.Path & "\header.bmp"
    End If
    
    If Len(Dir$(App.Path & "\frame.jpg")) <> 0 Then
        Kill App.Path & "\frame.jpg"
    End If
    
    If Len(Dir$(App.Path & "\frame.bmp")) <> 0 Then
        Kill App.Path & "\frame.bmp"
    End If
    
    cmdPlayBack.gifFileName = ""
    cmdConfigure.gifFileName = ""
    
    Call TidyUp_LaVolpe_Gifs
    
    Unload frmSelectPath
    
    End
End Sub

Private Sub Timer1_Timer()
    Dim cTIME As String
    Dim testTime1 As Date
    Dim testTime2 As Date
    
    On Local Error Resume Next
    
    
    If ActiveRecordings = 0 Then
        Frame2.Caption = "Recording"
    Else
        Frame2.Caption = "Recording (" & ActiveRecordings & ")"
    End If
    
    
    ' ctime and lasttesttime to check only every minute
    
    cTIME = Format$(Time$, "HH:MM")
    If cTIME <> LastTestTime Then
        LastTestTime = cTIME
        
        If CBool(ReadINI("General", "ConvertCapturesToAVI", INIFile)) Then
            If ReadINI("General", "ComputerName", INIFile) = CompName Then
                If ReadINI("General", "LastAVIConverted", INIFile) <> Format$(Date, "dd/mmm/yyyy") Then
                    ' got to here
                    ' auto convert to avi active
                    ' AVI codec set
                    ' Has not been run today
                    ' see if time to run
                    
                    testTime1 = CDate(Format$(Date, "dd/mmm/yyyy") & " " & ReadINI("General", "ConvertTime", INIFile) & ":00")
                    testTime2 = DateAdd("n", 10, testTime1)

                    If Now() >= testTime1 And Now() <= testTime2 Then
                        WriteINI "General", "LastAVIConverted", Format$(Date, "dd/mmm/yyyy"), INIFile
                        Timer1.Enabled = False
                        
                        Call Convert_All_To_AVI
                        
                        LastTestTime = cTIME
                    End If
                End If
            End If
        End If
    End If
    
        
        
     
    If LastTestDate <> Format$(Date, "dd/mmm/yyyy") Then
        LastTestDate = Format$(Date, "dd/mmm/yyyy")
        
        If Val(ReadINI("General", "KeepCapturesFor", INIFile)) <> 0 Then
            ' run the function that cleans up the capture folders
        
            Call Remove_Old_Captures
        End If
    End If
    
        
        
    Timer1.Enabled = True
        
End Sub


