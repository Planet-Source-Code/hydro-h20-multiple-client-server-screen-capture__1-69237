VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmConfigure 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure Capture Server"
   ClientHeight    =   7365
   ClientLeft      =   2340
   ClientTop       =   2640
   ClientWidth     =   8550
   Icon            =   "frmConfigure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7365
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmConfigure.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdUpdate1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Machine"
      TabPicture(1)   =   "frmConfigure.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line1(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Line4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdUpdate2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cboMachine"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame6"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame8"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame10"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame13"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      Begin VB.Frame Frame13 
         Caption         =   "Frame Compression Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -69960
         TabIndex        =   49
         Top             =   3960
         Width           =   2775
         Begin VB.Frame Frame15 
            Caption         =   "Compression Rate"
            Height          =   855
            Left            =   120
            TabIndex        =   52
            Top             =   1080
            Width           =   2415
            Begin VB.HScrollBar scrollCompRate 
               Height          =   255
               Left            =   120
               Max             =   95
               Min             =   55
               TabIndex        =   53
               Top             =   240
               Value           =   55
               Width           =   1695
            End
            Begin VB.Label Label14 
               Caption         =   "95%"
               Height          =   255
               Left            =   1920
               TabIndex        =   55
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label13 
               Caption         =   "Low ..................... High"
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   495
               Width           =   1695
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Compression Mode"
            Height          =   855
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   2415
            Begin VB.ComboBox cboCompMode 
               Height          =   315
               Left            =   120
               TabIndex        =   51
               Top             =   360
               Width           =   2175
            End
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Schedule Start/Stop Capture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -69960
         TabIndex        =   39
         Top             =   1800
         Width           =   2775
         Begin VB.Frame Frame12 
            Caption         =   "Stop"
            Height          =   615
            Left            =   360
            TabIndex        =   45
            Top             =   1320
            Width           =   1575
            Begin VB.ComboBox cboSchedMin 
               Height          =   315
               Index           =   1
               Left            =   840
               TabIndex        =   47
               Text            =   "01"
               ToolTipText     =   "Minute"
               Top             =   240
               Width           =   615
            End
            Begin VB.ComboBox cboSchedHour 
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   46
               Text            =   "01"
               ToolTipText     =   "Hour"
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label11 
               Caption         =   ":"
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
               Left            =   720
               TabIndex        =   48
               Top             =   240
               Width           =   135
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Start"
            Height          =   615
            Left            =   360
            TabIndex        =   41
            Top             =   600
            Width           =   1575
            Begin VB.ComboBox cboSchedHour 
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   43
               Text            =   "01"
               ToolTipText     =   "Hour"
               Top             =   240
               Width           =   615
            End
            Begin VB.ComboBox cboSchedMin 
               Height          =   315
               Index           =   0
               Left            =   840
               TabIndex        =   42
               Text            =   "01"
               ToolTipText     =   "Minute"
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label12 
               Caption         =   ":"
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
               Left            =   720
               TabIndex        =   44
               Top             =   240
               Width           =   135
            End
         End
         Begin VB.CheckBox chkAutoStart 
            Caption         =   "Automatically Start Capture"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Capture Area Setting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -74520
         TabIndex        =   24
         Top             =   1800
         Width           =   4455
         Begin VB.CheckBox chkCapMouse 
            Caption         =   "Capture Mouse Movements"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   2760
            Width           =   3015
         End
         Begin VB.CheckBox chkDTStamp 
            Caption         =   "Date/Time Stamp Capture Frames"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   2400
            Width           =   3015
         End
         Begin VB.CheckBox chkConvert256 
            Caption         =   "Convert Captures To 256 Colors"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   2040
            Width           =   3015
         End
         Begin VB.Frame Frame9 
            Caption         =   "Scale Down To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   2400
            TabIndex        =   34
            Top             =   240
            Width           =   1935
            Begin VB.TextBox txtTargetWidth 
               Height          =   285
               Left            =   1200
               TabIndex        =   36
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtTargetHeight 
               Height          =   285
               Left            =   1200
               TabIndex        =   35
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label10 
               Caption         =   "Width"
               Height          =   255
               Left            =   120
               TabIndex        =   38
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label9 
               Caption         =   "Height"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Source Capture Area"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   2175
            Begin VB.TextBox txtLeft 
               Height          =   285
               Left            =   1440
               TabIndex        =   29
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtTop 
               Height          =   285
               Left            =   1440
               TabIndex        =   28
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtHeight 
               Height          =   285
               Left            =   1440
               TabIndex        =   27
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtWidth 
               Height          =   285
               Left            =   1440
               TabIndex        =   26
               Top             =   1320
               Width           =   615
            End
            Begin VB.Label Label5 
               Caption         =   "Left Start"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Top Start"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Height"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label8 
               Caption         =   "Width"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   1320
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Capture Path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -72000
         TabIndex        =   21
         Top             =   960
         Width           =   4815
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Left            =   4440
            TabIndex        =   23
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblCapPath 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Capture Machine Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74520
         TabIndex        =   19
         Top             =   960
         Width           =   2415
         Begin VB.TextBox txtMachine 
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Capture Retention"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Days To Keep Captures For"
         Top             =   3840
         Width           =   3615
         Begin VB.CommandButton Command2 
            Caption         =   "Remove Captures"
            Height          =   495
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtDaysToKeep 
            Height          =   285
            Left            =   1680
            TabIndex        =   16
            Text            =   "0"
            ToolTipText     =   "We Recommend 14 Days"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Setting to 0 will keep captures for ever (not recommended)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label Label3 
            Caption         =   "Days To Keep"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.ComboBox cboMachine 
         Height          =   315
         Left            =   -72480
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   480
         Width           =   3135
      End
      Begin VB.CommandButton cmdUpdate2 
         Caption         =   "Update"
         Height          =   615
         Left            =   -67920
         TabIndex        =   12
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate1 
         Caption         =   "Update"
         Height          =   615
         Left            =   7080
         TabIndex        =   11
         Top             =   6240
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "AVI Setting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   3615
         Begin VB.CommandButton cmdConvertToAVI 
            Caption         =   "Convert MCP Files To AVI"
            Height          =   375
            Left            =   240
            TabIndex        =   57
            ToolTipText     =   "Convert Captures To AVI Format"
            Top             =   1800
            Width           =   3015
         End
         Begin VB.CommandButton cmdSetAVICompressor 
            Caption         =   "Set AVI Compression Settings"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            ToolTipText     =   "Configure AVI Codec To Use"
            Top             =   1320
            Width           =   3015
         End
         Begin VB.Frame Frame3 
            Height          =   735
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   3015
            Begin VB.ComboBox cboMinute 
               Height          =   315
               Left            =   2160
               TabIndex        =   10
               Text            =   "01"
               ToolTipText     =   "Minute"
               Top             =   240
               Width           =   615
            End
            Begin VB.ComboBox cboHour 
               Height          =   315
               Left            =   1440
               TabIndex        =   8
               Text            =   "01"
               ToolTipText     =   "Hour"
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label2 
               Caption         =   ":"
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
               Left            =   2040
               TabIndex        =   9
               Top             =   240
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "Time To Convert:"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.CheckBox chkAVIConvert 
            Caption         =   "Auto Convert Captures To AVI"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            ToolTipText     =   "Enable To Convert Captures To AVI Automatically"
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Capture Slots"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "If Enabled, Every Capture Will Be Broken In X Minute Slots"
         Top             =   480
         Width           =   3615
         Begin VB.ComboBox cboSlots 
            Height          =   315
            Left            =   2160
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkEnableSlots 
            Caption         =   "Enable Capture Slots"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "If Enabled, Every Capture Will Be Broken In X Minute Slots"
            Top             =   240
            Value           =   1  'Checked
            Width           =   1815
         End
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   -66960
         X2              =   -66960
         Y1              =   600
         Y2              =   6120
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   -74760
         X2              =   -66960
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   -74760
         X2              =   -74760
         Y1              =   600
         Y2              =   6120
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   -69240
         X2              =   -66960
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   -74775
         X2              =   -72600
         Y1              =   600
         Y2              =   600
      End
   End
End
Attribute VB_Name = "frmConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Dim AviSettings As AVI_COMPRESS_OPTIONS

Private Sub cboCompMode_KeyDown(KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    KeyCode = 0
End Sub

Private Sub cboCompMode_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    KeyAscii = 0
End Sub

Private Sub cboHour_KeyDown(KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    KeyCode = 0
End Sub

Private Sub cboHour_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    KeyAscii = 0
End Sub

Private Sub cboMachine_Click()
    Dim I As Integer
    Dim A As Integer
    Dim compMode As String
    Dim aStart As Boolean
    
    On Local Error Resume Next
    
    I = cboMachine.ListIndex + 1
    
    txtMachine = ReadINI("Location" & I, "Name", INIFile)
    lblCapPath = ReadINI("Location" & I, "CaptureAppPath", INIFile)
    txtLeft = ReadINI("Location" & I, "startLeft", INIFile)
    txtTop = ReadINI("Location" & I, "startTop", INIFile)
    txtWidth = ReadINI("Location" & I, "srcWidth", INIFile)
    txtHeight = ReadINI("Location" & I, "srcHeight", INIFile)
    txtTargetWidth = ReadINI("Location" & I, "targetWidth", INIFile)
    txtTargetHeight = ReadINI("Location" & I, "targetHeight", INIFile)
    
    aStart = CBool(ReadINI("Location" & I, "AutoStart", INIFile))
    If aStart Then
        chkAutoStart.Value = 1
    Else
        chkAutoStart.Value = 0
    End If
    Call chkAutoStart_Click
    
    aStart = CBool(ReadINI("Location" & I, "Convert256", INIFile))
    If aStart Then
        chkConvert256.Value = 1
    Else
        chkConvert256.Value = 0
    End If
    
    aStart = CBool(ReadINI("Location" & I, "DTStamp", INIFile))
    If aStart Then
        chkDTStamp.Value = 1
    Else
        chkDTStamp.Value = 0
    End If
    
    aStart = CBool(ReadINI("Location" & I, "CaptureMouse", INIFile))
    If aStart Then
        chkCapMouse.Value = 1
    Else
        chkCapMouse.Value = 0
    End If

    A = Val(Left$(ReadINI("Location" & I, "StartAt", INIFile), 2))
    cboSchedHour(0).ListIndex = A
    A = Val(Right$(ReadINI("Location" & I, "StartAt", INIFile), 2))
    A = A \ 15
    cboSchedMin(0).ListIndex = A
    
    A = Val(Left$(ReadINI("Location" & I, "StopAt", INIFile), 2))
    cboSchedHour(1).ListIndex = A
    A = Val(Right$(ReadINI("Location" & I, "StopAt", INIFile), 2))
    A = A \ 15
    cboSchedMin(1).ListIndex = A
    
    compMode = ReadINI("Location" & I, "compressionMode", INIFile)
    For A = 0 To (cboCompMode.ListCount - 1)
        If Left$(cboCompMode.List(A), 3) = compMode Then Exit For
    Next A
    cboCompMode.ListIndex = A
    
    scrollCompRate.Value = ReadINI("Location" & I, "compressionRATE", INIFile)
    
End Sub

Private Sub cboMinute_KeyDown(KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    KeyCode = 0
End Sub

Private Sub cboMinute_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    KeyAscii = 0
End Sub

Private Sub cboSchedHour_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    KeyCode = 0
End Sub

Private Sub cboSchedHour_KeyPress(Index As Integer, KeyAscii As Integer)
    On Local Error Resume Next
    KeyAscii = 0
End Sub

Private Sub cboSchedMin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    KeyCode = 0
End Sub

Private Sub cboSchedMin_KeyPress(Index As Integer, KeyAscii As Integer)
    On Local Error Resume Next
    KeyAscii = 0
End Sub

Private Sub cboSlots_KeyDown(KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    KeyCode = 0
End Sub

Private Sub cboSlots_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    KeyAscii = 0
End Sub

Private Sub chkAutoStart_Click()
    On Local Error Resume Next
    
    If chkAutoStart.Value = 0 Then
        Frame11.Enabled = False
        Frame12.Enabled = False
        cboSchedHour(0).Enabled = False
        cboSchedHour(1).Enabled = False
        cboSchedMin(0).Enabled = False
        cboSchedMin(1).Enabled = False
    Else
        Frame11.Enabled = True
        Frame12.Enabled = True
        cboSchedHour(0).Enabled = True
        cboSchedHour(1).Enabled = True
        cboSchedMin(0).Enabled = True
        cboSchedMin(1).Enabled = True
    End If
End Sub

Private Sub chkAVIConvert_Click()
    On Local Error Resume Next
    
    If chkAVIConvert.Value = 0 Then
        Frame3.Enabled = False
        Label1.Enabled = False
        cboHour.Enabled = False
        Label2.Enabled = False
        cboMinute.Enabled = False
        'cmdSetAVICompressor.Enabled = False
        'cmdConvertToAVI.Enabled = False
    Else
        Frame3.Enabled = True
        Label1.Enabled = True
        cboHour.Enabled = True
        Label2.Enabled = True
        cboMinute.Enabled = True
        'cmdSetAVICompressor.Enabled = True
        'cmdConvertToAVI.Enabled = True
    End If
End Sub

Private Sub chkEnableSlots_Click()
    On Local Error Resume Next
    If chkEnableSlots.Value = 0 Then
        cboSlots.Enabled = False
    Else
        cboSlots.Enabled = True
    End If
End Sub

Private Sub cmdConvertToAVI_Click()
    Dim r As Boolean
    
    If ReadINI("General", "ComputerName", INIFile) <> CompName Then
        MsgBox "AVI Compression Settings Have Not Been Configured", vbOKOnly + vbCritical, "AVI Codec Not Set"
        Exit Sub
    End If
    
    Me.Enabled = False
    r = Convert_All_To_AVI
    Me.Enabled = True
    
End Sub

Private Sub cmdSetAVICompressor_Click()
    Dim res As Long
    Dim ps As Long
    Dim pOpts As Long
    Dim szOutputAVIFile As String
    Dim s$
    Dim BMP As cDIB
    Dim strhdr As AVI_STREAM_INFO
    Dim pfile As Long 'ptr PAVIFILE
    Dim AviSettings1 As AVI_COMPRESS_OPTIONS

    Call AVIFileInit
    
    
    
    Dim psCompressed As Long 'ptr PAVISTREAM
    Dim BI As BITMAPINFOHEADER
    Dim I As Long
    
    'get an avi filename from user
    szOutputAVIFile = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "config.avi"

'    Open the file for writing
    res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)

    'Get the first bmp in the list for setting format
    Set BMP = New cDIB
    s$ = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "blue.bmp"
    If BMP.CreateFromFile(s$) <> True Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
        GoTo error
    End If

'   Fill in the header for the video stream
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
        .fccHandler = 0&                             '// default AVI handler
        .dwScale = 1
        .dwRate = 1                       '// fps
        .dwSuggestedBufferSize = BMP.SizeImage       '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, BMP.Width, BMP.Height)       '// rectangle for stream
    End With
    

'   And create the stream
    res = AVIFileCreateStream(pfile, ps, strhdr)
    
    pOpts = VarPtr(AviSettings1)
    res = AVISaveOptions(Me.hWnd, ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, 1, ps, pOpts)
    If res = 1 Then
        AviSettings = AviSettings1
        
        WriteINI "General", "ComputerName", CompName, INIFile
        Call Write_AVI_Settings

    End If
    
error:
    Set BMP = Nothing
    If (ps <> 0) Then Call AVIStreamClose(ps)
    If (pfile <> 0) Then Call AVIFileClose(pfile)
    Call AVIFileExit
    
    Kill szOutputAVIFile
    
End Sub

Private Sub cmdUpdate1_Click()
    Dim I As Integer
    Dim convertTime As String
    
    On Local Error Resume Next
    
    txtDaysToKeep.Text = Val(txtDaysToKeep.Text)
    
    If chkEnableSlots.Value = 1 Then
        I = Val(cboSlots.Text)
        I = I * 60
        WriteINI "General", "CaptureSlotDuration", CStr(I), INIFile
        CaptureSlotDuration = I
    Else
        WriteINI "General", "CaptureSlotDuration", "0", INIFile
        CaptureSlotDuration = 0
    End If


    convertTime = Format$(cboHour.Text, "00") & ":" & Format$(cboMinute.Text, "00")
    If chkAVIConvert.Value = 1 Then
        WriteINI "General", "ConvertCapturesToAVI", "True", INIFile
    Else
        WriteINI "General", "ConvertCapturesToAVI", "False", INIFile
    End If
    Call Write_AVI_Settings
    
    WriteINI "General", "ConvertTime", convertTime, INIFile

    WriteINI "General", "KeepCapturesFor", txtDaysToKeep.Text, INIFile

    MsgBox "General Settings Updated", vbInformation + vbOKOnly, "Updated"
    
End Sub

Private Sub Write_AVI_Settings()
    On Local Error Resume Next
    
    WriteINI "AVISettings", "cbFormat", CStr(AviSettings.cbFormat), INIFile
    WriteINI "AVISettings", "cbParms", CStr(AviSettings.cbParms), INIFile
    WriteINI "AVISettings", "dwBytesPerSecond", CStr(AviSettings.dwBytesPerSecond), INIFile
    WriteINI "AVISettings", "dwFlags", CStr(AviSettings.dwFlags), INIFile
    WriteINI "AVISettings", "dwInterleaveEvery", CStr(AviSettings.dwInterleaveEvery), INIFile
    WriteINI "AVISettings", "dwKeyFrameEvery", CStr(AviSettings.dwKeyFrameEvery), INIFile
    WriteINI "AVISettings", "dwQuality", CStr(AviSettings.dwQuality), INIFile
    WriteINI "AVISettings", "cbFormat", CStr(AviSettings.fccHandler), INIFile
    WriteINI "AVISettings", "fccType", CStr(AviSettings.fccType), INIFile
    WriteINI "AVISettings", "lpFormat", CStr(AviSettings.lpFormat), INIFile
    WriteINI "AVISettings", "lpParms", CStr(AviSettings.lpParms), INIFile
        
End Sub

Private Function CaptureFolderExists(capFolder As String) As Boolean
    Dim l As String
    
    On Local Error Resume Next
    
    l = Dir$(App.Path & "\*.*", vbDirectory)
    While Len(l) <> 0 And CaptureFolderExists = False
        If l <> "." And l <> ".." Then
            If (GetAttr(App.Path & "\" & l) And vbDirectory) = vbDirectory Then
                If l = capFolder Then
                    CaptureFolderExists = True
                End If
            End If
        End If
        l = Dir$
    Wend


End Function

Private Sub cmdUpdate2_Click()
    Dim ddt As Long
    Dim I As Integer
    Dim aStart As String
    Dim l As String
    Dim oldMachine As String
    Dim vFRM As frmViewer
    Dim A As Integer
    Dim B As Integer
    Dim ret As Boolean
    
    On Local Error Resume Next
    
    txtMachine.Text = Trim$(txtMachine.Text)
    If Len(txtMachine.Text) = 0 Then
        MsgBox "Must Enter A Machine Name", vbOKOnly + vbCritical, "Error"
        txtMachine.SetFocus
        Exit Sub
    End If
    
    If txtMachine.Text = "Not In Use" Then
        MsgBox "Must Enter A Machine Name", vbOKOnly + vbCritical, "Error"
        txtMachine.SetFocus
        Exit Sub
    End If
    
    If lblCapPath.Caption = "Not In Use" Then
        MsgBox "You Must Select A Capture Path", vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
       
    If Val(txtTargetHeight.Text) > Val(txtHeight.Text) Then
        MsgBox "Target Height Must Be Less or Equal To The Source Height", vbOKOnly + vbCritical, "Error"
        txtTargetHeight.SetFocus
        Exit Sub
    End If
    
    If Val(txtTargetWidth.Text) > Val(txtWidth.Text) Then
        MsgBox "Target Width Must Be Less or Equal To The Source Width", vbOKOnly + vbCritical, "Error"
        txtTargetWidth.SetFocus
        Exit Sub
    End If
        
        
    oldMachine = cboMachine.Text
    
    If cboMachine.Text <> txtMachine.Text Then
        ' machine name change
        
        ' ask question if to keep records
        ' if keep then rename folder else
        ' delete folder
        ' rename records path
        
        ' before any name change must ensure
        ' that the capture name does not already exist
        If CaptureFolderExists(txtMachine.Text) Then
            MsgBox "Sorry, But You Cannot Use This As A Machine Name!!", vbOKOnly + vbExclamation, "Invalid Name"
            txtMachine.SetFocus
            Exit Sub
        End If
        
        If cboMachine.Text <> "Not In Use" Then
            ' new folder name does not exist
            ' ensure the old folder does exist
            If CaptureFolderExists(cboMachine.Text) = True Then
                ddt = MsgBox("Machine Name Has Changed!!!" & vbCrLf & vbCrLf & "Would You Like To Keep The Records From " & UCase$(cboMachine.Text) & vbCrLf & vbCrLf & "Click Yes To Keep Recordings." & vbCrLf & "Click No To Delete The Recordings." & vbCrLf & "Click Cancel To Abort Name Change.", vbYesNoCancel + vbQuestion, "Machine Name Changed")
                If ddt = vbCancel Then
                    txtMachine.Text = cboMachine.Text
                    Exit Sub
                End If
        
                If ddt = vbNo Then
                    ' delete the folder
                    myDelTree App.Path & "\" & txtMachine.Text
                End If
        
                If ddt = vbYes Then
                    ' rename the folder
                    Name App.Path & "\" & cboMachine.Text As App.Path & "\" & txtMachine.Text
                End If

            End If
        End If
    End If
    I = cboMachine.ListIndex + 1
    cboMachine.List(I - 1) = txtMachine.Text
    WriteINI "Location" & I, "Name", txtMachine.Text, INIFile
    WriteINI "Location" & I, "CaptureAppPath", lblCapPath.Caption, INIFile
    WriteINI "Location" & I, "CapturePath", lblCapPath.Caption & "\Captures", INIFile
    WriteINI "Location" & I, "CaptureStart", lblCapPath.Caption & "\StartCapture.dat", INIFile
    WriteINI "Location" & I, "CaptureStop", lblCapPath.Caption & "\StopCapture.dat", INIFile
    WriteINI "Location" & I, "startLeft", txtLeft.Text, INIFile
    WriteINI "Location" & I, "startTop", txtTop.Text, INIFile
    WriteINI "Location" & I, "srcWidth", txtWidth.Text, INIFile
    WriteINI "Location" & I, "srcHeight", txtHeight.Text, INIFile
    WriteINI "Location" & I, "targetWidth", txtTargetWidth.Text, INIFile
    WriteINI "Location" & I, "targetHeight", txtTargetHeight.Text, INIFile
    If chkConvert256.Value = 1 Then
        aStart = "True"
    Else
        aStart = "False"
    End If
    WriteINI "Location" & I, "Convert256", aStart, INIFile
    
    If chkDTStamp.Value = 1 Then
        aStart = "True"
    Else
        aStart = "False"
    End If
    WriteINI "Location" & I, "DTStamp", aStart, INIFile
    
    If chkCapMouse.Value = 1 Then
        aStart = "True"
    Else
        aStart = "False"
    End If
    WriteINI "Location" & I, "CaptureMouse", aStart, INIFile

    If chkAutoStart.Value = 1 Then
        aStart = "True"
    Else
        aStart = "False"
    End If
    WriteINI "Location" & I, "AutoStart", aStart, INIFile
    
    aStart = Format$(cboSchedHour(0).Text, "00") & ":" & Format$(cboSchedMin(0).Text, "00")
    WriteINI "Location" & I, "StartAt", aStart, INIFile
    
    aStart = Format$(cboSchedHour(1).Text, "00") & ":" & Format$(cboSchedMin(1).Text, "00")
    WriteINI "Location" & I, "StopAt", aStart, INIFile
    
    aStart = Left$(cboCompMode.Text, 3)
    WriteINI "Location" & I, "compressionMode", aStart, INIFile
    
    aStart = scrollCompRate.Value
    WriteINI "Location" & I, "compressionRATE", aStart, INIFile
    
    A = 0
    For I = 0 To (Forms.Count - 1)
        If Forms(I).Tag = "~VIEWER~" Then
            If InStr(1, Forms(I).Label1, oldMachine, vbTextCompare) > 0 Then
                ' found view form
                Set vFRM = Forms(I)
                A = Left$(Forms(I).Label1, 2)
                Exit For
            End If
        End If
    Next I

    If A <> 0 Then
        ' found viewer form
        ret = Setup_Capture_Form(A, txtMachine.Text, vFRM)
        Set vFRM = Nothing
        
        If oldMachine = "Not In Use" Then
            If ret = False Then
                mdiMAIN.cboLocation.AddItem txtMachine.Text & " (OFF)"
            Else
                mdiMAIN.cboLocation.AddItem txtMachine.Text
            End If
        Else
            B = -1
            For I = 0 To (mdiMAIN.cboLocation.ListCount - 1)
                If InStr(1, mdiMAIN.cboLocation.List(I), oldMachine, vbTextCompare) > 0 Then
                    B = I
                    Exit For
                End If
            Next I
             
            If B >= 0 Then
                If ret = False Then
                    mdiMAIN.cboLocation.List(B) = txtMachine.Text & " (OFF)"
                Else
                    mdiMAIN.cboLocation.List(B) = txtMachine.Text
                End If
            End If
        End If
     End If
    
    MsgBox "Machine Details Saved", vbInformation + vbOKOnly, "Updated"
    
End Sub

Private Sub Command1_Click()
    On Local Error Resume Next
    
    'Load frmSelectPath
    
    frmSelectPath.Caption = "Select Capture Path For " & txtMachine
    frmSelectPath.lblOldPath.Caption = lblCapPath.Caption
    If Right$(frmSelectPath.lblOldPath.Caption, 1) <> "\" Then
        frmSelectPath.lblOldPath.Caption = frmSelectPath.lblOldPath.Caption & "\"
    End If
    frmSelectPath.ucNetworkTree1.Collapse_Tree ' collapse tree, makes it look neater for next path
    
    
    Me.Enabled = False
    
    frmSelectPath.Show 'vbModal
    DoEvents
    
    frmSelectPath.ucNetworkTree1.SelectedFolder = frmSelectPath.lblOldPath.Caption
    
    frmSelectPath.Timer1.Enabled = True
    DoEvents
  
    While IsFormVisible("frmSelectPath")
        DoEvents
    Wend
    
    
    If Len(frmSelectPath.lblNewPath) <> 0 Then
        lblCapPath.Caption = frmSelectPath.lblNewPath
    End If
    
    Me.Enabled = True
    txtMachine.SetFocus
    
End Sub

Private Sub Command2_Click()
    On Local Error Resume Next
    
    Call Remove_Old_Captures
End Sub

Private Sub Form_Load()
    Dim I As Integer
    Dim CaptureSlotDuration As Integer
    Dim ConvertCapturesToAVI As Boolean
    
    On Local Error Resume Next
           
    Label4.Caption = "Setting to 0 will keep captures for ever (not recommended)" & vbCrLf & "The captures will be removed overnight."
    
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    
    AviSettings.cbFormat = ReadINI("AVISettings", "cbFormat", INIFile)
    AviSettings.cbParms = ReadINI("AVISettings", "cbParms", INIFile)
    AviSettings.dwBytesPerSecond = ReadINI("AVISettings", "dwBytesPerSecond", INIFile)
    AviSettings.dwFlags = ReadINI("AVISettings", "dwFlags", INIFile)
    AviSettings.dwInterleaveEvery = ReadINI("AVISettings", "dwInterleaveEvery", INIFile)
    AviSettings.dwKeyFrameEvery = ReadINI("AVISettings", "dwKeyFrameEvery", INIFile)
    AviSettings.dwQuality = ReadINI("AVISettings", "dwQuality", INIFile)
    AviSettings.fccHandler = ReadINI("AVISettings", "cbFormat", INIFile)
    AviSettings.fccType = ReadINI("AVISettings", "fccType", INIFile)
    AviSettings.lpFormat = ReadINI("AVISettings", "lpFormat", INIFile)
    AviSettings.lpParms = ReadINI("AVISettings", "lpParms", INIFile)
    
    For I = 5 To 60 Step 5
        cboSlots.AddItem I & " mins"
    Next I
    cboSlots.ListIndex = 0
    
    CaptureSlotDuration = ReadINI("General", "CaptureSlotDuration", INIFile)
    If CaptureSlotDuration = 0 Then
        chkEnableSlots.Value = 0
    Else
        I = CaptureSlotDuration \ 60 ' convert to mins
        I = (I \ 5) - 1 ' convert to list index
        cboSlots.ListIndex = I
    End If
    
    For I = 0 To 23
        cboHour.AddItem Format$(I, "00")
        cboSchedHour(0).AddItem Format$(I, "00")
        cboSchedHour(1).AddItem Format$(I, "00")
    Next I
    cboHour.ListIndex = 0
    cboSchedHour(0).ListIndex = 0
    cboSchedHour(1).ListIndex = 0
    
    For I = 0 To 45 Step 15
        cboMinute.AddItem Format$(I, "00")
        cboSchedMin(0).AddItem Format$(I, "00")
        cboSchedMin(1).AddItem Format$(I, "00")
    Next I
    cboMinute.ListIndex = 0
    cboSchedMin(0).ListIndex = 0
    cboSchedMin(1).ListIndex = 0
    
    ConvertCapturesToAVI = CBool(ReadINI("General", "ConvertCapturesToAVI", INIFile))
    If ConvertCapturesToAVI Then
        chkAVIConvert.Value = 1
        I = Val(Left$(ReadINI("General", "ConvertTime", INIFile), 2))
        cboHour.ListIndex = I
        
        I = Val(Right$(ReadINI("General", "ConvertTime", INIFile), 2))
        I = I \ 15
        cboMinute.ListIndex = I
    Else
        chkAVIConvert.Value = 0
        Frame3.Enabled = False
    End If
    
    cboCompMode.AddItem "1:1 [High Quality]"
    cboCompMode.AddItem "1:2"
    cboCompMode.AddItem "2:1"
    cboCompMode.AddItem "2:2 [Normal]"
    cboCompMode.AddItem "1:3"
    cboCompMode.AddItem "3:1"
    cboCompMode.AddItem "2:3"
    cboCompMode.ListIndex = 0
    
    txtDaysToKeep = ReadINI("General", "KeepCapturesFor", INIFile)
    
    
    
    For I = 1 To 16
        cboMachine.AddItem ReadINI("Location" & I, "Name", INIFile)
    Next I
    cboMachine.ListIndex = 0
    
    
    
End Sub

Private Sub scrollCompRate_Change()
    On Local Error Resume Next
    Label14 = scrollCompRate.Value & "%"
End Sub

Private Sub txtDaysToKeep_Change()
    On Local Error Resume Next
    
    If Val(txtDaysToKeep.Text) = 0 Then
        Command2.Enabled = False
    Else
        Command2.Enabled = True
    End If
End Sub

Private Sub txtDaysToKeep_GotFocus()
    On Local Error Resume Next
    txtDaysToKeep.BackColor = vbYellow
End Sub

Private Sub txtDaysToKeep_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub txtDaysToKeep_LostFocus()
    On Local Error Resume Next
    txtDaysToKeep.BackColor = vbWhite
End Sub

Private Sub txtHeight_GotFocus()
    On Local Error Resume Next
    txtHeight.BackColor = vbYellow
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub txtHeight_LostFocus()
    On Local Error Resume Next
    txtHeight.BackColor = vbWhite
End Sub

Private Sub txtLeft_GotFocus()
    On Local Error Resume Next
    txtLeft.BackColor = vbYellow
End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub txtLeft_LostFocus()
    On Local Error Resume Next
    txtLeft.BackColor = vbWhite
End Sub

Private Sub txtMachine_GotFocus()
    On Local Error Resume Next
    txtMachine.BackColor = vbYellow
End Sub

Private Sub txtMachine_LostFocus()
    On Local Error Resume Next
    txtMachine.BackColor = vbWhite
End Sub

Private Sub txtTargetHeight_GotFocus()
    On Local Error Resume Next
    txtTargetHeight.BackColor = vbYellow
End Sub

Private Sub txtTargetHeight_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub txtTargetHeight_LostFocus()
    On Local Error Resume Next
    txtTargetHeight.BackColor = vbWhite
End Sub

Private Sub txtTargetWidth_GotFocus()
    On Local Error Resume Next
    txtTargetWidth.BackColor = vbYellow
End Sub

Private Sub txtTargetWidth_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub txtTargetWidth_LostFocus()
    On Local Error Resume Next
    txtTargetWidth.BackColor = vbWhite
End Sub

Private Sub txtTop_GotFocus()
    On Local Error Resume Next
    txtTop.BackColor = vbYellow
End Sub

Private Sub txtTop_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub txtTop_LostFocus()
    On Local Error Resume Next
    txtTop.BackColor = vbWhite
End Sub

Private Sub txtWidth_GotFocus()
    On Local Error Resume Next
    txtWidth.BackColor = vbYellow
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub txtWidth_LostFocus()
    On Local Error Resume Next
    txtWidth.BackColor = vbWhite
End Sub
