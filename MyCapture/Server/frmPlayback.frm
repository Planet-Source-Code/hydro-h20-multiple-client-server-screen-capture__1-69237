VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmPlayback 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Playback"
   ClientHeight    =   7200
   ClientLeft      =   3420
   ClientTop       =   2385
   ClientWidth     =   9990
   Icon            =   "frmPlayback.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7200
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9000
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   66
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdEXPORT 
      Caption         =   "Export To AVI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   6480
      Width           =   3375
   End
   Begin PicClip.PictureClip picEnabled 
      Left            =   4080
      Top             =   8400
      _ExtentX        =   6112
      _ExtentY        =   979
      _Version        =   393216
      Cols            =   7
      Picture         =   "frmPlayback.frx":0D4A
   End
   Begin PicClip.PictureClip picDisabled 
      Left            =   4080
      Top             =   7800
      _ExtentX        =   6112
      _ExtentY        =   979
      _Version        =   393216
      Cols            =   7
      Picture         =   "frmPlayback.frx":7234
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   10560
      TabIndex        =   59
      Top             =   2280
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   10440
      TabIndex        =   58
      Top             =   360
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   120
      TabIndex        =   54
      Top             =   3975
      Width           =   3375
   End
   Begin VB.CommandButton cmdMonthNext 
      BackColor       =   &H00FFFFC0&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   970
      Width           =   375
   End
   Begin VB.CommandButton cmdMonthPrev 
      BackColor       =   &H00FFFFC0&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   970
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Machine"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   5760
      TabIndex        =   61
      Top             =   6360
      Width           =   2055
      Begin VB.Label Label3 
         Height          =   255
         Left            =   1320
         TabIndex        =   63
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   1095
      End
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   255
      Left            =   3600
      TabIndex        =   60
      Top             =   5280
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   327682
   End
   Begin VB.Image imgEXPAND 
      Height          =   495
      Left            =   3000
      Picture         =   "frmPlayback.frx":D71E
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Maximise Picture"
      Top             =   240
      Width           =   495
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   855
      Left            =   3720
      TabIndex        =   65
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1720
      _cy             =   1508
   End
   Begin VB.Image imgArrow 
      Height          =   480
      Index           =   1
      Left            =   9120
      Picture         =   "frmPlayback.frx":DB60
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgArrow 
      Height          =   480
      Index           =   0
      Left            =   8400
      Picture         =   "frmPlayback.frx":DFA2
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBLUE 
      Height          =   420
      Left            =   7800
      Picture         =   "frmPlayback.frx":E3E4
      Top             =   8160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgPlayButtons 
      Height          =   495
      Index           =   6
      Left            =   7080
      Stretch         =   -1  'True
      ToolTipText     =   "Stop"
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image imgPlayButtons 
      Height          =   495
      Index           =   5
      Left            =   7680
      Stretch         =   -1  'True
      ToolTipText     =   "Speed Up"
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image imgPlayButtons 
      Height          =   495
      Index           =   4
      Left            =   5280
      Stretch         =   -1  'True
      ToolTipText     =   "Slow Down"
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image imgPlayButtons 
      Height          =   495
      Index           =   3
      Left            =   6480
      Stretch         =   -1  'True
      ToolTipText     =   "Pause"
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image imgPlayButtons 
      Height          =   495
      Index           =   2
      Left            =   5880
      Stretch         =   -1  'True
      ToolTipText     =   "Play"
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image imgPlayButtons 
      Height          =   495
      Index           =   1
      Left            =   8280
      Stretch         =   -1  'True
      ToolTipText     =   "Move To End"
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image imgPlayButtons 
      Height          =   495
      Index           =   0
      Left            =   4680
      Stretch         =   -1  'True
      ToolTipText     =   "Move To Start"
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Duration"
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
      Index           =   2
      Left            =   2400
      TabIndex        =   57
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Stopped"
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
      Index           =   1
      Left            =   1200
      TabIndex        =   56
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Started"
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
      Index           =   0
      Left            =   120
      TabIndex        =   55
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4935
      Left            =   3600
      Picture         =   "frmPlayback.frx":E6DD
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   41
      Left            =   3000
      TabIndex        =   53
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   40
      Left            =   2520
      TabIndex        =   52
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   39
      Left            =   2040
      TabIndex        =   51
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   38
      Left            =   1560
      TabIndex        =   50
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   37
      Left            =   1080
      TabIndex        =   49
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   36
      Left            =   600
      TabIndex        =   48
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   35
      Left            =   120
      TabIndex        =   47
      Top             =   3870
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   34
      Left            =   3000
      TabIndex        =   46
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   33
      Left            =   2520
      TabIndex        =   45
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   32
      Left            =   2040
      TabIndex        =   44
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   31
      Left            =   1560
      TabIndex        =   43
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   30
      Left            =   1080
      TabIndex        =   42
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   29
      Left            =   600
      TabIndex        =   41
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   28
      Left            =   120
      TabIndex        =   40
      Top             =   3390
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   27
      Left            =   3000
      TabIndex        =   39
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   26
      Left            =   2520
      TabIndex        =   38
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   25
      Left            =   2040
      TabIndex        =   37
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   24
      Left            =   1560
      TabIndex        =   36
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   23
      Left            =   1080
      TabIndex        =   35
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   22
      Left            =   600
      TabIndex        =   34
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   21
      Left            =   120
      TabIndex        =   33
      Top             =   2910
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   20
      Left            =   3000
      TabIndex        =   32
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   19
      Left            =   2520
      TabIndex        =   31
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   18
      Left            =   2040
      TabIndex        =   30
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   17
      Left            =   1560
      TabIndex        =   29
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   16
      Left            =   1080
      TabIndex        =   28
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   15
      Left            =   600
      TabIndex        =   27
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   14
      Left            =   120
      TabIndex        =   26
      Top             =   2430
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   13
      Left            =   3000
      TabIndex        =   25
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   12
      Left            =   2520
      TabIndex        =   24
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   11
      Left            =   2040
      TabIndex        =   23
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   10
      Left            =   1560
      TabIndex        =   22
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   9
      Left            =   1080
      TabIndex        =   21
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   8
      Left            =   600
      TabIndex        =   20
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   1950
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   6
      Left            =   3000
      TabIndex        =   18
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   5
      Left            =   2520
      TabIndex        =   17
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   4
      Left            =   2040
      TabIndex        =   16
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   3
      Left            =   1560
      TabIndex        =   15
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   2
      Left            =   1080
      TabIndex        =   14
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   600
      TabIndex        =   13
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1590
      Width           =   495
   End
   Begin VB.Label lblDayHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   6
      Left            =   2990
      TabIndex        =   11
      Top             =   1350
      Width           =   495
   End
   Begin VB.Label lblDayHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   2510
      TabIndex        =   10
      Top             =   1350
      Width           =   495
   End
   Begin VB.Label lblDayHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   2030
      TabIndex        =   9
      Top             =   1350
      Width           =   495
   End
   Begin VB.Label lblDayHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   1550
      TabIndex        =   8
      Top             =   1350
      Width           =   495
   End
   Begin VB.Label lblDayHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   1070
      TabIndex        =   7
      Top             =   1350
      Width           =   495
   End
   Begin VB.Label lblDayHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   1350
      Width           =   495
   End
   Begin VB.Label lblDayHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1350
      Width           =   495
   End
   Begin VB.Label lblMonthYear 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "September 2007"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1050
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   3855
      Left            =   120
      Top             =   960
      Width           =   3365
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF8080&
      Height          =   2620
      Left            =   105
      Top             =   3705
      Width           =   3405
   End
End
Attribute VB_Name = "frmPlayback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AVIExporting As Boolean

Dim streamRate As Double
Dim streamFN As Integer
Dim streamPLAYING As Boolean
Dim fPOS As Long
Dim AllowSliderChange As Boolean

Dim selectedDATE As Date
Dim dispMONTH As Integer
Dim dispYEAR As Integer
Dim prevSELECTED As Integer

Private Sub cmdEXPORT_Click()
    Dim aviFILE As String
    Dim fNAME As String
    Dim cssHEADER As cssHEADERtype
    Dim mcpFILE As String
    
    On Local Error Resume Next
    
    
    If ReadINI("General", "ComputerName", INIFile) <> CompName Then
        MsgBox "AVI Compression Settings Have Not Been Configured", vbOKOnly + vbCritical, "AVI Codec Not Set"
        Exit Sub
    End If
    
    AVIExporting = True
    Slider1.Value = 0
    
    
    
'    fNAME = GetFileName(List2.List(List1.ListIndex))
'    fNAME = Format$(selectedDATE, "yyyymmdd") & "~" & Combo1.Text & "~" & left$(fNAME, Len(fNAME) - 4) & ".avi"
'    aviFILE = App.Path & "\AVIExports\" & fNAME

    mcpFILE = List2.List(List1.ListIndex)
    cssHEADER = Read_Capture_Header(mcpFILE)
    aviFILE = JustPath(mcpFILE) & "Capture~" & Format$(cssHEADER.Started, "HHMMSS") & "~" & Format$(cssHEADER.Stopped, "HHMMSS") & ".avi"

    Create_Header_Frame Me, Combo1.Text, Format$(selectedDATE, "ddd dd/mmm/yyyy"), Format$(cssHEADER.Started, "HH:MM:SS"), Format$(cssHEADER.Stopped, "HH:MM:SS")

    Call AVIFileInit
    Call WriteAVI(aviFILE, 1)
    Call AVIFileExit
    AVIExporting = False
 
    Slider1.Value = 0

    If streamPLAYING Then
        Call imgPlayButtons_Click(6)
        streamPLAYING = False
    End If
    
    If streamFN <> 0 Then Close #streamFN
    Image1.Picture = imgBLUE.Picture
    streamFN = 0

    Kill mcpFILE
    
    
    Call Load_Clips
    List1.ListIndex = 0
    Call List1_Click
    
    MsgBox "Conversion Completed", vbOKOnly + vbInformation, "Completed"
    
End Sub

Private Sub WriteAVI(ByVal fileName As String, Optional ByVal FrameRate As Integer = 1)
    Dim s$
    Dim InitDir As String
    Dim szOutputAVIFile As String
    Dim res As Long
    Dim pfile As Long 'ptr PAVIFILE
    Dim BMP As cDIB
    Dim ps As Long 'ptr PAVISTREAM
    Dim psCompressed As Long 'ptr PAVISTREAM
    Dim strhdr As AVI_STREAM_INFO
    Dim BI As BITMAPINFOHEADER
    Dim opts As AVI_COMPRESS_OPTIONS
    Dim pOpts As Long
    Dim i As Long
    
    'get an avi filename from user
    szOutputAVIFile = fileName$
'    Open the file for writing
    res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo error

    'Get the first bmp in the list for setting format
    Set BMP = New cDIB
    s$ = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "frame.bmp"
    If BMP.CreateFromFile(s$) <> True Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
        GoTo error
    End If

'   Fill in the header for the video stream
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
        .fccHandler = 0&                             '// default AVI handler
        .dwScale = 1
        .dwRate = Val(FrameRate%)                        '// fps
        .dwSuggestedBufferSize = BMP.SizeImage       '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, BMP.Width, BMP.Height)       '// rectangle for stream
    End With
    
    'validate user input
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30

'   And create the stream
    res = AVIFileCreateStream(pfile, ps, strhdr)
    If (res <> AVIERR_OK) Then GoTo error


'    'get the compression options from the user
'    'Careful! this API requires a pointer to a pointer to a UDT
    pOpts = VarPtr(opts)
'    res = AVISaveOptions(lHwnd, ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, 1, ps, pOpts)
'    'returns TRUE if User presses OK, FALSE if Cancel, or error code
'    If res <> 1 Then 'In C TRUE = 1
'      Call AVISaveOptionsFree(1, pOpts)
'      GoTo error
'    End If
    
    
    
    opts.cbFormat = ReadINI("AVISettings", "cbFormat", INIFile)
    opts.cbParms = ReadINI("AVISettings", "cbParms", INIFile)
    opts.dwBytesPerSecond = ReadINI("AVISettings", "dwBytesPerSecond", INIFile)
    opts.dwFlags = ReadINI("AVISettings", "dwFlags", INIFile)
    opts.dwInterleaveEvery = ReadINI("AVISettings", "dwInterleaveEvery", INIFile)
    opts.dwKeyFrameEvery = ReadINI("AVISettings", "dwKeyFrameEvery", INIFile)
    opts.dwQuality = ReadINI("AVISettings", "dwQuality", INIFile)
    opts.fccHandler = ReadINI("AVISettings", "cbFormat", INIFile)
    opts.fccType = ReadINI("AVISettings", "fccType", INIFile)
    opts.lpFormat = ReadINI("AVISettings", "lpFormat", INIFile)
    opts.lpParms = ReadINI("AVISettings", "lpParms", INIFile)
    
    
    'make compressed stream
    res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If res <> AVIERR_OK Then GoTo error
    
    'set format of stream according to the bitmap
    With BI
        .biBitCount = BMP.BitCount
        .biClrImportant = BMP.ClrImportant
        .biClrUsed = BMP.ClrUsed
        .biCompression = BMP.Compression
        .biHeight = BMP.Height
        .biWidth = BMP.Width
        .biPlanes = BMP.Planes
        .biSize = BMP.SizeInfoHeader
        .biSizeImage = BMP.SizeImage
        .biXPelsPerMeter = BMP.XPPM
        .biYPelsPerMeter = BMP.YPPM
    End With
    
    'set the format of the compressed stream
    res = AVIStreamSetFormat(psCompressed, 0, ByVal BMP.PointerToBitmapInfo, BMP.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo error

    ' write 5 secs of header
    BMP.CreateFromFile App.Path & "\header.bmp"
    For i = 1 To 5
         res = AVIStreamWrite(psCompressed, i - 1, 1, BMP.PointerToBits, BMP.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
         If res <> AVIERR_OK Then GoTo error
    Next i


    AllowSliderChange = False
    For i = 1 To Slider1.Max
        Slider1.Value = i
        Call Display_Frame(False)
        DoEvents
        
        BMP.CreateFromFile (s$) 'load the bitmap (ignore errors)
        res = AVIStreamWrite(psCompressed, (i + 5) - 1, 1, BMP.PointerToBits, BMP.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
        If res <> AVIERR_OK Then GoTo error
    Next i
    AllowSliderChange = True
    
error:
'   Now close the file
    Set BMP = Nothing
    
    If (ps <> 0) Then Call AVIStreamClose(ps)

    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)

    If (pfile <> 0) Then Call AVIFileClose(pfile)

    Call AVIFileExit

    If (res <> AVIERR_OK) Then
      MsgBox "There was an error writing the file.", vbInformation, App.Title
    End If
End Sub

Private Sub cmdMonthNext_Click()
    On Local Error Resume Next
    dispMONTH = dispMONTH + 1
    If dispMONTH = 13 Then
        dispMONTH = 1
        dispYEAR = dispYEAR + 1
    End If
    
    Call DisplayCalendar
    Call Display_Availability
End Sub

Private Sub cmdMonthPrev_Click()
    On Local Error Resume Next
    dispMONTH = dispMONTH - 1
    If dispMONTH = 0 Then
        dispMONTH = 12
        dispYEAR = dispYEAR - 1
    End If
    
    Call DisplayCalendar
    Call Display_Availability
End Sub


Private Sub Combo1_Click()
    On Local Error Resume Next
    Call DisplayCalendar
    Call Display_Availability
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    KeyCode = 0
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    KeyAscii = 0
End Sub

Private Sub Play_Stream()
    On Local Error Resume Next
    
    streamPLAYING = True
    While streamPLAYING
        Call Display_Frame(False)
        
        WaitABit streamRate, True
        If streamPLAYING = False Then Exit Sub
        
        If Slider1.Value + 1 > Slider1.Max Then
            Slider1.Value = 1
            streamPLAYING = False
            Call Invert_Buttons
        Else
            AllowSliderChange = False
            Slider1.Value = Slider1.Value + 1
            AllowSliderChange = True
        End If
        
        DoEvents
    Wend

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim A As Long
    Dim t As Long
    
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

    Call DisableALL
    
    
    For i = 1 To LocationCount
        machinename = ReadINI("Location" & i, "Name", INIFile)
    
        If machinename <> "Not In Use" Then
            Combo1.AddItem machinename
        End If
    Next i
    Combo1.ListIndex = 0
    
    t = 1590
    For i = 0 To 41
        lblDay(i).Left = lblDayHeader(A).Left
        lblDay(i).Top = t
        
        A = A + 1
        If A > 6 Then
            A = 0
            t = lblDay(i).Top + lblDay(i).Height - 15
        End If
    Next i
    
    Shape1.Height = (t - Shape1.Top) + 25
    
    dispMONTH = Month(Now())
    dispYEAR = Year(Now())
    
    Call DisplayCalendar
    Call Display_Availability
End Sub

Private Sub DisableALL()
    Dim i As Integer
    
    On Local Error Resume Next
    
    For i = 0 To 6
        imgPlayButtons(i).Picture = picDisabled.GraphicCell(i)
        imgPlayButtons(i).Enabled = False
    Next i
    
End Sub

Private Sub EnableForPlay()
    Dim i As Integer
    
    For i = 0 To 6
        If i >= 3 Then
            imgPlayButtons(i).Picture = picDisabled.GraphicCell(i)
            imgPlayButtons(i).Enabled = False
        Else
            imgPlayButtons(i).Picture = picEnabled.GraphicCell(i)
            imgPlayButtons(i).Enabled = True
        End If
    Next i

End Sub

Private Sub DisplayCalendar()
    Dim dispDATE As Date
    Dim dispDAY As String
    Dim i As Integer
    
    dispDATE = CDate("01/" & dispMONTH & "/" & dispYEAR)
    lblMonthYear.Caption = Format$(dispDATE, "mmmm yyyy")
    dispDAY = Format$(dispDATE, "ddd")
    While dispDAY <> "Mon"
        dispDATE = dispDATE - 1
        dispDAY = Format$(dispDATE, "ddd")
    Wend
    
    For i = 0 To 41
        'lblDay(i).Appearance = 0
        lblDay(i).FontBold = False
        lblDay(i).BackColor = vbWhite
        lblDay(i).fontSIZE = 10
        lblDay(i).Tag = Format$(dispDATE, "dd/mm/yyyy")
        lblDay(i).Caption = Day(dispDATE)
        
        If Month(dispDATE) <> dispMONTH Or Year(dispDATE) <> dispYEAR Then
            lblDay(i).ForeColor = &HC0C0C0
        Else
            lblDay(i).ForeColor = &HFF0000
        End If
        dispDATE = dispDATE + 1
    Next i
    
    List1.Clear
    List2.Clear
    prevSELECTED = -1
    
End Sub

Private Sub Display_Availability()
    Dim i As Integer
    Dim dispDATE As Date
    Dim folderDATE As String
    Dim testFolder As String
    Dim l As String
    
    For i = 0 To 41
        dispDATE = CDate(lblDay(i).Tag)
        If Month(dispDATE) = dispMONTH And Year(dispDATE) = dispYEAR Then
            folderDATE = Format$(dispDATE, "yyyymmdd")
            testFolder = App.Path & "\" & Combo1.Text & "\"
            l = Dir$(testFolder, vbDirectory)
            Do While Len(l) <> 0
                If l <> "." And l <> ".." Then
                    If (GetAttr(testFolder & l) And vbDirectory) = vbDirectory Then
                        If l = folderDATE Then
                            lblDay(i).FontBold = True
                            lblDay(i).fontSIZE = 12
                            l = ""
                            GoTo eLOOP
                        End If
                    End If
                End If
                l = Dir$
eLOOP:
            Loop
        End If
    Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Local Error Resume Next
    
    ' we may be playing a stream, if so stop it
    If streamPLAYING Then
        Call imgPlayButtons_Click(6)
        Image1.Picture = imgBLUE.Picture
        streamPLAYING = False
    End If
    
    If streamFN <> 0 Then Close #streamFN
    streamFN = 0
    
End Sub

Private Sub Invert_Buttons()
    Dim i As Integer
    On Local Error Resume Next
    For i = 0 To 6
        If imgPlayButtons(i).Enabled Then
            imgPlayButtons(i).Enabled = False
            imgPlayButtons(i).Picture = picDisabled.GraphicCell(i)
        Else
            imgPlayButtons(i).Enabled = True
            imgPlayButtons(i).Picture = picEnabled.GraphicCell(i)
        End If
    Next i
End Sub

Private Sub imgEXPAND_Click()
    Dim i As Integer
    Dim l As Long
    Dim W As Long
    Dim Bid(0 To 6) As Integer
    
    On Local Error Resume Next
    
    Bid(0) = 0
    Bid(1) = 4
    Bid(2) = 2
    Bid(3) = 3
    Bid(4) = 6
    Bid(5) = 5
    Bid(6) = 1
    
Undo:
    If imgEXPAND.Tag = 0 Then
        imgEXPAND.Tag = 1
        imgEXPAND.ToolTipText = "Minimise Picture"
    Else
        imgEXPAND.Tag = 0
        imgEXPAND.ToolTipText = "Maximise Picture"
    End If
    
    imgEXPAND.Picture = imgArrow(imgEXPAND.Tag).Picture
    DoEvents
    
    Select Case imgEXPAND.Tag
        Case "0"
            ' small size
            Image1.Stretch = True
            Image1.Width = 6255
            Image1.Height = 4935
            Me.Width = 10170
            Me.Height = 7575
        Case "1"
            If wmp.Visible = False Then
                Image1.Stretch = False
                DoEvents
                If Image1.Width < 6255 Or Image1.Height < 4935 Then GoTo Undo
            Else
                Image1.Width = 12030
                Image1.Height = 9030
            End If
            
            Me.Width = Image1.Left + Image1.Width + 315
            Me.Height = Image1.Top + Image1.Height + 2400
    End Select
    
    'imgEXPAND.top = (Image1.top + Image1.Height) - (imgEXPAND.Height + 120)
    'imgEXPAND.left = (Image1.left + Image1.Width) - (imgEXPAND.Width + 120)
    
    
    ' adjust slider
    
    Slider1.Width = Image1.Width
    Slider1.Top = Image1.Top + Image1.Height + 115
    
    ' playbuttons
    W = (imgPlayButtons(0).Width * 7) + (105 * 6)
    l = (Image1.Width - W) \ 2
    l = l + Image1.Left
    For i = 0 To 6
        imgPlayButtons(Bid(i)).Top = Image1.Top + Image1.Height + 585
        imgPlayButtons(Bid(i)).Left = l
        l = l + imgPlayButtons(Bid(i)).Width + 105
    Next i
    
    Frame2.Top = Image1.Top + Image1.Height + 1185
    l = (Image1.Width - Frame2.Width) \ 2
    l = l + Image1.Left
    Frame2.Left = l

    If wmp.Visible Then
        wmp.Top = Image1.Top
        wmp.Left = Image1.Left
        wmp.Width = Image1.Width
        wmp.Height = Me.Height - 750
    End If
    
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2


End Sub

Private Sub imgPlayButtons_Click(Index As Integer)
    ' image play click
    
    Call Open_Stream
    
    Select Case Index
        Case 0 ' move to start
            Slider1.Value = 0
        Case 1 ' move to end
            Slider1.Value = Slider1.Max
        Case 2 ' play
            streamRate = 1
            Call Invert_Buttons
            Call Play_Stream
        Case 3 ' pause
            Call Invert_Buttons
            streamPLAYING = False
        Case 4 ' slow down
            If streamRate = 1.5 Then Exit Sub
            streamRate = streamRate + 0.5
            Select Case streamRate
                Case 0.5
                    Label3.Caption = "Rate x2"
                Case 1
                    Label3.Caption = "Rate x1"
                Case 1.5
                    Label3.Caption = "Rate x-2"
            End Select
        Case 5 ' speed up
            If streamRate = 0.5 Then Exit Sub
            streamRate = streamRate - 0.5
            Select Case streamRate
                Case 0.5
                    Label3.Caption = "Rate x2"
                Case 1
                    Label3.Caption = "Rate x1"
                Case 1.5
                    Label3.Caption = "Rate x-2"
            End Select
        Case 6 ' stop
            Call Invert_Buttons
            streamPLAYING = False
            Slider1.Value = 1
    End Select
    
End Sub

Private Sub Sync_Stream_To_Slider()
    Dim cssHEADER As cssHEADERtype
    Dim cssFRAME As cssFRAMEtype
    Dim fCOUNT As Long
    Dim A As Long
    Dim dummy As String * 1
    
    ' 1st read in header
    Get #streamFN, 1, cssHEADER
    
    fPOS = Len(cssHEADER) + 1
    
    If Slider1.Value = 1 Then Exit Sub

    
    For A = 1 To (Slider1.Value - 1)
        Get #streamFN, fPOS, cssFRAME
        fPOS = fPOS + Len(cssFRAME)
        fPOS = fPOS + (cssFRAME.jpgSize - 1)
        Get #streamFN, fPOS, dummy
        fPOS = fPOS + 1
    Next A
    
End Sub

Private Sub Display_Frame(RollBackFrame As Boolean)
    Dim cssFRAME As cssFRAMEtype
    Dim picDATA As String
    Dim pFN As Integer
    Dim dummy As String * 1
    
    On Local Error Resume Next
    
    Get #streamFN, , cssFRAME
    picDATA = Space$(cssFRAME.jpgSize)
    Get #streamFN, , picDATA
    
    If RollBackFrame = True Then
        Get #streamFN, fPOS - 1, dummy
    Else
        fPOS = fPOS + Len(cssFRAME) + cssFRAME.jpgSize
    End If
    
    pFN = FreeFile
    Open App.Path & "\frame.jpg" For Binary Access Write As #pFN
    Put #pFN, , picDATA
    Close #pFN
    
    Image1.Picture = LoadPicture(App.Path & "\frame.jpg")
    
    If AVIExporting Then
        Call SavePicture(Image1.Picture, App.Path & "\frame.bmp")
    End If
End Sub

Private Sub Open_Stream()
    On Local Error Resume Next
    
    If streamFN <> 0 Then Exit Sub
    
    streamFN = FreeFile
    Open List2.List(List1.ListIndex) For Binary Access Read As #streamFN
    
End Sub

Private Sub lblDay_Click(Index As Integer)
    Dim dispDATE As Date
    Dim folderDATE As String
    
    On Local Error Resume Next
    
    dispDATE = CDate(lblDay(Index).Tag)
    If Month(dispDATE) <> dispMONTH And Year(dispDATE) <> dispYEAR Then Exit Sub
    
    If lblDay(Index).FontBold = False Then Exit Sub
    
    If prevSELECTED >= 0 Then
        'lblDay(prevSELECTED).Appearance = 0
        lblDay(prevSELECTED).FontBold = True
        lblDay(prevSELECTED).BackColor = vbWhite
        lblDay(prevSELECTED).ForeColor = &HFF0000
    End If
    
    prevSELECTED = Index
    'lblDay(prevSELECTED).Appearance = 0
    lblDay(prevSELECTED).FontBold = True
    lblDay(prevSELECTED).ForeColor = vbWhite
    lblDay(prevSELECTED).BackColor = &HFF0000
        

    folderDATE = Format$(dispDATE, "yyyymmdd")
    File1.Path = App.Path & "\" & Combo1.Text & "\" & folderDATE
    File1.Pattern = "*.mcp;*.avi"
    selectedDATE = dispDATE
    
    Call Load_Clips
    
End Sub

Private Sub Load_Clips()
    Dim i As Integer
    Dim l As String
    Dim cssHEADER As cssHEADERtype
    
    On Local Error Resume Next
    
    File1.Refresh
    List1.Clear
    List2.Clear
    For i = 0 To (File1.ListCount - 1)
        l = File1.Path & "\" & File1.List(i)
        List2.AddItem l
        
        If Right$(l, 3) = "mcp" Then
            cssHEADER = Read_Capture_Header(l)
        Else
            l = File1.List(i)
            l = Replace_Text(l, "Capture~", "")
            l = Left$(l, Len(l) - 4)
            
            cssHEADER.Started = CDate(Format$(Date, "dd/mmm/yyyy") & " " & Left$(l, 2) & ":" & Mid$(l, 3, 2) & ":" & Mid$(l, 5, 2))
            cssHEADER.Stopped = CDate(Format$(Date, "dd/mmm/yyyy") & " " & Mid$(l, 8, 2) & ":" & Mid$(l, 10, 2) & ":" & Right$(l, 2))
            
            cssHEADER.frameCount = Abs(DateDiff("s", cssHEADER.Started, cssHEADER.Stopped)) + 1
        End If
        
        l = Format$(cssHEADER.Started, "HH:MM:SS") & "  " & Format$(cssHEADER.Stopped, "HH:MM:SS") & "   " & Seconds_To_Text(cssHEADER.frameCount)
        List1.AddItem l
    Next i

End Sub

Private Sub List1_Click()
    Dim cssHEADER As cssHEADERtype
    Dim l As String
    
    On Local Error Resume Next
    
    ' we may be playing a stream, if so stop it
    If streamPLAYING Then
        Call imgPlayButtons_Click(6)
        Image1.Picture = imgBLUE.Picture
        streamPLAYING = False
    End If
    
    If streamFN <> 0 Then Close #streamFN
    streamFN = 0
    
    l = List2.List(List1.ListIndex)
    
    If Right$(l, 3) = "mcp" Then
        wmp.Controls.stop
        wmp.Close
        cmdEXPORT.Enabled = True
        wmp.Visible = False
        wmp.Height = 0
        wmp.Width = 0
        DoEvents
        Call SetViewerArea(True)
        cssHEADER = Read_Capture_Header(List2.List(List1.ListIndex))
        Slider1.Min = 1
        Slider1.Max = cssHEADER.frameCount
        If cssHEADER.frameCount >= 3600 Then
            If cssHEADER.frameCount < 7200 Then
                Slider1.TickFrequency = 300 ' less than 2hrs tick every 5mins
            Else
                Slider1.TickFrequency = 600 ' more than 2hrs tick every 10mins
            End If
        ElseIf cssHEADER.frameCount >= 60 Then
            ' more than a min, but less than 1hour
            Slider1.TickFrequency = 30 'every 30secs
        Else
            Slider1.TickFrequency = 1
        End If
        Call Open_Stream
        AllowSliderChange = True
        Slider1.Value = 0
        streamRate = 1
        Select Case streamRate
            Case 0.5
                Label3.Caption = "Rate x2"
            Case 1
                Label3.Caption = "Rate x1"
            Case 1.5
                Label3.Caption = "Rate x-2"
        End Select
        Slider1.ToolTipText = "0"
        If imgEXPAND.Tag = 1 Then
            Call imgEXPAND_Click
        End If
        Call EnableForPlay
    Else
        cmdEXPORT.Enabled = False
        Call SetViewerArea(False)
        DoEvents
        wmp.Top = Image1.Top
        wmp.Left = Image1.Left
        wmp.Width = Image1.Width
        wmp.Height = Me.Height - 750
        wmp.Visible = True
        wmp.URL = l
        wmp.stretchToFit = True
        wmp.Controls.stop
        DoEvents
        If imgEXPAND.Tag = 1 Then
            Call imgEXPAND_Click
        End If
    End If
    
End Sub

Private Sub SetViewerArea(DispForMCP As Boolean)
    On Local Error Resume Next
    Image1.Visible = DispForMCP
    Slider1.Visible = DispForMCP
    Frame2.Visible = DispForMCP
'    imgEXPAND.Visible = DispForMCP
    imgPlayButtons(0).Visible = DispForMCP
    imgPlayButtons(1).Visible = DispForMCP
    imgPlayButtons(2).Visible = DispForMCP
    imgPlayButtons(3).Visible = DispForMCP
    imgPlayButtons(4).Visible = DispForMCP
    imgPlayButtons(5).Visible = DispForMCP
    imgPlayButtons(6).Visible = DispForMCP

End Sub

Private Sub Slider1_Change()
    On Local Error Resume Next
    Slider1.ToolTipText = Slider1.Value
    Label2.Caption = "Frame " & Slider1.Value
    If AllowSliderChange = False Then Exit Sub
    Call Sync_Stream_To_Slider
    Call Display_Frame(False)
End Sub

