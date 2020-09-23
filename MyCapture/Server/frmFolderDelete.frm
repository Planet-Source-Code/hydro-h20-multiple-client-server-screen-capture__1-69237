VERSION 5.00
Begin VB.Form frmFolderDelete 
   BorderStyle     =   0  'None
   Caption         =   "ted"
   ClientHeight    =   2580
   ClientLeft      =   5055
   ClientTop       =   4980
   ClientWidth     =   5655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin Server.ucGIF ucGIF1 
      Height          =   900
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   1588
      borderStyle     =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0   'False
      fileLen         =   13283
      fileData        =   "frmFolderDelete.frx":0000
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4320
      Top             =   240
   End
   Begin Server.ucBorder ucBorder1 
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2778
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Initialising..."
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
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4575
   End
End
Attribute VB_Name = "frmFolderDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ucGIF1.BackColor = Me.BackColor
    ucGIF1.Animate = True
    Me.Width = ucBorder1.Width
    Me.Height = ucBorder1.Height
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
End Sub

Private Sub Timer1_Timer()
    On Local Error Resume Next
    Call frmONTOP(Me, True)
End Sub
