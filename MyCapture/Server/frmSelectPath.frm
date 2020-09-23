VERSION 5.00
Begin VB.Form frmSelectPath 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   3090
   ClientTop       =   3825
   ClientWidth     =   8085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSelectPath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5520
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7560
      Top             =   4800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select"
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5520
      Width           =   735
   End
   Begin Server.ucNetworkTree ucNetworkTree1 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8705
      ShowPrinterShares=   0   'False
      AutoLoadTree    =   -1  'True
   End
   Begin VB.Label lblNewPath 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   5295
   End
   Begin VB.Label lblOldPath 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmSelectPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Local Error Resume Next
    Call ucNetworkTree1.Load_Tree
    ucNetworkTree1.SelectedFolder = lblNewPath
End Sub

Private Sub Command2_Click()
    Dim l As String
    Dim r As Long
    
    On Local Error Resume Next
    
    If Len(lblNewPath.Caption) = 0 Then
        Timer1.Enabled = False
        MsgBox "You Have Not Selected A Network Path.", vbOKOnly + vbExclamation + vbSystemModal, "No Path"
        Timer1.Enabled = True
        Exit Sub
    End If
    
    ' test path
    l = Dir$(lblNewPath.Caption & "*.*")
    If Err.Number <> 0 Then
        Timer1.Enabled = False
        MsgBox "There Appears To Be A Problem With This Path." & vbCrLf & vbCrLf & "Please Select Another Path.", vbOKOnly + vbExclamation + vbSystemModal, "Problem With Path"
        Timer1.Enabled = False
        Exit Sub
    End If
    
    l = Dir$(lblNewPath.Caption & "CaptureIt.exe")
    If Len(l) = 0 Then
        Timer1.Enabled = False
        r = MsgBox("CaptureIt Is Not Installed In This Path." & vbCrLf & vbCrLf & "Do You Still Want To Select This Path?", vbYesNo + vbQuestion + vbSystemModal, "Select This Path")
        If r = vbNo Then
            Timer1.Enabled = True
            Exit Sub
        End If
    End If
    
    Timer1.Enabled = False
    lblNewPath.Caption = Left$(lblNewPath.Caption, Len(lblNewPath.Caption) - 1)
    Me.Visible = False
End Sub

Private Sub Command3_Click()
    lblNewPath.Caption = ""
    Timer1.Enabled = False
    Me.Visible = False
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
End Sub

Private Sub Timer1_Timer()
    On Local Error Resume Next
    Call frmONTOP(Me, True)
End Sub

Private Sub ucNetworkTree1_FolderChanged()
    On Local Error Resume Next
    lblNewPath.Caption = ucNetworkTree1.SelectedFolder
End Sub

