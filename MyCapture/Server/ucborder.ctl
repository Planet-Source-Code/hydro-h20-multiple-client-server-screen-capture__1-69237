VERSION 5.00
Begin VB.UserControl ucBorder 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image imgLeftMiddle 
      Height          =   525
      Left            =   400
      Picture         =   "ucBorder.ctx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   45
   End
   Begin VB.Image imgRightMiddle 
      Height          =   525
      Left            =   2280
      Picture         =   "ucBorder.ctx":01E6
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   45
   End
   Begin VB.Image imgBottomMiddle 
      Height          =   45
      Left            =   1800
      Picture         =   "ucBorder.ctx":03CC
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   735
   End
   Begin VB.Image imgTopMiddle 
      Height          =   45
      Left            =   1560
      Picture         =   "ucBorder.ctx":05CA
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image imgBottomRightCorner 
      Height          =   45
      Left            =   3240
      Picture         =   "ucBorder.ctx":07C8
      Top             =   2520
      Width           =   45
   End
   Begin VB.Image imgBottomLeftCorner 
      Height          =   45
      Left            =   480
      Picture         =   "ucBorder.ctx":082E
      Top             =   2760
      Width           =   45
   End
   Begin VB.Image imgTopLeftCorner 
      Height          =   45
      Left            =   360
      Picture         =   "ucBorder.ctx":0894
      Top             =   840
      Width           =   45
   End
   Begin VB.Image imgTopRightCorner 
      Height          =   45
      Left            =   3240
      Picture         =   "ucBorder.ctx":08FA
      Top             =   480
      Width           =   45
   End
End
Attribute VB_Name = "ucBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Size_And_Move()

    ' top right corner
    imgTopRightCorner.Move 0, 0
    
    ' top left cornner
    imgTopLeftCorner.Move UserControl.Width - imgTopRightCorner.Width, 0
    
    ' bottom left corner
    imgBottomLeftCorner.Move 0, UserControl.Height - imgBottomLeftCorner.Height
    
    ' bottom right corner
    imgBottomRightCorner.Move UserControl.Width - imgBottomRightCorner.Width, UserControl.Height - imgBottomRightCorner.Height
    
    ' top middle
    imgTopMiddle.Move imgTopRightCorner.Width, 0, UserControl.Width - (imgTopLeftCorner.Width + imgTopRightCorner.Width)
    
    ' bottom middle
    imgBottomMiddle.Move imgBottomRightCorner.Width, UserControl.Height - imgBottomMiddle.Height, UserControl.Width - (imgBottomLeftCorner.Width + imgBottomRightCorner.Width)
    
    ' left middle
    imgLeftMiddle.Move 0, imgTopRightCorner.Height, imgLeftMiddle.Width, UserControl.Height - (imgTopLeftCorner.Height + imgBottomLeftCorner.Height)
    
    ' right middle
    imgRightMiddle.Move UserControl.Width - imgRightMiddle.Width, imgTopRightCorner.Height, imgRightMiddle.Width, UserControl.Height - (imgTopRightCorner.Height + imgBottomRightCorner.Height)
    
    
End Sub

Private Sub UserControl_Resize()
    Call Size_And_Move
End Sub
