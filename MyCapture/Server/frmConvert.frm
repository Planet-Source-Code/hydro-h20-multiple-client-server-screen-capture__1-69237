VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmConvert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCP to AVI Convertor"
   ClientHeight    =   7095
   ClientLeft      =   5805
   ClientTop       =   4110
   ClientWidth     =   6540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7440
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox lblFILE 
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmConvert.frx":0000
      Top             =   5520
      Width           =   6255
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   6360
      Width           =   2055
      Begin VB.Label Label2 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   3375
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   327682
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4935
      Left            =   120
      Picture         =   "frmConvert.frx":0006
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmConvert"
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

Public Sub cmdEXPORT_Click()
    Dim aviFILE As String
    Dim fNAME As String
    Dim cssHEADER As cssHEADERtype
    Dim mcpFILE As String
    Dim capMACHINE As String
    
    On Local Error Resume Next
    
    If ReadINI("General", "ComputerName", INIFile) <> CompName Then Exit Sub
    
    Call Open_Stream
    
    AllowSliderChange = True
    Slider1.Value = 0
    
    AVIExporting = True
    

    mcpFILE = lblFILE.Text
    cssHEADER = Read_Capture_Header(mcpFILE)
    aviFILE = JustPath(mcpFILE) & "Capture~" & Format$(cssHEADER.Started, "HHMMSS") & "~" & Format$(cssHEADER.Stopped, "HHMMSS") & ".avi"

    ' get machine name
    fNAME = mcpFILE
    fNAME = Replace_Text(fNAME, App.Path & "\", "")
    i = InStr(1, fNAME, "\", vbTextCompare)
    capMACHINE = left$(fNAME, i - 1)
    Create_Header_Frame Me, capMACHINE, Format$(cssHEADER.Started, "ddd dd/mmm/yyyy"), Format$(cssHEADER.Started, "HH:MM:SS"), Format$(cssHEADER.Stopped, "HH:MM:SS")

    
    Call AVIFileInit
    Call WriteAVI(aviFILE, 1)
    Call AVIFileExit
    AVIExporting = False
 
    Slider1.Value = 0
    streamPLAYING = False
    
    
    If streamFN <> 0 Then
        Close #streamFN
        streamFN = 0
    End If
    
    AVIConvertingNow = False
    Kill mcpFILE
    

End Sub

Private Sub WriteAVI(ByVal filename As String, Optional ByVal FrameRate As Integer = 1)
    Dim s$
    Dim InitDir As String
    Dim szOutputAVIFile As String
    Dim res As Long
    Dim pfile As Long 'ptr PAVIFILE
    Dim bmp As cDIB
    Dim ps As Long 'ptr PAVISTREAM
    Dim psCompressed As Long 'ptr PAVISTREAM
    Dim strhdr As AVI_STREAM_INFO
    Dim BI As BITMAPINFOHEADER
    Dim opts As AVI_COMPRESS_OPTIONS
    Dim pOpts As Long
    Dim i As Long
    
    'get an avi filename from user
    szOutputAVIFile = filename$
'    Open the file for writing
    res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo error

    'Get the first bmp in the list for setting format
    Set bmp = New cDIB
    s$ = App.Path & IIf(right$(App.Path, 1) <> "\", "\", "") & "frame.bmp"
    If bmp.CreateFromFile(s$) <> True Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
        GoTo error
    End If

'   Fill in the header for the video stream
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
        .fccHandler = 0&                             '// default AVI handler
        .dwScale = 1
        .dwRate = Val(FrameRate%)                        '// fps
        .dwSuggestedBufferSize = bmp.SizeImage       '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, bmp.Width, bmp.Height)       '// rectangle for stream
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
        .biBitCount = bmp.BitCount
        .biClrImportant = bmp.ClrImportant
        .biClrUsed = bmp.ClrUsed
        .biCompression = bmp.Compression
        .biHeight = bmp.Height
        .biWidth = bmp.Width
        .biPlanes = bmp.Planes
        .biSize = bmp.SizeInfoHeader
        .biSizeImage = bmp.SizeImage
        .biXPelsPerMeter = bmp.XPPM
        .biYPelsPerMeter = bmp.YPPM
    End With
    
    'set the format of the compressed stream
    res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo error

    ' write 5 secs of header
    bmp.CreateFromFile App.Path & "\header.bmp"
    For i = 1 To 5
         res = AVIStreamWrite(psCompressed, i - 1, 1, bmp.PointerToBits, bmp.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
         If res <> AVIERR_OK Then GoTo error
    Next i

    AllowSliderChange = False
    For i = 1 To Slider1.Max
        Slider1.Value = i
        Call Display_Frame(False)
        DoEvents
        
         bmp.CreateFromFile (s$) 'load the bitmap (ignore errors)
         res = AVIStreamWrite(psCompressed, (i + 5) - 1, 1, bmp.PointerToBits, bmp.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
         If res <> AVIERR_OK Then GoTo error
    Next i
    AllowSliderChange = True
    
error:
'   Now close the file
    Set bmp = Nothing
    
    If (ps <> 0) Then Call AVIStreamClose(ps)

    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)

    If (pfile <> 0) Then Call AVIFileClose(pfile)

    Call AVIFileExit

    If (res <> AVIERR_OK) Then
      MsgBox "There was an error writing the file.", vbInformation, App.Title
    End If
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
        Else
            AllowSliderChange = False
            Slider1.Value = Slider1.Value + 1
            AllowSliderChange = True
        End If
        
        DoEvents
    Wend

End Sub

Private Sub Form_Load()
    
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Local Error Resume Next
    
    
    If streamFN <> 0 Then Close #streamFN
    streamFN = 0
    
End Sub

Private Sub Sync_Stream_To_Slider()
    Dim cssHEADER As cssHEADERtype
    Dim cssFRAME As cssFRAMEtype
    Dim fCOUNT As Long
    Dim a As Long
    Dim dummy As String * 1
    
    ' 1st read in header
    Get #streamFN, 1, cssHEADER
    
    fPOS = Len(cssHEADER) + 1
    
    If Slider1.Value = 1 Then Exit Sub

    
    For a = 1 To (Slider1.Value - 1)
        Get #streamFN, fPOS, cssFRAME
        fPOS = fPOS + Len(cssFRAME)
        fPOS = fPOS + (cssFRAME.jpgSize - 1)
        Get #streamFN, fPOS, dummy
        fPOS = fPOS + 1
    Next a
    
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
    Dim cssHEADER As cssHEADERtype
    
    On Local Error Resume Next
    
    If streamFN <> 0 Then Exit Sub
    
    cssHEADER = Read_Capture_Header(lblFILE.Text)
    Slider1.Min = 1
    Slider1.Max = cssHEADER.frameCOUNT
    If cssHEADER.frameCOUNT >= 3600 Then
        If cssHEADER.frameCOUNT < 7200 Then
            Slider1.TickFrequency = 300 ' less than 2hrs tick every 5mins
        Else
            Slider1.TickFrequency = 600 ' more than 2hrs tick every 10mins
        End If
    ElseIf cssHEADER.frameCOUNT >= 60 Then
        ' more than a min, but less than 1hour
        Slider1.TickFrequency = 30 'every 30secs
    Else
        Slider1.TickFrequency = 1
    End If
        
    streamFN = FreeFile
    Open lblFILE.Text For Binary Access Read As #streamFN
    
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
    
End Sub

Private Sub Slider1_Change()
    On Local Error Resume Next
    Slider1.ToolTipText = Slider1.Value
    Label2.Caption = "Frame " & Slider1.Value
    If AllowSliderChange = False Then Exit Sub
    Call Sync_Stream_To_Slider
    Call Display_Frame(False)
End Sub

Private Sub Timer1_Timer()
    On Local Error Resume Next
    Call frmONTOP(Me, True)
End Sub
