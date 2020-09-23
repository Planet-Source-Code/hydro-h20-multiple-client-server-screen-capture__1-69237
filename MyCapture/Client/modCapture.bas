Attribute VB_Name = "modCapture"
Option Explicit
Option Compare Text

' API Stuff
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long
Public Const STRETCHMODE = vbPaletteModeNone   'You can find other modes in the "PaletteModeConstants" section of your Object Browser
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

' API For Capturing cursor STD only at the mo
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type PCURSORINFO
    cbSize As Long
    Flags As Long
    hCursor As Long
    ptScreenPos As POINTAPI
End Type
'To grab cursor shape -require at least win98 as per Microsoft documentation...
Public Declare Function GetCursorInfo Lib "user32.dll" (ByRef pci As PCURSORINFO) As Long
'To get a Handle to the cursor
Public Declare Function GetCursor Lib "user32" () As Long
'To draw cursor shape on bitmap
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
'to get the cursor position
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Const CursorIconSize As Integer = 9

Public Type CaptureAreaType
    startTop As Long
    startLeft As Long
    srcWidth As Long
    srcHeight As Long
    targetWidth As Long
    targetHeight As Long
    Compression As Integer
    DTStamp As Boolean
    qualityH As Integer
    qualityV As Integer
    Convert256 As Boolean
    CaptureMouse As Boolean
End Type

Public CaptureArea As CaptureAreaType
Public FrameCounter As Long
Public tmpBMPname As String
Public tmpBMP256name As String
Public tmpJPGname As String
Public LastTimeCaptured As String



Public Function CaptureDesktop()
    Dim dskDC As Long
    Dim m_Jpeg As cJPEGi
    Dim m_Image As cImage
    Dim MyPic As StdPicture
    Dim impBMP As String
    Dim Point As POINTAPI
    Dim pcin As PCURSORINFO
    Dim ret
    
    On Local Error Resume Next
    
    dskDC = GetDC(0)
    'desktop capturing
    BitBlt frmCapture.Picture1.hdc, 0, 0, CaptureArea.srcWidth, CaptureArea.srcHeight, dskDC, CaptureArea.startLeft, CaptureArea.startTop, vbSrcCopy
        
    'free memory, otherwise after 3-5 minutes everything will crush!
    ReleaseDC 0, dskDC
    
    ' stamp the date and time onto image
    If CaptureArea.DTStamp Then
        frmCapture.Picture1.CurrentX = 10
        frmCapture.Picture1.CurrentY = 10
        frmCapture.Picture1.Print Format$(Date, "ddd dd/mmm/yyyy") & " " & LastTimeCaptured
        frmCapture.Picture1.Refresh
    End If
            
    If CaptureArea.CaptureMouse Then
        'now to get the icon of mouse and paint on form the mouse
        GetCursorPos Point
        pcin.hCursor = GetCursor
        pcin.cbSize = Len(pcin)
        ret = GetCursorInfo(pcin)
        DrawIcon frmCapture.Picture1.hdc, Point.x - CursorIconSize - CaptureArea.startLeft, Point.y - CursorIconSize - CaptureArea.startTop, pcin.hCursor
    End If
    
    frmCapture.picRSetMode.Cls
    If CaptureArea.srcHeight <> CaptureArea.targetHeight Or CaptureArea.srcWidth <> CaptureArea.targetWidth Then
        Call SetStretchBltMode(frmCapture.picRSetMode.hdc, STRETCHMODE)
        Call StretchBlt(frmCapture.picRSetMode.hdc, 0, 0, CaptureArea.targetWidth, CaptureArea.targetHeight, frmCapture.Picture1.hdc, 0, 0, CaptureArea.srcWidth, CaptureArea.srcHeight, vbSrcCopy)
    Else
        BitBlt frmCapture.picRSetMode.hdc, 0, 0, CaptureArea.srcWidth, CaptureArea.srcHeight, frmCapture.Picture1.hdc, CaptureArea.startLeft, CaptureArea.startTop, vbSrcCopy
    End If
    frmCapture.picRSetMode.Refresh
            
            
    SavePicture frmCapture.picRSetMode.Image, tmpBMPname
    impBMP = tmpBMPname
    
    ' convert to 256 here
    If CaptureArea.Convert256 Then
        Convert256 tmpBMPname, tmpBMP256name
        impBMP = tmpBMP256name
    End If
    
    ' convert BMP to JPG
    Set m_Jpeg = New cJPEGi
    Set m_Image = New cImage
    
    Set MyPic = LoadPicture(impBMP)
    
    m_Image.CopyStdPicture MyPic
            
    m_Jpeg.Quality = CaptureArea.Compression
    m_Jpeg.SetSamplingFrequencies CaptureArea.qualityH, CaptureArea.qualityV, 1, 1, 1, 1
    m_Jpeg.SampleHDC m_Image.hdc, m_Image.Width, m_Image.Height

    tmpJPGname = App.Path & "\Captures\cap" & Format$(FrameCounter, "00000000") & "~" & Format$(Date, "ddmmyyyy") & "~" & Format$(Time$, "HHMMSS") & ".jpg"
    If Len(Dir$(tmpJPGname)) <> 0 Then Kill tmpJPGname
    
    m_Jpeg.SaveFile tmpJPGname

    Set m_Image = Nothing
    Set m_Jpeg = Nothing
    Set MyPic = Nothing
            

End Function

Private Sub Convert256(bmpFILE As String, bmp256File As String)
Dim m_cDIB As cDIBSection
Dim oPic As StdPicture
Dim eD As EDSSColourDepthConstants
Dim cDIBSave As New cDIBSectionSave ' ' call class module for save

    On Local Error Resume Next
    
    Set m_cDIB = New cDIBSection
    Set oPic = LoadPicture(bmpFILE)
   
    m_cDIB.CreateFromPicture oPic
    eD = edss256Colour
    
    
    cDIBSave.Save bmp256File, m_cDIB, , eD, edssSystemDefault
    
    Set m_cDIB = Nothing
    Set oPic = Nothing
    Set cDIBSave = Nothing
    
    
End Sub
