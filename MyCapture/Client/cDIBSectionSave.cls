VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSectionSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO2
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 1) As RGBQUAD
End Type
Private Type BITMAPINFO16
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 15) As RGBQUAD
End Type
Private Type BITMAPINFO256
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type

Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Declare Function CreateDIBSection2 Lib "gdi32" Alias "CreateDIBSection" _
    (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO2, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long
Private Declare Function CreateDIBSection16 Lib "gdi32" Alias "CreateDIBSection" _
    (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO16, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long
Private Declare Function CreateDIBSection256 Lib "gdi32" Alias "CreateDIBSection" _
    (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO256, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long

Private Const BITMAPTYPE As Integer = &H4D42
Private Type BITMAPFILEHEADER
   bfType As Integer '- type  ="BM" i.e &H4D42 - 2
   bfSize As Long ' - size in bytes of file - 6
   bfReserved1 As Integer ' - reserved, must be 0 - 8
   bfReserved2 As Integer ' - reserved, must be 0 - 10
   bfOffBits As Long ' offset from this structure to the bitmap bits - 14
End Type

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const CREATE_ALWAYS = 2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_BEGIN = 0
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Long, lpBI As BITMAPINFO256, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO256, ByVal wUsage As Long) As Long


Public Enum EDSSColourDepthConstants
   edss2Colour
   edss16Colour
   edss256Colour
   edssTrueColour
End Enum

Public Enum EDSSColourReductionConstants
   edssSystemDefault
   edssUsePalette
   edssGeneratePalette
End Enum

Public Function Save( _
      ByVal sFileName As String, _
      ByRef cDIB As cDIBSection, _
      Optional ByRef cPal As cPalette = Nothing, _
      Optional ByVal eOutputColourDepth As EDSSColourDepthConstants = edss256Colour, _
      Optional ByVal eColourReductionMethod As EDSSColourReductionConstants = edssSystemDefault _
   ) As Boolean
Dim tBIH As BITMAPINFOHEADER

   If eOutputColourDepth = edssTrueColour Then
      ' Just directly save the DIBSection as is:
      Save = cDIB.SavePicture(sFileName)
      
   Else
      ' We must create a new DIBSection of the correct colour depth to save.
      ' Also, we may need to perform a colour depth reduction before saving
      
      ' First create the DIBSection
      Dim tBI2 As BITMAPINFO2, tBI16 As BITMAPINFO16, tBI256 As cDIBSectionSave.BITMAPINFO256
      Dim hDib As Long, lptr As Long, bSuccess As Boolean
      Dim lHDC As Long, hBmpOld As Long, lHDCWOrk As Long, i As Long
      Dim cDIBWork As cDIBSection
      Dim cColReduce As New cColourReduceDIB
      Dim bBltIn As Boolean
               
      ' Do any colour reduction as required:
      Select Case eColourReductionMethod
      Case edssSystemDefault
         bBltIn = True
         
      Case edssUsePalette
         Set cDIBWork = New cDIBSection
         cDIBWork.Create cDIB.Width, cDIB.Height
         cColReduce.ApplyPalette cDIB, cDIBWork, cPal
      
      Case edssGeneratePalette
         Select Case eOutputColourDepth
         Case edss256Colour
            ' Create optimal palette using octree quantisation:
            Set cPal = New cPalette
            cPal.CreateOptimal cDIB
            Set cDIBWork = New cDIBSection
            cDIBWork.Create cDIB.Width, cDIB.Height
            cColReduce.ApplyPalette cDIB, cDIBWork, cPal, False
            
         Case edss16Colour
            ' Use a default palette
            Set cPal = New cPalette
            cPal.Create16Colour
            Set cDIBWork = New cDIBSection
            cDIBWork.Create cDIB.Width, cDIB.Height
            cColReduce.ApplyPalette cDIB, cDIBWork, cPal
            
         Case edss2Colour
            ' Use a default palette
            Set cDIBWork = New cDIBSection
            cDIBWork.Create cDIB.Width, cDIB.Height
            cColReduce.BlackAndWhite cDIB, cDIBWork
         
         End Select
      End Select
                              
      ' Create our output DIB section (with appropriate palette):
      Select Case eOutputColourDepth
      Case edss2Colour
         pbCreate2ColourDIBSection hDib, tBI2, lptr, cDIB.Width, cDIB.Height
      Case edss16Colour
         pbCreate16ColourDIBSection hDib, tBI16, lptr, cDIB.Width, cDIB.Height
      Case edss256Colour
         pbCreate256ColourDIBSection hDib, tBI256, lptr, cDIB.Width, cDIB.Height, cPal
      End Select
      
      If Not hDib = 0 Then
                              
         ' Blit the appropriate true colour DIB into our DIB:
         lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
         lHDCWOrk = CreateCompatibleDC(lHDC)
         DeleteDC lHDC
         hBmpOld = SelectObject(lHDCWOrk, hDib)
         cDIB.PaintPicture lHDCWOrk
         
         If bBltIn Then
            cDIB.LoadPictureBlt lHDCWOrk
         End If
         
         SelectObject lHDCWOrk, hBmpOld
         DeleteDC lHDCWOrk
                  
                  
         ' Save the bitmap we created:
         Select Case eOutputColourDepth
         Case edss2Colour
            bSuccess = SaveToBitmap2(tBI2, lptr, sFileName)
         Case edss16Colour
            bSuccess = SaveToBitmap16(tBI16, lptr, sFileName)
         Case edss256Colour
            bSuccess = SaveToBitmap256(tBI256, lptr, sFileName)
         End Select
      
         DeleteObject hDib
      End If
      
   End If
   
End Function
Private Function pbCreate2ColourDIBSection(ByRef hDib As Long, ByRef tBI As BITMAPINFO2, ByRef lptr As Long, ByVal lWidth As Long, ByVal lheight As Long) As Boolean
Dim lScanSize As Long
Dim lHDC As Long
   lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   With tBI.bmiHeader
       .biSize = Len(tBI.bmiHeader)
       .biWidth = lWidth
       .biHeight = lheight
       .biPlanes = 1
       .biBitCount = 1
       .biCompression = BI_RGB
       lScanSize = lWidth \ 8
       lScanSize = lScanSize + lScanSize Mod 4
       .biSizeImage = lScanSize * .biHeight
   End With
   With tBI.bmiColors(1)
      .rgbBlue = &HFF: .rgbRed = &HFF: .rgbGreen = &HFF
   End With
   hDib = CreateDIBSection2( _
           lHDC, _
           tBI, _
           DIB_RGB_COLORS, _
           lptr, _
           0, 0)
   pbCreate2ColourDIBSection = (hDib <> 0)
   DeleteDC lHDC
End Function
Private Function pbCreate16ColourDIBSection(ByRef hDib As Long, ByRef tBI As BITMAPINFO16, ByRef lptr As Long, ByVal lWidth As Long, ByVal lheight As Long) As Boolean
Dim lScanSize As Long
Dim lHDC As Long
Dim i As Long
   lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   With tBI.bmiHeader
       .biSize = Len(tBI.bmiHeader)
       .biWidth = lWidth
       .biHeight = lheight
       .biPlanes = 1
       .biBitCount = 4
       .biCompression = BI_RGB
       lScanSize = lWidth \ 2
       lScanSize = lScanSize + lScanSize Mod 4
       .biSizeImage = lScanSize * .biHeight
   End With
   Dim cP As New cPalette
   cP.Create16Colour
   For i = 0 To 15
      With tBI.bmiColors(i)
         .rgbBlue = cP.Red(i + 1)
         .rgbGreen = cP.Green(i + 1)
         .rgbRed = cP.Blue(i + 1)
      End With
   Next
   hDib = CreateDIBSection16( _
           lHDC, _
           tBI, _
           DIB_RGB_COLORS, _
           lptr, _
           0, 0)
   pbCreate16ColourDIBSection = (hDib <> 0)
   DeleteDC lHDC
End Function
Private Function pbCreate256ColourDIBSection(ByRef hDib As Long, ByRef tBI As BITMAPINFO256, ByRef lptr As Long, ByVal lWidth As Long, ByVal lheight As Long, Optional ByRef cP As cPalette = Nothing) As Boolean
Dim lScanSize As Long
Dim lHDC As Long
Dim i As Long
Dim iMax As Long
   lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   With tBI.bmiHeader
       .biSize = Len(tBI.bmiHeader)
       .biWidth = lWidth
       .biHeight = lheight
       .biPlanes = 1
       .biBitCount = 8
       .biCompression = BI_RGB
       lScanSize = (lWidth + lWidth Mod 4)
       .biSizeImage = lScanSize * .biHeight
   End With
   ' Halftone palette:
   If cP Is Nothing Then
      Set cP = New cPalette
      cP.CreateHalfTone
   End If
   iMax = 255
   If iMax >= cP.Count Then
      iMax = cP.Count - 1
   End If
   For i = 0 To iMax
      With tBI.bmiColors(i)
         .rgbBlue = cP.Blue(i + 1)
         .rgbGreen = cP.Green(i + 1)
         .rgbRed = cP.Red(i + 1)
      End With
   Next
   hDib = CreateDIBSection256( _
           lHDC, _
           tBI, _
           DIB_RGB_COLORS, _
           lptr, _
           0, 0)
   pbCreate256ColourDIBSection = (hDib <> 0)
   DeleteDC lHDC
End Function
      
      
Private Function SaveToBitmap256(ByRef tBI As cDIBSectionSave.BITMAPINFO256, ByVal lPtrBits As Long, ByVal sFileName As String)
Dim tBH As BITMAPFILEHEADER
Dim tRGBQ As RGBQUAD
Dim hFile As Long
Dim lBytesWritten As Long
Dim lSize As Long
Dim lR As Long
Dim bErr As Boolean
Dim hMem As Long, lptr As Long
Dim lErr As Long

   ' Prepare the BITMAPFILEHEADER
   With tBH
      .bfType = BITMAPTYPE
      .bfOffBits = 14 + Len(tBI)
      .bfSize = .bfOffBits + tBI.bmiHeader.biSizeImage
   End With
   hFile = CreateFile(sFileName, _
                 GENERIC_READ Or GENERIC_WRITE, _
                  ByVal 0&, _
                  ByVal 0&, _
                  CREATE_ALWAYS, _
                  FILE_ATTRIBUTE_NORMAL, _
                  0)
   lErr = Err.LastDllError
   If (hFile = INVALID_HANDLE_VALUE) Then
      ' error
      Err.Raise 17, App.EXEName & ".cDIBSection256", ApiError(lErr)
   Else
      
      ' Writing the BITMAPFILEINFOHEADER is somewhat painful
      ' due to non-byte alignment of structure...
      hMem = GlobalAlloc(GPTR, 14)
      lptr = GlobalLock(hMem)
      CopyMemory ByVal lptr, tBH.bfType, 2
      CopyMemory ByVal lptr + 2, tBH.bfSize, 4
      CopyMemory ByVal lptr + 6, 0&, 4
      CopyMemory ByVal lptr + 10, tBH.bfOffBits, 4
      lSize = 14
      lR = WriteFile(hFile, ByVal lptr, lSize, lBytesWritten, ByVal 0&)
      GlobalUnlock hMem
      GlobalFree hMem
      
      ' Add the BITMAPINFOHEADER and colour palette:
      bErr = FileErrHandler(lR, lSize, lBytesWritten)
      If Not bErr Then
         lSize = Len(tBI)
         lR = WriteFile(hFile, tBI, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      If Not bErr Then
         ' Its easy to write the bitmap data, though...
         lSize = tBI.bmiHeader.biSizeImage
         lR = WriteFile(hFile, ByVal lPtrBits, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      
      CloseHandle hFile
      SaveToBitmap256 = Not (bErr)
   End If

End Function

      
Private Function SaveToBitmap16(ByRef tBI As BITMAPINFO16, ByVal lPtrBits As Long, ByVal sFileName As String)
Dim tBH As BITMAPFILEHEADER
Dim tRGBQ As RGBQUAD
Dim hFile As Long
Dim lBytesWritten As Long
Dim lSize As Long
Dim lR As Long
Dim bErr As Boolean
Dim hMem As Long, lptr As Long
Dim lErr As Long

   ' Prepare the BITMAPFILEHEADER
   With tBH
      .bfType = BITMAPTYPE
      .bfOffBits = 14 + Len(tBI)
      .bfSize = .bfOffBits + tBI.bmiHeader.biSizeImage
   End With
   hFile = CreateFile(sFileName, _
                 GENERIC_READ Or GENERIC_WRITE, _
                  ByVal 0&, _
                  ByVal 0&, _
                  CREATE_ALWAYS, _
                  FILE_ATTRIBUTE_NORMAL, _
                  0)
   lErr = Err.LastDllError
   If (hFile = INVALID_HANDLE_VALUE) Then
      ' error
      Err.Raise 17, App.EXEName & ".cDIBSection256", ApiError(lErr)
   Else
      
      ' Writing the BITMAPFILEINFOHEADER is somewhat painful
      ' due to non-byte alignment of structure...
      hMem = GlobalAlloc(GPTR, 14)
      lptr = GlobalLock(hMem)
      CopyMemory ByVal lptr, tBH.bfType, 2
      CopyMemory ByVal lptr + 2, tBH.bfSize, 4
      CopyMemory ByVal lptr + 6, 0&, 4
      CopyMemory ByVal lptr + 10, tBH.bfOffBits, 4
      lSize = 14
      lR = WriteFile(hFile, ByVal lptr, lSize, lBytesWritten, ByVal 0&)
      GlobalUnlock hMem
      GlobalFree hMem
      
      ' Add the BITMAPINFOHEADER and colour palette:
      bErr = FileErrHandler(lR, lSize, lBytesWritten)
      If Not bErr Then
         lSize = Len(tBI)
         lR = WriteFile(hFile, tBI, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      If Not bErr Then
         ' Its easy to write the bitmap data, though...
         lSize = tBI.bmiHeader.biSizeImage
         lR = WriteFile(hFile, ByVal lPtrBits, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      
      CloseHandle hFile
      SaveToBitmap16 = Not (bErr)
   End If

End Function

Private Function SaveToBitmap2(ByRef tBI As BITMAPINFO2, ByVal lPtrBits As Long, ByVal sFileName As String)
Dim tBH As BITMAPFILEHEADER
Dim tRGBQ As RGBQUAD
Dim hFile As Long
Dim lBytesWritten As Long
Dim lSize As Long
Dim lR As Long
Dim bErr As Boolean
Dim hMem As Long, lptr As Long
Dim lErr As Long

   ' Prepare the BITMAPFILEHEADER
   With tBH
      .bfType = BITMAPTYPE
      .bfOffBits = 14 + Len(tBI)
      .bfSize = .bfOffBits + tBI.bmiHeader.biSizeImage
   End With
   hFile = CreateFile(sFileName, _
                 GENERIC_READ Or GENERIC_WRITE, _
                  ByVal 0&, _
                  ByVal 0&, _
                  CREATE_ALWAYS, _
                  FILE_ATTRIBUTE_NORMAL, _
                  0)
   lErr = Err.LastDllError
   If (hFile = INVALID_HANDLE_VALUE) Then
      ' error
      Err.Raise 17, App.EXEName & ".cDIBSection256", ApiError(lErr)
   Else
      
      ' Writing the BITMAPFILEINFOHEADER is somewhat painful
      ' due to non-byte alignment of structure...
      hMem = GlobalAlloc(GPTR, 14)
      lptr = GlobalLock(hMem)
      CopyMemory ByVal lptr, tBH.bfType, 2
      CopyMemory ByVal lptr + 2, tBH.bfSize, 4
      CopyMemory ByVal lptr + 6, 0&, 4
      CopyMemory ByVal lptr + 10, tBH.bfOffBits, 4
      lSize = 14
      lR = WriteFile(hFile, ByVal lptr, lSize, lBytesWritten, ByVal 0&)
      GlobalUnlock hMem
      GlobalFree hMem
      
      ' Add the BITMAPINFOHEADER and colour palette:
      bErr = FileErrHandler(lR, lSize, lBytesWritten)
      If Not bErr Then
         lSize = Len(tBI)
         lR = WriteFile(hFile, tBI, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      If Not bErr Then
         ' Its easy to write the bitmap data, though...
         lSize = tBI.bmiHeader.biSizeImage
         lR = WriteFile(hFile, ByVal lPtrBits, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      
      CloseHandle hFile
      SaveToBitmap2 = Not (bErr)
   End If

End Function
Private Function ApiError(ByVal e As Long) As String
    Dim s As String, c As Long
    s = String(256, 0)
    c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS, _
                      0, e, 0&, s, Len(s), ByVal 0)
    If c Then ApiError = Left$(s, c)
End Function

Private Function FileErrHandler(ByVal lR As Long, ByVal lSize As Long, ByVal lBytes As Long) As Boolean
   If (lR = 0) Or Not (lSize = lBytes) Then
      'Err.Raise
      FileErrHandler = True
   End If
End Function


