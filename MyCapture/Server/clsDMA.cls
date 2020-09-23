VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'Class to implement direct memory read/writes for bitmaps.
'for 256 colour pictures, 1 byte = 1 pixel
'for 24bit true colour pictures, 3 bytes = 1 pixel where x*3=Blue, x*3+1=Green and x*3+2=Red
'for 16bit hi colour, 1 pixel = 2 bytes

Option Explicit
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As _
    Any, ByVal Bytes As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Type SCANLINE
    minX As Long
    maxX As Long
End Type


Private SA As SAFEARRAY2D
Private BMP As BITMAP
Private Data() As Byte
'local variable(s) to hold property value(s)
Private mvarBytesPerPixel As Byte 'local copy
Private mvarColor As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarUBoundX As Long 'local copy
Private mvarUBoundY As Long 'local copy

'local variable(s) to hold property value(s)
Private mvarErrorMsg As String 'local copy

Public Property Get ErrorMsg() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorMsg
    ErrorMsg = mvarErrorMsg
End Property

Public Property Get UBoundY() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UboundY
    UBoundY = mvarUBoundY
End Property

Public Property Get UBoundX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UboundX
    UBoundX = mvarUBoundX
End Property

Public Property Let color(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Color = 5
    mvarColor = vData
End Property

Public Property Get color() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Color
    color = mvarColor
End Property

Public Property Get BytesPerPixel() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BytesPerPixel
    BytesPerPixel = mvarBytesPerPixel
End Property

Public Function LoadPicArray(ByVal p As StdPicture) As Boolean
'returns true if function works.

If p.Handle Then
    If GetObjectAPI(p.Handle, Len(BMP), BMP) Then    'retrieve bitmap information about p
        If BMP.bmWidth Then
            mvarBytesPerPixel = BMP.bmWidthBytes \ BMP.bmWidth
            If (mvarBytesPerPixel > 0) And (mvarBytesPerPixel < 4) Then
                ' make the local matrix point to bitmap pixels
                With SA
                  .cbElements = 1
                  .cDims = 2
                  .Bounds(0).lLbound = 0
                  .Bounds(0).cElements = BMP.bmHeight
                  .Bounds(1).lLbound = 0
                  .Bounds(1).cElements = BMP.bmWidthBytes
                  .pvData = BMP.bmBits   'Set SA to point to the BMP bitmap data
                End With
                ' copy bitmap data into byte array
                CopyMemory ByVal VarPtrArray(Data), VarPtr(SA), 4
                mvarUBoundX = UBound(Data, 1) \ mvarBytesPerPixel
                mvarUBoundY = UBound(Data, 2)
                LoadPicArray = True
            Else
                mvarErrorMsg = "Bitmap bytes/pixel must be 1-3 instead of " & CStr(mvarBytesPerPixel)
            End If
        Else
            mvarErrorMsg = "Bitmap width cannot be zero!"
        End If
    Else
        mvarErrorMsg = "Unable to retrieve Bitmap"
    End If
Else
    mvarErrorMsg = "No picture was loaded!"
End If
Set p = Nothing
End Function

Public Sub GetData(A() As Byte)
'causes array "a" to point to picture data
'the user MUST call ReleaseData() when done!
CopyMemory ByVal VarPtrArray(A), VarPtr(SA), 4
End Sub

Public Sub ReleaseData(A() As Byte)
CopyMemory ByVal VarPtrArray(A), 0&, 4
End Sub

Public Sub Clear()
ZeroMemory Data(0, 0), (mvarUBoundX + 1) * mvarBytesPerPixel * (mvarUBoundY + 1)
End Sub

Public Sub DrawPixel(ByVal x&, ByVal Y&, Optional ByVal C& = -1)
Dim ndx&
If C = -1 Then C = mvarColor
Select Case mvarBytesPerPixel
    Case 1: Data(x, Y) = C And &HFF
    Case 2
        Data(x + x, Y) = (C \ 256) And &HFF
        Data(x + x + 1, Y) = C And &HFF
    Case 3
        ndx = x + x + x
        Data(ndx, Y) = (C \ 65536) And &HFF
        Data(ndx + 1, Y) = (C \ 256) And &HFF
        Data(ndx + 2, Y) = C And &HFF
    Case 4
        ndx = x + x + x + x
        Data(ndx, Y) = (C \ 16777216) And &HFF
        Data(ndx + 1, Y) = (C \ 65536) And &HFF
        Data(ndx + 2, Y) = (C \ 256) And &HFF
        Data(ndx + 3, Y) = C And &HFF
End Select
End Sub

Public Function ReadPixel(ByVal x&, ByVal Y&) As Long
Dim ndx&
ndx = x * BytesPerPixel
Select Case mvarBytesPerPixel
    Case 1: ReadPixel = Data(x, Y)
    Case 2: ReadPixel = Data(ndx, Y) * 256& + Data(ndx + 1, Y)
    Case 3: ReadPixel = (Data(ndx, Y) * 256& + Data(ndx + 1, Y)) * 256& + Data(ndx + 2, Y)
    Case 4: ReadPixel = ((Data(ndx, Y) * 256& + Data(ndx + 1, Y)) * 256& + Data(ndx + 2, Y)) * 256& + Data(ndx + 3, Y)
End Select
End Function


Public Sub GreyScale(Optional BackColor As Long = -1)
    Dim Y1 As Long
    Dim X1 As Long
    Dim C As Long
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
    Dim AveCol As Integer
    Dim A As Integer



    For Y1 = 0 To mvarUBoundY
       For X1 = 0 To mvarUBoundX
            C = ReadPixel(X1, Y1)
                            
            If C <> BackColor Then
                RGBfromLONG C, r, g, b
                
                AveCol = 0
                A = 0
                AveCol = AveCol + g
                A = A + 1
                If AveCol <= 0 Then AveCol = 0
                AveCol = (AveCol / A)
                               
                DrawPixel X1, Y1, RGB(AveCol, AveCol, AveCol)
            End If

        Next X1
    Next Y1

End Sub

Public Sub Darken(percent As Integer)
'Darken a picture
Dim newVal As Integer
Dim h As Integer, rRed As Integer, rBlue As Integer, rGreen As Integer
Dim W As Integer
Dim K As Integer
Dim C As Long
Dim icRed As Long
Dim icBlue As Long
Dim icGreen As Long
Dim opRed As Long
Dim opBlue As Long
Dim opGreen As Long

newVal = percent * -5

For h = 0 To mvarUBoundY
    For W = 0 To mvarUBoundX
        C = ReadPixel(W, h)
        RGBfromLONG C, rRed, rBlue, rGreen
        opRed = rRed
        opBlue = rBlue
        opGreen = rGreen
        rRed = rRed + newVal
        If rRed > -1 And icRed < 256 Then opRed = rRed
        
        rGreen = rGreen + newVal
        If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
        rBlue = rBlue + newVal
        If rBlue > -1 And rBlue < 256 Then opBlue = rBlue
        If rRed <> 1000 Then
            If opRed < 0 Then opRed = 0
            If opGreen < 0 Then opGreen = 0
            If opBlue < 0 Then opBlue = 0
           C = RGB(opRed, opGreen, opBlue)
           DrawPixel W, h, C
        End If
    Next W
Next h
End Sub

Public Sub Invert()
'Invert the image of a picturebox
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Integer, opBlue As Integer, opGreen As Integer

For h = 0 To mvarUBoundY
    For W = 0 To mvarUBoundX

        C = ReadPixel(W, h)
        RGBfromLONG C, opRed, opBlue, opGreen
        opRed = 255 - opRed
        opGreen = 255 - opGreen
        opBlue = 255 - opBlue
        C = RGB(opRed, opGreen, opBlue)
        DrawPixel W, h, C
    Next W
Next h
End Sub

Public Sub Flash(Optional FlashVal As Integer = 10)
    Dim Y1 As Long
    Dim X1 As Long
    Dim C As Long
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
    Dim K As Integer
    Dim op As Integer
    
    For Y1 = 0 To mvarUBoundY
        For X1 = 0 To mvarUBoundX
            For K = 0 To FlashVal
                C = ReadPixel(X1, Y1)
                RGBfromLONG C, r, b, g
                op = CoefCol(r, g, b)
                C = RGB(op, op, op)
                DrawPixel X1, Y1, C
            Next K
        Next X1
    Next Y1
End Sub

Private Function CoefCol(r As Integer, g As Integer, b As Integer) As Integer
CoefCol = (r + g + b) / 2.45
End Function

Private Function RGBfromLONG(LongCol As Long, rRed As Integer, rBlue As Integer, rGreen As Integer)
' Get The Red, Blue And Green Values Of A Colour From The Long Value
Dim Blue As Double
Dim Green As Double
Dim Red As Double

Blue = Fix((LongCol / 256) / 256)
Green = Fix((LongCol - ((Blue * 256) * 256)) / 256)
Red = Fix(LongCol - ((Blue * 256) * 256) - (Green * 256))

rRed = Red
rBlue = Blue
rGreen = Green

End Function

Public Sub DrawLine(ByVal X1&, ByVal Y1&, ByVal X2&, ByVal Y2&, Optional ByVal C& = -1)
'good old bresenhams line drawing algorithm
Dim dx&, dy&, x&, Y&, xinc1&, xinc2&, yinc1&, yinc2&, ndx&
Dim den&, num&, numadd&, numpixels&, curpixel&
Dim c1 As Byte, c2 As Byte, c3 As Byte, c4 As Byte

If C = -1 Then C = mvarColor
dx = Abs(X2 - X1)   '  // The difference between the x's
dy = Abs(Y2 - Y1)   '  // The difference between the y's
x = X1              '  // Start x off at the first pixel
Y = Y1              '  // Start y off at the first pixel

If (X2 >= X1) Then  '  // The x-values are increasing
  xinc1 = 1
  xinc2 = 1
Else                '  // The x-values are decreasing
  xinc1 = -1
  xinc2 = -1
End If

If (Y2 >= Y1) Then  '  // The y-values are increasing
  yinc1 = 1
  yinc2 = 1
Else                '  // The y-values are decreasing
  yinc1 = -1
  yinc2 = -1
End If

If (dx >= dy) Then  '  // There is at least one x-value for every y-value
  xinc1 = 0         '  // Don't change the x when numerator >= denominator
  yinc2 = 0         '  // Don't change the y for every iteration
  den = dx
  num = dx / 2
  numadd = dy
  numpixels = dx    '  // There are more x-values than y-values
Else                '  // There is at least one y-value for every x-value
  xinc2 = 0         '  // Don't change the x for every iteration
  yinc1 = 0         '  // Don't change the y when numerator >= denominator
  den = dy
  num = dy / 2
  numadd = dx
  numpixels = dy    '  // There are more y-values than x-values
End If
Select Case BytesPerPixel
    Case 4
        ndx = x * 4
        c1 = C And &HFF: c2 = (C \ 256) And &HFF: c3 = (C \ 65536) And &HFF: c4 = (C \ 16777216) And &HFF
        For curpixel = 0 To numpixels
          Data(ndx, Y) = c4 '// Draw the current pixel
          Data(ndx + 1, Y) = c3
          Data(ndx + 2, Y) = c2
          Data(ndx + 3, Y) = c1
          num = num + numadd ' // Increase the numerator by the top of the fraction
          If num >= den Then ' // Check if numerator >= denominator
            num = num - den '  // Calculate the new numerator value
            x = x + xinc1   '  // Change the x as appropriate
            Y = Y + yinc1   '  // Change the y as appropriate
          End If
          x = x + xinc2     '  // Change the x as appropriate
          Y = Y + yinc2     '  // Change the y as appropriate
        Next curpixel
    Case 3
        ndx = x * 3
        c1 = C And &HFF: c2 = (C \ 256) And &HFF: c3 = (C \ 65536) And &HFF
        For curpixel = 0 To numpixels
          Data(ndx, Y) = c3 '// Draw the current pixel
          Data(ndx + 1, Y) = c2
          Data(ndx + 2, Y) = c1
          num = num + numadd ' // Increase the numerator by the top of the fraction
          If num >= den Then ' // Check if numerator >= denominator
            num = num - den '  // Calculate the new numerator value
            x = x + xinc1   '  // Change the x as appropriate
            Y = Y + yinc1   '  // Change the y as appropriate
          End If
          x = x + xinc2     '  // Change the x as appropriate
          Y = Y + yinc2     '  // Change the y as appropriate
        Next curpixel
    Case 2
        c1 = C And &HFF: c2 = (C \ 256) And &HFF
        For curpixel = 0 To numpixels
          Data(x + x, Y) = c2 '// Draw the current pixel
          Data(x + x + 1, Y) = c1
          num = num + numadd ' // Increase the numerator by the top of the fraction
          If num >= den Then ' // Check if numerator >= denominator
            num = num - den '  // Calculate the new numerator value
            x = x + xinc1   '  // Change the x as appropriate
            Y = Y + yinc1   '  // Change the y as appropriate
          End If
          x = x + xinc2     '  // Change the x as appropriate
          Y = Y + yinc2     '  // Change the y as appropriate
        Next curpixel
    Case 1
        c1 = C And &HFF
        For curpixel = 0 To numpixels
          Data(x, Y) = c1    ' // Draw the current pixel
          num = num + numadd ' // Increase the numerator by the top of the fraction
          If num >= den Then ' // Check if numerator >= denominator
            num = num - den '  // Calculate the new numerator value
            x = x + xinc1   '  // Change the x as appropriate
            Y = Y + yinc1   '  // Change the y as appropriate
          End If
          x = x + xinc2     '  // Change the x as appropriate
          Y = Y + yinc2     '  // Change the y as appropriate
        Next curpixel
End Select
End Sub

Public Sub DrawCircle(ByVal X1&, ByVal Y1&, ByVal r&, Optional ByVal C& = -1)
'bresenhams circle drawing algorithm
Dim x As Long
Dim Y As Long
Dim d As Long

'Start the coordinates at the top of the circle (90 degrees)
x = 0: Y = r: d = 1 - r
'Move the x-coordinate to the right, until the x- and
'y-coordinates meet (signifying the point at 45 degrees).
Do While Y >= x
    DrawPixel X1 + x, Y1 + Y, C
    DrawPixel X1 + x, Y1 - Y, C
    DrawPixel X1 - x, Y1 + Y, C
    DrawPixel X1 - x, Y1 - Y, C
    DrawPixel X1 + Y, Y1 + x, C
    DrawPixel X1 + Y, Y1 - x, C
    DrawPixel X1 - Y, Y1 + x, C
    DrawPixel X1 - Y, Y1 - x, C
    If d < 0 Then
        d = d + (x + x) + 3
    Else
        d = d + 2 * (x - Y) + 5
        Y = Y - 1
    End If
    'Move to the next x-coordinate
    x = x + 1
Loop

End Sub

Private Sub Class_Terminate()
' MUST be called to free up memory
CopyMemory ByVal VarPtrArray(Data), 0&, 4
End Sub

Public Sub DeSpeckle(Optional ByVal X1& = 0, Optional ByVal Y1& = 0, Optional ByVal X2& = 0, Optional ByVal Y2& = 0)
Dim x&, Y&, C&, d&
If X2 = 0 Then X2 = UBound(Data, 1)
If Y2 = 0 Then Y2 = UBound(Data, 2)
For x = X1 + 1 To X2 - 1
    For Y = Y1 + 1 To Y2 - 1
        C = ReadPixel(x, Y)
        If C = 0 Then
            d = ReadPixel(x - 1, Y)
            If d Then
                C = C + d \ 4
                d = ReadPixel(x + 1, Y)
                If d Then
                    C = C + d \ 4
                    d = ReadPixel(x, Y + 1)
                    If d Then
                        C = C + d \ 4
                        d = ReadPixel(x, Y - 1)
                        If d Then
                            C = C + d \ 4
                            DrawPixel x, Y, C
                        End If
                    End If
                End If
            End If
        End If
    Next Y
Next x
End Sub

Public Sub FillPolygon(px() As Long, py() As Long, Optional ByVal FillColor& = -1)
'fills a convex polygon defined by the pts array with color
'pt list MUST be in order either clockwise or counterclockwise
Dim i&, npts&, miny&, maxy&, maxX&, minX&, start&, j&, x#, dx#, b#, ndx&
Dim sl() As SCANLINE

If FillColor = -1 Then FillColor = mvarColor
npts = UBound(px)
If UBound(py) <> npts Then Exit Sub
miny = py(0): maxX = px(0)
maxy = miny: minX = maxX
For i = 1 To npts
    If py(i) > maxy Then maxy = py(i)
    If py(i) < miny Then miny = py(i)
    If px(i) > maxX Then maxX = px(i)
    If px(i) < minX Then minX = px(i)
Next i
ReDim sl(miny To maxy)
For i = miny To maxy   'initialize with impossible values
    sl(i).maxX = minX
    sl(i).minX = maxX
Next i
start = npts
For i = 0 To npts
    'calculate all the points on the line between pts(start) and pts(i)
    If py(i) = py(start) Then 'horizontal line
        j = py(i)
        If px(i) > px(start) Then
            If px(i) > sl(j).maxX Then sl(j).maxX = px(i)
            If px(start) < sl(j).minX Then sl(j).minX = px(start)
        Else
            If px(start) > sl(j).maxX Then sl(j).maxX = px(start)
            If px(i) < sl(j).minX Then sl(j).minX = px(i)
        End If
    Else
        dx = (px(i) - px(start)) / (py(i) - py(start))
        b = px(i) - py(i) * dx
        If py(i) > py(start) Then
            For j = py(start) To py(i)
                x = dx * j + b
                If x > sl(j).maxX Then sl(j).maxX = x
                If x < sl(j).minX Then sl(j).minX = x
            Next j
        Else
            For j = py(i) To py(start)
                x = dx * j + b
                If x > sl(j).maxX Then sl(j).maxX = x
                If x < sl(j).minX Then sl(j).minX = x
            Next j
        End If
    End If
    start = i
Next i
Select Case mvarBytesPerPixel
Case 4
    For i = miny To maxy
        For j = sl(i).minX To sl(i).maxX
            ndx = j * 3
            Data(ndx, i) = (FillColor \ 16777216) And &HFF
            Data(ndx + 1, i) = (FillColor \ 65536) And &HFF
            Data(ndx + 2, i) = (FillColor \ 256) And &HFF
            Data(ndx + 3, i) = FillColor And &HFF
        Next j
    Next i
Case 3
    For i = miny To maxy
        For j = sl(i).minX To sl(i).maxX
            Data(j * 3, i) = (FillColor \ 65536) And &HFF
            Data(j * 3 + 1, i) = (FillColor \ 256) And &HFF
            Data(j * 3 + 2, i) = FillColor And &HFF
        Next j
    Next i
Case 2
    For i = miny To maxy
        For j = sl(i).minX To sl(i).maxX
            Data(j + j, i) = (FillColor \ 256) And &HFF
            Data(j + j + 1, i) = FillColor And &HFF
        Next j
    Next i
Case 1
    For i = miny To maxy
        For j = sl(i).minX To sl(i).maxX
            Data(j, i) = FillColor And &HFF
        Next j
    Next i
End Select
Erase sl
End Sub

Public Sub FloodFill(ByVal x&, ByVal Y&, Optional ByVal FillColor& = -1)
Dim tc&
If FillColor = -1 Then FillColor = mvarColor
tc = ReadPixel(x, Y) 'test colour
If tc = FillColor Then Exit Sub
Select Case mvarBytesPerPixel
    Case 4: FloodFill4 x, Y, tc, FillColor
    Case 3: FloodFill3 x, Y, tc, FillColor
    Case 2: FloodFill2 x, Y, tc, FillColor
    Case 1: FloodFill1 x, Y, tc, FillColor
End Select
End Sub

Private Sub FloodFill4(ByVal x&, ByVal Y&, TestColor&, FillColor&)
'starting at point x,y it will flood fill all points of that color
Dim xx&, yy&, maxX&, maxy&, nx&, ny&, i&, lnear&
Dim fR As Byte, fG As Byte, fB As Byte
Dim tR As Byte, tG As Byte, tB As Byte
Dim aPts() As POINTAPI
Dim npts As Long
Dim siz As Long
Dim bOK As Boolean

fR = (FillColor \ 65536) And &HFF
fG = (FillColor \ 256) And &HFF
fB = FillColor And &HFF
tR = (TestColor \ 65536) And &HFF
tG = (TestColor \ 256) And &HFF
tB = TestColor And &HFF
maxX = UBound(Data, 1): maxy = UBound(Data, 2)
siz = 1000
ReDim aPts(siz)
lnear = siz - 5
aPts(npts).x = x: aPts(npts).Y = Y
npts = 1
While npts
    npts = npts - 1
    xx = aPts(npts).x: yy = aPts(npts).Y
    Data(xx * 3, yy) = fB: Data(xx * 3 + 1, yy) = fG: Data(xx * 3 + 2, yy) = fR
    ny = yy
    For i = 1 To 4
        bOK = False
        Select Case i
            Case 1: If xx > 0 Then nx = (xx - 1) * 3: bOK = True
            Case 2: If xx < maxX Then nx = (xx + 1) * 3: bOK = True
            Case 3: If yy > 0 Then nx = xx * 3: ny = yy - 1: bOK = True
            Case 4: If yy < maxy Then ny = yy + 1: bOK = True
        End Select
        If bOK Then
            If Data(nx, ny) = tB Then
                If Data(nx + 1, ny) = tG Then
                    If Data(nx + 2, ny) = tR Then
                        aPts(npts).x = nx \ 3: aPts(npts).Y = ny
                        npts = npts + 1
                    End If
                End If
            End If
        End If
    Next i
    If npts > lnear Then  'if we are getting close to the limit of the array, increase the size
        siz = siz + 1000
        lnear = lnear + 1000
        ReDim Preserve aPts(siz)
    End If
Wend
Erase aPts
End Sub

Private Sub FloodFill3(ByVal x&, ByVal Y&, TestColor&, FillColor&)
'starting at point x,y it will flood fill all points of that color
Dim xx&, yy&, maxX&, maxy&, nx&, ny&, i&, lnear&
Dim fR As Byte, fG As Byte, fB As Byte
Dim tR As Byte, tG As Byte, tB As Byte
Dim aPts() As POINTAPI
Dim npts As Long
Dim siz As Long
Dim bOK As Boolean

fR = (FillColor \ 65536) And &HFF
fG = (FillColor \ 256) And &HFF
fB = FillColor And &HFF
tR = (TestColor \ 65536) And &HFF
tG = (TestColor \ 256) And &HFF
tB = TestColor And &HFF
maxX = UBound(Data, 1): maxy = UBound(Data, 2)
siz = 1000
ReDim aPts(siz)
lnear = siz - 5
aPts(npts).x = x: aPts(npts).Y = Y
npts = 1
While npts
    npts = npts - 1
    xx = aPts(npts).x: yy = aPts(npts).Y
    Data(xx * 3, yy) = fB: Data(xx * 3 + 1, yy) = fG: Data(xx * 3 + 2, yy) = fR
    ny = yy
    For i = 1 To 4
        bOK = False
        Select Case i
            Case 1: If xx > 0 Then nx = (xx - 1) * 3: bOK = True
            Case 2: If xx < maxX Then nx = (xx + 1) * 3: bOK = True
            Case 3: If yy > 0 Then nx = xx * 3: ny = yy - 1: bOK = True
            Case 4: If yy < maxy Then ny = yy + 1: bOK = True
        End Select
        If bOK Then
            If Data(nx, ny) = tB Then
                If Data(nx + 1, ny) = tG Then
                    If Data(nx + 2, ny) = tR Then
                        aPts(npts).x = nx \ 3: aPts(npts).Y = ny
                        npts = npts + 1
                    End If
                End If
            End If
        End If
    Next i
    If npts > lnear Then  'if we are getting close to the limit of the array, increase the size
        siz = siz + 1000
        lnear = lnear + 1000
        ReDim Preserve aPts(siz)
    End If
Wend
Erase aPts
End Sub

Private Sub FloodFill2(ByVal x&, ByVal Y&, TestColor&, FillColor&)
'starting at point x,y it will flood fill all points of that color
Dim xx&, yy&, maxX&, maxy&, nx&, ny&, i&, lnear&
Dim fc1 As Byte, fc2 As Byte, tc1 As Byte, tc2 As Byte
Dim aPts() As POINTAPI
Dim npts As Long
Dim siz As Long
Dim bOK As Boolean

fc1 = (FillColor \ 256) And &HFF
fc2 = FillColor And &HFF
tc1 = (TestColor \ 256) And &HFF
tc2 = TestColor And &HFF
maxX = UBound(Data, 1): maxy = UBound(Data, 2)
siz = 1000
ReDim aPts(siz)
lnear = siz - 5
aPts(npts).x = x: aPts(npts).Y = Y
npts = 1
While npts
    npts = npts - 1
    xx = aPts(npts).x: yy = aPts(npts).Y
    Data(xx + xx, yy) = fc1: Data(xx + xx + 1, yy) = fc2
    ny = yy
    For i = 1 To 4
        bOK = False
        Select Case i
            Case 1: If xx > 0 Then nx = (xx - 1) * 2: bOK = True
            Case 2: If xx < maxX Then nx = (xx + 1) * 2: bOK = True
            Case 3: If yy > 0 Then nx = xx * 2: ny = yy - 1: bOK = True
            Case 4: If yy < maxy Then ny = yy + 1: bOK = True
        End Select
        If bOK Then
            If Data(nx, ny) = tc1 Then
                If Data(nx + 1, ny) = tc2 Then
                    aPts(npts).x = nx \ 2: aPts(npts).Y = ny
                    npts = npts + 1
                End If
            End If
        End If
    Next i
    If npts > lnear Then  'if we are getting close to the limit of the array, increase the size
        siz = siz + 1000
        lnear = lnear + 1000
        ReDim Preserve aPts(siz)
    End If
Wend
Erase aPts
End Sub

Private Sub FloodFill1(ByVal x&, ByVal Y&, TestColor&, FillColor&)
'starting at point x,y it will flood fill all points of that color
Dim xx&, yy&, maxX&, maxy&, nx&, ny&, i&, lnear&
Dim fc As Byte, tc As Byte
Dim aPts() As POINTAPI
Dim npts As Long
Dim siz As Long
Dim bOK As Boolean

fc = FillColor And &HFF
tc = TestColor And &HFF
maxX = UBound(Data, 1): maxy = UBound(Data, 2)
siz = 1000
ReDim aPts(siz)
lnear = siz - 5
aPts(npts).x = x: aPts(npts).Y = Y
npts = 1
While npts
    npts = npts - 1
    xx = aPts(npts).x: yy = aPts(npts).Y
    Data(xx, yy) = fc
    ny = yy
    For i = 1 To 4
        bOK = False
        Select Case i
            Case 1: If xx > 0 Then nx = (xx - 1): bOK = True
            Case 2: If xx < maxX Then nx = (xx + 1): bOK = True
            Case 3: If yy > 0 Then nx = xx: ny = yy - 1: bOK = True
            Case 4: If yy < maxy Then ny = yy + 1: bOK = True
        End Select
        If bOK Then
            If Data(nx, ny) = tc Then
                aPts(npts).x = nx: aPts(npts).Y = ny
                npts = npts + 1
            End If
        End If
    Next i
    If npts > lnear Then  'if we are getting close to the limit of the array, increase the size
        siz = siz + 1000
        lnear = lnear + 1000
        ReDim Preserve aPts(siz)
    End If
Wend
Erase aPts
End Sub

