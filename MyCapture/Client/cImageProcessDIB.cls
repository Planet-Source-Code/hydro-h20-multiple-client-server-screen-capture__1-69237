VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColourReduceDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

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
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Event InitProgress(ByVal lMax As Long)
Public Event Progress(ByVal lPosition As Long)
Public Event Complete(ByVal lTimeMs As Long)

Private Type tPalItem
   rgbRed As Byte
   rgbGreen As Byte
   rgbBlue As Byte
   lColorRef As Long
End Type
   

Public Sub BlackAndWhite( _
        ByRef cFrom As cDIBSection, _
        ByRef cTo As cDIBSection _
    )
'
Dim pict() As Byte
Dim pict2() As Byte
Dim sa As SAFEARRAY2D
Dim sa2 As SAFEARRAY2D
Dim x As Long, y As Long
Dim i As Long, iCoeff As Long
Dim lTIme As Long
Dim xMax As Long, yMax As Long
Dim lError As Long
Dim lNew As Long
Dim iC As Long, iC2 As Long

    lTIme = timeGetTime()
       
    GrayScale cFrom
       
    ' have the local matrix point to bitmap pixels
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
    ' Pict now stores the To buffer
        
    ' have the local matrix point to bitmap pixels
    With sa2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cFrom.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cFrom.BytesPerScanLine
        .pvData = cFrom.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4
    ' Pict2 now stores the From buffer


    yMax = cTo.Height - 1
    xMax = (cTo.Width - 1) * 3
    
    RaiseEvent InitProgress(xMax)
    For x = 0 To xMax Step 3
        For y = 0 To yMax
            ' Apply a simple threshold:
            If (pict2(x, y) > 128) Then
                iC = iC + 1
                pict(x, y) = 255
                pict(x + 1, y) = 255
                pict(x + 2, y) = 255
                lError = (255 - pict2(x, y)) - 128
            Else
                iC2 = iC2 + 1
                pict(x, y) = 0
                pict(x + 1, y) = 0
                pict(x + 2, y) = 0
                ' Black tolerance:
                If (pict2(x, y) > 16) Then
                    lError = pict2(x, y)
                Else
                    lError = 0
                End If
            End If
            
            ' Diffuse the error:
            If (x < xMax - 3) Then
                lNew = pict2(x + 3, y) + (lError * 7) \ 16
                If (lNew > 255) Then lNew = 255
                If (lNew < 0) Then lNew = 0
                pict2(x + 3, y) = lNew
                pict2(x + 4, y) = lNew
                pict2(x + 5, y) = lNew
            End If
            If (y < yMax) Then
                For i = -3 To 3 Step 3
                    If (x + i) > 0 And (x + i) < xMax Then
                        Select Case i
                        Case -3
                            iCoeff = 3
                        Case 0
                            iCoeff = 5
                        Case 3
                            iCoeff = 1
                        End Select
                        lNew = pict2(x + i, y + 1) + (lError * iCoeff) \ 16
                        If (lNew > 255) Then lNew = 255
                        If (lNew < 0) Then lNew = 0
                        pict2(x + i, y + 1) = lNew
                        pict2(x + i + 1, y + 1) = lNew
                        pict2(x + i + 2, y + 1) = lNew
                    End If
                Next i
            End If
        Next y
        RaiseEvent Progress(x)
    Next x
    
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    CopyMemory ByVal VarPtrArray(pict2), 0&, 4
    
    cFrom.LoadPictureBlt cTo.hdc
    RaiseEvent Complete(timeGetTime - lTIme)
    
    
End Sub

Public Sub ApplyPalette( _
      ByRef cFrom As cDIBSection, _
      ByRef cTo As cDIBSection, _
      ByRef cPal As cPalette, _
      Optional ByVal bDiffuseError As Boolean = True _
   )
'
Dim pict() As Byte
Dim pict2() As Byte
Dim sa As SAFEARRAY2D
Dim sa2 As SAFEARRAY2D
Dim x As Long, y As Long
Dim i As Long, iCoeff As Long, j As Long
Dim lTIme As Long
Dim xMax As Long, yMax As Long
Dim lErrorRed As Long, lErrorBlue As Long, lErrorGreen As Long
Dim lNewRed As Long, lNewBlue As Long, lNewGreen As Long
Dim lIndex As Long
Dim iC As Long, iC2 As Long


    lTIme = timeGetTime()
       
    ' have the local matrix point to bitmap pixels
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
    ' Pict now stores the To buffer
        
    ' have the local matrix point to bitmap pixels
    With sa2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cFrom.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cFrom.BytesPerScanLine
        .pvData = cFrom.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4
    ' Pict2 now stores the From buffer
   
    yMax = cTo.Height - 1
    xMax = (cTo.Width - 1) * 3
    
    RaiseEvent InitProgress(xMax)
    For x = 0 To xMax Step 3
        For y = 0 To yMax
            ' Get nearest colour:
            
            lIndex = cPal.ClosestIndex(pict2(x + 2, y), pict2(x + 1, y), pict2(x, y))
                                 
            pict(x + 2, y) = cPal.Red(lIndex)
            pict(x + 1, y) = cPal.Green(lIndex)
            pict(x, y) = cPal.Blue(lIndex)
                        
            If bDiffuseError Then
               lErrorRed = -1 * (CLng(pict(x + 2, y)) - pict2(x + 2, y))
               lErrorGreen = -1 * (CLng(pict(x + 1, y)) - pict2(x + 1, y))
               lErrorBlue = -1 * (CLng(pict(x, y)) - pict2(x, y))
               
               ' Diffuse the error:
               'Debug.Print lErrorRed, lErrorGreen, lErrorBlue
               If Abs(lErrorRed) + Abs(lErrorGreen) + Abs(lErrorBlue) > 3 Then
                  If (x < xMax - 3) Then
                      lNewBlue = pict2(x + 3, y) + (lErrorBlue * 7) \ 16
                      lNewGreen = pict2(x + 4, y) + (lErrorGreen * 7) \ 16
                      lNewRed = pict2(x + 5, y) + (lErrorRed * 7) \ 16
                      Range lNewBlue, 0, 255
                      Range lNewGreen, 0, 255
                      Range lNewRed, 0, 255
                      pict2(x + 3, y) = lNewBlue
                      pict2(x + 4, y) = lNewGreen
                      pict2(x + 5, y) = lNewRed
                  End If
                  If (y < yMax) Then
                      For i = -3 To 3 Step 3
                          If (x + i) > 0 And (x + i) < xMax Then
                              Select Case i
                              Case -3
                                  iCoeff = 0
                              Case 0
                                  iCoeff = 4
                              Case 3
                                  iCoeff = 0
                              End Select
                              lNewBlue = pict2(x + i, y + 1) + (lErrorBlue * iCoeff) \ 16
                              lNewGreen = pict2(x + i + 1, y + 1) + (lErrorGreen * iCoeff) \ 16
                              lNewRed = pict2(x + i + 2, y + 1) + (lErrorRed * iCoeff) \ 16
                              Range lNewBlue, 0, 255
                              Range lNewGreen, 0, 255
                              Range lNewRed, 0, 255
                              pict2(x + i, y + 1) = lNewBlue
                              pict2(x + i + 1, y + 1) = lNewGreen
                              pict2(x + i + 2, y + 1) = lNewRed
                          End If
                      Next i
                  End If
               End If
            End If
        Next y
        RaiseEvent Progress(x)
    Next x
    
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    CopyMemory ByVal VarPtrArray(pict2), 0&, 4
    
    Debug.Print iC, iC2
    cFrom.LoadPictureBlt cTo.hdc
    RaiseEvent Complete(timeGetTime - lTIme)
    
    
End Sub
   
Private Sub Range( _
      ByRef lIn As Long, _
      ByVal lMin As Long, _
      ByVal lMax As Long _
   )
   If (lIn < lMin) Then
      lIn = lMin
   ElseIf (lIn > lMax) Then
      lIn = lMax
   End If
End Sub

Public Sub GrayScale( _
        ByRef cTo As cDIBSection _
    )
'
Dim bDib() As Byte
Dim x As Long, y As Long
Dim xMax As Long, yMax As Long
Dim bContinue As Boolean
Dim lB As Long, lG As Long, lR As Long
Dim lGray As Long
Dim lTIme As Long
Dim tSA As SAFEARRAY2D

    lTIme = timeGetTime()
        
    ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
        
    yMax = cTo.Height - 1
    xMax = cTo.Width - 1
    
    RaiseEvent InitProgress(xMax)
    For x = 0 To (xMax * 3) Step 3
        For y = 0 To yMax
            lB = bDib(x, y)
            lG = bDib(x + 1, y)
            lR = bDib(x + 2, y)
                
            'But now all people *should* use the most accurate, it means ITU standard:
            lGray = (222 * lR + 707 * lG + 71 * lB) / 1000
            
            bDib(x, y) = lGray
            bDib(x + 1, y) = lGray
            bDib(x + 2, y) = lGray
        Next y
        RaiseEvent Progress(x)
    Next x
    
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    
    RaiseEvent Complete(timeGetTime - lTIme)
    
End Sub
