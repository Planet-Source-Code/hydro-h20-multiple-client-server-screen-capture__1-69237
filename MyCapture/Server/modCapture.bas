Attribute VB_Name = "modCapture"
Option Compare Text
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long


Public ActiveRecordings As Integer
Public myCaptureFolders() As String
Public AVIConvertingNow As Boolean
Public ViewerID As Integer
Public ViewMode As Integer
Public CaptureSlotDuration As Integer
Public CompName As String

Public Const LocationCount = 16

Public Type cssHEADERtype
    frameCount As Long
    Started As Date
    Stopped As Date
End Type

Public Type cssFRAMEtype
    jpgSize As Long
End Type

Public Type CaptureSettings
    InUse As Boolean
    Available As Boolean
    Capturing As Boolean
    CaptureFrame As Long
End Type

Public CaptureLocations(1 To 16) As CaptureSettings

Public Sub test()
    Dim cHEADER As cssHEADERtype
    
    cHEADER = Read_Capture_Header("C:\0~File\VBProjects\MyCapture\NewCode\Server\a64x2\NewCapture.mcp")
    

End Sub

Public Function Read_Capture_Header(capFile As String) As cssHEADERtype
    Dim cssHEADER As cssHEADERtype
    Dim jFN As Integer

    jFN = FreeFile
    Open capFile For Binary Access Read As #jFN
    Get #jFN, , cssHEADER
    Close #jFN
    
    Read_Capture_Header = cssHEADER
        
End Function

Public Function Close_Capture(capFile As String, capMACHINE As String) As Boolean
    Dim newFilename As String
    Dim cssHEADER As cssHEADERtype
    Dim newFolder As String
    
    On Local Error Resume Next
    
    cssHEADER = Read_Capture_Header(capFile)
    newFolder = App.Path & "\" & capMACHINE & "\" & Format$(cssHEADER.Started, "yyyymmdd")
    MkDir newFolder
    newFilename = newFolder & "\Capture~" & Format$(cssHEADER.Started, "HHMMSS") & "~" & Format$(cssHEADER.Stopped, "HHMMSS") & ".mcp"
    
    Name capFile As newFilename

End Function

Public Function Write_Frame(capFile As String, capFrame As String, capMACHINE As String) As Boolean
    Dim cssHEADER As cssHEADERtype
    Dim cssFRAME As cssFRAMEtype
    Dim jFN As Integer
    Dim capDT As Date
    Dim tFILE As String
    Dim tTIME As String
    Dim tDATE As String
    Dim jpgFN As Integer
    Dim buffer As String

    If CaptureSlotDuration > 0 Then
        If Len(Dir$(capFile)) <> 0 Then
            cssHEADER = Read_Capture_Header(capFile)
            If cssHEADER.frameCount = CaptureSlotDuration Then
                Call Close_Capture(capFile, capMACHINE)
            End If
        End If
    End If

    tFILE = Left$(capFrame, Len(capFrame) - 4)
    tTIME = Right$(tFILE, 6)
    tFILE = Left$(tFILE, Len(tFILE) - 7)
    tDATE = Right$(tFILE, 8)
    tTIME = Left$(tTIME, 2) & ":" & Mid$(tTIME, 3, 2) & ":" & Right$(tTIME, 2)
    tDATE = Left$(tDATE, 2) & "/" & Mid$(tDATE, 3, 2) & "/" & Right$(tDATE, 4)
    capDT = CDate(tDATE & " " & tTIME)
    
    
    jFN = FreeFile
    If Len(Dir$(capFile)) = 0 Then
        Open capFile For Binary Access Read Write As #jFN
        cssHEADER.Started = capDT
        cssHEADER.Stopped = capDT
        cssHEADER.frameCount = 0
        Put #jFN, , cssHEADER
        Close #jFN
    End If
    
    Open capFile For Binary Access Read Write As #jFN
    Get #jFN, , cssHEADER
    
    cssHEADER.frameCount = cssHEADER.frameCount + 1
    cssHEADER.Stopped = capDT
    Put #jFN, 1, cssHEADER
    
    cssFRAME.jpgSize = FileLen(capFrame)
    Put #jFN, FileLen(capFile) + 1, cssFRAME
    
    
    jpgFN = FreeFile
    buffer = Space$(cssFRAME.jpgSize)
    
    ' 1st need to load jpg into a buffer
    Open capFrame For Binary Access Read As #jpgFN
    Get #jpgFN, , buffer
    Close #jpgFN
    
    Put #jFN, , buffer
    Close #jFN
    
    buffer = ""

End Function

Public Sub Main()
    Dim i As Integer
    Dim LocationName As String
    Dim ret As Boolean
    
    On Local Error Resume Next
    
    INIFile = App.Path & "\Capture.ini"
    
    frmSplash.Show
    DoEvents
    
    CaptureSlotDuration = ReadINI("General", "CaptureSlotDuration", INIFile)
    CompName = GetComputerName
    
    Load mdiMAIN
    DoEvents
    
    mdiMAIN.cboLocation.Clear
    For i = 1 To LocationCount
        LocationName = ReadINI("Location" & i, "Name", INIFile)
        If LocationName <> "Not In Use" Then
            ret = Add_Capture_Window(i, LocationName)
            If ret = False Then
                mdiMAIN.cboLocation.AddItem LocationName & " (OFF)"
            Else
                mdiMAIN.cboLocation.AddItem LocationName
            End If
        Else
            ret = Add_Capture_Window(i, LocationName)
        End If
    Next i
    DoEvents
    
    frmSplash.lblProgress.Caption = "Loading Network Settings..."
    frmSplash.SetFocus
    DoEvents
    Load frmSelectPath
    
    ViewMode = 3
    ViewerID = 1
    Call Arrange_Capture_Forms(ViewMode, ViewerID)
        
    mdiMAIN.Frame1.Enabled = True
    mdiMAIN.imgSINGLE.Enabled = True
    mdiMAIN.imgFour.Enabled = True
    mdiMAIN.imgAll.Enabled = True
    mdiMAIN.Frame2.Enabled = True
    mdiMAIN.Frame3.Enabled = True
    mdiMAIN.cboLocation.Enabled = True
    mdiMAIN.cmdStartRec.Enabled = True
    mdiMAIN.cmdStopRec.Enabled = True

    With mdiMAIN.cmdPlayBack
        .Enabled = True
        .Hover = True
    End With
    
    With mdiMAIN.cmdConfigure
        .Enabled = True
        .Hover = True
    End With
    
    Unload frmSplash
    mdiMAIN.WindowState = 2
    mdiMAIN.cboLocation.ListIndex = 0
    mdiMAIN.Show
    DoEvents
        
End Sub

Public Function Add_Capture_Window(CaptureID As Integer, CaptureName As String) As Boolean
    Dim vFRM As New frmViewer
    
    On Local Error Resume Next
    
    Load vFRM

    frmSplash.lblProgress = "Checking Capture Machine ... " & CaptureName
    frmSplash.SetFocus
    DoEvents
        
    Add_Capture_Window = Setup_Capture_Form(CaptureID, CaptureName, vFRM)
    
End Function

Public Function Setup_Capture_Form(CaptureID As Integer, CaptureName As String, capForm As frmViewer) As Boolean
    Dim CapPath As String
    Dim A As Integer
    Dim Machine As String
    Dim l As String
    Dim pingR As Boolean
    
    On Local Error Resume Next
    
    capForm.Label1.Caption = Format$(CaptureID, "00") & " .. " & CaptureName
    capForm.Tag = "~VIEWER~"
    capForm.timerAuto.Enabled = False
    capForm.Timer1.Enabled = False
    
    If CaptureName <> "Not In Use" Then
        CapPath = ReadINI("Location" & CaptureID, "CaptureAppPath", INIFile)
        
        Machine = CapPath
        Machine = Right$(Machine, Len(Machine) - 2)
        A = InStr(1, Machine, "\", vbTextCompare)
        Machine = Left$(Machine, A - 1)
        capForm.lblMachine = Machine
        capForm.lblMachine = CaptureName
               
        MkDir App.Path & "\" & CaptureName
        
        pingR = Ping(Machine)
        If pingR = False Then
            capForm.Label1.Caption = capForm.Label1.Caption & " (OFF)"
        Else
            ' machine exists, but does the capture path
            Err.Clear
            l = Dir$(CapPath & "\*.*")
            If Err.Number <> 0 Or Len(l) = 0 Then
                capForm.Label1.Caption = capForm.Label1.Caption & " (OFF)"
            Else
                Err.Clear
                capForm.File1.Path = ReadINI("Location" & CaptureID, "CapturePath", INIFile) & "\"
                If Err.Number <> 0 Then
                    capForm.Label1.Caption = capForm.Label1.Caption & " (OFF)"
                Else
                    capForm.File1.Pattern = "*.jpg"
                    capForm.Timer1.Enabled = True
                    
                    If CBool(ReadINI("Location" & CaptureID, "AutoStart", INIFile)) Then
                        capForm.timerAuto.Enabled = True
                    End If
                    Setup_Capture_Form = True
                End If
            End If
        End If
    End If

End Function

Public Sub Arrange_Capture_Forms(showMODE As Integer, showID As Integer)
    Dim UseableWidth As Long
    Dim UseableHeight As Long
    Dim vWidth As Long
    Dim vHeight As Long
    Dim i As Integer
    Dim A As Integer
    Dim vID As Integer
    Dim lPos As Long
    Dim tPOS As Long
    Dim fCNT As Integer
    
    ' showMODE 's
    ' 1 - full screen
    ' 2 - 4 windows
    ' 3 - 16 windows
    
    ' if mode 1, show id will be the window to show
    ' all other windows to be main invisible
    
    ' if mode 2, show id will be the 1st window id
    ' show four windows, e.g. showid = 1,    windows 1,2,3,4
    '                         showid = 5,    windows 5,6,7,8
    '                         showid = 9,    windows 9,10,11,12
    '                         showid = 13,   windows 13,14,15,16
    
    ' if mode3, the show id is irrevalant, all windows will be shown
    
    
    UseableHeight = mdiMAIN.Height - 575
    UseableWidth = mdiMAIN.Width - mdiMAIN.Picture2.Width - 200
    
    Select Case showMODE
        Case 1
            vWidth = UseableWidth
            vHeight = UseableHeight
            A = showID
        Case 2
            vWidth = UseableWidth \ 2
            vHeight = UseableHeight \ 2
            A = showID + 3
        Case 3
            vWidth = UseableWidth \ 4
            vHeight = UseableHeight \ 4
            A = 16
    End Select
    
    For i = 0 To (Forms.Count - 1)
        If Forms(i).Tag = "~VIEWER~" Then
            vID = Left$(Forms(i).Label1, 2)
            
            If vID >= showID And vID <= A Then
                Forms(i).Left = lPos
                Forms(i).Top = tPOS
                Forms(i).Width = vWidth
                Forms(i).Height = vHeight
                
                fCNT = fCNT + 1
                
                Select Case showMODE
                    Case 1
                        ' do nothing
                    Case 2
                        If fCNT = 2 Then
                            lPos = 0
                            tPOS = tPOS + vHeight
                        Else
                            lPos = lPos + vWidth
                        End If
                    Case 3
                        If fCNT = 4 Or fCNT = 8 Or fCNT = 12 Then
                            lPos = 0
                            tPOS = tPOS + vHeight
                        Else
                            lPos = lPos + vWidth
                        End If
                End Select
                Forms(i).Visible = True
            Else
                Forms(i).Visible = False
            End If
        End If
    Next i

   
End Sub

Public Function Seconds_To_Text(ByVal pSECS As Long) As String
    Dim h As Integer
    Dim m As Integer
    Dim s As Integer
    Dim hms As String
    
    If pSECS >= 3600 Then
        h = Int(pSECS / 3600)
        pSECS = pSECS - (3600 * h)
    End If
    
    If pSECS >= 60 Then
        m = Int(pSECS / 60)
        pSECS = pSECS - (60 * m)
    End If
    
    s = pSECS
    
    If s <> 0 Then
        hms = s & "s"
    End If
    
    If m <> 0 Then
        If Len(hms) <> 0 Then hms = " " & hms
        hms = m & "m" & hms
    End If
    
    If h <> 0 Then
        If Len(hms) <> 0 Then hms = " " & hms
        hms = h & "h" & hms
    End If
    
    Seconds_To_Text = hms
    
End Function

Public Function WaitABit(WaitInSecs As Double, EventsOn As Boolean, Optional fileCheck As String) As Boolean
    Dim start As Double
    Dim Delay As Double
    Dim rt As Double
    
    On Local Error Resume Next
    
    Delay = 0
    start = GetTickCount / 1000
    While Delay < WaitInSecs
        rt = GetTickCount / 1000
        If rt < start Then
            start = rt
        End If
        Delay = Delay + (rt - start)
        start = rt
        If Len(fileCheck) <> 0 Then
            If Len(Dir$(fileCheck)) = 0 Then Delay = WaitInSecs + 1
        End If
        If EventsOn Then DoEvents
    Wend
End Function

Public Function GetFileName(pfName As String) As String
    Dim i As Integer
    
    On Local Error Resume Next

    For i = Len(pfName) To 1 Step -1
        If Mid$(pfName, i, 1) = "\" Then Exit For
        GetFileName = Mid$(pfName, i, 1) & GetFileName
    Next i
    
End Function

Public Function Replace_Text(searchSTRING As String, searchFOR As String, replaceWITH As String) As String
    Dim newSTRING As String
    Dim i As Integer
    Dim lPART As String
    Dim rPART As String
    
    On Local Error Resume Next
    
    newSTRING = searchSTRING
    i = InStr(1, newSTRING, searchFOR, vbBinaryCompare)
    While i <> 0
        lPART = Left$(newSTRING, i - 1)
        rPART = Right$(newSTRING, Len(newSTRING) - ((i - 1) + Len(searchFOR)))
        newSTRING = lPART & replaceWITH & rPART
        i = InStr(i + Len(replaceWITH), newSTRING, searchFOR, vbBinaryCompare)
    Wend
    Replace_Text = newSTRING
    
End Function

Public Function Start_Capture(ByVal capMACHINE As String) As Boolean
    Dim i As Integer
    Dim capString As String
    Dim StartCapFN As Integer
    Dim StartCapFile As String
    Dim dummyFile As String
    
    On Local Error Resume Next
    
    If InStr(1, capMACHINE, " (OFF)", vbTextCompare) > 0 Then
        Exit Function
    End If
    
    capMACHINE = Replace_Text(capMACHINE, " (OFF)", "")
    
    For i = 1 To 16
        If ReadINI("Location" & i, "Name", INIFile) = capMACHINE Then Exit For
    Next i
    
    capString = "startLeft=" & ReadINI("Location" & i, "startLeft", INIFile) & vbCrLf
    capString = capString & "startTop=" & ReadINI("Location" & i, "startTop", INIFile) & vbCrLf
    capString = capString & "srcWidth=" & ReadINI("Location" & i, "srcWidth", INIFile) & vbCrLf
    capString = capString & "srcHeight=" & ReadINI("Location" & i, "srcHeight", INIFile) & vbCrLf
    capString = capString & "Rate=" & ReadINI("Location" & i, "rate", INIFile) & vbCrLf
    capString = capString & "targetWidth=" & ReadINI("Location" & i, "targetWidth", INIFile) & vbCrLf
    capString = capString & "targetHeight=" & ReadINI("Location" & i, "targetHeight", INIFile) & vbCrLf
    capString = capString & "Quality=" & ReadINI("Location" & i, "CompressionMode", INIFile) & vbCrLf
    capString = capString & "Compression=" & ReadINI("Location" & i, "CompressionRATE", INIFile) & vbCrLf
    capString = capString & "DTStamp=" & ReadINI("Location" & i, "DTStamp", INIFile) & vbCrLf
    capString = capString & "Convert256=" & ReadINI("Location" & i, "Convert256", INIFile) & vbCrLf
    capString = capString & "CaptureMouse=" & ReadINI("Location" & i, "CaptureMouse", INIFile)
    
    StartCapFile = ReadINI("Location" & i, "CaptureStart", INIFile)
    dummyFile = App.Path & "\" & Format$(Date, "ddmmyyyy") & Format$(Time$, "HHMMSS") & ".dat"
    StartCapFN = FreeFile
    Open dummyFile For Output As #StartCapFN
    Print #StartCapFN, capString
    Close #StartCapFN
    
    Name dummyFile As StartCapFile
    
End Function

Public Function Stop_Capture(ByVal capMACHINE As String) As Boolean
    Dim i As Integer
    Dim StopCapFN As Integer
    Dim StopCapFile As String
    Dim dummyFile As String
    
    On Local Error Resume Next
    
    If InStr(1, capMACHINE, " (OFF)", vbTextCompare) > 0 Then
        Exit Function
    End If
    
    capMACHINE = Replace_Text(capMACHINE, " (OFF)", "")
    
    For i = 1 To 16
        If ReadINI("Location" & i, "Name", INIFile) = capMACHINE Then Exit For
    Next i
    
    StopCapFile = ReadINI("Location" & i, "CaptureStop", INIFile)
    dummyFile = App.Path & "\" & Format$(Date, "ddmmyyyy") & Format$(Time$, "HHMMSS") & ".dat"
    StopCapFN = FreeFile
    Open dummyFile For Output As #StopCapFN
    Close #StopCapFN
    
    Name dummyFile As StopCapFile
    
End Function

Public Function JustPath(Path As String) As String
    Dim Cnt As Integer
    
    On Local Error Resume Next
    
    Cnt = 1

    Do Until Mid$(Path, Len(Path) - Cnt, 1) = "\"
        Cnt = Cnt + 1
    Loop
    JustPath = Left$(Path, Len(Path) - Cnt)
End Function

Public Function Build_myCaptureFolders(Optional EventsOn As Boolean = False) As Boolean
    Dim i As Integer
    Dim LocationName As String
    Dim capCNT As Integer
    Dim l As String
    Dim folderLOC As String
    
    On Local Error Resume Next
    
    ReDim myCaptureFolders(0 To 0)
    
    For i = 1 To LocationCount
        LocationName = ReadINI("Location" & i, "Name", INIFile)
        If LocationName <> "Not In Use" Then
            folderLOC = App.Path & "\" & LocationName & "\"
            l = Dir$(folderLOC, vbDirectory)
            While Len(l) <> 0
                If l <> "." And l <> ".." Then
                    If (GetAttr(folderLOC & l) And vbDirectory) = vbDirectory Then
                        capCNT = capCNT + 1
                        ReDim Preserve myCaptureFolders(0 To capCNT)
                        myCaptureFolders(capCNT) = folderLOC & l
                    End If
                End If
                l = Dir$
                If EventsOn Then DoEvents
            Wend
        End If
    Next i
    
End Function

Public Function Remove_Old_Captures() As Boolean
    Dim i As Integer
    Dim capFolder As String
    Dim fDATE As String
    Dim cDAYS As Integer
    Dim dispFolder As String
    Dim A As Integer
    
    On Local Error Resume Next
    
    frmFolderDelete.Show
    DoEvents
    
    Call Build_myCaptureFolders(True)
    
    If UBound(myCaptureFolders) = 0 Then
        frmFolderDelete.Timer1.Enabled = False
        Unload frmFolderDelete
        Call TidyUp_LaVolpe_Gifs
        Exit Function
    End If
    
    cDAYS = ReadINI("General", "KeepCapturesFor", INIFile)
    
    For i = 1 To UBound(myCaptureFolders)
        dispFolder = myCaptureFolders(i)
         ' know I could use instrrev but want to keep it vb5 compatible
        For A = Len(dispFolder) To 1 Step -1
            If Mid$(dispFolder, A, 1) = "\" And A <> (Len(dispFolder) - 8) Then Exit For
        Next A
        dispFolder = Right$(dispFolder, Len(dispFolder) - A)
        frmFolderDelete.Label1 = dispFolder
        frmFolderDelete.Label1.Refresh
        DoEvents
        
        capFolder = Right$(myCaptureFolders(i), 8)
        fDATE = CDate(Right$(capFolder, 2) & "/" & Mid$(capFolder, 5, 2) & "/" & Left$(capFolder, 4))
        If CDate(fDATE) < Date - cDAYS Then
            frmFolderDelete.Label1 = dispFolder & "...Deleting"
            frmFolderDelete.Label1.Refresh
            DoEvents
            myDelTree myCaptureFolders(i)
        End If
    Next i
        
    frmFolderDelete.Timer1.Enabled = False
    Unload frmFolderDelete
    
    Call TidyUp_LaVolpe_Gifs
    
End Function

Public Function Convert_All_To_AVI() As Boolean
    Dim i As Integer
    Dim mcpFILES() As String
    Dim mcpCNT As Integer
    Dim l As String
        
    On Local Error Resume Next
    
    Call Build_myCaptureFolders
    
    If UBound(myCaptureFolders) = 0 Then Exit Function
    
    ReDim mcpFILES(0 To 0)
    
    For i = 1 To UBound(myCaptureFolders)
        l = Dir$(myCaptureFolders(i) & "\*.mcp")
        While Len(l) <> 0
            mcpCNT = mcpCNT + 1
            ReDim Preserve mcpFILES(0 To mcpCNT)
            mcpFILES(mcpCNT) = myCaptureFolders(i) & "\" & l
            l = Dir$
        Wend
    Next i
    
    If mcpCNT = 0 Then Exit Function
    
    
    frmConvert.Show
    frmConvert.Timer1.Enabled = True
    DoEvents
    
    For i = 1 To mcpCNT
        frmConvert.lblFILE = mcpFILES(i)
        AVIConvertingNow = True
        DoEvents
        Call frmConvert.cmdEXPORT_Click
        While AVIConvertingNow
            DoEvents
        Wend
    Next i
    
    frmConvert.Timer1.Enabled = False
    Unload frmConvert
    
End Function

Public Sub myDelTree(ByVal vDir As Variant)
    Dim FSO, FS

    On Local Error Resume Next
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FS = FSO.deletefolder(vDir, True)
    Set FSO = Nothing
End Sub

Public Function Create_Header_Frame(fCAP As Form, capMACHINE As String, capDATE As String, capFROM As String, capTO As String) As Boolean
    Dim capText As String
    Dim fontSIZE As Integer
    Dim fontWIDTH As Long
    Dim yPOS As Long
    Dim xPOS As Long
    Dim capSPACER As Long
    Dim fontHEIGHT As Long
    
    On Local Error Resume Next
    
    ' 1st size picture to 1st frame of capture
    fCAP.Picture1 = LoadPicture(App.Path & "\frame.jpg")
    
    ' set background all black
    fCAP.Picture1 = LoadPicture("")
    
    fCAP.Picture1.Cls
    fCAP.Picture1.BackColor = vbBlack
    fCAP.Picture1.FillColor = vbBlack
    fCAP.Picture1.FillStyle = 1

    ' determine largest text
    If Len(capMACHINE) > Len(capText) Then capText = capMACHINE
    If Len(capDATE) > Len(capText) Then capText = capDATE
    If Len(capFROM & " - " & capTO) > Len(capText) Then capText = capFROM & " - " & capTO
    
    ' determine largest font size
    fCAP.Picture1.FontBold = True
    
    fontSIZE = 24
    fCAP.Picture1.fontSIZE = fontSIZE
    fontWIDTH = fCAP.Picture1.TextWidth(capText)
    If fontWIDTH <= fCAP.Picture1.Width Then GoTo fontSizeSelected
    
    fontSIZE = 18
    fCAP.Picture1.fontSIZE = fontSIZE
    fontWIDTH = fCAP.Picture1.TextWidth(capText)
    If fontWIDTH <= fCAP.Picture1.Width Then GoTo fontSizeSelected
        
    fontSIZE = 14
    fCAP.Picture1.fontSIZE = fontSIZE
    fontWIDTH = fCAP.Picture1.TextWidth(capText)
    If fontWIDTH <= fCAP.Picture1.Width Then GoTo fontSizeSelected
        
    fontSIZE = 12
    fCAP.Picture1.fontSIZE = fontSIZE
    fontWIDTH = fCAP.Picture1.TextWidth(capText)
    If fontWIDTH <= fCAP.Picture1.Width Then GoTo fontSizeSelected
    
    fontSIZE = 10
    fCAP.Picture1.fontSIZE = fontSIZE
    fontWIDTH = fCAP.Picture1.TextWidth(capText)
    If fontWIDTH <= fCAP.Picture1.Width Then GoTo fontSizeSelected

    fontSIZE = 8
    fCAP.Picture1.fontSIZE = fontSIZE

fontSizeSelected:
    ' determine top start and the space between each line
    fontHEIGHT = fCAP.Picture1.TextHeight(capText) * 3
    capSPACER = (fCAP.Picture1.Height - fontHEIGHT) \ 4
    yPOS = capSPACER

    ' print the capMACHINE
    xPOS = (fCAP.Picture1.Width - fCAP.Picture1.TextWidth(capMACHINE)) \ 2
    fCAP.Picture1.CurrentX = xPOS
    fCAP.Picture1.CurrentY = yPOS
    fCAP.Picture1.Print capMACHINE
    yPOS = yPOS + capSPACER
    
    ' print date
    xPOS = (fCAP.Picture1.Width - fCAP.Picture1.TextWidth(capDATE)) \ 2
    fCAP.Picture1.CurrentX = xPOS
    fCAP.Picture1.CurrentY = yPOS
    fCAP.Picture1.Print capDATE
    yPOS = yPOS + capSPACER
    
    ' print date
    xPOS = (fCAP.Picture1.Width - fCAP.Picture1.TextWidth(capFROM & " - " & capTO)) \ 2
    fCAP.Picture1.CurrentX = xPOS
    fCAP.Picture1.CurrentY = yPOS
    fCAP.Picture1.Print capFROM & " - " & capTO
        
    fCAP.Picture1.Refresh

    SavePicture fCAP.Picture1.Image, App.Path & "\header.bmp"
    

End Function

Public Function frmONTOP(otForm As Form, ONTop As Boolean, Optional iLeft, Optional iTOP, Optional iWidth, Optional iHeight) As Boolean
    Dim OnTopNotOnTop As Long
    Dim rc
    On Local Error Resume Next
    If ONTop = True Then
        OnTopNotOnTop = -1
    Else
        OnTopNotOnTop = -2
    End If
    If IsMissing(iLeft) Then
        With otForm
            iLeft = .Left / Screen.TwipsPerPixelX
            iTOP = .Top / Screen.TwipsPerPixelY
            iWidth = .Width / Screen.TwipsPerPixelX
            iHeight = .Height / Screen.TwipsPerPixelY
        End With
    End If
    rc = SetWindowPos(otForm.hwnd, OnTopNotOnTop, iLeft, iTOP, iWidth, iHeight, 0)
    frmONTOP = ONTop
End Function


Public Function IsFormVisible(FormName As String) As Boolean
    Dim i As Integer
    Dim fNAME As String
    Dim SName As String
    
    On Local Error Resume Next
    
    SName = Trim$(FormName)
    For i = 0 To Forms.Count - 1
        fNAME = Forms(i).Name
        If fNAME = SName Then
            If Forms(i).Visible = True Then IsFormVisible = True
            Exit For
        End If
    Next i
End Function

Public Sub TidyUp_LaVolpe_Gifs()
    Dim l As String
    
    On Local Error Resume Next
    
    l = Dir$(App.Path & "\~tLV*.gif")
    While Len(l) <> 0
        If Left$(l, 4) = "~tLV" Then
            Kill App.Path & "\" & l
        End If
        l = Dir$
    Wend
    

End Sub
