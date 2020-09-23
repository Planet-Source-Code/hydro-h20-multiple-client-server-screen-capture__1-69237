VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl ucNetworkTree 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   ScaleHeight     =   3600
   ScaleWidth      =   5145
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   720
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.DriveListBox driDrives 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   1770
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   3122
      _Version        =   327682
      Indentation     =   0
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ucnetworktree.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ucnetworktree.ctx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ucnetworktree.ctx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ucnetworktree.ctx":1BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ucnetworktree.ctx":1ED8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ucnetworktree.ctx":2B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ucnetworktree.ctx":2E44
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ucnetworktree.ctx":315E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ucnetworktree.ctx":3DB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ucnetworktree.ctx":4AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ucnetworktree.ctx":5738
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ucNetworkTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' API's
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Const imgComputer = 1
Private Const imgNetwork = 2
Private Const imgHDShare = 10 '3
Private Const imgHD = 4
Private Const imgFolder = 5
Private Const imgFloppy = 8
Private Const imgCD = 7
Private Const imgRAM = 8
Private Const imgNetComp = 9

Private uc_ShowComputer As Boolean
Private uc_ShowNetwork As Boolean
Private uc_ShowPrinterShares As Boolean
Private uc_SelectedFolder As String
Private uc_ListPrintersOnly As Boolean
Private uc_AutoRefreshTree As Boolean ' only applicable if uc_ShowNetwork is true
Private uc_AutoRefreshFreq As Integer
Private uc_AutoLoadTree As Boolean
Private RefreshCount As Long
Private PreviousNet As String

Public Event FolderChanged()

'==========================================================================
' For Networks
'==========================================================================
'CNetworkEnum
'
'Utility class to perform network enumeration functions.
'Can locate all servers, printers and shares on a given network.

Private Const RESOURCE_CONNECTED As Long = &H1&
Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCE_REMEMBERED As Long = &H3&

Private Const RESOURCEDISPLAYTYPE_DIRECTORY& = &H9
Private Const RESOURCEDISPLAYTYPE_DOMAIN& = &H1
Private Const RESOURCEDISPLAYTYPE_FILE& = &H4
Private Const RESOURCEDISPLAYTYPE_GENERIC& = &H0
Private Const RESOURCEDISPLAYTYPE_GROUP& = &H5
Private Const RESOURCEDISPLAYTYPE_NETWORK& = &H6
Private Const RESOURCEDISPLAYTYPE_ROOT& = &H7
Private Const RESOURCEDISPLAYTYPE_SERVER& = &H2
Private Const RESOURCEDISPLAYTYPE_SHARE& = &H3
Private Const RESOURCEDISPLAYTYPE_SHAREADMIN& = &H8

Private Const RESOURCETYPE_ANY As Long = &H0&
Private Const RESOURCETYPE_DISK As Long = &H1&
Private Const RESOURCETYPE_PRINT As Long = &H2&
Private Const RESOURCETYPE_UNKNOWN As Long = &HFFFF&

Private Const RESOURCEUSAGE_ALL As Long = &H0&
Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&
Private Const RESOURCEUSAGE_CONTAINER As Long = &H2&
Private Const RESOURCEUSAGE_RESERVED As Long = &H80000000

Private Const NO_ERROR = 0
Private Const ERROR_MORE_DATA = 234
Private Const RESOURCE_ENUM_ALL As Long = &HFFFF

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    pLocalName As Long
    pRemoteName As Long
    pComment As Long
    pProvider As Long
End Type

Private Type NETRESOURCE_EXTENDED
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    sLocalName As String
    sRemoteName As String
    sComment As String
    sProvider As String
End Type

'WNet API resources
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function VarPtrAny Lib "vb40032.dll" Alias "VarPtr" (lpObject As Any) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (lpTo As Any, lpFrom As Any, ByVal lLen As Long)
Private Declare Sub CopyMemByPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpTo As Long, ByVal lpFrom As Long, ByVal lLen As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Filtered immediate access storage
Private sUserName As String
Private sMachineName As String
Private sServerList As String
Private sPrinterList As String
Private sShareList As String
Private sDirectoryList As String
Private sDomainList As String
Private sFileList As String
Private sGenericList As String
Private sGroupList As String
Private sNetworkList As String
Private sRootList As String
Private sShareAdminList As String

'Sets the resource types that will be enumerated (disk,printer etc)
'Private lResType As Long
Private Const lResType = 0

'Limitation
Private Const MAX_RESOURCES = 256
Private Const NOT_A_CONTAINER = -1


'==========================================================================
' Properties
'==========================================================================
Public Property Let ShowComputer(ByRef new_ShowComputer As Boolean)
    uc_ShowComputer = new_ShowComputer
    PropertyChanged "ShowComputer"
    If Not IsDesignTime Then
        Build_Tree
    End If
End Property

Public Property Get ShowComputer() As Boolean
    ShowComputer = uc_ShowComputer
End Property

Public Property Let ShowNetwork(ByRef new_ShowNetwork As Boolean)
    uc_ShowNetwork = new_ShowNetwork
    PropertyChanged "ShowNetwork"
    If Not IsDesignTime Then
        Build_Tree
    End If
End Property

Public Property Get ShowNetwork() As Boolean
    ShowNetwork = uc_ShowNetwork
End Property

Public Property Get SelectedFolder() As String
    SelectedFolder = uc_SelectedFolder
End Property

Public Property Let SelectedFolder(ByRef new_SelectedFolder As String)
    ' traverse tree as far as it can
    Call Search_Folder_Tree(new_SelectedFolder)
    PropertyChanged "SelectedFolder"
End Property

Public Property Get ShowPrinterShares() As Boolean
    ShowPrinterShares = uc_ShowPrinterShares
End Property

Public Property Let ShowPrinterShares(ByRef new_ShowPrinterShares As Boolean)
    uc_ShowPrinterShares = new_ShowPrinterShares
    PropertyChanged "ShowPrinterShares"
    If Not IsDesignTime Then
        Build_Tree
    End If
End Property

Public Property Get ListPrintersOnly() As Boolean
    ListPrintersOnly = uc_ListPrintersOnly
End Property

Public Property Let ListPrintersOnly(ByRef new_ListPrintersOnly As Boolean)
    uc_ListPrintersOnly = new_ListPrintersOnly
    PropertyChanged "ListPrintersOnly"
End Property

Public Property Let AutoRefreshTree(ByRef new_AutoRefreshTree As Boolean)
    uc_AutoRefreshTree = new_AutoRefreshTree
    PropertyChanged "AutoRefreshTree"
    If Not IsDesignTime Then
        If uc_AutoRefreshTree Then
            Timer2.Enabled = True
        Else
            Timer2.Enabled = False
            RefreshCount = 0
        End If
    End If
End Property

Public Property Get AutoRefreshTree() As Boolean
    AutoRefreshTree = uc_AutoRefreshTree
End Property

Public Property Let AutoRefreshFrequency(ByRef new_AutoFreq As Integer)
    If new_AutoFreq < 1 Or new_AutoFreq > 120 Then Exit Property

    uc_AutoRefreshFreq = new_AutoFreq
    PropertyChanged "AutoRefreshFrequency"
End Property

Public Property Get AutoRefreshFrequency() As Integer
    AutoRefreshFrequency = uc_AutoRefreshFreq
End Property

Public Property Let AutoLoadTree(ByRef new_AutoLoad As Boolean)
    uc_AutoLoadTree = new_AutoLoad
    PropertyChanged "AutoLoadTree"
    If Not IsDesignTime And uc_AutoLoadTree Then
        Call Build_Tree
    End If
End Property

Public Property Get AutoLoadTree() As Boolean
    AutoLoadTree = uc_AutoLoadTree
End Property

Public Sub Load_Tree()
    On Local Error Resume Next
    Call Build_Tree
End Sub

Private Sub Timer2_Timer()
    Dim netPATH As String
    
    On Local Error Resume Next
    
    If uc_ShowNetwork = False Then Exit Sub
    
    RefreshCount = RefreshCount + 1

    If RefreshCount < (uc_AutoRefreshFreq * 60) Then Exit Sub
    
    RefreshCount = 0
    
    Call Get_Network_Details
    
    If PreviousNet = sShareList Then Exit Sub
    Debug.Print "P=" & PreviousNet
    Debug.Print "s=" & sShareList
    
    Call Build_Tree(True)
    
    If Len(uc_SelectedFolder) <> 0 Then
        Call Search_Folder_Tree(uc_SelectedFolder)
    End If
        
End Sub

Private Sub TreeView1_Expand(ByVal Node As ComctlLib.Node)
    Dim i As Integer
    Dim strRelative As String
    Dim l As String
    Dim folderLIST() As String
    Dim folderCNT As Integer
    Dim subFOUND As Boolean
    
        

    On Local Error Resume Next
        
    If Node.Child.Text = "" Then
        MousePointer = vbHourglass
                
        TreeView1.Nodes.Remove Node.Child.Index
        strRelative = Node.Key
                
        ReDim folderLIST(0 To 0)
        folderCNT = 0
        l = Dir$(strRelative & "*.*", vbDirectory)
        While Len(l) <> 0
            If l <> "." And l <> ".." Then
                If (GetAttr(strRelative & l) And vbDirectory) = vbDirectory Then
                    folderCNT = folderCNT + 1
                    ReDim Preserve folderLIST(0 To folderCNT)
                    folderLIST(folderCNT) = l
                End If
            End If
            l = Dir$
        Wend
        
        For i = 1 To folderCNT
            TreeView1.Nodes.Add strRelative, tvwChild, strRelative & LCase$(folderLIST(i)) & "\", folderLIST(i), imgFolder
            
            ' see if any other folders down the tree
            l = Dir$(strRelative & folderLIST(i) & "\*.*", vbDirectory)
            subFOUND = False
            While Len(l) <> 0 And subFOUND = False
                If l <> "." And l <> ".." Then
                    If (GetAttr(strRelative & folderLIST(i) & "\" & l) And vbDirectory) = vbDirectory Then
                        subFOUND = True
                    End If
                End If
                l = Dir$
            Wend
            
            If subFOUND Then
                TreeView1.Nodes.Add strRelative & LCase$(folderLIST(i)) & "\", tvwChild
            End If
        Next i
        
        MousePointer = vbDefault
        DoEvents
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
    On Local Error Resume Next
    
    If Node.Key = "MYCOMPUTER" Or Node.Key = "NETWORK" Then
        uc_SelectedFolder = ""
    Else
        uc_SelectedFolder = Node.Key
    End If
    RaiseEvent FolderChanged
        
End Sub

Private Sub UserControl_Initialize()
    uc_ShowComputer = True
    uc_ShowNetwork = True
    uc_ShowPrinterShares = True
    uc_ListPrintersOnly = False
    uc_AutoRefreshTree = False
    uc_AutoRefreshFreq = 1
    uc_AutoLoadTree = False
    If Not IsDesignTime Then
        If uc_AutoLoadTree Then Call Build_Tree
    End If
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    TreeView1.Move 0, 0, UserControl.Width, UserControl.Height
    Label1.Move (UserControl.Width - Label1.Width) \ 2, (UserControl.Height - Label1.Height) \ 2
    
End Sub

Private Sub Build_Tree(Optional NetAlreadyLoaded As Boolean = False)
    Dim i As Integer
    Dim strPATH As String
    Dim intDriveType As Integer
    Dim iconID As Integer
    Dim strMACHINE As String
    Dim strSHARE As String
    Dim prevMACHINE As String
    Dim a As Integer
    Dim sOK As Boolean
    Dim netROOTShown As Boolean
    Dim tmr2 As Boolean
    
    On Local Error Resume Next
    
    tmr2 = Timer2.Enabled
    Timer2.Enabled = False
    MousePointer = vbHourglass
    TreeView1.Nodes.Clear
    TreeView1.Visible = False
    Label1.Visible = True
    DoEvents
    
    If uc_ShowComputer Then
        ' add root node for mycomputer
        TreeView1.Nodes.Add , , "MYCOMPUTER", "My Computer", imgComputer
        
        For i = 0 To driDrives.ListCount - 1
            strPATH = LCase$(left(driDrives.List(i), 1)) & ":\"
            
            intDriveType = GetDriveType(strPATH)

            Select Case intDriveType
                Case 1: iconID = imgHD
                Case 2: iconID = imgFloppy
                Case 3: iconID = imgHD
                Case 4: iconID = imgHDShare
                Case 5: iconID = imgCD
                Case 6: iconID = imgRAM
            End Select

            TreeView1.Nodes.Add "MYCOMPUTER", tvwChild, strPATH, driDrives.List(i), iconID
            TreeView1.Nodes.Add strPATH, tvwChild
        Next i
    End If
    
    
    If uc_ShowNetwork Then
        If NetAlreadyLoaded = False Then
            Call Get_Network_Details
        End If
        
        If Len(sShareList) > 0 Then
            PreviousNet = sShareList
            
            TreeView1.Nodes.Add , , "NETWORK", "Network Places", imgNetwork

            List1.Clear
            i = InStr(1, sShareList, "|", vbTextCompare)
            While i <> 0
                List1.AddItem left$(sShareList, i - 1)
                
                sShareList = right$(sShareList, Len(sShareList) - i)
                i = InStr(1, sShareList, "|", vbTextCompare)
            Wend
            
            prevMACHINE = ""
            For i = 0 To (List1.ListCount - 1)
                strMACHINE = List1.List(i)
                strMACHINE = right$(strMACHINE, Len(strMACHINE) - 2) ' removes \\
                a = InStr(1, strMACHINE, "\", vbTextCompare)
                strSHARE = right$(strMACHINE, Len(strMACHINE) - a)
                strMACHINE = left$(strMACHINE, a - 1)
                strMACHINE = LCase$(strMACHINE)
                
                If strMACHINE <> prevMACHINE Then
                    netROOTShown = False
'                    TreeView1.Nodes.Add "NETWORK", tvwChild, "\\" & strMACHINE & "\", UCase$(strMACHINE), imgNetComp
                            
                    prevMACHINE = strMACHINE
                End If
                              
                If uc_ListPrintersOnly Then
                    If InStr(1, strSHARE, "[Printer]", vbTextCompare) > 0 Then
                        If netROOTShown = False Then
                            TreeView1.Nodes.Add "NETWORK", tvwChild, "\\" & strMACHINE & "\", UCase$(strMACHINE), imgNetComp
                            netROOTShown = True
                        End If
                    
                        strSHARE = left$(strSHARE, Len(strSHARE) - 9)
                        TreeView1.Nodes.Add "\\" & strMACHINE & "\", tvwChild, "\\" & strMACHINE & "\" & LCase$(strSHARE) & "\", strSHARE, 11
                    End If
                Else
                    If netROOTShown = False Then
                        TreeView1.Nodes.Add "NETWORK", tvwChild, "\\" & strMACHINE & "\", UCase$(strMACHINE), imgNetComp
                        netROOTShown = True
                    End If
                
                    If InStr(1, strSHARE, "[Printer]", vbTextCompare) = 0 Then
                        TreeView1.Nodes.Add "\\" & strMACHINE & "\", tvwChild, "\\" & strMACHINE & "\" & LCase$(strSHARE) & "\", strSHARE, imgHDShare
                        TreeView1.Nodes.Add "\\" & strMACHINE & "\" & LCase$(strSHARE) & "\", tvwChild
                    ElseIf uc_ShowPrinterShares Then
                        strSHARE = left$(strSHARE, Len(strSHARE) - 9)
                        TreeView1.Nodes.Add "\\" & strMACHINE & "\", tvwChild, "\\" & strMACHINE & "\" & LCase$(strSHARE) & "\", strSHARE, 11
                    End If
                End If
            Next i
        End If
    End If
    
    Label1.Visible = False
    TreeView1.Visible = True
    Timer2.Enabled = tmr2
    MousePointer = vbDefault
    DoEvents
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        uc_ShowComputer = .ReadProperty("ShowComputer", True)
        uc_ShowNetwork = .ReadProperty("ShowNetwork", True)
        uc_ShowPrinterShares = .ReadProperty("ShowPrinterShares", True)
        uc_ListPrintersOnly = .ReadProperty("ListPrintersOnly", False)
        uc_AutoRefreshTree = .ReadProperty("AutoRefreshTree", False)
        uc_AutoRefreshFreq = .ReadProperty("AutoRefreshFrequency", 1)
        uc_AutoLoadTree = .ReadProperty("AutoLoadTree", False)
    End With
    If Not IsDesignTime Then
        If uc_AutoLoadTree Then
            Call Build_Tree
        End If
        If uc_AutoRefreshTree Then
            Timer2.Enabled = True
        Else
            Timer2.Enabled = False
        End If
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ShowComputer", uc_ShowComputer, True
        .WriteProperty "ShowNetwork", uc_ShowNetwork, True
        .WriteProperty "ShowPrinterShares", uc_ShowPrinterShares, True
        .WriteProperty "ListPrintersOnly", uc_ListPrintersOnly, False
        .WriteProperty "AutoRefreshTree", uc_AutoRefreshTree, False
        .WriteProperty "AutoRefreshFrequency", uc_AutoRefreshFreq, 1
        .WriteProperty "AutoLoadTree", uc_AutoLoadTree, False
    End With
End Sub

Private Sub Search_Folder_Tree(ByVal sFOLDER As String)
    Dim rootNODE As String
    Dim i As Integer
    Dim NodeSearch As ComctlLib.Node
    Dim NodeOld As ComctlLib.Node
    Dim cFolder As String
    
    On Local Error Resume Next
    
    sFOLDER = LCase$(sFOLDER)
    
    If left$(sFOLDER, 2) = "\\" Then
        ' NETWORK
        If uc_ShowNetwork = False Then Exit Sub
        rootNODE = "NETWORK"
        i = 3
    Else
        ' MYCOMPUTER
        If uc_ShowComputer = False Then Exit Sub
        rootNODE = "MYCOMPUTER"
        i = 1
    End If


    Set NodeSearch = TreeView1.Nodes(rootNODE)
    Set NodeOld = NodeSearch
    Call TreeView1_Expand(NodeSearch)
    
    i = InStr(i, sFOLDER, "\", vbTextCompare)
    While i <> 0
        cFolder = left$(sFOLDER, i)
        ' check to see if folder exists
        Err.Clear
        Set NodeOld = NodeSearch
        Set NodeSearch = TreeView1.Nodes(cFolder)
        If Err.Number <> 0 Then
            ' cannot navigate to here exit
            Set NodeSearch = NodeOld
            GoTo ExitSearch
        End If
        NodeSearch.EnsureVisible
        Call TreeView1_Expand(NodeSearch)
        i = InStr(i + 1, sFOLDER, "\", vbTextCompare)
    Wend
    
ExitSearch:
    Call TreeView1_NodeClick(NodeSearch)
End Sub


'=================================================================
' Network stuff
'=================================================================
Private Sub Get_Network_Details()
Dim bFirstTime As Boolean
Dim lReturn As Long
Dim hEnum As Long
Dim lCount As Long
Dim lMin As Long
Dim lLength As Long
Dim l As Long
Dim lBufferSize As Long
Dim lLastIndex As Long

'Stores the results of the enumeration
Dim uNetApi(0 To MAX_RESOURCES) As NETRESOURCE
Dim uNet() As NETRESOURCE_EXTENDED

bFirstTime = True
Do
  'Create an enumeration using the required resource type
  If bFirstTime Then
    lReturn = WNetOpenEnum(RESOURCE_GLOBALNET, lResType, RESOURCEUSAGE_ALL, ByVal 0&, hEnum)
    bFirstTime = False
  Else
    If uNet(lLastIndex).dwUsage And RESOURCEUSAGE_CONTAINER Then
      lReturn = WNetOpenEnum(RESOURCE_GLOBALNET, lResType, RESOURCEUSAGE_ALL, uNet(lLastIndex), hEnum)
    Else
      lReturn = NOT_A_CONTAINER
      hEnum = 0
    End If
    lLastIndex = lLastIndex + 1
  End If
  
  'Make sure that we have a good enumeration
  If lReturn = NO_ERROR Then
    lCount = RESOURCE_ENUM_ALL
    'Work through the enumeration until we run out
    Do
      lBufferSize = UBound(uNetApi) * Len(uNetApi(0)) / 2
      lReturn = WNetEnumResource(hEnum, lCount, uNetApi(0), lBufferSize)
      If lCount > 0 Then
        ReDim Preserve uNet(0 To lMin + lCount - 1) As NETRESOURCE_EXTENDED
        For l = 0 To lCount - 1
          
          'Each Resource will appear here as uNet(i)
          uNet(lMin + l).dwScope = uNetApi(l).dwScope
          uNet(lMin + l).dwType = uNetApi(l).dwType
          uNet(lMin + l).dwDisplayType = uNetApi(l).dwDisplayType
          uNet(lMin + l).dwUsage = uNetApi(l).dwUsage
          
          'Get the name
          If uNetApi(l).pLocalName Then
            lLength = lstrlen(uNetApi(l).pLocalName)
            uNet(lMin + l).sLocalName = Space$(lLength)
            CopyMem ByVal uNet(lMin + l).sLocalName, ByVal uNetApi(l).pLocalName, lLength
          End If
          
          
          'Get the remote name
          If uNetApi(l).pRemoteName Then
            lLength = lstrlen(uNetApi(l).pRemoteName)
            uNet(lMin + l).sRemoteName = Space$(lLength)
            CopyMem ByVal uNet(lMin + l).sRemoteName, ByVal uNetApi(l).pRemoteName, lLength
          End If
          
          'Get any comment associated with it
          If uNetApi(l).pComment Then
            lLength = lstrlen(uNetApi(l).pComment)
            uNet(lMin + l).sComment = Space$(lLength)
            CopyMem ByVal uNet(lMin + l).sComment, ByVal uNetApi(l).pComment, lLength
          End If
          
          'Get the provider information
          If uNetApi(l).pProvider Then
            lLength = lstrlen(uNetApi(l).pProvider)
            uNet(lMin + l).sProvider = Space$(lLength)
            CopyMem ByVal uNet(lMin + l).sProvider, ByVal uNetApi(l).pProvider, lLength
          End If
        Next l
      End If
      lMin = lMin + lCount
    Loop While lReturn = ERROR_MORE_DATA
  End If
  
  'Check if we have a successfully opened Enumeration
  If hEnum Then
      l = WNetCloseEnum(hEnum)
  End If
  
Loop While lLastIndex < lMin

'Decode the results
Call DecodeEnum(uNet)

End Sub

'Decodes the network array into a useful set of values
Private Sub DecodeEnum(uNet() As NETRESOURCE_EXTENDED)
  
Dim l As Long

If UBound(uNet) > 0 Then
  'Get some local information
  Call DecodeLocalInfo
  
  'Parse the network enumeration
  For l = 0 To UBound(uNet)
    'TODO: Include comments? uNet(l).sComment
    Select Case uNet(l).dwDisplayType
    Case RESOURCEDISPLAYTYPE_DIRECTORY&
      sDirectoryList = sDirectoryList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_DOMAIN
      sDomainList = sDomainList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_FILE
      sFileList = sFileList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_GENERIC
      sGenericList = sGenericList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_GROUP
      sGroupList = sGroupList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_NETWORK&
      sNetworkList = sNetworkList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_ROOT&
      sRootList = sRootList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_SERVER
      sServerList = sServerList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_SHARE
      If uNet(l).dwType = 1 Then
        sShareList = sShareList + uNet(l).sRemoteName + "|"
      Else
        sShareList = sShareList + uNet(l).sRemoteName + "[Printer]|"
      End If
    Case RESOURCEDISPLAYTYPE_SHAREADMIN&
      sShareAdminList = sShareAdminList + uNet(l).sRemoteName + "|"
    End Select
  Next l
End If

End Sub

Private Sub DecodeLocalInfo()

On Error Resume Next

'Create a buffer
sUserName = String(255, Chr(0))

'Get the username
Call GetUserName(sUserName, 255)

'Strip the rest of the buffer
sUserName = left(sUserName, InStr(sUserName, Chr(0)) - 1)

'Create a buffer
sMachineName = String(255, Chr(0))
Call GetComputerName(sMachineName, 255)

'Remove the unnecessary chr(0)'s
sMachineName = left$(sMachineName, InStr(1, sMachineName, Chr(0)) - 1)

End Sub

Private Function IsDesignTime() As Boolean
    On Local Error GoTo errDesign
    
    If Ambient.UserMode Then
    End If
    
    If Not Ambient.UserMode Then GoTo errDesign
    
    IsDesignTime = False
    Exit Function
    
errDesign:
    IsDesignTime = True

End Function

Public Sub Collapse_Tree()
    Dim i As Long
    
    On Local Error Resume Next
    
    For i = (TreeView1.Nodes.Count - 1) To 0 Step -1
        If TreeView1.Nodes.Item(i).Expanded Then
            TreeView1.Nodes.Item(i).Expanded = False
        End If
    Next i

End Sub
