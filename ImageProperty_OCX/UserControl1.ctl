VERSION 5.00
Begin VB.UserControl ImageProperty 
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   Begin VB.CheckBox chkAttrib 
      BackColor       =   &H80000009&
      Caption         =   "Hidden"
      Height          =   195
      Index           =   0
      Left            =   3480
      TabIndex        =   29
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   0
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1181
         Left            =   120
         ScaleHeight     =   77
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   102
         TabIndex        =   8
         Top             =   120
         Width           =   1560
      End
      Begin VB.CheckBox chkAttrib 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   195
      End
      Begin VB.CheckBox chkAttrib 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1200
         TabIndex        =   6
         Top             =   3240
         Width           =   195
      End
      Begin VB.CheckBox chkAttrib 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   2160
         TabIndex        =   5
         Top             =   3240
         Width           =   195
      End
      Begin VB.CheckBox chkAttrib 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   3360
         TabIndex        =   4
         Top             =   3240
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   5280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   3
         Top             =   2880
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   385
         TabIndex        =   1
         Top             =   1440
         Width           =   5775
         Begin VB.Label lblLocation 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
            Height          =   735
            Left            =   75
            TabIndex        =   2
            Top             =   75
            Width           =   5655
         End
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Created                :"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified               :"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Accessed     :"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 00:00:00"
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   25
         Top             =   2400
         Width           =   1545
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 00:00:00"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   24
         Top             =   2640
         Width           =   1545
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 00:00:00"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   23
         Top             =   2880
         Width           =   1545
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Left            =   1800
         TabIndex        =   22
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Left            =   1800
         TabIndex        =   21
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lbSize 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size  "
         Height          =   195
         Left            =   1800
         TabIndex        =   20
         Top             =   600
         Width           =   390
      End
      Begin VB.Label lblDepth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Depth          "
         Height          =   195
         Left            =   1800
         TabIndex        =   19
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lbDim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dimension  "
         Height          =   195
         Left            =   1800
         TabIndex        =   18
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   2880
         TabIndex        =   17
         Top             =   120
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   2880
         TabIndex        =   15
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   2880
         TabIndex        =   14
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   2880
         TabIndex        =   13
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hidden"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   3240
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System"
         Height          =   195
         Left            =   1440
         TabIndex        =   11
         Top             =   3240
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Read Only"
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Top             =   3240
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Archive"
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   3240
         Width           =   540
      End
   End
End
Attribute VB_Name = "ImageProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type SHFILEINFO
    hIcon                                         As Long
    iIcon                                         As Long
    dwAttributes                                  As Long
    szDisplayName                                 As String * 260
    szTypeName                                    As String * 80
End Type
Public Enum Display
    Picture = 0
    Icon = 1
End Enum
#If False Then
Private Picture, Icon
#End If

'Private Const SHGFI_DISPLAYNAME               As Long = &H200 '  get display name

Private Const SHGFI_TYPENAME                  As Long = &H400 '  get type name
Private Const SHGFI_ICONLOCATION              As Long = &H1000
Private Const SHGFI_SYSICONINDEX              As Long = &H4000
Private Type TYPERECT
    Left                                          As Long
    Top                                           As Long
    Right                                         As Long
    Bottom                                        As Long
End Type
Private Enum Appearance
    Flat = 0
    HalfRaised = 1
    Raised = 2
    Sunken = 3
    Etched = 4
    Bump = 5
    Line = 6
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Flat, HalfRaised, Raised, Sunken, Etched, Bump, Line
#End If

Private Const BDR_RAISEDOUTER                 As Long = &H1
Private Const BDR_SUNKENOUTER                 As Long = &H2
Private Const BDR_RAISEDINNER                 As Long = &H4
Private Const BDR_SUNKENINNER                 As Long = &H8
Private Const EDGE_RAISED                     As Double = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
'Private Const EDGE_SUNKEN                     As Double = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED                     As Double = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP                       As Double = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const BF_LEFT                         As Long = &H1
Private Const BF_TOP                          As Long = &H2
Private Const BF_RIGHT                        As Long = &H4
Private Const BF_BOTTOM                       As Long = &H8
'Private Const BF_DIAGONAL                     As Long = &H10
'Private Const BF_MIDDLE                       As Long = &H800
'Private Const BF_SOFT                         As Long = &H1000
'Private Const BF_ADJUST                       As Long = &H2000
Private Const BF_FLAT                         As Long = &H4000
'Private Const BF_MONO                         As Long = &H8000
'Private Const BF_TOPLEFT                      As Double = (BF_TOP Or BF_LEFT)
'Private Const BF_TOPRIGHT                     As Double = (BF_TOP Or BF_RIGHT)
'Private Const BF_BOTTOMLEFT                   As Double = (BF_BOTTOM Or BF_LEFT)
'Private Const BF_BOTTOMRIGHT                  As Double = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT                         As Double = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Type SHFILEOPSTRUCT
    hwnd                                          As Long
    wFunc                                         As Long
    pFrom                                         As String
    pTo                                           As String
    fFlags                                        As Integer
    fAnyOperationsAborted                         As Long
    hNameMappings                                 As Long
    lpszProgressTitle                             As String    '  only used if FOF_SIMPLEPROGRESS
End Type
Private Const FILE_ATTRIBUTE_ARCHIVE          As Long = &H20
'Private Const FILE_ATTRIBUTE_COMPRESSED       As Long = &H800
'Private Const FILE_ATTRIBUTE_DIRECTORY        As Long = &H10
Private Const FILE_ATTRIBUTE_HIDDEN           As Long = &H2
'Private Const FILE_ATTRIBUTE_NORMAL           As Long = &H80
Private Const FILE_ATTRIBUTE_READONLY         As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM           As Long = &H4
'Private Const FILE_ATTRIBUTE_TEMPORARY        As Long = &H100
Private Const MAX_PATH                        As Integer = 260
Private Const DI_MASK                         As Long = &H1
Private Const DI_IMAGE                        As Long = &H2
Private Const DI_NORMAL                       As Double = DI_MASK Or DI_IMAGE
Private Type FileTime
    dwLowDateTime                                 As Long
    dwHighDateTime                                As Long
End Type
Private Type WIN32_FIND_DATA
    dwFilechkattrib                               As Long
    ftCreationTime                                As FileTime
    ftLastAccessTime                              As FileTime
    ftLastWriteTime                               As FileTime
    nFileSizeHigh                                 As Long
    nFileSizeLow                                  As Long
    dwReserved0                                   As Long
    dwReserved1                                   As Long
    cFileName                                     As String * MAX_PATH
    cAlternate                                    As String * 14
End Type
Private Type SYSTEMTIME
    wYear                                         As Integer
    wMonth                                        As Integer
    wDayOfWeek                                    As Integer
    wDay                                          As Integer
    wHour                                         As Integer
    wMinute                                       As Integer
    wSecond                                       As Integer
    wMilliseconds                                 As Integer
End Type

'Default Property Values:
Private Const m_def_FileName                  As String = ""
Private Const m_def_ForeColor                 As Integer = 0
Private Const m_def_BackColor                 As Long = vbButtonFace 'Integer = 0
Private Const m_def_Enabled                   As Integer = 0
Private Const m_def_BackStyle                 As Integer = 0
Private Const m_def_BorderStyle               As Integer = 0
Private Const m_def_AutoSize                  As Integer = 0

'Property Variables:
Private m_FileName                            As String
Private m_Backcolor                           As OLE_COLOR
Private DateCreated                           As String
Private DateModified                          As String
Private DateLastAccessed                      As String
Private FileAttribut                          As String
Private Location                              As String
Private TitleName                             As String
Private m_ForeColor                           As OLE_COLOR
Private m_Enabled                             As Boolean
Private m_Font                                As Font
Private m_BackStyle                           As Integer
Private m_BorderStyle                         As Integer
Private m_AutoSize                            As Boolean
Private m_DisplayFileAssociated               As Display
Private Const m_def_DisplayFileAssociated     As Long = Display.Picture

'Event Declarations:
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Paint()

Private PWidth                                As Integer
Private PHeight                               As Integer
Private Declare Function SHGetFileInfoA Lib "shell32" (ByVal pszPath As String, _
                                                       ByVal dwFileAttributes As Long, _
                                                       psfi As SHFILEINFO, _
                                                       ByVal cbFileInfo As Long, _
                                                       ByVal uFlags As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, _
                                                qrc As TYPERECT, _
                                                ByVal edge As Long, _
                                                ByVal grfFlags As Long) As Boolean
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, _
                                                                                     ByVal pszBuf As String, _
                                                                                     ByRef cchBuf As Long) As String
Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, _
                                                                                ByVal lpszTitle As String, _
                                                                                ByVal cbBuf As Integer) As Integer
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                                                              lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FileTime, _
                                                              lpSystemTime As SYSTEMTIME) As Long
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, _
                                                                                                 ByVal lpIconPath As String, _
                                                                                                 lpiIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, _
                                                  ByVal xLeft As Long, _
                                                  ByVal yTop As Long, _
                                                  ByVal hIcon As Long, _
                                                  ByVal cxWidth As Long, _
                                                  ByVal cyWidth As Long, _
                                                  ByVal istepIfAniCur As Long, _
                                                  ByVal hbrFlickerFreeDraw As Long, _
                                                  ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoSize() As Boolean

    AutoSize = m_AutoSize

End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)

    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture4,Picture4,-1,BackColor
Public Property Get BackColor() As OLE_COLOR

    BackColor = m_Backcolor 'Picture4.BackColor
    PaintControl Picture4, Etched, m_Backcolor, vbBlack, "", False
    PaintControl Picture3, Etched, m_Backcolor, vbBlack, "", False
 
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    m_Backcolor = New_BackColor
    PropertyChanged "BackColor"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer

    BackStyle = m_BackStyle

End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)

    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer

    BorderStyle = m_BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)

    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"

End Property

Public Sub CariInfoGambar()



Dim leng As Long
Dim f     As SHFILEINFO
Dim cType As CFileInfo

    On Error GoTo Pesan
    leng = Len(FileName)
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = ""
    lblSize.Caption = ""
    Set cType = New CFileInfo
    If Right$(FileName, 1) = "\" Then
        cType.FileName = FileName
    Else 'NOT RIGHT$(FILENAME,...
        cType.FileName = FileName
    End If
    cType.GetImageFileInfo
   
    With cType
    
        If .TypeOfImage <> UNKNOWN Then
            Label3.Caption = .Depth
          
            SHGetFileInfoA FileName, 0, f, leng, SHGFI_TYPENAME Or SHGFI_ICONLOCATION Or SHGFI_SYSICONINDEX
            Label1.Caption = f.szTypeName 'strType
            Label2.Caption = .Width & " x " & .Height
            PWidth = .Width
            PHeight = .Height
        End If
    End With 'cType
    
    GetFileInfo
Pesan:
    If Err.Number <> 0 Then
        'Path not Found
        If Err.Number = 53 Then
            Exit Sub
        End If
        MsgBox Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "CariInfoGambar"
        
    End If

End Sub



Public Property Get DisplayFileAssociated() As Display

    DisplayFileAssociated = m_DisplayFileAssociated

End Property

Public Property Let DisplayFileAssociated(ByVal New_DisplayFileAssociated As Display)

    m_DisplayFileAssociated = New_DisplayFileAssociated
    PropertyChanged "DisplayFileAssociated"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean

    Enabled = m_Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    m_Enabled = New_Enabled
    PropertyChanged "Enabled"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FileName() As String

    On Error GoTo Pesan
    FileName = m_FileName
Pesan:
    If Err.Number <> 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbInformation + vbOKOnly, "FILENAME"
       
    End If

End Property

Public Property Let FileName(ByVal New_FileName As String)

Dim leng As Long
Dim f As SHFILEINFO
    On Error GoTo Pesan
    m_FileName = New_FileName
    If LenB(FileName) Then
    
        CariInfoGambar
        GetFileInfo
        GetIcon
        leng = Len(m_FileName)
        SHGetFileInfoA FileName, 0, f, leng, SHGFI_TYPENAME Or SHGFI_ICONLOCATION Or SHGFI_SYSICONINDEX
        Label1.Caption = f.szTypeName
    End If
Pesan:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "FileName"
        Exit Property
    End If
    PropertyChanged "FileName"

End Property

'Public DisplayFileAssociated As Display
Private Function FindFile(sFileName As String) As WIN32_FIND_DATA

Dim Win32Data         As WIN32_FIND_DATA
Dim plngFirstFileHwnd As Long

  
    On Error GoTo Pesan
    ' Find file and get file data
    plngFirstFileHwnd = FindFirstFile(sFileName, Win32Data)
    If plngFirstFileHwnd = 0 Then
        FindFile.cFileName = "Error"
    Else 'NOT PLNGFIRSTFILEHWND...
        FindFile = Win32Data
    End If
    FindClose plngFirstFileHwnd
   
Pesan:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "FindFile"
       
    End If

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font

    Set Font = m_Font

End Property

Public Property Set Font(ByVal New_Font As Font)

    Set m_Font = New_Font
    PropertyChanged "Font"

End Property


Public Property Get ForeColor() As OLE_COLOR

    ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)

    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"

End Property

Private Function FormatSize(ByVal Amount As Long) As String
Dim Buffer As String
Dim Result As String

    On Error GoTo Pesan
    Buffer = Space$(255) 'Fill buffer
    Result = StrFormatByteSize(Amount, Buffer, Len(Buffer)) 'Format file size
    If InStr(Result, vbNullChar) > 1 Then
        FormatSize = Left$(Result, InStr(Result, vbNullChar) - 1)
    End If
Pesan:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "FormatSize"
       
    End If

End Function

Private Sub Get_Picture(picDest As PictureBox, _
                        ByVal MaxSize As Integer, _
                        ByVal sPath As String)


Dim typRect As TYPERECT
Dim w       As Integer
Dim h       As Integer

  
    On Error GoTo Pesan
    
    With picDest
        .Picture = LoadPicture(sPath)
        w = .Width
        h = .Height
    End With 'picDest
    If w > MaxSize Or h > MaxSize Then
        If w >= h Then
            picDest.Width = MaxSize
            picDest.Height = (h / w) * MaxSize
        Else 'NOT W...
            picDest.Height = MaxSize
            picDest.Width = (w / h) * MaxSize
        End If
    Else 'NOT W...
        picDest.Width = w
        picDest.Height = h
    End If
    picDest.AutoRedraw = True
    picDest.PaintPicture picDest, 0, 0, picDest.Width, picDest.Height
    EFCBevel picDest, 7, vbWhite
    Set Picture2.Picture = Picture2.Image
    PaintControl Picture2, Bump, , , , False
   
Pesan:
    If Err.Number <> 0 Then
        If Err.Number = 13 Then
            Exit Sub
        End If
        MsgBox Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "GETPICTURE"
       
    End If

End Sub

Private Function GetExtension(ByVal FullFilePath As String) As String

Dim p As Long

    On Error GoTo Pesan
    If Len(FullFilePath) > 0 Then
        p = InStrRev(FullFilePath, ".")
        If p > 0 Then
            If p < Len(FullFilePath) Then
                GetExtension = Mid$(FullFilePath, p + 1)
            End If
        End If
    End If
Pesan:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "GetExtension"
      
    End If

End Function

Private Sub GetFileInfo()

Dim lSize                As Long
Dim FileTime             As SYSTEMTIME
Dim FileData             As WIN32_FIND_DATA

Dim TimeStamp            As String
   

Dim chk                  As Integer


 
    On Error GoTo FileInfoError
    For chk = 0 To 4
        chkAttrib(chk).Value = 0
    Next chk
    Screen.MousePointer = vbHourglass 'Set mouse pointer to hourglass
    FileData = FindFile(FileName) 'Find file and get data
    TimeStamp = FileDateTime(FileName)
    '------GET FILE NAME AND PATH------'
    lblName.Caption = GetFTitle(FileName)  'Get file title

    TitleName = GetFTitle(FileName)
    Location = FileName
    lblLocation.Caption = "Location: " & FileName  'Get file location
   
    '------GET FILE SIZE------'
    If FileData.nFileSizeHigh = 0 Then 'Get file size
        lSize = FileData.nFileSizeLow
        lblSize.Caption = FormatSize(lSize)  'Format size
     
    Else 'NOT FILEDATA.NFILESIZEHIGH...
        lSize = FileData.nFileSizeHigh
        lblSize.Caption = FormatSize(lSize)  'Format size

    End If
    '------GET FILE DATES------'
    ' Created
    FileTimeToSystemTime FileData.ftCreationTime, FileTime
    lblDate(0).Caption = FileTime.wDay & "/" & FileTime.wMonth & "/" & FileTime.wYear & " " & FileTime.wHour & ":" & FileTime.wMinute & ":" & FileTime.wSecond
    ' Modified
    FileTimeToSystemTime FileData.ftLastWriteTime, FileTime
    lblDate(1).Caption = FileTime.wDay & "/" & FileTime.wMonth & "/" & FileTime.wYear & " " & FileTime.wHour & ":" & FileTime.wMinute & ":" & FileTime.wSecond
    DateModified = "Date Modified: " & Format$(TimeStamp, "dddddd") & "  " & Format$(TimeStamp, "h:mm AM/PM")  'FileTime.wDay & "/" & FileTime.wMonth & "/" & FileTime.wYear & " " & FileTime.wHour & ":" & FileTime.wMinute & ":" & FileTime.wSecond
    ' Accessed
    FileTimeToSystemTime FileData.ftLastAccessTime, FileTime
    lblDate(2).Caption = ""
    lblDate(2).Caption = FileTime.wDay & "/" & FileTime.wMonth & "/" & FileTime.wYear

    '------GET FILE ATTRIBUTES------'
    ' Hidden
    If (FileData.dwFilechkattrib And FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN Then
        chkAttrib(1).Value = 1
        FileAttribut = "Hidden"
    End If
    ' System
    If (FileData.dwFilechkattrib And FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM Then
        chkAttrib(2).Value = 1
        FileAttribut = "System"
    End If
    ' Read Only
    If (FileData.dwFilechkattrib And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY Then
        chkAttrib(3).Value = 1
        FileAttribut = "Read Only"
    End If
    ' Archive
    If (FileData.dwFilechkattrib And FILE_ATTRIBUTE_ARCHIVE) = FILE_ATTRIBUTE_ARCHIVE Then
        chkAttrib(4).Value = 1
        FileAttribut = "Archive"
    End If
    Screen.MousePointer = vbDefault 'Set mouse pointer to default
FileInfoError:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "GetFileInfo"
        Exit Sub
    End If
    Screen.MousePointer = vbDefault 'Set mouse pointer to default

End Sub

Private Function GetFTitle(strFileName As String) As String

Dim cbBuf As String

    On Error GoTo GFTError
    cbBuf = String$(250, vbNullChar) 'Fill buffer with null chars
    GetFileTitle strFileName, cbBuf, Len(cbBuf) 'Get file title
    GetFTitle = Left$(cbBuf, InStr(1, cbBuf, vbNullChar) - 1) 'Extract file title from buffer
GFTError:
    If Err.Number <> 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbInformation + vbOKOnly, "ADA YANG SALAH"
       
    End If

End Function

Private Sub GetIcon()



Dim I     As String

Dim lIcon As Long
    On Error GoTo Pesan
    
    With Picture1
        .Visible = False
        .AutoRedraw = True
        .Cls
        .AutoRedraw = True
        .BackColor = m_Backcolor
    End With 'Picture1
    I = GetExtension(FileName)
    ' Extract assocciate icon from file
    lIcon = ExtractAssociatedIcon(App.hInstance, FileName, 0&)
    DrawIconEx Picture1.hDC, 0, 0, lIcon, 0, 0, 0, 0, DI_NORMAL 'Draw icon in picturebox
    Set Picture1.Picture = Picture1.Image
    Picture1.Visible = True
    'Picture2.Picture = Picture1.Image
    DestroyIcon lIcon 'Destroy icon
    If I = "bmp" Or I = "BMP" Or I = "Bmp" Or I = "JPG" Or I = "Jpg" Or I = "jpg" Then
        If PWidth > PHeight Then
            Get_Picture Picture2, 100, FileName
            
        Else 'NOT PWIDTH...
            Get_Picture Picture2, 75, FileName
        End If
    End If
Pesan:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "GET ICON"
        
    End If

End Sub

Private Sub Picture1_Click()

    ShellExecute hwnd, "Open", m_FileName, "", App.Path, 1

End Sub


Private Sub UserControl_Initialize()

Dim N As Integer
     m_Backcolor = m_def_BackColor

    For N = 0 To 4
        chkAttrib(N).BackColor = vbWhite
    Next N '
    ' GetFileInfo
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = ""
    Picture1.Visible = True
    Picture1.Width = 32 * 15
    Picture1.Height = 32 * 15

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

'    m_BackColor = m_def_BackColor

    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_AutoSize = m_def_AutoSize
    m_FileName = m_def_FileName

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)



    With PropBag
        ' m_FileName = PropBag.ReadProperty("FileName", Nothing)
        '    m_BackColor = PropBag.ReadProperty("BackColor", vbBlue)
        m_ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
        m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
        Set m_Font = .ReadProperty("Font", Ambient.Font)
        m_BackStyle = .ReadProperty("BackStyle", m_def_BackStyle)
        m_BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
        m_AutoSize = .ReadProperty("AutoSize", m_def_AutoSize)
        m_FileName = .ReadProperty("FileName", m_def_FileName)
        m_DisplayFileAssociated = .ReadProperty("DisplayFileAssociated", m_def_DisplayFileAssociated)
        m_Backcolor = .ReadProperty("BackColor", &H8000000F)
    End With 'PropBag

End Sub

Private Sub UserControl_Resize()

    If UserControl.Width < 400 * 15 Or UserControl.Width > 400 * 15 Then
        UserControl.Width = 400 * 15
    End If
    If UserControl.Height < 240 * 15 Or UserControl.Height > 240 * 15 Then
        UserControl.Height = 240 * 15
    End If
    
    With Picture1
        .Width = 32
        .Height = 32
        .Left = Picture4.ScaleWidth - .ScaleWidth - 5
        .Top = Picture4.ScaleHeight - .ScaleHeight - 5
        'Picture1.BackColor = vbWhite
        'PaintControl Picture4, HalfRaised, vbButtonFace, vbBlack, "", False
    End With 'Picture1
    PaintControl Picture2, Bump, vbWhite, vbBlack, "", False
    
End Sub

Private Sub UserControl_Show()

    PaintControl Picture4, Etched, m_Backcolor, vbBlack, "", False
    PaintControl Picture3, Etched, m_Backcolor, vbBlack, "", False

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)



    With PropBag
        'Call PropBag.WriteProperty("FileName", m_FileName, Nothing)
        '    Call PropBag.WriteProperty("BackColor", m_BackColor, vbRed)
        .WriteProperty "ForeColor", m_ForeColor, m_def_ForeColor
        
        .WriteProperty "Enabled", m_Enabled, m_def_Enabled
        
        .WriteProperty "Font", m_Font, Ambient.Font
        
        .WriteProperty "BackStyle", m_BackStyle, m_def_BackStyle
        
        .WriteProperty "BorderStyle", m_BorderStyle, m_def_BorderStyle
        
        .WriteProperty "AutoSize", m_AutoSize, m_def_AutoSize
        
        .WriteProperty "FileName", m_FileName, m_def_FileName
        
        .WriteProperty "DisplayFileAssociated", m_DisplayFileAssociated, m_def_DisplayFileAssociated
        
        .WriteProperty "BackColor", m_Backcolor, &H8000000F
        
    End With 'PropBag

End Sub

'Public Sub Iconix(picBox As PictureBox)
'Dim lIcon As Long
'
'    On Error GoTo Pesan
'
'    With picBox
'        .AutoRedraw = True
'        .Cls
'        .AutoRedraw = False
'        ' Extract assocciate icon from file
'    End With 'picBox
'    lIcon = ExtractAssociatedIcon(App.hInstance, FileName, 0&)
'    DrawIconEx picBox.hDC, 0, 0, lIcon, 0, 0, 0, 0, DI_NORMAL 'Draw icon in picturebox
'    'Set Picture1.Picture = Picture1.Image
'    'picBox.Picture = Picture1.Image
'    DestroyIcon lIcon 'Destroy icon
'Pesan:
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "ICONIX"
'
'    End If
'
'End Sub




