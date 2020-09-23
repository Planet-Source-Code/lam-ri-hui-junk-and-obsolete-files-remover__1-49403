VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Junk and Obsolete Files Remover"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      ToolTipText     =   "Enter the drive you want to clean"
      Top             =   120
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do It Now"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      ToolTipText     =   "Click this button to start cleaning. This process may take a few minutes depending to your harddisk's size."
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   $"Main.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   6975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(C) Lam Ri Hui 2003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3960
      TabIndex        =   5
      Top             =   1680
      Width           =   2145
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "e.g. C:\"
      Height          =   195
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Drive's name : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   1710
      Left            =   120
      Picture         =   "Main.frx":00CD
      Top             =   120
      Width           =   1710
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Function FindFilesAPI(path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)
On Error Resume Next

    Dim FileName As String ' Walking filename variable...
    Dim DirName As String ' SubDirectory Name
    Dim dirNames() As String ' Buffer for directory name entries
    Dim nDir As Integer ' Number of directories in this path
    Dim I As Integer ' For-loop counter...
    Dim hSearch As Long ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    If Right(path, 1) <> "\" Then path = path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        DoEvents
        DirName = StripNulls(WFD.cFileName)
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                DirCount = DirCount + 1
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            FileName = StripNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") Then
                FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                FileCount = FileCount + 1
                List1.AddItem path & FileName
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
            DoEvents
        Wend
        Cont = FindClose(hSearch)
    End If
    ' If there are sub-directories...
    If nDir > 0 Then
        ' Recursively walk into them...
        For I = 0 To nDir - 1
            FindFilesAPI = FindFilesAPI + FindFilesAPI(path & dirNames(I) & "\", SearchStr, FileCount, DirCount)
            DoEvents
        Next I
    End If
End Function

Private Sub Command1_Click()
On Error Resume Next
    'make sure the textbox has drive's path entered
    If Text1.Text = "" Then
    MsgBox "Please enter a drive."
    Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    'remove files with the following extensions
    JunkRemove "*.~*"
    JunkRemove "~*.*"
    JunkRemove "*.*~"
    JunkRemove "*.---"
    JunkRemove "*.tmp"
    JunkRemove "*._mp"
    JunkRemove "*.old"
    JunkRemove "*.bak"
    JunkRemove "*.chk"
    JunkRemove "*.gid"
    JunkRemove "0???????.nch"
    JunkRemove "*.~*"
    JunkRemove "*.dmp"
    
    MsgBox "Process done."
    Screen.MousePointer = vbDefault
End Sub

Private Sub JunkRemove(Extension As String)
On Error Resume Next
    Dim SearchPath As String, FindStr As String
    Dim FileSize As Long
    Dim NumFiles As Integer, NumDirs As Integer
    Dim I
    Dim FTD As New Collection
    Dim FileName
    List1.Clear

    'search files with specified extension
    SearchPath = Text1.Text
    FindStr = Extension
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
    For I = 0 To List1.ListCount
        FTD.Add List1.List(I)
    Next I
    
    'delete the files
    For I = 1 To FTD.Count - 1
            Kill FTD.Item(I)
    Next I

End Sub

