VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FrmDatabaseBackup 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Database Backup"
   ClientHeight    =   4275
   ClientLeft      =   1290
   ClientTop       =   3375
   ClientWidth     =   7800
   Icon            =   "FrmDatabaseBackup.frx":0000
   LinkMode        =   1  'Source
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicProg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   379
      ScaleHeight     =   480
      ScaleWidth      =   7050
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1665
      Width           =   7080
      Begin MSComctlLib.ProgressBar Prg1 
         Height          =   315
         Left            =   75
         TabIndex        =   6
         Top             =   90
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6319
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   525
      Left            =   379
      TabIndex        =   4
      Top             =   2505
      Width           =   7080
      Begin VB.Label LblErr 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   210
         Width           =   6825
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   379
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   334
      Top             =   495
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   5621
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   315
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Database Backup"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   379
      TabIndex        =   7
      Top             =   510
      Width           =   7080
   End
End
Attribute VB_Name = "FrmDatabaseBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FO_DELETE = &H3
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

'Private Sub Command1_Click()
'
'        If hakUpdate(Me.Name) = "0" Then
'                   LblErrMsg = DisplayMsg(3008)
'                   Me.MousePointer = vbDefault:
'                   Exit Sub
'        End If
'
'    Dim lngHandle As Long, SHDirOp As SHFILEOPSTRUCT, lngLong As Long
'    Dim ft1 As FILETIME, ft2 As FILETIME, SysTime As SYSTEMTIME
'   ' On Error GoTo skip
''    'Set the dialog's title
''    CDBox.DialogTitle = "Choose a file ..."
''    'Raise an error when the user pressed cancel
''    CDBox.CancelError = True
''    'Show the 'Open File'-dialog
''    CDBox.ShowOpen
'
'
''    'Create a new directory
''    CreateDirectory "C:\KPD-Team", ByVal &H0
'
'    Dim ls_directoryName As String
'    Dim ls_clientDirectoryName As String
'    Dim ls_time As String
'
'    Dim ls_fileName As String
'    Dim ls_filefullName As String
'
'
'    Me.MousePointer = vbHourglass
'    Prg1.Max = 4
'    Prg1.Value = 1
'
'    ls_time = Format(Now, "yyyyMMdd hhmmss")
'    ls_directoryName = App.path + "\BackupDB\"
'    ls_clientDirectoryName = "c:\" + "BackupDB\"
'    MakeSureDirectoryPathExists ls_directoryName
'
'    ls_fileName = ls_time & " " & gs_DBName & " "
'
'    ls_filefullName = ls_directoryName & ls_fileName
'
'    Prg1.Value = 2
'
'    Dim ls_sql As String
'
'    ls_sql = "Backup database " & gs_DBName & " to disk='" & Trim(ls_filefullName) & ".bak'"
'        Db.Execute ls_sql
'
''    Dim oZip As CGZipFiles
''    Set oZip = New CGZipFiles
''    oZip.ZipFileName = "\" & Trim(ls_fileName) & ".zip"
''    oZip.AddFile Trim(ls_filefullName) & ".bak"
''    If oZip.MakeZipFile <> 0 Then
''    'MsgBox oZip.GetLastMessage
''    End If
''    Set oZip = Nothing
'
'    Prg1.Value = 3
'
'    MakeSureDirectoryPathExists ls_clientDirectoryName
'
'    'CopyFile App.Path & Trim(ls_fileName) & ".zip", ls_clientDirectoryName & Trim(ls_fileName) & ".zip", 0
'    CopyFile Trim(ls_filefullName) & ".bak", ls_clientDirectoryName & Trim(ls_fileName) & ".bak", 0
'
'skip:
'   Prg1.Value = 4
'
'   If Trim(Err.Description) <> "" Then
'        LblErr = "[0000] Database Backup Failed !"
'   Else
'        LblErr = "[0000] Database Backup Success !"
'   End If
'        Me.MousePointer = vbDefault
'End Sub

Private Sub Command1_Click()
   Me.MousePointer = vbHourglass
           
   Dim ls_time As String
   Dim ls_fileName As String
   Dim ls_filefullName As String
   Dim ls_sql As String
   
   Dim strINI As String
   Dim cFiles As New collection
   Dim strDir As String
   
   If hakUpdate(Me.Name) = "0" Then
      LblErr = DisplayMsg(3008)
      Me.MousePointer = vbDefault:
      Exit Sub
   End If
      
   Prg1.Max = 4
   Prg1.Value = 1
   MakeSureDirectoryPathExists gvDBBackupServer
   
   ' set path and filename value
   ls_time = Format(Now, "yyyyMMdd hhmmss")
   ls_fileName = ls_time & " " & gs_DBName & " "
   ls_fileName = Trim$(ls_fileName)
   ls_filefullName = gvDBBackupServer & ls_fileName
   ls_filefullName = Trim$(ls_filefullName)
    
   Prg1.Value = 2
   
   ' Execute backup database
   ls_sql = "Backup database " & gs_DBName & " to disk='" & ls_filefullName & ".bak'"
   Db.Execute ls_sql
   
   ' Copy backup from server to local
   MakeSureDirectoryPathExists gvDBBackupLocal
   CopyFile ls_filefullName & ".bak", gvDBBackupLocal & ls_fileName & ".bak", 0
   
   ' Compress backup file
   strDir = Dir(gvDBBackupLocal & "\*.*")
   Do While strDir <> ""
      cFiles.Add strDir
      strDir = Dir()
   Loop
   Call CreateZipFile(gvDBBackupServer, cFiles, ls_fileName & ".bak", ls_filefullName & ".zip")
   
   Prg1.Value = 3
   
   ' Copy zip file from server to local
   MakeSureDirectoryPathExists gvDBBackupLocal
   CopyFile ls_filefullName & ".zip", gvDBBackupLocal & ls_fileName & ".zip", 0

skip:
   Prg1.Value = 4

   If Trim(err.Description) <> "" Then
      LblErr = DisplayMsg("0064")
   Else
      LblErr = DisplayMsg("0065")
   End If
   Me.MousePointer = vbDefault
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErr.Caption = ErrMsg
End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
IsiCombo
LblErr.Caption = ""
End Sub

Sub IsiCombo()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CmdSubMenu_Click()
  frmMainMenu.Show
  Unload Me
End Sub
