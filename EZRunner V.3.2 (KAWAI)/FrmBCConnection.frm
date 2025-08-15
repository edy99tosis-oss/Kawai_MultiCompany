VERSION 5.00
Begin VB.Form FrmBCConnection 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Connection Setting"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6315
   Icon            =   "FrmBCConnection.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTestConn 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Test Connection"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "FFTT*/"
      Top             =   3480
      Width           =   1605
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Apply"
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "FFTT*/"
      Top             =   3480
      Width           =   1005
   End
   Begin VB.CommandButton CmdSubMenu 
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "TFFT*/"
      Top             =   3480
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Tag             =   "TFTF*/"
      Top             =   720
      Width           =   5805
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Tag             =   "TTFF*/"
         Top             =   750
         Width           =   3135
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   1965
         Width           =   3135
      End
      Begin VB.TextBox txtUserId 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox txtDbName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   1140
         Width           =   3135
      End
      Begin VB.TextBox txtServerName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   14
         Tag             =   "TTFF*/"
         Top             =   780
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   1995
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Id"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   1620
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   385
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Setting Connection To Ceisa"
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
      Index           =   3
      Left            =   0
      TabIndex        =   13
      Tag             =   "TTTF*/"
      Top             =   120
      Width           =   6450
   End
End
Attribute VB_Name = "FrmBCConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RS As New ADODB.Recordset
Public ConnStr As String

Private Sub cmdApply_Click()
'Dim strSQL As String
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BCConnectionString"
    
    cmd.Parameters.append cmd.CreateParameter("ServerName", adVarChar, adParamInput, 100, txtServerName.Text)
    cmd.Parameters.append cmd.CreateParameter("Port", adVarChar, adParamInput, 100, txtPort.Text)
    cmd.Parameters.append cmd.CreateParameter("DatabaseName", adVarChar, adParamInput, 100, txtDbName.Text)
    cmd.Parameters.append cmd.CreateParameter("UserId", adVarChar, adParamInput, 100, txtUserId.Text)
    cmd.Parameters.append cmd.CreateParameter("Password", adVarChar, adParamInput, 100, fc_Encrypt(txtPassword.Text))
    cmd.Parameters.append cmd.CreateParameter("LastUser", adVarChar, adParamInput, 100, userLogin)
    
    Set RS = cmd.Execute
    
     MsgBox "Apply Connection Success", vbInformation, "Information"
    
End Sub

Private Sub CmdSubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub cmdTestConn_Click()
    'LblErrMsg = ""
    Me.MousePointer = vbHourglass
    KoneksiMysql
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim sql As String
Dim RS As New Recordset

    sql = "SELECT * FROM Connection_Mysql"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        txtServerName = Trim(RS("ServerName"))
        txtPort = Trim(RS("Port"))
        txtDbName = Trim(RS("DatabaseName"))
        txtUserId = Trim(RS("UserId"))
        txtPassword = fc_Decrypt(Trim(RS("Password")))
    End If
End Sub

Private Sub KoneksiMysql()
Dim db_name As String
Dim db_server As String
Dim db_port As String
Dim db_user As String
Dim db_pass As String
Dim Conn As New ADODB.Connection
Dim sql As String
Dim RS As New Recordset

'//error traping
On Error GoTo buat_koneksi_Error

    sql = "SELECT * FROM Connection_Mysql"
    Set RS = Db.Execute(sql)
    
    '/variable localhost
    db_name = Trim(RS("DatabaseName"))
    db_server = Trim(RS("ServerName"))
    db_port = Trim(RS("Port"))
    db_user = Trim(RS("UserId"))
    db_pass = fc_Decrypt(Trim(RS("Password")))


'/variable kawai 3
'db_name = "tpbdb"
'db_server = "172.16.10.223"
'db_port = "3306"
'db_user = "beacukai"
'db_pass = "beacukai"
'/buat connection string
ConnStr = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_user & ";PWD=" & db_pass & ";PORT=" & db_port & ""
'/buka koneksi
With Conn
    .ConnectionString = ConnStr
    .Open
   MsgBox "Test Connection Success", vbInformation, "Information"
End With
'___________________________________________________________
On Error GoTo 0
Exit Sub

buat_koneksi_Error:
    MsgBox "Connection Failed", vbCritical, "Change Apply Setting Failed"
End Sub

Private Sub txtServerName_Change()
    'LblErrMsg = ""
End Sub


