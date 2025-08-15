VERSION 5.00
Begin VB.Form Frm_SerialNo_Information 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Serial No Information"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8910
   Icon            =   "Frm_SerialNo_Information.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   975
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   8355
      Begin VB.TextBox txtSerialNoTo 
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
         Index           =   1
         Left            =   4800
         MaxLength       =   25
         TabIndex        =   1
         Top             =   360
         Width           =   1530
      End
      Begin VB.TextBox txtSerialNoFrom 
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
         Index           =   0
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   0
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial No To"
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
         Left            =   3480
         TabIndex        =   10
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label lblSerialNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial No From"
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
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Submit"
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
      Index           =   4
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   8355
      Begin VB.Label LblErrMsg 
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
         Height          =   240
         Left            =   105
         TabIndex        =   5
         Top             =   195
         Width           =   8130
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1140
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   6840
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No Information"
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
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   8565
   End
End
Attribute VB_Name = "Frm_SerialNo_Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tampung As String
Dim NoPengajuan As String
Dim RS As New ADODB.Recordset

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    
    up_Clear
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
End Sub

Private Sub up_Clear()
    txtSerialNoFrom(0).Text = ""
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub


Private Sub cmdSearch_Click(Index As Integer)
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    
    If txtSerialNoFrom(0).Text = "" Then
        LblErrMsg.Caption = "Please input Serial No !"
        txtSerialNoFrom(0).SetFocus
        Exit Sub
    End If
    
    LblErrMsg.Caption = ""
                
    Me.MousePointer = vbHourglass
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_SerialNoInformation_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("SerialNoFrom", adVarChar, adParamInput, 10, txtSerialNoFrom(0).Text)
    cmd.Parameters.append cmd.CreateParameter("SerialNoTo", adVarChar, adParamInput, 10, txtSerialNoTo(1).Text)

    
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
       LblErrMsg = "[000] - Serial Number " & txtSerialNoFrom(0).Text & " to " & txtSerialNoTo(1).Text & " Already order at Po Number : " & Trim(RS("Po_no")) & ""
    Else
       LblErrMsg = "[000] - Serial Number " & txtSerialNoFrom(0).Text & " to " & txtSerialNoTo(1).Text & " Available "
    End If
    
    Me.MousePointer = vbDefault
    
    RS.Close
End Sub


Private Sub txtSerialNoFrom_Change(Index As Integer)
    LblErrMsg.Caption = ""
End Sub
