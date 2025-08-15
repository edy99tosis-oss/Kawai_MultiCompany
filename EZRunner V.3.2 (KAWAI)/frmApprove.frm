VERSION 5.00
Begin VB.Form frmApprove 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Approval Sign Code"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   Icon            =   "frmApprove.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   7620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
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
      Left            =   6023
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3525
      Width           =   1290
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
      Left            =   308
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3525
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Approval Sign Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   308
      TabIndex        =   8
      Top             =   1050
      Width           =   7005
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1170
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   1
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   0
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pajak"
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
         Left            =   1980
         TabIndex        =   12
         Top             =   1215
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice"
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
         Left            =   1980
         TabIndex        =   11
         Top             =   855
         Width           =   630
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DO"
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
         Left            =   1980
         TabIndex        =   9
         Top             =   480
         Width           =   270
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   540
      Left            =   308
      TabIndex        =   6
      Top             =   2865
      Width           =   7005
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
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   195
         Width           =   6780
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   5475
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   330
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Approval Sign Code"
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
      Height          =   390
      Index           =   0
      Left            =   308
      TabIndex        =   10
      Top             =   322
      Width           =   7005
   End
End
Attribute VB_Name = "frmApprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim i As Integer
Dim rsCompany As New ADODB.Recordset

Sub Kosong()
    Text1(0) = ""
    Text1(1) = ""
    Text1(2) = ""
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    sql = "select DO_SignCode as DO,Invoice_SignCode as invoice, FakturPajak_SignCode as Pajak, Last_Update, Last_User " & _
        "from Company_Profile"
    If rsCompany.State <> adStateClosed Then rsCompany.Close
    rsCompany.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    Call Kosong
    If Not (rsCompany.EOF) Then
        For i = 0 To 2
            If IsNull(rsCompany(i)) Then
                Text1(i) = ""
            Else
                Text1(i) = rsCompany(i)
            End If
        Next i
    End If
End Sub

Private Sub Command1_Click()
Dim baru As Integer
    baru = 0
    rsCompany.filter = ""
    rsCompany.Requery
    
    If rsCompany.EOF Then rsCompany.AddNew: baru = 1
    rsCompany(0) = Trim(Text1(0))
    rsCompany(1) = Trim(Text1(1))
    rsCompany(2) = Trim(Text1(2))
    rsCompany(3) = Now
    rsCompany(4) = userLogin
    rsCompany.update
    
    If baru = 1 Then
        LblErrMsg = DisplayMsg(1000)
    Else
        LblErrMsg = DisplayMsg(1101)
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub
