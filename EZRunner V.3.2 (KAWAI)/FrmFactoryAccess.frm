VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.dll"
Begin VB.Form FrmFactoryAccess 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Factory Access"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5355
   Icon            =   "FrmFactoryAccess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrNext 
      Left            =   2520
      Top             =   2880
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Cancel"
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
      TabIndex        =   3
      Top             =   2760
      Width           =   1230
   End
   Begin VB.CommandButton CmdSubmit 
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   4920
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
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   225
         Width           =   4635
      End
   End
   Begin VB.Frame fraCompanyList 
      BackColor       =   &H00FDDFE3&
      Height          =   1185
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   4920
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Search"
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
         Left            =   7620
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   630
         Width           =   1035
      End
      Begin MSForms.OptionButton OBFactory2 
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   375
         BackColor       =   16637923
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "661;450"
         Value           =   "0"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton OBFactory1 
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   375
         BackColor       =   16637923
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "661;450"
         Value           =   "0"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PT KAWAI INDONESIA PLANT-3                        "
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
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   3075
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PT KAWAI INDONESIA PLANT-3 NEW"
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
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   3165
      End
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Factory Access"
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
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   5370
   End
End
Attribute VB_Name = "FrmFactoryAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pCompanyCode As String
Private selectedFactory As String

Private Sub cmdCancel_Click()
    frmLogin.Show
End Sub

Private Sub CmdSubmit_Click()
     On Error GoTo ErrHandler

    Dim sql As String
    Dim selectedFactory As String
    Dim userID As String
    Dim FrmMainMenuCaption As String

    userID = Trim(frmLogin.txtUser.Text)

    ' === Tentukan Factory_Code berdasarkan pilihan ===
    If OBFactory1.Value = True Then
        selectedFactory = "00000"
        FrmMainMenuCaption = "PT KAWAI INDONESIA PLANT-3"
    ElseIf OBFactory2.Value = True Then
        selectedFactory = "11111"
        FrmMainMenuCaption = "PT KAWAI INDONESIA PLANT-3 NEW"
    Else
        MsgBox "Silakan pilih salah satu company terlebih dahulu.", vbExclamation, "Peringatan"
        Exit Sub
    End If

    ' === Update privilege ===
    sql = "UPDATE dbo.App_FactoryPrivilege " & _
          "SET Show = 1, UpdateDate = GETDATE() " & _
          "WHERE UserID = '" & userID & "' " & _
          "AND Factory_Code = '" & selectedFactory & "'"
          
    Db.Execute sql

    ' === Tampilkan pesan sukses ===
    LblErrMsg.Caption = "Factory selected successfully"
    DoEvents
    Call Delay(1)

    ' === Reset supaya login berikutnya selalu pilih factory lagi ===
    frmLogin.NeedFactorySelection = True

    ' === Lanjut ke Main Menu ===
    pCompanyCode = selectedFactory
    frmMainMenu.Caption = "EZ Runner ver.3 - Main Menu | " & FrmMainMenuCaption
    frmMainMenu.loadtree
    frmMainMenu.Show
    DoEvents
    Me.Hide

    Exit Sub

ErrHandler:
    MsgBox "Terjadi kesalahan saat submit company: " & err.Description, vbCritical, "Error"
End Sub

Private Sub tmrNext_Timer()
    ' === Timer jalan setelah 3 detik ===
    tmrNext.Enabled = False

    ' === Lanjut ke Main Menu ===
    pCompanyCode = selectedFactory
    frmMainMenu.loadtree
    frmMainMenu.Show
    DoEvents
    Me.Hide
End Sub

Private Sub Form_Load()
 If gb_Simulation = True Then Call up_InitSimulation(Me)
 If gb_Simulation = True Then OBFactory1.BackColor = RGB(204, 255, 204)
 If gb_Simulation = True Then OBFactory2.BackColor = RGB(204, 255, 204)
 
 
   Dim ctl As Control
   
    ' Reset semua OptionButton
    For Each ctl In Me.Controls
        If TypeOf ctl Is OptionButton Then ctl.Value = False
    Next ctl

    ' Timer default disabled
    tmrNext.Enabled = False
    tmrNext.Interval = 1000  ' 1 detik
    
    LblErrMsg.Caption = ""
       
End Sub
