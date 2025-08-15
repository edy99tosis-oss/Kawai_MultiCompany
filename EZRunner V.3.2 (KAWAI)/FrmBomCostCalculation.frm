VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBomCostCalculation 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BOM Cost Calculation"
   ClientHeight    =   4380
   ClientLeft      =   6195
   ClientTop       =   3390
   ClientWidth     =   8070
   Icon            =   "FrmBomCostCalculation.frx":0000
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   ScaleHeight     =   5373.782
   ScaleMode       =   0  'User
   ScaleWidth      =   8010.443
   Begin VB.PictureBox PicProg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   7365
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2280
      Width           =   7395
      Begin MSComctlLib.ProgressBar Prg1 
         Height          =   315
         Left            =   105
         TabIndex        =   9
         Top             =   90
         Width           =   7140
         _ExtentX        =   12594
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   7395
      Begin MSComCtl2.DTPicker dt 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   315
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMM yyyy"
         Format          =   294191107
         UpDown          =   -1  'True
         CurrentDate     =   37831
      End
      Begin VB.Label LblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
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
         TabIndex        =   7
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ca&ncel"
      Enabled         =   0   'False
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
      Index           =   1
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
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
      Index           =   0
      Left            =   6495
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   360
      TabIndex        =   1
      Top             =   2910
      Width           =   7395
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
         TabIndex        =   2
         Top             =   195
         Width           =   7050
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
      TabIndex        =   0
      Top             =   3720
      Width           =   1140
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   5880
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   240
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Cost Calculation"
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
      Left            =   360
      TabIndex        =   12
      Top             =   480
      Width           =   7275
   End
   Begin VB.Label lblKet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   375
      TabIndex        =   11
      Top             =   2700
      Width           =   60
   End
End
Attribute VB_Name = "FrmBomCostCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoCon As New ADODB.Connection

Dim i As Integer, sql As String
Dim blnCancel As Boolean
Dim inputParent As String, lotno As String
Dim tglProd As String, qtyParent As Double
Dim factoryCD As String, lineCD As String
Dim TampungDt As Byte

Dim totProd As Double, totOffQty As Double
Dim startDaily As String

Private Sub Command1_Click(Index As Integer)
    If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Exit Sub

    Me.MousePointer = vbHourglass
    
    Select Case Index
    'Case 0: MRPCalculation
    Case 1: blnCancel = (MsgBox("Are you sure want to cancel transfer proccess?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation") = vbYes)
    End Select
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    Dim Ret As String, NC As Long, TempPWD As String

    dt = Format(Now, "MMM yyyy")
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    Ret = String(255, 0)
    NC = GetPrivateProfileString("StartDaily", "Date", "", Ret, 255, IniFile)
    If NC <> 0 Then Ret = Left$(Ret, NC)
    startDaily = Ret
End Sub

Private Sub dt_change()
    Call dt_Click
    TampungDt = dt.Month
    LblErrMsg.Caption = ""
    lblKet.Caption = ""
End Sub

Private Sub dt_Click()
    If dt.Month = 1 And Val(TampungDt) = 12 Then dt.Year = dt.Year + 1
    If dt.Month = 12 And Val(TampungDt) = 1 Then dt.Year = dt.Year - 1
End Sub
Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub


