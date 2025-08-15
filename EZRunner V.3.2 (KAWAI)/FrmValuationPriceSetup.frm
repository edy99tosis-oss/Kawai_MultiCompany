VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmValuationPriceSetup 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valuation Price Setup"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   Icon            =   "FrmValuationPriceSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   690
      TabIndex        =   5
      Top             =   3270
      Width           =   7665
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
         Left            =   90
         TabIndex        =   6
         Top             =   180
         Width           =   7455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2085
      Left            =   705
      TabIndex        =   3
      Top             =   960
      Width           =   7665
      Begin VB.Frame Frame3 
         BackColor       =   &H00FDDFE3&
         Height          =   1005
         Left            =   270
         TabIndex        =   9
         Top             =   870
         Width           =   7125
         Begin VB.ComboBox cbTerm 
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
            ItemData        =   "FrmValuationPriceSetup.frx":0E42
            Left            =   4140
            List            =   "FrmValuationPriceSetup.frx":0E4C
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   570
            Width           =   1335
         End
         Begin VB.OptionButton OptBook 
            BackColor       =   &H00FDDFE3&
            Caption         =   "Book Keeping Exchange Rate"
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
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   3045
         End
         Begin VB.OptionButton optDaily 
            BackColor       =   &H00FDDFE3&
            Caption         =   "Daily Exchange Rate"
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
            Left            =   120
            TabIndex        =   10
            Top             =   210
            Width           =   2235
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FDDFE3&
            Caption         =   "Term :"
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
            Left            =   3390
            TabIndex        =   12
            Top             =   585
            Width           =   675
         End
      End
      Begin MSForms.ComboBox CbCurr 
         Height          =   315
         Left            =   3180
         TabIndex        =   8
         Top             =   450
         Width           =   1305
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2302;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valuation Price Currency Code : "
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
         Left            =   240
         TabIndex        =   4
         Top             =   510
         Width           =   2835
      End
   End
   Begin VB.CommandButton cmdSubmit 
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
      Left            =   7215
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4050
      Width           =   1140
   End
   Begin VB.CommandButton cmdsub 
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
      Left            =   690
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4020
      Width           =   1140
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   6525
      TabIndex        =   7
      Top             =   255
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Valuation Price Setup"
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
      Left            =   60
      TabIndex        =   2
      Top             =   255
      Width           =   8970
   End
End
Attribute VB_Name = "FrmValuationPriceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim basecurr, term, sql, ratecls As String
Dim RS As New ADODB.Recordset

Private Sub cbCurr_Change()
    cbCurr_Click
End Sub

Private Sub cbCurr_Click()
LblErrMsg.Caption = ""
End Sub

Private Sub cbTerm_Change()
    cbCurr_Click
End Sub

Private Sub cbTerm_Click()
   cbCurr_Click
End Sub

Private Sub cmdsub_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub CmdSubmit_Click()
    On Error GoTo errHandler
    'basecurr = Str(0) + Trim$(Str(cbCurr.ListIndex + 1))
    term = Trim$(Str(cbTerm.ListIndex + 1))
    Db.Execute "update company_profile set valuationPrice_baseCurrency = '" & uf_GetCurrencyCode(Trim$(CbCurr)) & "', " & _
                     " ValuationPrice_ExchTerm ='" & term & "', " & _
                     " Rate_Cls = '" & IIf(optDaily.Value = True, "0", "1") & "' " & _
                     " where company_code = '00000' "
    LblErrMsg.Caption = DisplayMsg(1000) '"Data Save Success !"
    Exit Sub
errHandler:
    LblErrMsg.Caption = DisplayMsg("0057")    '"Data Cannot be Saved !"
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    RS.Open "select ValuationPrice_BaseCurrency, ValuationPrice_ExchTerm, Rate_Cls from Company_Profile", Db, adOpenForwardOnly, adLockReadOnly
    If IsNull(RS(0)) = False Then
        basecurr = Trim$(RS(0))
        term = RS(1)
    End If
    
    If Trim(RS!Rate_Cls) = "0" Then
     optDaily.Value = True
    Else
     OptBook.Value = True
    End If
    
    RS.Close
    CbCurr.clear
    Call up_FillCombo(CbCurr, "curr_cls", "description,curr_cls")
    CbCurr.ListWidth = 40
    CbCurr.ColumnWidths = "40 pt;0 pt"
    CbCurr.ListIndex = Int(basecurr - 1)
    cbTerm.ListIndex = term - 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RS.State = adStateOpen Then RS.Close
    Set RS = Nothing
End Sub
