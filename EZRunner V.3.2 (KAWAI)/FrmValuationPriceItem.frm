VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmValuationPriceItem 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Valuation Price Master (By Product Code)"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmValuationPriceItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10680
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "..."
      Height          =   300
      Index           =   1
      Left            =   2310
      TabIndex        =   35
      Top             =   8497
      Width           =   300
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1410
      Left            =   480
      TabIndex        =   28
      Top             =   1440
      Width           =   14175
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   4950
         TabIndex        =   34
         Top             =   772
         Width           =   300
      End
      Begin VB.CommandButton CmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
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
         Left            =   9990
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   1140
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   5430
         X2              =   9780
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Cls"
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
         Left            =   270
         TabIndex        =   32
         Top             =   390
         Width           =   855
      End
      Begin VB.Label LblGroupFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   5430
         TabIndex        =   31
         Top             =   390
         Width           =   960
      End
      Begin MSForms.ComboBox CboGroupFilter 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   330
         Width           =   1335
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2355;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox CboProductFilter 
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Top             =   765
         Width           =   2955
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "5212;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblProductFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   5430
         TabIndex        =   30
         Top             =   825
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
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
         Left            =   270
         TabIndex        =   29
         Top             =   825
         Width           =   1155
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   5430
         X2              =   9780
         Y1              =   1065
         Y2              =   1065
      End
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
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
      Left            =   12180
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9720
      Width           =   1155
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   480
      TabIndex        =   17
      Top             =   9000
      Width           =   14115
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
         Left            =   45
         TabIndex        =   18
         Top             =   180
         Width           =   13935
      End
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9720
      Width           =   1155
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9720
      Width           =   1185
   End
   Begin VB.TextBox TxtAmount 
      Alignment       =   1  'Right Justify
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
      Left            =   12960
      MaxLength       =   25
      TabIndex        =   8
      Tag             =   "Amount"
      Top             =   8490
      Width           =   1515
   End
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
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
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9720
      Width           =   1155
   End
   Begin VB.TextBox LblProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
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
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   8550
      Width           =   2115
   End
   Begin VB.TextBox LblCost 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
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
      Left            =   6255
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8550
      Width           =   1725
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4815
      Left            =   480
      TabIndex        =   13
      Top             =   3000
      Width           =   14130
      _cx             =   24924
      _cy             =   8493
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   10932991
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483624
      BackColorAlternate=   -2147483624
      GridColor       =   12582912
      GridColorFixed  =   12582912
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   275
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComCtl2.DTPicker StartDate 
      Height          =   330
      Left            =   8190
      TabIndex        =   5
      Top             =   8490
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
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
      CustomFormat    =   "dd MMM yyyy"
      Format          =   152174595
      CurrentDate     =   37891
   End
   Begin MSMask.MaskEdBox mask 
      Height          =   315
      Left            =   9960
      TabIndex        =   6
      Top             =   8490
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker EndDate 
      Height          =   330
      Left            =   9960
      TabIndex        =   14
      Top             =   8490
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   582
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
      CustomFormat    =   "dd MMM yyyy"
      Format          =   152174595
      CurrentDate     =   37860
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12840
      TabIndex        =   33
      Top             =   690
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "End Date"
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
      Index           =   6
      Left            =   10320
      TabIndex        =   27
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Amount"
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
      Index           =   5
      Left            =   12960
      TabIndex        =   26
      Top             =   8160
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Curr"
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
      Left            =   11880
      TabIndex        =   25
      Top             =   8160
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cost Cls"
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
      Index           =   3
      Left            =   5040
      TabIndex        =   24
      Top             =   8160
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Description"
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
      Left            =   6480
      TabIndex        =   23
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Start Date"
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
      Left            =   8190
      TabIndex        =   22
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Description"
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
      Left            =   2730
      TabIndex        =   21
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Product Code"
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
      Left            =   630
      TabIndex        =   20
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Valuation Price Master (By Product Code)"
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
      Index           =   0
      Left            =   480
      TabIndex        =   19
      Top             =   690
      Width           =   14235
   End
   Begin MSForms.ComboBox CboCost 
      Height          =   315
      Left            =   5040
      TabIndex        =   4
      Tag             =   "Cost Cls"
      Top             =   8490
      Width           =   1035
      VariousPropertyBits=   746604571
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1826;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   6240
      X2              =   7980
      Y1              =   8790
      Y2              =   8790
   End
   Begin MSForms.ComboBox CboCurr 
      Height          =   315
      Left            =   11880
      TabIndex        =   7
      Tag             =   "Currency"
      Top             =   8490
      Width           =   915
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "1614;556"
      TextColumn      =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   495
      Left            =   480
      Top             =   8400
      Width           =   14130
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   480
      Top             =   8040
      Width           =   14130
   End
   Begin MSForms.ComboBox CboProduct 
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Tag             =   "Product Code"
      Top             =   8490
      Width           =   1635
      VariousPropertyBits=   746604571
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "2884;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2700
      X2              =   4800
      Y1              =   8790
      Y2              =   8790
   End
End
Attribute VB_Name = "FrmValuationPriceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ClsProc As New ClsProc
Dim ColS As Integer
Dim ColProCD As Integer, ColProDes As Integer
Dim cOlcost As Integer, ColCostDes As Integer
Dim ColStart As Integer, ColEnd As Integer
Dim ColCurr As Integer, ColCurrDes As Integer
Dim ColAmount As Integer
Dim nilKosong As Boolean
Dim ubah As Boolean, hapus As Boolean
Dim UbahEndDate As Boolean, NonValidDate As Boolean
Dim HakU As Integer, i As Integer
Dim SDate, EDate, StartDateAwal, EndDateAkhir

Private Sub cmdBrowser_Click(Index As Integer)
 Me.MousePointer = vbHourglass
 
 Select Case Index
 Case 0:
  frm_BrowseItem.getItemCode = CboProductFilter.Text
  frm_BrowseItem.Show 1
  CboProductFilter.Text = frm_BrowseItem.getItemCode
 Case 1:
  If CboProduct.Enabled = True And CboProduct.locked = False Then
   frm_BrowseItem.getItemCode = CboProduct.Text
   frm_BrowseItem.Show 1
   CboProduct.Text = frm_BrowseItem.getItemCode
  End If
 End Select
 
 Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    HakU = hakUpdate(Me.Name)
        
    Call IsiCboGroup
    Call IsiCboProduct
    Call IsiCboCost
    Call up_FillCombo(cbocurr, "curr_cls")
    
    Call SetCol
    Call Kosong(1)
    Call Header
    ubah = False
    
End Sub

Sub SetCol()
    ColS = 0
    ColProCD = 1: ColProDes = 2
    cOlcost = 3: ColCostDes = 4
    ColStart = 5: ColEnd = 6
    ColCurr = 7: ColCurrDes = 8
    ColAmount = 9
End Sub

Sub IsiCboGroup()
    Dim RsGroup As Recordset
    Dim SqlGroup As String
    
    SqlGroup = "Select Group_Cls, Description From Group_Cls"
    
    Set RsGroup = Db.Execute(SqlGroup)
    
    With CboGroupFilter
        .clear
        .columnCount = 2
        .ColumnWidths = "90pt;200pt"
        .ListWidth = 290
        .ListRows = 15
        
        .AddItem
        .List(0, 0) = strAll
        .List(0, 1) = strAll
        
        i = 1
        Do While Not RsGroup.EOF
            .AddItem
            .List(i, 0) = Trim(RsGroup("Group_Cls"))
            .List(i, 1) = Trim(RsGroup("Description"))
                        
            RsGroup.MoveNext
            i = i + 1
        Loop
    End With
    Set RsGroup = Nothing
End Sub

Sub IsiCboProduct()
    Dim RsProduct As Recordset
    Dim sqlProduct As String
    
    sqlProduct = ""
    
    sqlProduct = "Select Item_Code, Item_Name From Item_Master where use_endday >= convert(char(8), getdate(), 112) "
        
    Set RsProduct = Db.Execute(sqlProduct)
    
    With CboProduct
        .clear
        .columnCount = 2
        .ColumnWidths = "130pt;200pt"
        .ListWidth = 330
        .ListRows = 15
                        
        i = 0
        Do While Not RsProduct.EOF
            .AddItem
            .List(i, 0) = Trim(RsProduct("Item_Code"))
            .List(i, 1) = Trim(RsProduct("Item_Name"))
                        
            RsProduct.MoveNext
            i = i + 1
        Loop
    End With
    Set RsProduct = Nothing
End Sub

Sub IsiCboProductFilter(filter As String)
    Dim RsProduct As Recordset
    Dim sqlProduct As String, SqlFilter As String
    
    LblProductFilter = ""
    sqlProduct = "": SqlFilter = ""
    
    If filter <> "" And filter <> strAll Then _
        SqlFilter = " And Group_Cls = '" & Trim(filter) & "'"
    
    sqlProduct = "Select Item_Code, Item_Name From Item_Master Where use_endday >= convert(char(8), getdate(), 112) "
    sqlProduct = sqlProduct + SqlFilter
    
    Set RsProduct = Db.Execute(sqlProduct)
    
    With CboProductFilter
        .clear
        .columnCount = 2
        .ColumnWidths = "130pt;200pt"
        .ListWidth = 330
        .ListRows = 15
        
        .AddItem
        .List(0, 0) = strAll
        .List(0, 1) = strAll
        
        i = 1
        Do While Not RsProduct.EOF
            .AddItem
            .List(i, 0) = Trim(RsProduct("Item_Code"))
            .List(i, 1) = Trim(RsProduct("Item_Name"))
                        
            RsProduct.MoveNext
            i = i + 1
        Loop
    End With
    Set RsProduct = Nothing
End Sub

Sub IsiCboCost()
    Dim RsCost As Recordset
    Dim sQlcost As String
    
    sQlcost = "Select Cost_Cls, Description From InventoryCost_Master"
    Set RsCost = Db.Execute(sQlcost)
    
    With CboCost
        .clear
        .columnCount = 2
        .ColumnWidths = "90pt;200pt"
        .ListWidth = 290
        .ListRows = 15
        
        i = 0
        Do While Not RsCost.EOF
            .AddItem
            .List(i, 0) = Trim(RsCost("Cost_Cls"))
            .List(i, 1) = Trim(RsCost("Description"))
                        
            RsCost.MoveNext
            i = i + 1
        Loop
    End With
    Set RsCost = Nothing
End Sub

Sub Header()
    With grid
        .clear
        .ColS = 10: .Rows = 1
        
        .TextMatrix(0, ColS) = ""
        .TextMatrix(0, ColProCD) = "Product Code"
        .TextMatrix(0, ColProDes) = "Description"
        .TextMatrix(0, cOlcost) = "Cost Cls"
        .TextMatrix(0, ColCostDes) = "Description"
        .TextMatrix(0, ColStart) = "Start Date"
        .TextMatrix(0, ColEnd) = "End Date"
        .TextMatrix(0, ColCurr) = "Currency"
        .ColHidden(ColCurr) = True
        .TextMatrix(0, ColCurrDes) = "Currency"
        .TextMatrix(0, ColAmount) = "Amount"
                
        .ColWidth(ColS) = 300
        .ColWidth(ColProCD) = 2500
        .ColWidth(ColProDes) = 2500
        .ColWidth(cOlcost) = 1200
        .ColWidth(ColCostDes) = 1500
        .ColWidth(ColStart) = 1500
        .ColWidth(ColEnd) = 1500
        .ColWidth(ColCurr) = 900
        .ColWidth(ColCurrDes) = 1300
        .ColWidth(ColAmount) = 1700
        
        .ColAlignment(ColS) = flexAlignLeftCenter
        .ColAlignment(ColProCD) = flexAlignLeftCenter
        .ColAlignment(ColProDes) = flexAlignLeftCenter
        .ColAlignment(cOlcost) = flexAlignLeftCenter
        .ColAlignment(ColCostDes) = flexAlignLeftCenter
        .ColAlignment(ColStart) = flexAlignLeftCenter
        .ColAlignment(ColEnd) = flexAlignLeftCenter
        .ColAlignment(ColCurr) = flexAlignLeftCenter
        .ColAlignment(ColCurrDes) = flexAlignLeftCenter
        .ColAlignment(ColAmount) = flexAlignRightCenter
                
        Call ClsProc.AlignHeader(grid)
        .RowHeightMax = 250
        .EditMaxLength = 1
    End With
End Sub

Sub IsiGrid()
    Dim rsGrid As New ADODB.Recordset
    Dim tglAwal, tglAkhir
    
    Call KosongBawah
    With grid
        Call Header
        
        If Trim(CboGroupFilter.Text) = "" Then
            LblErrMsg = DisplayMsg(8057) '"Please Select Finish Good Part Cls"
            CboGroupFilter.SetFocus
            Exit Sub
        End If
        
        CboGroupFilter.Text = CboGroupFilter.Text
        If CboGroupFilter.MatchFound = False Then
            LblErrMsg = DisplayMsg(8057) '"Record with This Finish Good Part Cls Not found"
            CboGroupFilter.SetFocus
            Exit Sub
        End If
        
        If Trim(CboProductFilter.Text) = "" Then
            LblErrMsg = DisplayMsg(1009) '"Please Select Product Code"
            CboProductFilter.SetFocus
            Exit Sub
        End If
        
        CboProductFilter.Text = CboProductFilter.Text
        If CboProductFilter.MatchFound = False Then
            LblErrMsg = DisplayMsg(1009) ' "Record with This Product Code Not found"
            CboProductFilter.SetFocus
            Exit Sub
        End If
        
        sql = SqlSintaks
        Me.MousePointer = vbHourglass
        Set rsGrid = Db.Execute(sql)
        If rsGrid.EOF Or rsGrid.BOF Then
            Set rsGrid = Nothing
            LblErrMsg.Caption = DisplayMsg(4006)
            grid.clear
            Header
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        Do While Not rsGrid.EOF
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, ColS) = ""
            .Cell(flexcpBackColor, .Rows - 1, ColS) = vbWhite
            .TextMatrix(.Rows - 1, ColProCD) = Trim(rsGrid("Item_Code"))
            .TextMatrix(.Rows - 1, ColProDes) = Trim(rsGrid("Item_Desc"))
            .TextMatrix(.Rows - 1, cOlcost) = Trim(rsGrid("Cost_Cls"))
            .TextMatrix(.Rows - 1, ColCostDes) = Trim(rsGrid("Cost_Desc"))
            tglAwal = Mid(rsGrid("Start_Date"), 5, 2) & "/" & Right(rsGrid("Start_Date"), 2) & "/" & Left(rsGrid("Start_Date"), 4)
            tglAkhir = IIf(IsNull(rsGrid("End_date")), "99/99/9999", Mid(rsGrid("End_Date"), 5, 2) & "/" & Right(rsGrid("End_Date"), 2) & "/" & Left(rsGrid("End_Date"), 4))
            .TextMatrix(.Rows - 1, ColStart) = Format(tglAwal, "dd MMM yyyy")
            .TextMatrix(.Rows - 1, ColEnd) = Format(tglAkhir, "dd MMM yyyy")
            .TextMatrix(.Rows - 1, ColCurr) = Trim(rsGrid("Currency_Code"))
            .TextMatrix(.Rows - 1, ColCurrDes) = Trim(rsGrid("Currency_Desc"))
            .TextMatrix(.Rows - 1, ColAmount) = Format(rsGrid("Amount"), gs_formatAmount)
            
            rsGrid.MoveNext
        Loop
        Set rsGrid = Nothing
    End With
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
    Call IsiGrid
End Sub

Private Sub CmdSubmit_Click()
    Dim RS As New ADODB.Recordset
    Dim tanya
    
    If hakUpdate(Me.Name) = 0 Then _
            LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    hapus = False
    With grid
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "D" Then
                If IsEmpty(tanya) Then tanya = MsgBox("Do You Really Want To Delete This Data ?", vbQuestion & vbYesNo, "Confirmation")
                If tanya = vbYes Then
                    sql = "Delete From InventoryCost_Item Where Item_Code = '" & Trim(.TextMatrix(i, ColProCD)) & "' And " & _
                           "Cost_Cls = '" & Trim(.TextMatrix(i, cOlcost)) & "' And " & _
                           "Start_Date = '" & Format(.TextMatrix(i, ColStart), "yyyymmdd") & "'"
                    Db.Execute sql
                
                    hapus = True
                    SDate = Format(.TextMatrix(i, ColStart), "yyyymmdd")
                    EDate = Format(.TextMatrix(i, ColEnd), "yyyymmdd")
                    
                    cektgl
                Else
                    Exit For
                End If
            End If
        Next i
        
        If (hapus) Then grid.clear: Kosong: Header: IsiGrid: LblErrMsg = DisplayMsg(1201): Exit Sub
        
        'Validasi Untuk Insert dan Update
        If Trim(CboProduct.Text) = "" Then
            LblErrMsg = DisplayMsg(1009) '"Please Select Product Code"
            Exit Sub
        End If
        
        CboProduct.Text = CboProduct.Text
        If CboProduct.MatchFound = False Then
            LblErrMsg = DisplayMsg(1009) '"Record with This Product Code Not found"
            Exit Sub
        End If
        
        If Trim(CboCost.Text) = "" Then
            LblErrMsg = DisplayMsg(8053) '"Please Select Cost Cls"
            Exit Sub
        End If
        
        CboCost.Text = CboCost.Text
        If CboCost.MatchFound = False Then
            LblErrMsg = DisplayMsg(8053) ' "Record with This Cost Cls Not found"
            Exit Sub
        End If
        
        If mask.Text <> "99/99/9999" Then
             If IsDate(mask.Text) = False Then
                LblErrMsg.Caption = DisplayMsg(4065) '"End Date is not valid"
                mask.SetFocus
                Exit Sub
             End If
             
             If CDate(StartDate) > CDate(EndDate) Then
                LblErrMsg.Caption = DisplayMsg(4068) '"Start Date must be lower than " & Format(EndDate, "dd MMM yyyy")
                EndDate.SetFocus
                Exit Sub
             End If
        End If
        
        If Trim(cbocurr.Text) = "" Then
            LblErrMsg = DisplayMsg(1028) '"Please Select Currency"
            Exit Sub
        End If
        
        cbocurr.Text = cbocurr.Text
        If cbocurr.MatchFound = False Then
            LblErrMsg = DisplayMsg(4005) '"Record with This Currency Not found"
            Exit Sub
        End If
        
        
        If txtamount.Text = "" Or IsNumeric(txtamount) = False Then
          txtamount.SetFocus
          LblErrMsg = DisplayMsg(1094)  '"Please Input Amount"
          Exit Sub
        End If
        
        If CDbl(txtamount.Text) > gd_MaxAmount Then
          txtamount.SetFocus
          LblErrMsg = DisplayMsg(4051) & " " & gd_MaxAmount
          Exit Sub
        End If
        txtamount = Format(txtamount, gs_formatAmount)
        
        'End Validasi
        
        sql = "Select * From InventoryCost_Item"
        If RS.State <> adStateClosed Then RS.Close
        RS.Open sql, Db, 1, 3
        
        RS.filter = "Item_Code = '" & Trim(CboProduct.Text) & "' " & _
                    "And Cost_Cls = '" & Trim(CboCost.Text) & "' " & _
                    "And Start_Date = '" & Format(StartDate.Value, "yyyymmdd") & "' "
        If ubah = False Then
            If Not (RS.EOF And RS.BOF) Then
                LblErrMsg = DisplayMsg(1023): StartDate.SetFocus: Exit Sub
            Else
                cektgl
                If NonValidDate Then Exit Sub
                RS.AddNew
            End If
        Else
            RS.filter = "Item_Code = '" & Trim(CboProduct.Text) & "' " & _
                    "And Cost_Cls = '" & Trim(CboCost.Text) & "' " & _
                    "And Start_Date = '" & StartDateAwal & "' "
        End If
        
        cektgl
        If NonValidDate Then Exit Sub
        
        RS("Item_Code") = Trim(CboProduct.Text)
        RS("Cost_Cls") = Trim(CboCost.Text)
        RS("Start_Date") = Format(StartDate.Value, "yyyymmdd")
        If mask.Text = "99/99/9999" Then
           If UbahEndDate = True Then
             RS("End_Date") = EndDateAkhir
           Else
             RS("End_Date") = "99999999"
           End If
        Else
            RS("End_Date") = Format(mask.Text, "yyyyddmm")
        End If
        
        RS("Currency_Code") = Trim(cbocurr.Column(0))
        RS("Amount") = CDbl(txtamount)
        RS("Last_Update") = Now
        RS("Last_User") = userLogin
        
        RS.update
        RS.Close
      
        .clear
        Call KosongBawah
        LblErrMsg = DisplayMsg(IIf((ubah = False), 1000, 1101))
        ubah = False
        Call IsiGrid
        CboProduct.locked = False
        CboCost.locked = False
    End With
End Sub

Sub cektgl()
    Dim rsBefore As New Recordset
    Dim rsAfter As New Recordset
    Dim Tgl, TempDate
    
    NonValidDate = False
    UbahEndDate = False
    
    If hapus Then
        sql = "Select * From InventoryCost_Item Where " & _
                "Item_Code = '" & Trim(CboProduct.Text) & "' " & _
                "And Cost_Cls = '" & Trim(CboCost.Text) & "' " & _
                "And Start_Date < '" & SDate & "' Order By Start_Date, End_Date"
        If rsBefore.State <> adStateClosed Then rsBefore.Close
        rsBefore.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    
        sql = "Select * From InventoryCost_Item Where " & _
                "Item_Code = '" & Trim(CboProduct.Text) & "' " & _
                "And Cost_Cls = '" & Trim(CboCost.Text) & "' " & _
                "And Start_Date > '" & SDate & "' Order By Start_Date, End_Date"
        If rsAfter.State <> adStateClosed Then rsAfter.Close
        rsAfter.Open sql, Db, adOpenKeyset, adLockOptimistic
    
        If Not (rsBefore.BOF And rsBefore.EOF) Then
            rsBefore.MoveLast
            If Not (rsAfter.BOF And rsAfter.EOF) Then
                rsAfter.MoveFirst
                Tgl = Mid(rsAfter("Start_Date"), 5, 2) & "/" & Right(rsAfter("Start_Date"), 2) & "/" & Left(rsAfter("Start_Date"), 4)
                TempDate = Format(DateAdd("d", -1, CDate(Tgl)), "yyyymmdd")
        
                sql = "Update InventoryCost_Item Set End_Date = '" & TempDate & "', Last_Update = getdate(), Last_User = '" & userLogin & "' Where " & _
                      "Item_Code = '" & rsBefore("Item_Code") & "' " & _
                      "And Cost_Cls = '" & rsBefore("Cost_Cls") & "' " & _
                      "And Start_Date = '" & rsBefore("Start_Date") & "' "
                Db.Execute sql
        
            Else
                sql = "Update InventoryCost_Item Set End_Date = '99999999', Last_Update = getdate(), Last_User = '" & userLogin & "' Where " & _
                      "Item_Code = '" & rsBefore("Item_Code") & "' " & _
                      "And Cost_Cls = '" & rsBefore("Cost_Cls") & "' " & _
                      "And Start_Date = '" & rsBefore("Start_Date") & "' "
                Db.Execute sql
            End If
        End If
        Exit Sub
    End If
    
    If ubah = False Then
        SDate = Format(StartDate.Value, "yyyymmdd")
        EDate = Format(mask.Text, "yyyymmdd")
        
        sql = "Select * From InventoryCost_Item Where " & _
                "Item_Code = '" & Trim(CboProduct.Text) & "' " & _
                "And Cost_Cls = '" & Trim(CboCost.Text) & "' " & _
                "And Start_Date < '" & SDate & "' Order By Start_Date, End_Date"
        If rsBefore.State <> adStateClosed Then rsBefore.Close
        rsBefore.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    
        sql = "Select * From InventoryCost_Item Where " & _
                "Item_Code = '" & Trim(CboProduct.Text) & "' " & _
                "And Cost_Cls = '" & Trim(CboCost.Text) & "' " & _
                "And Start_Date > '" & SDate & "' Order By Start_Date, End_Date"
        If rsAfter.State <> adStateClosed Then rsAfter.Close
        rsAfter.Open sql, Db, adOpenKeyset, adLockOptimistic
    
        If Not (rsAfter.BOF And rsAfter.EOF) Then
            rsAfter.MoveFirst
            
            Tgl = Mid(rsAfter("Start_Date"), 5, 2) & "/" & Right(rsAfter("Start_Date"), 2) & "/" & Left(rsAfter("Start_Date"), 4)
            TempDate = Format(CDate(Tgl), "yyyymmdd")
            
            If EDate = "99/99/9999" Then
                UbahEndDate = True
                EndDateAkhir = Format(DateAdd("d", -1, CDate(Tgl)), "yyyymmdd")
            Else
                If (EDate >= TempDate) Then
                    LblErrMsg.Caption = DisplayMsg(8054) & " " & Format(CDate(Tgl), "dd MMM yyyy")
                    NonValidDate = True
                    EndDate.SetFocus
                    mask.SetFocus
                    Exit Sub
                End If
            End If
        End If
    
    
        If Not (rsBefore.BOF And rsBefore.EOF) Then
            rsBefore.MoveLast
            TempDate = Format(DateAdd("d", -1, CDate(StartDate.Value)), "yyyymmdd")
            
            sql = "Update InventoryCost_Item Set End_Date = '" & TempDate & "', Last_Update = getdate(), Last_User = '" & userLogin & "' Where " & _
                    "Item_Code = '" & rsBefore("Item_Code") & "' " & _
                    "And Cost_Cls = '" & rsBefore("Cost_Cls") & "' " & _
                    "And Start_Date = '" & rsBefore("Start_Date") & "' "
            Db.Execute sql
        End If
        Exit Sub
    Else
    
        SDate = Format(StartDate.Value, "yyyymmdd")
        EDate = Format(mask.Text, "yyyymmdd")
        
        sql = "Select * From InventoryCost_Item Where " & _
                "Item_Code = '" & Trim(CboProduct.Text) & "' " & _
                "And Cost_Cls = '" & Trim(CboCost.Text) & "' " & _
                "And Start_Date < '" & StartDateAwal & "' Order By Start_Date, End_Date"
        If rsBefore.State <> adStateClosed Then rsBefore.Close
        rsBefore.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    
        sql = "Select * From InventoryCost_Item Where " & _
                "Item_Code = '" & Trim(CboProduct.Text) & "' " & _
                "And Cost_Cls = '" & Trim(CboCost.Text) & "' " & _
                "And Start_Date > '" & StartDateAwal & "' Order By Start_Date, End_Date"
        If rsAfter.State <> adStateClosed Then rsAfter.Close
        rsAfter.Open sql, Db, adOpenKeyset, adLockOptimistic
        
            
        If Not (rsAfter.BOF And rsAfter.EOF) Then
            rsAfter.MoveFirst
            
            Tgl = Mid(rsAfter("Start_Date"), 5, 2) & "/" & Right(rsAfter("Start_Date"), 2) & "/" & Left(rsAfter("Start_Date"), 4)
            TempDate = Format(CDate(Tgl), "yyyymmdd")
            
            If EDate = "99/99/9999" Then
                UbahEndDate = True
                EndDateAkhir = Format(DateAdd("d", -1, CDate(Tgl)), "yyyymmdd")
            Else
                If (EDate >= TempDate) Then
                    LblErrMsg.Caption = DisplayMsg(8054) & " " & Format(CDate(Tgl), "dd MMM yyyy")
                    NonValidDate = True
                    EndDate.SetFocus
                    mask.SetFocus
                    Exit Sub
                End If
            End If
        End If
        
        If Not (rsBefore.BOF And rsBefore.EOF) Then
            rsBefore.MoveLast
            Tgl = Mid(rsBefore("Start_Date"), 5, 2) & "/" & Right(rsBefore("Start_Date"), 2) & "/" & Left(rsBefore("Start_Date"), 4)
            TempDate = Format(CDate(Tgl), "yyyymmdd")
    
            If (SDate <= TempDate) Then
                LblErrMsg.Caption = DisplayMsg(8055) & " " & Format(CDate(Tgl), "dd MMM yyyy")
                NonValidDate = True
                StartDate.SetFocus
                Exit Sub
            Else
            
            TempDate = Format(DateAdd("d", -1, CDate(StartDate.Value)), "yyyymmdd")
            sql = "Update InventoryCost_Item Set End_Date = '" & TempDate & "', Last_Update = getdate(), Last_User = '" & userLogin & "' Where " & _
                    "Item_Code = '" & rsBefore("Item_Code") & "' " & _
                    "And Cost_Cls = '" & rsBefore("Cost_Cls") & "' " & _
                    "And Start_Date = '" & rsBefore("Start_Date") & "' "
            Db.Execute sql
            
            End If
        End If
        
        Exit Sub
    End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Cell(flexcpBackColor, Row, Col) <> vbWhite Then Cancel = True
End Sub

Private Sub grid_Click()
    nilKosong = True
    With grid
        LblErrMsg = ""
        
        If .Row > 0 Then
            If .Cell(flexcpBackColor, .Row, .Col) = vbWhite Then .FocusRect = flexFocusInset Else .FocusRect = flexFocusNone
        End If
    End With
    nilKosong = False
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With grid
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") _
            And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then
                KeyAscii = 0
                Exit Sub
        End If
    End With
End Sub

Public Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim setRow As Integer
    
    If nilKosong Then Exit Sub
    
    With grid
        If Col = ColS Then
            If .TextMatrix(Row, ColS) = "S" Then
                CboProduct = Trim(.TextMatrix(Row, ColProCD))
                lblProduct = Trim(.TextMatrix(Row, ColProDes))
                CboCost = Trim(.TextMatrix(Row, cOlcost))
                LblCost = Trim(.TextMatrix(Row, ColCostDes))
                StartDate.Value = Format(Trim(.TextMatrix(Row, ColStart)), "dd MMM yyyy")
                StartDateAwal = Format(Trim(.TextMatrix(Row, ColStart)), "yyyymmdd")
                mask.Text = Format(.TextMatrix(Row, ColEnd), "dd/mm/yyyy")
                If Trim(.TextMatrix(Row, ColEnd)) <> "99/99/9999" Then
                    EndDate = Format(Trim(.TextMatrix(Row, ColEnd)), "dd MMM yyyy")
                    mask.Text = Format(.TextMatrix(Row, ColEnd), "dd/MM/yyyy")
                End If
                cbocurr.Text = Trim(.TextMatrix(Row, ColCurrDes))
                txtamount.Text = Trim(.TextMatrix(Row, ColAmount))
                Call kosongColGrid(Row)
                CboProduct.locked = True
                CboCost.locked = True
                ubah = True
            ElseIf .TextMatrix(Row, ColS) = "D" Then
                Call kosongColGrid(, "S")
                CboProduct.locked = True
                CboCost.locked = True
                ubah = False
            Else
                CboProduct.locked = False
                CboCost.locked = False
            End If
        End If
    End With
End Sub

Private Sub kosongColGrid(Optional Row As Long, Optional Kolom As String)
    With grid
        .Col = 0
        If Kolom <> "" Then
            For i = 1 To .Rows - 1
               If .Text = Kolom Then .Text = ""
               If .TextMatrix(i, 0) <> "D" Then .TextMatrix(i, 0) = ""
            Next i
            KosongBawah
        Else
            For i = 1 To .Rows - 1
               If .TextMatrix(i, 0) <> "" Then .TextMatrix(i, 0) = ""
            Next i
            .TextMatrix(Row, 0) = "S"
        End If
    End With
End Sub

Sub Kosong(Optional stAwal As Byte)
    If stAwal = 1 Then
        LblErrMsg = ""
        LblGroupFilter = ""
        LblProductFilter = ""
        CboGroupFilter.ListIndex = -1
        CboProductFilter.Text = ""
        CboProductFilter.ListIndex = -1
    End If
    CboProduct.locked = False
    CboCost.locked = False
    Call KosongBawah
End Sub

Sub KosongBawah()
    lblProduct = ""
    LblCost = ""
    StartDate.Value = Format(DateValue(Year(Now) & "/" & Month(Now) & "/01"), "dd MMM YYYY")
    EndDate.Value = Format(Now, "dd MMM YYYY")
    mask.Text = "99/99/9999"
    txtamount.Text = ""
    CboProduct.ListIndex = -1
    CboCost.ListIndex = -1
    cbocurr.ListIndex = -1
End Sub

Function SqlSintaks() As String
    Dim SqlS As String
    Dim SqlGroup As String, sqlProduct As String
    SqlGroup = "": sqlProduct = ""
    
    If UCase(Trim(CboGroupFilter.Text)) <> strAll Then
        SqlGroup = " And IM.Group_Cls = '" & Trim(CboGroupFilter.Column(0)) & "' "
    End If
    
    If UCase(Trim(CboProductFilter.Text)) <> strAll Then
        sqlProduct = " And IM.Item_Code = '" & Trim(CboProductFilter.Text) & "' "
    End If
    
    SqlS = "Select ICI.Item_Code, IM.Item_Name Item_Desc, ICI.Cost_Cls, ICM.Description Cost_Desc, Start_Date, End_Date, " & _
            "Currency_Code, (Select Description From Curr_Cls Where Curr_Cls = Currency_Code) Currency_Desc, " & _
            "Amount " & _
            "From InventoryCost_Master ICM " & _
            "Inner Join InventoryCost_Item ICI on ICM.Cost_Cls = ICI.Cost_Cls " & _
            "Inner Join Item_Master IM on ICI.Item_Code = IM.Item_Code "
    SqlS = SqlS + SqlGroup + sqlProduct
    SqlS = SqlS + "Order by ICI.Item_Code, ICI.Cost_Cls, Start_Date, Currency_Code "
    
    SqlSintaks = SqlS
End Function

Private Sub CboGroupFilter_Change()
    LblErrMsg = ""
    LblGroupFilter.Caption = ""
    CboProductFilter.clear
End Sub

Private Sub CboGroupFilter_Click()
    If CboGroupFilter.ListIndex <> -1 Then
        LblGroupFilter.Caption = CboGroupFilter.Column(1)
        Call IsiCboProductFilter(CboGroupFilter.Column(0))
        CboProductFilter.ListIndex = -1
    End If
    Call Header
End Sub

Private Sub CboGroupFilter_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call CboGroupFilter_Click
End Sub

Private Sub CboProdustFilter_Change()
    LblErrMsg.Caption = ""
    LblProductFilter.Caption = ""
End Sub

Private Sub CboProductFilter_Click()
    If CboProductFilter.ListIndex <> -1 Then _
        LblProductFilter.Caption = CboProductFilter.Column(1)
    Call Header
End Sub

Private Sub CboProductFilter_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call CboProductFilter_Click
End Sub

Private Sub CboProduct_Change()
    LblErrMsg = ""
    lblProduct = ""
End Sub

Private Sub CboProduct_Click()
    LblErrMsg = ""
    If CboProduct.ListIndex <> -1 Then
        lblProduct = CboProduct.Column(1)
    End If
End Sub

Private Sub CboProduct_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call CboProduct_Click
End Sub

Private Sub CboProduct_LostFocus()
    Call CboProduct_Click
End Sub

Private Sub CboCost_Change()
    LblErrMsg = ""
    LblCost = ""
End Sub

Private Sub CboCost_Click()
    LblErrMsg = ""
    If CboCost.ListIndex <> -1 Then
        LblCost = CboCost.Column(1)
    End If
End Sub

Private Sub CboCost_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call CboCost_Click
End Sub

Private Sub CboCost_LostFocus()
    Call CboCost_Click
End Sub

Private Sub StartDate_Change()
    If mask.Text <> "99/99/9999" Then
        LblErrMsg.Caption = ""
        If CDate(StartDate.Value) > CDate(EndDate.Value) Then
           LblErrMsg.Caption = DisplayMsg(4068) '"Start Date must be lower than " & Format(EndDate, "dd MMM yyyy")
           Exit Sub
        End If
    End If
End Sub

Private Sub StartDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then StartDate_Change
End Sub

Private Sub EndDate_Change()
    LblErrMsg.Caption = ""
    mask.Text = Format(EndDate, "dd/mm/yyyy")
    If CDate(EndDate) < CDate(StartDate) Then
       LblErrMsg.Caption = DisplayMsg(4066) '"End Date must be higher than " & Format(StartDate, "dd MMM yyyy")
       Exit Sub
    End If
End Sub

Private Sub Enddate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then EndDate_Change
End Sub

Private Sub mask_LostFocus()
    If IsDate(mask.Text) = True Then EndDate.Value = CDate(Format(mask.Text, "dd mm yyyy"))
End Sub

Private Sub TxtAmount_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> Asc(".") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    
End Sub


Private Sub cmdCancel_Click()
    Call Kosong
    Call IsiGrid
End Sub

Private Sub cmdClear_Click()
    Call Kosong(1)
    Call Header
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then Unload Me Else LblErrMsg.Caption = ErrMsg
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

