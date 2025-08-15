VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmProductionScheduleCalculation 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Production Schedule Calculation"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   15120
   Icon            =   "FrmProductionScheduleCalculation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chC 
      BackColor       =   &H00FDDFE3&
      Height          =   210
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   195
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13043
      TabIndex        =   13
      Top             =   270
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H0080FFFF&
      Caption         =   "Calculate"
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
      Left            =   12540
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10260
      Width           =   1140
   End
   Begin VB.CommandButton cmdDetail 
      BackColor       =   &H0080FFFF&
      Caption         =   "To Detail"
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
      Left            =   13770
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10260
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   293
      TabIndex        =   16
      Top             =   9510
      Width           =   14625
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
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   14265
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
      Left            =   293
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10260
      Width           =   1140
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H0080FFFF&
      Caption         =   "Generate Prod. Schedule"
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10260
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2370
      Left            =   293
      TabIndex        =   15
      Top             =   900
      Width           =   14625
      Begin VB.CommandButton cmdSearch 
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
         Left            =   3255
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1815
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtStart 
         Height          =   315
         Left            =   2070
         TabIndex        =   1
         Top             =   690
         Width           =   1500
         _ExtentX        =   2646
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtEnd 
         Height          =   315
         Left            =   4530
         TabIndex        =   2
         Top             =   690
         Width           =   1500
         _ExtentX        =   2646
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   ": Skip"
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
         Left            =   11490
         TabIndex        =   27
         Top             =   1395
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   ": Processed"
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
         Left            =   11490
         TabIndex        =   26
         Top             =   1890
         Width           =   1005
      End
      Begin VB.Shape shapeProcessed 
         BackColor       =   &H003CF4CA&
         BackStyle       =   1  'Opaque
         Height          =   435
         Left            =   10875
         Shape           =   5  'Rounded Square
         Top             =   1770
         Width           =   495
      End
      Begin VB.Shape shapeSkip 
         BackColor       =   &H00C0FBFC&
         BackStyle       =   1  'Opaque
         Height          =   435
         Left            =   10875
         Shape           =   5  'Rounded Square
         Top             =   1275
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
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
         Left            =   210
         TabIndex        =   25
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3780
         TabIndex        =   24
         Top             =   750
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "Calculate"
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
         Left            =   210
         TabIndex        =   23
         Top             =   1905
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   22
         Top             =   1537
         Width           =   795
      End
      Begin MSForms.ComboBox cboCalculate 
         Height          =   315
         Left            =   2070
         TabIndex        =   5
         Top             =   1845
         Width           =   855
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "1508;556"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboGenerate 
         Height          =   315
         Left            =   2070
         TabIndex        =   4
         Top             =   1470
         Width           =   855
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "1508;556"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
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
         Left            =   210
         TabIndex        =   21
         Top             =   750
         Width           =   1275
      End
      Begin MSForms.ComboBox cboPo 
         Height          =   315
         Left            =   2070
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2355
         VariousPropertyBits=   612386843
         MaxLength       =   35
         DisplayStyle    =   3
         Size            =   "4154;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "PO No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   210
         TabIndex        =   20
         Top             =   1155
         Width           =   1755
      End
      Begin VB.Line Line2 
         X1              =   4530
         X2              =   7920
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblCustomer 
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "lblCustomer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4530
         TabIndex        =   19
         Top             =   345
         Width           =   3495
      End
      Begin MSForms.ComboBox cboCustomer 
         Height          =   315
         Left            =   2070
         TabIndex        =   0
         Top             =   300
         Width           =   2355
         VariousPropertyBits=   612386843
         MaxLength       =   5
         DisplayStyle    =   3
         Size            =   "4154;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Customer Code"
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
         Left            =   2910
         TabIndex        =   18
         Top             =   330
         Width           =   1515
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid griD 
      Height          =   5835
      Left            =   300
      TabIndex        =   8
      Top             =   3450
      Width           =   14625
      _cx             =   25797
      _cy             =   10292
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
      BackColor       =   12632256
      ForeColor       =   -2147483640
      BackColorFixed  =   10932991
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483624
      BackColorAlternate=   12632256
      GridColor       =   8421504
      GridColorFixed  =   8421504
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   5
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   275
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
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
      WordWrap        =   -1  'True
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
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   120
      Left            =   300
      TabIndex        =   28
      Top             =   9360
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   212
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Production Schedule Calculation"
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
      Left            =   293
      TabIndex        =   14
      Top             =   270
      Width           =   14580
   End
End
Attribute VB_Name = "FrmProductionScheduleCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lb_col_Check As Byte
Dim lb_col_Description As Byte
Dim lb_col_PartNumber As Byte
Dim lb_col_Customer As Byte
Dim lb_col_PoNo As Byte
Dim lb_col_Qty As Byte
Dim lb_col_DeliveryDate As Byte
Dim lb_col_Currency As Byte
Dim lb_col_Price As Byte
Dim lb_col_Amount As Byte
Dim lb_col_Calculate As Byte
Dim lb_col_Generate As Byte
Dim lb_col_ProdCls As Byte
Dim lb_col_OrderSeqNo As Byte
Dim lb_col_Lot As Byte
Dim lb_col_Lead As Byte

Dim lc_Green As ColorConstants
Dim lc_Yellow As ColorConstants

Dim lb_afterInit As Boolean
Dim lb_flagCheckGrid As Boolean

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    prgBar.Visible = False
    Call S_iniSialisasi
    Call S_heaDer
    Call S_setCombo
    
    lb_afterInit = True
    If cboCustomer.ListCount > 0 Then cboCustomer.ListIndex = 0
End Sub

Private Sub CmdSubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub cmdDetail_Click()
    Dim li_baris As Long
    Dim li_ActiveRow As Long
    
    li_ActiveRow = 0
    
    For li_baris = grid.FixedRows To grid.Rows - 1
        If grid.Cell(flexcpChecked, li_baris, lb_col_Check) = flexChecked Then
            li_ActiveRow = li_baris
            Exit For
        End If
    Next li_baris
    
    FrmProductionScheduleCalculationDetail.DtStart.Value = DtStart.Value
    FrmProductionScheduleCalculationDetail.DtEnd.Value = DtEnd.Value
    FrmProductionScheduleCalculationDetail.cboCustomer.Text = IIf(li_ActiveRow <> 0, grid.TextMatrix(li_ActiveRow, lb_col_Customer), Trim$(cboCustomer.Text))
    If li_ActiveRow = 0 Then FrmProductionScheduleCalculationDetail.cboCustomer_Change
    FrmProductionScheduleCalculationDetail.cboPo.Text = IIf(li_ActiveRow <> 0, grid.TextMatrix(li_ActiveRow, lb_col_PoNo), Trim$(cboPo.Text))
    FrmProductionScheduleCalculationDetail.CboItemCode.Text = IIf(li_ActiveRow <> 0, grid.TextMatrix(li_ActiveRow, lb_col_PartNumber), "ALL")
    FrmProductionScheduleCalculationDetail.cmdsubmenu.Caption = "&Back"
    Me.Hide
    FrmProductionScheduleCalculationDetail.Show
    FrmProductionScheduleCalculationDetail.cmdSearch.SetFocus
    SendKeys (vbCr)
End Sub

Private Sub S_iniSialisasi()
    lb_col_Check = 0
    lb_col_PartNumber = 1
    lb_col_Description = 2
    lb_col_Customer = 3
    lb_col_PoNo = 4
    lb_col_Qty = 5
    lb_col_DeliveryDate = 6
    lb_col_Calculate = 7
    lb_col_Generate = 8
    lb_col_Currency = 9
    lb_col_Price = 10
    lb_col_Amount = 11
    lb_col_ProdCls = 12
    lb_col_OrderSeqNo = 13
    lb_col_Lot = 14
    lb_col_Lead = 15
    
    DtStart.Value = Format(Now, "1 MMM yyyy")
    DtEnd.Value = Format(Now, "dd MMM yyyy")
    
'    lc_Green = RGB(202, 244, 104)
    lc_Green = RGB(202, 244, 60)
'    lc_Yellow = RGB(250, 247, 141)
    lc_Yellow = RGB(252, 251, 192)
End Sub

Private Sub S_heaDer()
Dim li_i As Integer
With grid
    .clear
    .Rows = 2
    .ColS = 16
    .FixedRows = 2
    .FrozenCols = 3
    .MergeCells = flexMergeFixedOnly
    .SelectionMode = flexSelectionByRow
    .FocusRect = flexFocusLight
    
    .TextMatrix(0, lb_col_Check) = " "
    .TextMatrix(0, lb_col_PartNumber) = "Part Number"
    .TextMatrix(0, lb_col_Description) = "Description"
    .TextMatrix(0, lb_col_Customer) = "Customer"
    .TextMatrix(0, lb_col_PoNo) = "PO No"
    .TextMatrix(0, lb_col_Qty) = "Qty"
    .TextMatrix(0, lb_col_DeliveryDate) = "Delivery" & vbLf & "Date"
    .TextMatrix(0, lb_col_Currency) = "Curr"
    .TextMatrix(0, lb_col_Price) = "Price"
    .TextMatrix(0, lb_col_Amount) = "Amount"
    .TextMatrix(0, lb_col_Calculate) = "Calculate"
    .TextMatrix(0, lb_col_Generate) = "Generate"
    .TextMatrix(0, lb_col_ProdCls) = "ProdCls"
    .TextMatrix(0, lb_col_OrderSeqNo) = "Order" & vbLf & "SeqNo"
    .TextMatrix(0, lb_col_Lot) = "Lot"
    .TextMatrix(0, lb_col_Lead) = "Lead"
    
    For li_i = 0 To .ColS - 1
        .TextMatrix(1, li_i) = .TextMatrix(0, li_i)
        .MergeCol(li_i) = True
    Next li_i
         
    .ColWidth(lb_col_Check) = 270
    .ColWidth(lb_col_PartNumber) = 1500
    .ColWidth(lb_col_Description) = 2400
    .ColWidth(lb_col_Customer) = 1300
    .ColWidth(lb_col_PoNo) = 2300
    .ColWidth(lb_col_Qty) = 1000
    .ColWidth(lb_col_DeliveryDate) = 1500
    .ColWidth(lb_col_Calculate) = 900
    .ColWidth(lb_col_Generate) = 900
    .ColWidth(lb_col_Currency) = 700
    .ColWidth(lb_col_Price) = 1800
    .ColWidth(lb_col_Amount) = 2000
    .ColWidth(lb_col_ProdCls) = 900
    .ColWidth(lb_col_OrderSeqNo) = 900
    .ColWidth(lb_col_Lot) = 900
    .ColWidth(lb_col_Lead) = 900
    
    .ColHidden(lb_col_ProdCls) = True
    .ColHidden(lb_col_OrderSeqNo) = True
    .ColHidden(lb_col_Lot) = True
    .ColHidden(lb_col_Lead) = True
    
    .ColAlignment(lb_col_DeliveryDate) = flexAlignCenterCenter
    .ColAlignment(lb_col_Currency) = flexAlignCenterCenter
    .ColAlignment(lb_col_Calculate) = flexAlignCenterCenter
    .ColAlignment(lb_col_Generate) = flexAlignCenterCenter
    
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .ColS - 1) = flexAlignCenterCenter
    
End With
End Sub

Private Sub S_setCombo()
    Call S_addCboCust
'    Call S_addCboPo
    
    cboGenerate.AddItem "ALL"
    cboGenerate.AddItem "YES"
    cboGenerate.AddItem "NO"
    cboGenerate.ListIndex = 0
    
    cboCalculate.AddItem "ALL"
    cboCalculate.AddItem "YES"
    cboCalculate.AddItem "NO"
    cboCalculate.ListIndex = 0
End Sub

Private Sub S_addCboCust()
Dim ls_Que As String
Dim RsCust As New Recordset
Dim i As Integer

    ls_Que = "select trade_code, trade_name, address1 from trade_master where trade_cls in ('2', '4')"
    
    RsCust.CursorLocation = adUseClient
    RsCust.Open ls_Que, Db, adOpenKeyset, adLockReadOnly
    
    With cboCustomer
        .clear
        .columnCount = 3
        .ColumnWidths = "50pt;300pt;0pt"
        .ListWidth = 350
        .ListRows = 15
        
        i = 0
        If RsCust.RecordCount > 0 Then
            .AddItem
            .List(i, 0) = "ALL"
            .List(i, 1) = "ALL"
            i = i + 1
            While Not RsCust.EOF
                .AddItem
                .List(i, 0) = Trim(RsCust("Trade_code"))
                .List(i, 1) = IIf(IsNull(RsCust("trade_name")), " ", Trim(RsCust("Trade_Name")))
                .List(i, 2) = IIf(IsNull(RsCust("address1")), " ", Trim(RsCust("Address1")))
                RsCust.MoveNext
                i = i + 1
            Wend
        End If
    End With
    
    If RsCust.State = adStateOpen Then RsCust.Close
    Set RsCust = Nothing
End Sub

Private Sub S_addCboPo()
Dim ls_Que As String
Dim rsPO As New Recordset
Dim i As Long

    ls_Que = " declare @dtStart datetime " & vbCrLf & _
                " declare @dtEnd datetime " & vbCrLf & _
                "  " & vbCrLf & _
                " set @dtStart = '" & Format(DtStart, "yyyy-mm-dd") & "' " & vbCrLf & _
                " set @dtEnd = '" & Format(DtEnd, "yyyy-mm-dd") & "' " & vbCrLf & _
                "  " & vbCrLf & _
                " select distinct om.Po_No " & vbCrLf & _
                " from orderentry_master om " & vbCrLf & _
                " inner join orderentry_detail od " & vbCrLf & _
                " on om.cust_Code = od.Cust_Code " & vbCrLf & _
                "   and om.Po_No = od.Po_no " & vbCrLf & _
                " where datediff(d,@dtStart,od.delivery_date)>=0 " & vbCrLf & _
                "     and datediff(d,@dtEnd,od.delivery_date)<=0 " & vbCrLf & _
                "     and isnull(om.fix_cls,'0')='0' "
    
    If cboCustomer.ListIndex <> 0 Then
        ls_Que = ls_Que & vbCrLf & _
                "   and om.Cust_Code = '" & Trim$(cboCustomer.Text) & "' "
    End If
    
    rsPO.CursorLocation = adUseClient
    rsPO.Open ls_Que, Db, adOpenKeyset, adLockReadOnly
    
    With cboPo
        .clear
        .ColumnWidths = "115pt"
        .ListWidth = 115
        .ListRows = 15
        
        i = 0
        If rsPO.RecordCount > 0 Then
            .AddItem "ALL"
            i = i + 1
            While Not rsPO.EOF
                .AddItem Trim(rsPO("PO_No") & "")
                rsPO.MoveNext
                i = i + 1
            Wend
        End If
        If .ListCount > 1 Then .ListIndex = 0
    End With
    
    If rsPO.State = adStateOpen Then rsPO.Close
    Set rsPO = Nothing
End Sub

Private Sub cboCustomer_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboPo_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboCustomer_Change()
    If lb_afterInit Then
        If cboCustomer.MatchFound Then
            LblCustomer = cboCustomer.Column(1)
            LblErrMsg = ""
        Else
            LblCustomer = ""
            LblErrMsg = DisplayMsg(4072) 'Customer Code is not found
        End If
        S_addCboPo
    End If
    grid.Rows = grid.FixedRows
End Sub

Private Sub cboPo_Change()
    If lb_afterInit Then
        If cboPo.Text <> "" And Not cboPo.MatchFound Then
            LblErrMsg = DisplayMsg(1210) 'Po No is not found
        End If
        grid.Rows = grid.FixedRows
    End If
End Sub

Private Sub dtStart_Change()
    If lb_afterInit Then
        Call S_addCboPo
        grid.Rows = grid.FixedRows
    End If
End Sub

Private Sub dtEnd_Change()
    If lb_afterInit Then
        Call S_addCboPo
        grid.Rows = grid.FixedRows
    End If
End Sub

Private Sub cboCalculate_Change()
    If lb_afterInit Then
        grid.Rows = grid.FixedRows
    End If
End Sub

Private Sub cboGenerate_Change()
    If lb_afterInit Then
        grid.Rows = grid.FixedRows
    End If
End Sub

Private Function F_headerValidation() As Boolean
    If Not cboCustomer.MatchFound Then
        LblErrMsg = DisplayMsg(1045) 'please select cust code
        cboCustomer.SetFocus
        Exit Function
    End If
    If Not cboPo.MatchFound Then
        LblErrMsg = DisplayMsg(1048) 'please select po No
        cboPo.SetFocus
        Exit Function
    End If
    F_headerValidation = True
End Function

Private Sub cmdSearch_Click()
    LblErrMsg = ""
    If F_headerValidation Then
        Call S_fillGrid
    End If
End Sub

Private Sub S_fillGrid()
Dim ls_Que As String
Dim rsGrid As New ADODB.Recordset

    ls_Que = " declare @dtStart datetime " & vbCrLf & _
                " declare @dtEnd datetime " & vbCrLf & _
                "  " & vbCrLf & _
                " set @dtStart = '" & Format(DtStart, "yyyy-mm-dd") & "' " & vbCrLf & _
                " set @dtEnd = '" & Format(DtEnd, "yyyy-mm-dd") & "' " & vbCrLf & _
                "  " & vbCrLf & _
                " select od.item_Code, im.item_Name, od.Cust_Code, od.Po_No, od.Qty, od.Delivery_Date, " & vbCrLf & _
                "             od.Currency_Code, cc.description Currency, od.Price, od.Amount,  " & vbCrLf & _
                "             Calculate = case when isnull(od.calculate_Cls,'0') = '0' then 'No' else 'Yes' end, " & vbCrLf & _
                "             Generate = case when isnull(od.generate_Cls,'0') = '0' then 'No' else 'Yes' end, " & vbCrLf & _
                "             im.production_Cls, od.seq_no, im.lot_Qty, im. Product_ReadTime " & vbCrLf & _
                " from orderentry_detail od  " & vbCrLf & _
                " inner join orderentry_master om  " & vbCrLf & _
                " on od.po_no=om.po_no   " & vbCrLf & _
                " inner join item_master im  " & vbCrLf & _
                " on od.item_code = im.item_code  " & vbCrLf & _
                " left join Curr_Cls cc " & vbCrLf & _
                " on od.currency_code = cc.curr_cls " & vbCrLf & _
                " where datediff(d,@dtStart,od.delivery_date)>=0 " & vbCrLf & _
                "     and datediff(d,@dtEnd,od.delivery_date)<=0 " & vbCrLf & _
                "     and isnull(om.fix_cls,'0')='0' "
    
    If cboCustomer.ListIndex <> 0 Then
        ls_Que = ls_Que & vbCrLf & _
                "   and om.Cust_Code = '" & Trim$(cboCustomer.Text) & "' "
    End If
    If cboPo.ListIndex <> 0 Then
        ls_Que = ls_Que & vbCrLf & _
                "   and od.po_no = '" & Trim$(cboPo.Text) & "' "
    End If
    If cboCalculate.ListIndex <> 0 Then
        ls_Que = ls_Que & vbCrLf & _
                "   and isnull(od.calculate_Cls,'0') = '" & IIf(UCase(Trim$(cboCalculate.Text & "")) = "YES", 1, 0) & "' "
    End If
    If cboGenerate.ListIndex <> 0 Then
        ls_Que = ls_Que & vbCrLf & _
                "   and isnull(od.generate_Cls,'0') = '" & IIf(UCase(Trim$(cboGenerate.Text & "")) = "YES", 1, 0) & "' "
    End If
        
    ls_Que = ls_Que & vbCrLf & _
            " order by od.item_code, om.cust_code, od.delivery_date, om.po_no "

    Set rsGrid = Db.Execute(ls_Que)
    
    grid.Rows = grid.FixedRows
    If Not (rsGrid.BOF Or rsGrid.EOF) Then
        With grid
            While Not rsGrid.EOF
                .Rows = .Rows + 1
                .Cell(flexcpChecked, .Rows - 1, lb_col_Check) = flexUnchecked
                .TextMatrix(.Rows - 1, lb_col_PartNumber) = Trim$(rsGrid("item_code") & "")
                .TextMatrix(.Rows - 1, lb_col_Description) = Trim$(rsGrid("item_name") & "")
                .TextMatrix(.Rows - 1, lb_col_Customer) = Trim$(rsGrid("cust_code") & "")
                .TextMatrix(.Rows - 1, lb_col_PoNo) = Trim$(rsGrid("po_no") & "")
                .TextMatrix(.Rows - 1, lb_col_Qty) = Format(Val(rsGrid("qty")), gs_formatQty)
                .TextMatrix(.Rows - 1, lb_col_DeliveryDate) = Format(rsGrid("delivery_Date"), "dd MMM yyyy")
                .TextMatrix(.Rows - 1, lb_col_Currency) = Trim$(rsGrid("currency") & "")
                .TextMatrix(.Rows - 1, lb_col_Price) = Format(Val(rsGrid("price")), gs_formatPrice)
                .TextMatrix(.Rows - 1, lb_col_Amount) = Format(Val(rsGrid("amount")), gs_formatAmount)
                .TextMatrix(.Rows - 1, lb_col_Calculate) = Trim$(rsGrid("calculate") & "")
                .TextMatrix(.Rows - 1, lb_col_Generate) = Trim$(rsGrid("generate") & "")
                .TextMatrix(.Rows - 1, lb_col_ProdCls) = Trim$(rsGrid("production_cls") & "")
                .TextMatrix(.Rows - 1, lb_col_OrderSeqNo) = Trim$(rsGrid("Seq_No") & "")
                .TextMatrix(.Rows - 1, lb_col_Lot) = Format(Val(rsGrid("lot_qty")), gs_formatQty)
                .TextMatrix(.Rows - 1, lb_col_Lead) = Trim$(rsGrid("product_readTime") & "")
                
                If UCase(Trim$(rsGrid("generate") & "")) <> UCase("YES") Then
                    .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .ColS - 1) = vbWhite
                Else
'                    If UCase(Trim$(rsGrid("calculate") & "")) = UCase("TRUE") Then
'                        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, lb_col_Generate) = lc_Green
'                    End If
                End If
                
                rsGrid.MoveNext
            Wend
        End With
    Else
        LblErrMsg = DisplayMsg("4006") 'No data that you want to search !
    End If
    
    If rsGrid.State = adStateOpen Then rsGrid.Close
    Set rsGrid = Nothing
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = lb_col_Check Then
        chC.Value = IIf(F_checkC, 1, 0)
        lb_flagCheckGrid = False
    End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> lb_col_Check Then Cancel = True
End Sub

Private Sub griD_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = lb_col_Check Then Cancel = True
End Sub

Private Function F_totalCheck() As Integer
With grid
    Dim li_i As Integer
    Dim li_Temp As Integer
    
    li_Temp = 0
    For li_i = .FixedRows To .Rows - 1
        If .Cell(flexcpChecked, li_i, lb_col_Check) = flexChecked Then li_Temp = li_Temp + 1
    Next li_i
End With
    F_totalCheck = li_Temp
End Function

Private Sub S_lockControl(Flag As Boolean)
    cboCustomer.Enabled = Not Flag
    DtStart.Enabled = Not Flag
    DtEnd.Enabled = Not Flag
    cboPo.Enabled = Not Flag
    cboCalculate.Enabled = Not Flag
    cboGenerate.Enabled = Not Flag
    cmdSearch.Enabled = Not Flag
    cmdCalculate.Enabled = Not Flag
    cmdGenerate.Enabled = Not Flag
    cmdsubmenu.Enabled = Not Flag
    CmdDetail.Enabled = Not Flag
    grid.Enabled = Not Flag
End Sub

Private Sub cmdCalculate_Click()
    Dim lsi_i As Integer
    Dim lss_msG As String
    Dim lsi_totalCheck As Integer
    
    On Local Error GoTo errHandler
    
    lsi_totalCheck = F_totalCheck
    If lsi_totalCheck < 1 Then
        LblErrMsg = DisplayMsg("1211")    'No data to calculate !
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    LblErrMsg = ""
    Call S_lockControl(True)
    
    prgBar.Value = 0
    prgBar.Visible = True
    prgBar.Max = lsi_totalCheck
    
    Dim cls_Calc As New cls_Prod_Sched_Calc
    cls_Calc.StartDate = Format(DtStart, "yyyy-mm-dd")
    cls_Calc.EndDate = Format(DtEnd, "yyyy-mm-dd")
    
    With grid
        For lsi_i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lsi_i, lb_col_Check) = flexChecked Then
                If UCase(.TextMatrix(lsi_i, lb_col_Generate)) <> "YES" Then
'                    prgBar.Refresh
                    prgBar.Value = prgBar.Value + 1
                    
                    lss_msG = cls_Calc.CF_processItem(.TextMatrix(lsi_i, lb_col_Customer), .TextMatrix(lsi_i, lb_col_PartNumber), .TextMatrix(lsi_i, lb_col_OrderSeqNo), _
                                                                        .TextMatrix(lsi_i, lb_col_PoNo), .TextMatrix(lsi_i, lb_col_ProdCls), CDbl(.TextMatrix(lsi_i, lb_col_Lot)), _
                                                                        CDate(.TextMatrix(lsi_i, lb_col_DeliveryDate)), CDbl(.TextMatrix(lsi_i, lb_col_Qty)), .TextMatrix(lsi_i, lb_col_Lead)) & ""
                End If
                If lss_msG <> "" Then
                    Exit For
                End If
            Else
            End If
        Next lsi_i
        
        For lsi_i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lsi_i, lb_col_Check) = flexChecked Then
                .Cell(flexcpChecked, lsi_i, lb_col_Check) = flexUnchecked
                .TextMatrix(lsi_i, lb_col_Calculate) = "Yes"
                If UCase(.TextMatrix(lsi_i, lb_col_Generate)) = "YES" Then
'                    .Cell(flexcpBackColor, lsi_i, 0, lsi_i, .ColS - 1) = lc_Yellow
                Else
                    .Cell(flexcpBackColor, lsi_i, 0, lsi_i, .ColS - 1) = lc_Green
                End If
            Else
'                .Cell(flexcpBackColor, lsi_i, 0, lsi_i, .ColS - 1) = .BackColor
                If UCase(.TextMatrix(lsi_i, lb_col_Generate)) = "YES" Then  'Or UCase(.TextMatrix(lsi_i, lb_col_Calculate)) = "YES"
                    .Cell(flexcpBackColor, lsi_i, 0, lsi_i, .ColS - 1) = .BackColor
                Else
                    .Cell(flexcpBackColor, lsi_i, 0, lsi_i, .ColS - 1) = vbWhite
                End If
            End If
        Next lsi_i
    End With
    
    If lss_msG = "" Then
        LblErrMsg = DisplayMsg("1212") 'Production schedule calculation success !
    Else
        LblErrMsg = lss_msG
    End If
    
normalExit:
    prgBar.Visible = False
    Call S_lockControl(False)
    cmdCalculate.SetFocus
    Set cls_Calc = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg = err.Description
    err.clear
    Resume normalExit
End Sub

Function F_checkC() As Boolean
Dim li_baris As Integer
With grid
    For li_baris = .FixedRows To .Rows - 1
        If .Cell(flexcpChecked, li_baris, lb_col_Check) = flexUnchecked Then
            F_checkC = False
            lb_flagCheckGrid = True
            Exit Function
        End If
    Next li_baris
End With
F_checkC = True
End Function

Private Sub chC_Click()
Dim li_baris As Integer
With grid
    If Not lb_flagCheckGrid Then
    
        For li_baris = .FixedRows To .Rows - 1
                .Cell(flexcpChecked, li_baris, lb_col_Check) = IIf(chC.Value = 1, flexChecked, flexUnchecked)
        Next li_baris
    End If
End With
End Sub
