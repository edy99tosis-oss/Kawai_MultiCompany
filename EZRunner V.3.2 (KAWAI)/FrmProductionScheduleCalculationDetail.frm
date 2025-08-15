VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmProductionScheduleCalculationDetail 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Production Schedule Calculation Detail"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmProductionScheduleCalculationDetail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chG 
      BackColor       =   &H00FDDFE3&
      Height          =   210
      Left            =   630
      TabIndex        =   11
      Top             =   3945
      Width           =   195
   End
   Begin VB.CheckBox chC 
      BackColor       =   &H00FDDFE3&
      Height          =   210
      Left            =   360
      TabIndex        =   10
      Top             =   3945
      Width           =   195
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   120
      Left            =   300
      TabIndex        =   33
      Top             =   9360
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   212
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13043
      TabIndex        =   16
      Top             =   270
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Left            =   13785
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10215
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   293
      TabIndex        =   19
      Top             =   9465
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
         TabIndex        =   20
         Top             =   240
         Width           =   14265
      End
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   293
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10215
      Width           =   1140
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H0080FFFF&
      Caption         =   "Generate Prod. Schedule"
      Height          =   375
      Left            =   11445
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10215
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   293
      TabIndex        =   18
      Top             =   900
      Width           =   14625
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4470
         TabIndex        =   34
         Top             =   1837
         Width           =   300
      End
      Begin VB.CheckBox chItemInformation 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Show Item Information"
         Height          =   210
         Left            =   6285
         TabIndex        =   7
         Top             =   1485
         Width           =   2370
      End
      Begin VB.CheckBox chPartMaterial 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Show Part/Material"
         Height          =   210
         Left            =   6285
         TabIndex        =   6
         Top             =   1185
         Width           =   2370
      End
      Begin VB.CheckBox chDate 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Show valued date only"
         Height          =   210
         Left            =   6285
         TabIndex        =   5
         Top             =   892
         Value           =   1  'Checked
         Width           =   2370
      End
      Begin VB.ListBox lstUncalculated 
         Height          =   1620
         ItemData        =   "FrmProductionScheduleCalculationDetail.frx":0E42
         Left            =   8835
         List            =   "FrmProductionScheduleCalculationDetail.frx":0E49
         TabIndex        =   9
         Top             =   570
         Width           =   2550
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
         Height          =   375
         Left            =   6285
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1800
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtStart 
         Height          =   315
         Left            =   2070
         TabIndex        =   1
         Top             =   840
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
         Top             =   840
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
      Begin VB.Shape sHape 
         BackColor       =   &H003CF4CA&
         BackStyle       =   1  'Opaque
         Height          =   435
         Index           =   1
         Left            =   11550
         Shape           =   5  'Rounded Square
         Top             =   285
         Width           =   495
      End
      Begin VB.Shape sHape 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   1  'Opaque
         Height          =   435
         Index           =   2
         Left            =   11550
         Shape           =   5  'Rounded Square
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   ": Plan Production Schedule"
         Height          =   195
         Left            =   12165
         TabIndex        =   32
         Top             =   900
         Width           =   2295
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   ": Plan Delivery"
         Height          =   195
         Left            =   12165
         TabIndex        =   31
         Top             =   405
         Width           =   1275
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   ": Plan Consumption"
         Height          =   195
         Left            =   12165
         TabIndex        =   30
         Top             =   1395
         Width           =   1680
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   ": Plan Purchase"
         Height          =   195
         Left            =   12165
         TabIndex        =   29
         Top             =   1890
         Width           =   1335
      End
      Begin VB.Shape sHape 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   435
         Index           =   4
         Left            =   11550
         Shape           =   5  'Rounded Square
         Top             =   1770
         Width           =   495
      End
      Begin VB.Shape sHape 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         Height          =   435
         Index           =   3
         Left            =   11550
         Shape           =   5  'Rounded Square
         Top             =   1275
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
         Height          =   255
         Left            =   210
         TabIndex        =   28
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   180
         Left            =   3780
         TabIndex        =   27
         Top             =   907
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "Uncalculated Order(s) :"
         Height          =   180
         Left            =   8835
         TabIndex        =   26
         Top             =   300
         Width           =   2100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
         Height          =   195
         Left            =   210
         TabIndex        =   25
         Top             =   1890
         Width           =   915
      End
      Begin MSForms.ComboBox cboItemCode 
         Height          =   315
         Left            =   2070
         TabIndex        =   4
         Top             =   1830
         Width           =   2355
         VariousPropertyBits=   746604571
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "4154;556"
         MatchEntry      =   1
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
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   900
         Width           =   1275
      End
      Begin MSForms.ComboBox cboPo 
         Height          =   315
         Left            =   2070
         TabIndex        =   3
         Top             =   1335
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
         Height          =   165
         Left            =   210
         TabIndex        =   23
         Top             =   1410
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
         Height          =   225
         Left            =   4530
         TabIndex        =   22
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
         Height          =   255
         Left            =   2910
         TabIndex        =   21
         Top             =   330
         Width           =   1515
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid griD 
      Height          =   5940
      Left            =   300
      TabIndex        =   12
      Top             =   3345
      Width           =   14625
      _cx             =   25797
      _cy             =   10477
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
      Rows            =   25
      Cols            =   10
      FixedRows       =   3
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
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Production Schedule Calculation Detail"
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
      TabIndex        =   17
      Top             =   270
      Width           =   14580
   End
End
Attribute VB_Name = "FrmProductionScheduleCalculationDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lb_col_C As Byte
Dim lb_col_G As Byte
Dim lb_col_PartNumber As Byte
Dim lb_col_Description As Byte
Dim lb_col_Line As Byte
Dim lb_col_Lead As Byte
Dim lb_col_Lot As Byte
Dim lb_col_MinStock As Byte
Dim lb_col_MaxStock As Byte
Dim lb_col_Stock As Byte
Dim lb_col_End As Byte

Dim lb_col_OrderSeqNo As Byte
Dim lb_col_StockAwal As Byte
Dim lb_col_Edited As Byte
Dim lb_col_Value As Byte
Dim lb_col_ProdCls As Byte

Dim lb_Pos_CustCode As Byte
Dim lb_Pos_Po As Byte
Dim lb_Pos_OrderSeq As Byte
Dim lb_Pos_PlanItem As Byte
Dim lb_Pos_PlanSeq As Byte
Dim lb_Pos_PlanCls As Byte
Dim lb_Pos_PlanQty As Byte
Dim lb_Pos_PlanDate As Byte
Dim lb_Pos_PlanItemLineCode As Byte
Dim lb_Pos_PlanItemUnitCls As Byte
Dim lb_Pos_DailyProd As Byte
Dim lb_Pos_PartReceipt As Byte
Dim lb_Pos_FactoryCode As Byte
Dim lb_Pos_MinStock As Byte
Dim lb_Pos_MaxStock As Byte
Dim lb_Pos_LotQty As Byte
Dim lb_Pos_ReqQty As Byte

Dim lc_Green As ColorConstants
Dim lc_Pink As ColorConstants
Dim lc_Cyan As ColorConstants
Dim lc_Blue As ColorConstants
Dim lc_Gray As ColorConstants

Dim lb_afterInit As Boolean
Dim li_StaticCols As Integer
Dim lb_flagCheckGridC As Boolean
Dim lb_flagCheckGridG As Boolean
Dim l_StartDate As Date
Dim l_EndDate As Date
Dim lb_statusFillGrid As Boolean
Dim lb_statusFillCbo As Boolean

Dim MinDate As Date

Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = CboItemCode.Text
 frm_BrowseItem.Show 1
 CboItemCode.Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    Call S_iniSialisasi
    Call S_staticHeader
    Call S_setCombo
    
    lb_afterInit = True
    If cboCustomer.ListCount > 0 Then cboCustomer.ListIndex = 0
End Sub

Private Sub CmdSubMenu_Click()
    If cmdsubmenu.Caption = "&Back" Then
        Unload Me
        FrmProductionScheduleCalculation.Show
    Else
        Unload Me
        frmMainMenu.Show
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub S_iniSialisasi()
    lb_col_C = 0
    lb_col_G = 1
    lb_col_PartNumber = 2
    lb_col_Description = 3
    lb_col_Line = 4
    lb_col_Lead = 5
    lb_col_Lot = 6
    'blank column
    lb_col_MinStock = 8
    'blank column
    lb_col_MaxStock = 10
    lb_col_Stock = 11
    lb_col_OrderSeqNo = 12   'col cadangan
    lb_col_StockAwal = 13
    lb_col_Edited = 14
    lb_col_Value = 15
    lb_col_ProdCls = 16
    lb_col_End = 17
    
    li_StaticCols = 18
    
    lb_Pos_CustCode = 0
    lb_Pos_Po = 1
    lb_Pos_OrderSeq = 2
    lb_Pos_PlanItem = 3
    lb_Pos_PlanSeq = 4
    lb_Pos_PlanCls = 5
    lb_Pos_PlanQty = 6
    lb_Pos_PlanDate = 7
    lb_Pos_PlanItemLineCode = 8
    lb_Pos_PlanItemUnitCls = 9
    lb_Pos_DailyProd = 10
    lb_Pos_PartReceipt = 11
    lb_Pos_FactoryCode = 12
    lb_Pos_MinStock = 13
    lb_Pos_MaxStock = 14
    lb_Pos_LotQty = 15
    lb_Pos_ReqQty = 16
    
    DtStart.Value = Format(Now, "1 MMM yyyy")
    DtEnd.Value = Format(Now, "dd MMM yyyy")
    
    lc_Green = RGB(202, 244, 60)
    lc_Pink = RGB(255, 192, 255)
    lc_Cyan = RGB(192, 255, 255)
    lc_Blue = RGB(128, 128, 255)
    lc_Gray = RGB(223, 223, 223)
    
    sHape(1).BackColor = lc_Green
    sHape(2).BackColor = lc_Pink
    sHape(3).BackColor = lc_Cyan
    sHape(4).BackColor = lc_Blue
    
    prgBar.Visible = False
End Sub

Private Sub S_staticHeader()
Dim li_i As Integer
With grid
    .clear
    .Rows = 3
    .ColS = li_StaticCols
    .FixedRows = 3
    .FrozenCols = lb_col_Stock + 1
'    .MergeCells = flexMergeFixedOnly
    .MergeCells = flexMergeFree
    .SelectionMode = flexSelectionFree
    .FocusRect = flexFocusLight
    .AllowSelection = False
    
    .RowHeight(0) = .RowHeightMax
    .TextMatrix(0, lb_col_C) = "C"
    .TextMatrix(0, lb_col_G) = "G"
    .TextMatrix(0, lb_col_PartNumber) = "Part Number"
    .TextMatrix(0, lb_col_Description) = "Description"
    .TextMatrix(0, lb_col_Line) = "Line"
    .TextMatrix(0, lb_col_Lead) = "Lead" & vbLf & "day(s)"
    .TextMatrix(0, lb_col_Lot) = "Lot"
    .TextMatrix(0, lb_col_MinStock) = "Min. Stock"
    .TextMatrix(0, lb_col_MaxStock) = "Max. Stock"
    .TextMatrix(0, lb_col_Stock) = "Stock"
    .TextMatrix(0, lb_col_OrderSeqNo) = "Order" & vbLf & "SeqNo"
    .TextMatrix(0, lb_col_StockAwal) = "Plan" & vbLf & "SeqNo"
    .TextMatrix(0, lb_col_Edited) = "Edited"
    .TextMatrix(0, lb_col_Value) = "Value"
    .TextMatrix(0, lb_col_ProdCls) = "ProdCls"
    .TextMatrix(0, lb_col_End) = "End"
    
    .RowHeight(1) = .RowHeight(0)
    .RowHeight(2) = .RowHeight(0)
    
    For li_i = 0 To .ColS - 1
        .TextMatrix(1, li_i) = .TextMatrix(0, li_i)
        .TextMatrix(2, li_i) = .TextMatrix(0, li_i)
        .MergeCol(li_i) = True
    Next li_i
         
    .ColWidth(lb_col_C) = 270
    .ColWidth(lb_col_G) = 270
    .ColWidth(lb_col_PartNumber) = 2000
    .ColWidth(lb_col_Description) = 1700
    .ColWidth(lb_col_Line) = 1000
    .ColWidth(lb_col_Lead) = 700
    .ColWidth(lb_col_Lot) = 700
    .ColWidth(lb_col_MinStock) = 700
    .ColWidth(lb_col_MaxStock) = 700
    .ColWidth(lb_col_Stock) = 700
    .ColWidth(lb_col_OrderSeqNo) = 1000
    .ColWidth(lb_col_StockAwal) = 1000
    .ColWidth(lb_col_Edited) = 900
    .ColWidth(lb_col_Value) = 900
    .ColWidth(lb_col_ProdCls) = 900
    .ColWidth(lb_col_End) = 900
    
    .OutlineBar = flexOutlineBarSimpleLeaf
    .OutlineCol = lb_col_PartNumber
    
    .ColHidden(lb_col_OrderSeqNo) = True
    .ColHidden(lb_col_StockAwal) = True
    .ColHidden(lb_col_Edited) = True
    .ColHidden(lb_col_Value) = True
    .ColHidden(lb_col_ProdCls) = True
    .ColHidden(lb_col_Line) = True
    .ColHidden(lb_col_Lead) = True
    .ColHidden(lb_col_Lot) = True
    .ColHidden(lb_col_Lot + 1) = True 'blank column
    .ColHidden(lb_col_MinStock) = True
    .ColHidden(lb_col_MinStock + 1) = True 'blank column
    .ColHidden(lb_col_MaxStock) = True
    .ColHidden(lb_col_End) = True
    
    .ColAlignment(lb_col_Line) = flexAlignCenterCenter
    .ColAlignment(lb_col_Lead) = flexAlignCenterCenter
'    .ColAlignment(lb_col_Edited) = flexAlignCenterCenter
'    .ColAlignment(lb_col_value) = flexAlignCenterCenter
    
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .ColS - 1) = flexAlignCenterCenter
    
End With
End Sub

Private Sub S_setCombo()
    Call S_addCboCust
'    cboPartMaterial.AddItem "YES"   '"INCLUDE"
'    cboPartMaterial.AddItem "NO"    '"EXCLUDE"
'    cboPartMaterial.ListIndex = 1
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
On Local Error GoTo errHandler
    Me.MousePointer = vbHourglass
    lb_statusFillCbo = True

    ls_Que = " declare @dtStart datetime " & vbCrLf & _
                " declare @dtEnd datetime " & vbCrLf & _
                "  " & vbCrLf & _
                " set @dtStart = '" & Format(DtStart, "yyyy-mm-dd") & "' " & vbCrLf & _
                " set @dtEnd = '" & Format(DtEnd, "yyyy-mm-dd") & "' " & vbCrLf & _
                "  " & vbCrLf & _
                " select distinct om.Po_No, isnull(od.calculate_cls,'0') Calc_Cls " & vbCrLf & _
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
        
        lstUncalculated.clear
        
        i = 0
        If rsPO.RecordCount > 0 Then
            .AddItem "ALL"
            i = i + 1
            While Not rsPO.EOF
                If rsPO("calc_Cls") = "1" Then
                    .AddItem Trim$(rsPO("PO_No") & "")
                Else
                    lstUncalculated.AddItem Trim$(rsPO("PO_No") & "")
                End If
                rsPO.MoveNext
                i = i + 1
            Wend
        End If
    End With
    
normalExit:
    lb_statusFillCbo = False
    If cboPo.ListCount > 1 Then cboPo.ListIndex = 0 Else cboPo.clear
    Me.MousePointer = vbDefault
    If rsPO.State = adStateOpen Then rsPO.Close
    Set rsPO = Nothing
    Exit Sub
errHandler:
    LblErrMsg = err.Description
    err.clear
    Resume normalExit
End Sub

Public Sub S_addItemCode()
Dim ls_Que As String
Dim RsItem As New Recordset
Dim i As Long
On Local Error GoTo errHandler
    
    If Not lb_statusFillCbo Xor lb_afterInit Then Exit Sub
    
    Me.MousePointer = vbHourglass

    ls_Que = " declare @dtStart datetime " & vbCrLf & _
                " declare @dtEnd datetime " & vbCrLf & _
                "  " & vbCrLf & _
                " set @dtStart = '" & Format(DtStart, "yyyy-mm-dd") & "' " & vbCrLf & _
                " set @dtEnd = '" & Format(DtEnd, "yyyy-mm-dd") & "' " & vbCrLf & _
                "  " & vbCrLf & _
                " select distinct oed.item_code item_Code, im.Item_Name " & vbCrLf & _
                " from orderentry_detail oed " & vbCrLf & _
                " inner join orderentry_master om " & vbCrLf & _
                " on om.cust_Code = oed.Cust_Code " & vbCrLf & _
                "   and om.Po_No = oed.Po_no " & vbCrLf & _
                "  inner join item_master im " & vbCrLf & _
                "  on im.item_code = oed.item_code " & vbCrLf & _
                " where datediff(d,@dtStart,oed.delivery_date)>=0 " & vbCrLf & _
                "     and datediff(d,@dtEnd,oed.delivery_date)<=0 " & vbCrLf & _
                "     and isnull(om.fix_cls,'0')='0' "
    
    If cboCustomer.ListIndex <> 0 Then
        ls_Que = ls_Que & vbCrLf & _
                "   and om.Cust_Code = '" & Trim$(cboCustomer.Text) & "' "
    End If
    
    If cboPo.ListIndex > 0 Then
        ls_Que = ls_Que & vbCrLf & _
                "   and om.Po_No = '" & Trim$(cboPo.Text) & "' "
    End If
    
    RsItem.CursorLocation = adUseClient
    RsItem.Open ls_Que, Db, adOpenKeyset, adLockReadOnly
    
    With CboItemCode
        .clear
        .columnCount = 2
        .TextColumn = 1
        .ListWidth = 250
        .ColumnWidths = "80 pt;170 pt"
        .ListRows = 15
               
        i = 0
        If RsItem.RecordCount > 0 Then
            .AddItem "ALL"
            i = i + 1
            While Not RsItem.EOF
                .AddItem ""
                .List(i, 0) = Trim$(RsItem("item_Code"))
                .List(i, 1) = Trim$(RsItem("item_Name"))
                RsItem.MoveNext
                i = i + 1
            Wend
        End If
        If .ListCount > 1 Then .ListIndex = 0 Else .clear
    End With
    
normalExit:
    Me.MousePointer = vbDefault
    If RsItem.State = adStateOpen Then RsItem.Close
    Set RsItem = Nothing
    Exit Sub
errHandler:
    LblErrMsg = err.Description
    err.clear
    Resume normalExit
End Sub

Private Sub cboCustomer_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboPo_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboItemCode_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Public Sub cboCustomer_Change()
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
    grid.ColS = li_StaticCols
End Sub

Private Sub cboPo_Change()
    If lb_afterInit Then
        If cboPo.Text <> "" And Not cboPo.MatchFound Then
            LblErrMsg = DisplayMsg(1210) 'Po No is not found
        Else
            S_addItemCode
        End If
        grid.Rows = grid.FixedRows
        grid.ColS = li_StaticCols
    End If
End Sub

Private Sub dtStart_Change()
    If lb_afterInit Then
        Call S_addCboPo
        grid.Rows = grid.FixedRows
        grid.ColS = li_StaticCols
    End If
End Sub

Private Sub dtEnd_Change()
    If lb_afterInit Then
        Call S_addCboPo
        grid.Rows = grid.FixedRows
        grid.ColS = li_StaticCols
    End If
End Sub

Private Sub CboItemCode_Change()
    If lb_afterInit Then
        grid.Rows = grid.FixedRows
        grid.ColS = li_StaticCols
    End If
End Sub

Private Sub chDate_Click()
    If lb_afterInit Then
        grid.Rows = grid.FixedRows
        grid.ColS = li_StaticCols
    End If
End Sub

Private Sub chPartMaterial_Click()
    If lb_afterInit Then
        grid.Rows = grid.FixedRows
        grid.ColS = li_StaticCols
    End If
End Sub

Private Sub chItemInformation_Click()
    If lb_afterInit Then
        grid.Rows = grid.FixedRows
        grid.ColS = li_StaticCols
        grid.ColHidden(lb_col_Line) = IIf(chItemInformation.Value = 0, True, False)
        grid.ColHidden(lb_col_Lead) = IIf(chItemInformation.Value = 0, True, False)
        grid.ColHidden(lb_col_Lot) = IIf(chItemInformation.Value = 0, True, False)
        grid.ColHidden(lb_col_MinStock) = IIf(chItemInformation.Value = 0, True, False)
        grid.ColHidden(lb_col_MaxStock) = IIf(chItemInformation.Value = 0, True, False)
    End If
End Sub

Private Function F_headerValidation() As Boolean
    If Not cboCustomer.MatchFound Then
        LblErrMsg = DisplayMsg("1045") 'please select cust code
        cboCustomer.SetFocus
        Exit Function
    End If
    If Not cboPo.MatchFound Then
        LblErrMsg = DisplayMsg("1048") 'please select po No
        cboPo.SetFocus
        Exit Function
    End If
    If Not CboItemCode.MatchFound Then
        LblErrMsg = DisplayMsg("8082") 'Please select Item Code !
        CboItemCode.SetFocus
        Exit Function
    End If
    F_headerValidation = True
End Function

Private Sub cmdSearch_Click()
    Me.MousePointer = vbHourglass
    prgBar.Value = 0
    prgBar.Visible = True
    LblErrMsg = ""
    If F_headerValidation Then
        Call S_lockControl(True)
        If S_dynamicHeader Then
            Call S_fillGrid
            chC.Value = 1
            chG.Value = IIf(F_checkG, 1, 0)
            If chDate.Value = 1 Then Call S_hiddenBlankDate
        End If
        Call S_lockControl(False)
        cmdSearch.SetFocus
    End If
    prgBar.Visible = False
    Me.MousePointer = vbDefault
End Sub

Private Function F_getPoList() As String
Dim l_i As Integer
Dim strPo As String
    If cboPo.ListIndex = 0 Then
        For l_i = 1 To cboPo.ListCount - 1
            If strPo = "" Then
                strPo = "'" & cboPo.List(l_i, 0) & "'"
            Else
                strPo = strPo & ",'" & cboPo.List(l_i, 0) & "'"
            End If
        Next l_i
    Else
        strPo = "'" & cboPo.Text & "'"
    End If
    F_getPoList = strPo
End Function

Private Function F_getCustList() As String
Dim l_i As Integer
Dim strCust As String
    If cboCustomer.ListIndex = 0 Then
        For l_i = 1 To cboCustomer.ListCount - 1
            If strCust = "" Then
                strCust = "'" & cboCustomer.List(l_i, 0) & "'"
            Else
                strCust = strCust & ",'" & cboCustomer.List(l_i, 0) & "'"
            End If
        Next l_i
    Else
        strCust = "'" & cboCustomer.Text & "'"
    End If
    F_getCustList = strCust
End Function

Private Function S_dynamicHeader() As Boolean
Dim l_Que As String
Dim rsHeader As New ADODB.Recordset
Dim l_hari As Integer

    grid.ColS = li_StaticCols
    
    l_Que = " select start_Date = min(plan_date), end_Date = max(plan_Date) " & vbCrLf & _
            " from ( " & vbCrLf & _
            " select pcd.plan_Date, " & vbCrLf & _
            " im.finishgoodpart_Cls, im.production_cls, im.group_cls, " & vbCrLf & _
            " im.Lot_Qty, im.manufacture_Code Line_Code, im.item_Name, isnull(im.product_readtime,0) Lead_Time, " & vbCrLf & _
            " isnull(oed.generate_Cls,'0') Generate_Cls, oed.delivery_Date " & vbCrLf & _
            " from productioncalculate_detail pcd " & vbCrLf & _
            " inner join item_master im " & vbCrLf & _
            " on pcd.planitem_code = im.item_code " & vbCrLf & _
            " inner join orderentry_detail oed " & vbCrLf & _
            " on pcd.cust_Code = oed.cust_Code " & vbCrLf & _
            "     and pcd.po_No = oed.po_No " & vbCrLf & _
            "     and pcd.seq_No = oed.seq_No " & vbCrLf & _
            " where pcd.po_no in (" & F_getPoList & ") " & vbCrLf & _
            "     and pcd.cust_Code in (" & F_getCustList & ") "
    
    If chPartMaterial.Value <> 1 Then
        l_Que = l_Que & vbCrLf & _
                "   and isnull(im.group_cls,'03') <> '03' " 'Material = 03
    End If
        
    If CboItemCode.ListIndex <> 0 Then
        l_Que = l_Que & vbCrLf & _
                "   and oed.item_Code = '" & CboItemCode.Text & "' "
    End If
        
        l_Que = l_Que & vbCrLf & _
                "   )tbData "
            
    Set rsHeader = Db.Execute(l_Que)
    If Trim$(rsHeader("start_Date") & "") <> "" And Trim$(rsHeader("start_Date") & "") <> "" Then
        l_StartDate = rsHeader("start_Date")
        l_EndDate = rsHeader("end_Date")
    Else
        S_dynamicHeader = False
        LblErrMsg = DisplayMsg("4006") 'No data that you want to search !
        GoTo normalExit
    End If
    
    With grid
    For l_hari = 0 To DateDiff("d", l_StartDate, l_EndDate)
        .ColS = .ColS + 1
        .ColHidden(.ColS - 1) = True
        
        .ColS = .ColS + 1
        .ColWidth(.ColS - 1) = 1000
        .TextMatrix(0, .ColS - 1) = Format(DateAdd("d", l_hari, l_StartDate), "MMMM yyyy")
        .TextMatrix(1, .ColS - 1) = Format(DateAdd("d", l_hari, l_StartDate), "dd - ddd")
        .TextMatrix(2, .ColS - 1) = "REQ" '"IN"
        .ColDataType(.ColS - 1) = flexDTString
        .ColData(.ColS - 1) = Format(DateAdd("d", l_hari, l_StartDate), "yyyy-mm-dd") & ";i"
        .ColAlignment(.ColS - 1) = flexAlignRightCenter
        
        .ColS = .ColS + 1
        .ColWidth(.ColS - 1) = 1000
        .TextMatrix(0, .ColS - 1) = Format(DateAdd("d", l_hari, l_StartDate), "MMMM yyyy")
        .TextMatrix(1, .ColS - 1) = Format(DateAdd("d", l_hari, l_StartDate), "dd - ddd")
        .TextMatrix(2, .ColS - 1) = "PLAN" '"OUT"
        .ColDataType(.ColS - 1) = flexDTString
        .ColData(.ColS - 1) = Format(DateAdd("d", l_hari, l_StartDate), "yyyy-mm-dd") & ";o"
        .ColAlignment(.ColS - 1) = flexAlignRightCenter
        
    Next l_hari
    .MergeRow(0) = True
    .MergeRow(1) = True
    .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .ColS - 1) = flexAlignCenterCenter
    End With
    
    S_dynamicHeader = True
    
normalExit:
    If rsHeader.State = adStateOpen Then rsHeader.Close
    Set rsHeader = Nothing
    Exit Function
End Function

Private Sub S_fillGrid()
Dim ls_Que As String
Dim ls_lastItem As String
Dim ls_lastPo As String
Dim ls_lastCust As String
Dim ls_seqNo As String
Dim li_Hari As Integer
Dim li_TempCol As Integer
Dim rsGrid As New ADODB.Recordset
Dim li_totalRows As Integer

grid.Rows = grid.FixedRows
lb_statusFillGrid = True

    ls_Que = " declare @LastClosing datetime " & vbCrLf & _
                " declare @pDate datetime " & vbCrLf & _
                " declare @selisih numeric(3,0) " & vbCrLf & _
                "  " & vbCrLf & _
                " set @pDate = '" & Format(l_StartDate, "yyyy-mm-dd") & "' " & vbCrLf & _
                "  " & vbCrLf & _
                " select top 1 @lastClosing = cast(inventory_year as char(4)) + '-' + cast(inventory_month as char(2)) + '-01' " & vbCrLf & _
                " from inventory_Control " & vbCrLf & _
                " order by closingDate desc " & vbCrLf & _
                "  " & vbCrLf & _
                " set @selisih = datediff(m,@lastClosing,@pdate) "
    
    ls_Que = ls_Que & vbCrLf & _
                "  " & vbCrLf & _
                " select tbUtama.*, Stock_Awal = isnull(tbStock.Stock_Awal,0), " & vbCrLf & _
                " Req_Qty = isnull(( " & vbCrLf & _
                "     select plan_qty from productioncalculate_detail pcd " & vbCrLf & _
                "    Where pcd.cust_code = tbUtama.cust_code " & vbCrLf & _
                "    and pcd.po_no = tbUtama.po_no " & vbCrLf & _
                "    and pcd.seq_no = tbUtama.seq_no " & vbCrLf & _
                "    and pcd.planitem_code = tbUtama.planitem_code " & vbCrLf & _
                "    and pcd.plan_seqno in ('1', '3')), 0) "
                
    ls_Que = ls_Que & vbCrLf & _
                " from ( " & vbCrLf & _
                "             select pcd.Cust_Code, pcd.PO_No, pcd.Seq_No, pcd.PlanItem_Code, pcd.Plan_SeqNo, pcd.Plan_Cls, " & vbCrLf & _
                "                         pcd.Plan_Date, pcd.Plan_Qty, Generate_Cls = isnull(pcd.generate_Cls,'0'), " & vbCrLf & _
                "                         im.finishgoodpart_Cls, im.production_cls, im.group_cls, im.WH_Code, im.Lot_Qty, im.manufacture_Code, " & vbCrLf & _
                "                         im.Line_Code, im.item_Name, im.Unit_Cls, im.item_Code, im.makebuy_Cls, Lead_Time = isnull(im.product_readtime,0),            " & vbCrLf & _
                "                         oed.delivery_Date, ml.line_Name, tm.Trade_Name, im.Min_Stock, im.Max_Stock,  " & vbCrLf & _
                "                         Daily_Prod_SeqNo = isnull(dp.seq_no,0), Part_Receipt_SeqNo = isnull(pr.seq_no,0)  " & vbCrLf & _
                "             from productioncalculate_detail pcd  " & vbCrLf & _
                "             inner join item_master im  " & vbCrLf & _
                "             on pcd.planitem_code = im.item_code  " & vbCrLf & _
                "             inner join orderentry_detail oed  " & vbCrLf & _
                "             on pcd.cust_Code = oed.cust_Code  " & vbCrLf & _
                "                 and pcd.po_No = oed.po_No  " & vbCrLf & _
                "                 and pcd.seq_No = oed.seq_No  " & vbCrLf & _
                "             left join manufacture_line ml  " & vbCrLf & _
                "             on im.manufacture_Code = ml.manufacture_Code  " & vbCrLf & _
                "                 and im.line_code = ml.line_Code  " & vbCrLf & _
                "             left join trade_Master tm  " & vbCrLf & _
                "             on pcd.Cust_Code = tm.Trade_Code  "
    
    ls_Que = ls_Que & vbCrLf & _
                "             left join daily_production dp  " & vbCrLf & _
                "             on dp.item_code = pcd.planitem_code  " & vbCrLf & _
                "                 and dp.planCust_code = pcd.Cust_Code  " & vbCrLf & _
                "                 and dp.planPo_no = pcd.Po_No  " & vbCrLf & _
                "                 and dp.planpo_Seqno = pcd.seq_no  " & vbCrLf & _
                "                 and dp.plan_Seqno = pcd.plan_seqno  " & vbCrLf & _
                "                 and isnull(auto_cls,'0') = '1'  " & vbCrLf & _
                "             left join part_Receipt pr  " & vbCrLf & _
                "             on pr.dailySeq_no = dp.seq_No  " & vbCrLf & _
                "                 and pr.Receipt_Cls = 'P1'  " & vbCrLf & _
                "             where pcd.po_no in (" & F_getPoList & ")  " & vbCrLf & _
                "                 and pcd.cust_Code in (" & F_getCustList & ")  "
    
    If chPartMaterial.Value <> 1 Then
        ls_Que = ls_Que & vbCrLf & _
                "   and isnull(im.group_cls,'03') <> '03' " 'Material = 03
    End If
    
    If CboItemCode.ListIndex <> 0 Then
        ls_Que = ls_Que & vbCrLf & _
                "   and oed.item_Code = '" & CboItemCode.Text & "' "
    End If
    
    ls_Que = ls_Que & vbCrLf & _
                "             ) tbUtama " & vbCrLf & _
                " left join ( " & vbCrLf & _
                "                 select WH_Code, item_Code, Stock_Awal = Stock_Awal - Total_Order + Remain_PO - Total_Pcd + Total_DP " & vbCrLf & _
                "                 from ( " & vbCrLf & _
                "                             select  im.wh_Code, im.item_Code,  " & vbCrLf & _
                "                                          Stock_Awal = isnull( " & vbCrLf & _
                "                                                                                 case @selisih  " & vbCrLf & _
                "                                                                                     when 0 then a.lm_premonth " & vbCrLf & _
                "                                                                                     when 1 then a.tm_premonth " & vbCrLf & _
                "                                                                                     when 2 then a.nm_premonth " & vbCrLf & _
                "                                                                                     else " & vbCrLf & _
                "                                                                                         case when @selisih < 0 then a.hs_premonth else a.nm_premonth end " & vbCrLf & _
                "                                                                                     end, 0), " & vbCrLf & _
                "                                         Total_Order = isnull(b.total_Order,0), " & vbCrLf & _
                "                                         Total_Pcd = isnull(c.total_Pcd,0), " & vbCrLf & _
                "                                         Remain_PO = isnull(d.Remain_Po,0), " & vbCrLf & _
                "                                         total_DP = isnull(e.total_DP,0) " & vbCrLf & _
                "                             from item_master im "
                      
    ls_Que = ls_Que & vbCrLf & _
                "                             left join   ( " & vbCrLf & _
                "                                       select warehouse_Code, item_code, sum(hs_premonth) hs_premonth, sum(lm_premonth) lm_premonth, sum(tm_premonth) tm_premonth, sum(nm_premonth) nm_premonth from (" & vbCrLf & _
                "                                                 select warehouse_Code, item_code, 0 hs_premonth, lm_premonth, tm_premonth, nm_premonth " & vbCrLf & _
                "                                                 from stock_master " & vbCrLf & _
                "                                                 union all " & vbCrLf & _
                "                                                 select warehouse_Code, item_code, premonth, 0 lm_premonth, 0 tm_premonth, 0 nm_premonth " & vbCrLf & _
                "                                                 from stock_history " & vbCrLf & _
                "                                                 where year(@pdate) = stock_Year " & vbCrLf & _
                "                                                 and month(@pdate) = stock_Month " & vbCrLf & _
                "                                ) tblstockawal group by warehouse_Code, item_code " & vbCrLf & _
                "                                                 ) a " & vbCrLf & _
                "                             on im.item_Code = a.item_Code " & vbCrLf & _
                "                                 and im.wh_Code = a.warehouse_Code "

    ls_Que = ls_Que & vbCrLf & _
                "                             left join ( " & vbCrLf & _
                "                                             select item_Code, total_Order = isnull(sum(qty),0) " & vbCrLf & _
                "                                             from orderentry_Detail oed " & vbCrLf & _
                "                                             where datediff(d,@pDate, Delivery_Date)<0 " & vbCrLf & _
                "                                                 and datediff(m,@pDate, Delivery_Date)=0 " & vbCrLf & _
                "                                             group by item_Code " & vbCrLf & _
                "                                             ) b " & vbCrLf & _
                "                             on im.item_code = b.item_code "
                      
    ls_Que = ls_Que & vbCrLf & _
                "                             left join ( " & vbCrLf & _
                "                                                 select PlanItem_Code,  " & vbCrLf & _
                "                                                             total_Pcd = isnull( sum(    case plan_Cls   when '1' then plan_qty when '2' then -plan_qty " & vbCrLf & _
                "                                                                                                                                         when '3' then plan_qty when '4' then -plan_qty  end " & vbCrLf & _
                "                                                                                                             ) " & vbCrLf & _
                "                                                                                                 ,0) " & vbCrLf & _
                "                                                 from productioncalculate_Detail pcd " & vbCrLf & _
                "                                                 where datediff(d,@pDate, Plan_Date)<0 " & vbCrLf & _
                "                                                     and datediff(m,@pDate, Plan_Date)=0 " & vbCrLf & _
                "                                                 group by planItem_Code " & vbCrLf & _
                "                                                 ) c " & vbCrLf & _
                "                             on im.item_code = c.Planitem_Code "
    ls_Que = ls_Que & vbCrLf & _
                "                             left join ( " & vbCrLf & _
                "                                                 select pod.item_Code, Remain_Po = isnull(sum(pod.Qty),0) - isnull(sum(pr.qty),0) " & vbCrLf & _
                "                                                 from purchaseorder_master pom " & vbCrLf & _
                "                                                 inner join purchaseorder_Detail pod " & vbCrLf & _
                "                                                 on pom.po_no = pod.po_no " & vbCrLf & _
                "                                                 left join part_receipt pr " & vbCrLf & _
                "                                                 on pom.supplier_code = pr.supplier_code " & vbCrLf & _
                "                                                 and pom.po_no = pr.po_no " & vbCrLf & _
                "                                                 and pod.item_Code = pr.item_code " & vbCrLf & _
                "                                                 where datediff(d,@pDate, pom.delivery_date)<0 " & vbCrLf & _
                "                                                     and datediff(m,@pDate, pom.delivery_date)=0  " & vbCrLf & _
                "                                                 group by pod.item_code " & vbCrLf & _
                "                                                 ) d " & vbCrLf & _
                "                             on im.item_code = d.item_code "
    ls_Que = ls_Que & vbCrLf & _
                "                             left join ( " & vbCrLf & _
                "                                                 select factory_code, Line_Code, Item_Code, total_DP = sum(Qty) " & vbCrLf & _
                "                                                 from daily_Production dp " & vbCrLf & _
                "                                                 where datediff(d,@pDate, dp.schedule_date)<0 " & vbCrLf & _
                "                                                     and datediff(m,@pDate, dp.schedule_date)=0  " & vbCrLf & _
                "                                                 group by factory_code, Line_Code, Item_Code " & vbCrLf & _
                "                                                 ) e " & vbCrLf & _
                "                             on im.item_code = e.item_Code " & vbCrLf & _
                "                                 and im.manufacture_Code = e.factory_Code " & vbCrLf & _
                "                                 and im.Line_code = e.line_code " & vbCrLf & _
                "                             ) tbData " & vbCrLf & _
                "                 )tbStock " & vbCrLf & _
                " on tbUtama.item_Code = tbStock.item_Code " & vbCrLf & _
                "     and tbUtama.WH_Code = tbStock.WH_Code " & vbCrLf & _
                " order by  tbUtama.finishgoodpart_Cls, tbUtama.production_cls, tbUtama.makebuy_Cls, /*tbUtama.group_cls, */ " & vbCrLf & _
                "                 tbUtama.line_code desc, tbUtama.planitem_Code, tbUtama.Cust_Code, tbUtama.delivery_Date, tbUtama.Po_No "

    rsGrid.CursorLocation = adUseClient
    rsGrid.Open ls_Que, Db, adOpenKeyset, adLockReadOnly
    
    grid.Rows = grid.FixedRows
    If Not (rsGrid.BOF Or rsGrid.EOF) Then
        li_totalRows = rsGrid.RecordCount
        prgBar.Value = 0
        prgBar.Max = li_totalRows
        With grid
            While Not rsGrid.EOF
'                DoEvents
                prgBar.Value = prgBar.Value + 1
                
                If UCase(Trim$(rsGrid("planitem_Code") & "")) <> ls_lastItem Then
                    If .Rows <> .FixedRows Then
                        'bikin baris E.STOCK
                        .Rows = .Rows + 1
                        .RowHeight(.Rows - 1) = .RowHeightMax
                        .TextMatrix(.Rows - 1, lb_col_Stock) = "END"
                        
                        .MergeRow(.Rows - 1) = True
                        
                        Call S_calcStock(.Rows - 2, li_StaticCols, False)
                        
                        If .Cell(flexcpBackColor, .Rows - 2, 0) = lc_Gray Then .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .ColS - 1) = lc_Gray
                        .RowOutlineLevel(.Rows - 1) = 0
                        .IsSubtotal(.Rows - 1) = True
                        
                        'bikin baris kosong hitam2
                        .Rows = .Rows + 1
                        .RowHeight(.Rows - 1) = 50
                        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .ColS - 1) = vbGrayed
                        .RowOutlineLevel(.Rows - 1) = 0
                        .IsSubtotal(.Rows - 1) = True
                        
                    End If
                    
'PARENT DATA
'++++++++++++++++++
                    .Rows = .Rows + 1
                    .RowHeight(.Rows - 1) = .RowHeightMax
                    If Trim$(rsGrid("production_Cls") & "") = "01" Then
                        .Cell(flexcpChecked, .Rows - 1, lb_col_C) = flexChecked
                    Else
                        .Cell(flexcpChecked, .Rows - 1, lb_col_C) = flexChecked
                        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .ColS - 1) = lc_Gray
                    End If
                    .TextMatrix(.Rows - 1, lb_col_PartNumber) = Trim$(rsGrid("planitem_code") & "")
                    .TextMatrix(.Rows - 1, lb_col_Description) = Trim$(rsGrid("item_Name") & "")
                    .TextMatrix(.Rows - 1, lb_col_Line) = Trim$(rsGrid("line_Name") & "")
                    .TextMatrix(.Rows - 1, lb_col_Lead) = Trim$(rsGrid("Lead_Time") & "")
                    .TextMatrix(.Rows - 1, lb_col_Lot) = Format(Val(rsGrid("Lot_Qty")), gs_formatQty)
                    .TextMatrix(.Rows - 1, lb_col_MinStock) = Format(Val(rsGrid("Min_Stock")), gs_formatQty)
                    .TextMatrix(.Rows - 1, lb_col_MaxStock) = Format(Val(rsGrid("Max_Stock")), gs_formatQty)
                    .TextMatrix(.Rows - 1, lb_col_Stock) = "BEGIN"
                    .TextMatrix(.Rows - 1, li_StaticCols) = Format(Val(rsGrid("Stock_Awal")), gs_formatQty)
                    .Cell(flexcpAlignment, .Rows - 1, li_StaticCols, .Rows - 1, .ColS - 1) = flexAlignLeftCenter
                    .RowOutlineLevel(.Rows - 1) = 0
                    .IsSubtotal(.Rows - 1) = True
                    .MergeRow(.Rows - 1) = True
                End If
'++++++++++++++++++
                
                
'CHILD DATA
'++++++++++++++++++
                If ls_lastItem <> UCase(Trim$(rsGrid("planitem_code") & "")) Or ls_lastCust <> UCase(Trim$(rsGrid("cust_Code") & "")) _
                Or ls_lastPo <> UCase(Trim$(rsGrid("po_No") & "")) Or ls_seqNo <> UCase(Trim$(rsGrid("seq_no") & "")) Then
                    
                    .Rows = .Rows + 1
                    .RowHeight(.Rows - 1) = .RowHeightMax
                    If Trim$(rsGrid("production_Cls") & "") = "01" Then
                        .Cell(flexcpChecked, .Rows - 1, lb_col_G) = IIf(rsGrid("generate_Cls") = 1, flexChecked, flexUnchecked)
                        .TextMatrix(.Rows - 1, lb_col_G) = .Rows - 1
                    Else
                        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .ColS - 1) = lc_Gray
                    End If
                    .TextMatrix(.Rows - 1, lb_col_PartNumber) = Trim$(rsGrid("po_no") & "") '& " (" & UCase(Trim$(rsGrid("trade_Name") & "")) & ")"
                    .TextMatrix(.Rows - 1, lb_col_Description) = Format(rsGrid("delivery_Date"), "dd MMM yyyy")
                    
                    ' plan_Cls [IN] / [PLAN] = 2,4
                    ' plan_Cls [OUT] / [REQ] = 1,3
                    
                    For li_Hari = 0 To ((.ColS - li_StaticCols) / 3)
                        If Trim$(rsGrid("plan_Cls") & "") = "2" Or Trim$(rsGrid("plan_Cls") & "") = "4" Then
                            li_TempCol = li_StaticCols + ((3 * li_Hari) + 2)
                        Else
                            li_TempCol = li_StaticCols + (3 * li_Hari) + 1
                        End If
                        If Format(rsGrid("plan_Date"), "yyyy-mm-dd") = Split(.ColData(li_TempCol), ";")(0) Then
                            .TextMatrix(.Rows - 1, li_TempCol) = Format(Val(rsGrid("Plan_Qty")), gs_formatQty)
                            .Cell(flexcpBackColor, .Rows - 1, li_TempCol, .Rows - 1, li_TempCol) = sHape(Trim$(rsGrid("plan_Cls") & "")).BackColor
                            Exit For
                        End If
                    Next li_Hari

                    .RowOutlineLevel(.Rows - 1) = 1
                    .IsSubtotal(.Rows - 1) = True
                Else
                    For li_Hari = 0 To ((.ColS - li_StaticCols) / 3)
                        If Trim$(rsGrid("plan_Cls") & "") = "2" Or Trim$(rsGrid("plan_Cls") & "") = "4" Then
                            li_TempCol = li_StaticCols + ((3 * li_Hari) + 2)
                        Else
                            li_TempCol = li_StaticCols + (3 * li_Hari) + 1
                        End If
                        If Format(rsGrid("plan_Date"), "yyyy-mm-dd") = Split(.ColData(li_TempCol), ";")(0) Then
                            .TextMatrix(.Rows - 1, li_TempCol) = Format(Val(rsGrid("Plan_Qty")), gs_formatQty)
                            .Cell(flexcpBackColor, .Rows - 1, li_TempCol, .Rows - 1, li_TempCol) = sHape(Trim$(rsGrid("plan_Cls") & "")).BackColor
                            If Trim$(rsGrid("plan_Cls") & "") = "2" Or Trim$(rsGrid("plan_Cls") & "") = "4" Then
                                'parameter HARUS sesuai urutan di inisialisasi
                                .TextMatrix(.Rows - 1, lb_col_Value) = Trim$(rsGrid("Cust_Code") & "") & ";" & _
                                                                                        Trim$(rsGrid("Po_No") & "") & ";" & _
                                                                                        Trim$(rsGrid("Seq_No") & "") & ";" & _
                                                                                        Trim$(rsGrid("planitem_Code") & "") & ";" & _
                                                                                        Trim$(rsGrid("plan_Seqno") & "") & ";" & _
                                                                                        Trim$(rsGrid("plan_Cls") & "") & ";" & _
                                                                                        Trim$(rsGrid("plan_Qty") & "") & ";" & _
                                                                                        Format(rsGrid("plan_Date"), "yyyy-mm-dd") & ";" & _
                                                                                        Trim$(rsGrid("Line_code") & "") & ";" & _
                                                                                        Trim$(rsGrid("Unit_Cls") & "") & ";" & _
                                                                                        Trim$(rsGrid("Daily_Prod_Seqno") & "") & ";" & _
                                                                                        Trim$(rsGrid("Part_Receipt_Seqno") & "") & ";" & _
                                                                                        Trim$(rsGrid("manufacture_Code") & "") & ";" & _
                                                                                        Trim$(rsGrid("Min_Stock") & "") & ";" & _
                                                                                        Trim$(rsGrid("Max_Stock") & "") & ";" & _
                                                                                        Trim$(rsGrid("Lot_Qty") & "") & ";" & _
                                                                                        Trim$(rsGrid("Req_Qty") & "") & ";"
                      
                            End If
                            Exit For
                        End If
                    Next li_Hari
                End If
'++++++++++++++++++
                
                ls_lastItem = UCase(Trim$(rsGrid("planitem_code") & ""))
                ls_lastCust = UCase(Trim$(rsGrid("cust_Code") & ""))
                ls_lastPo = UCase(Trim$(rsGrid("po_No") & ""))
                ls_seqNo = UCase(Trim$(rsGrid("seq_No") & ""))
                
                rsGrid.MoveNext
            Wend
'---------------------------------------------
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = .RowHeightMax
            .TextMatrix(.Rows - 1, lb_col_Stock) = "END"
            
            .MergeRow(.Rows - 1) = True
            
            Call S_calcStock(.Rows - 2, li_StaticCols, False)
            
            If .Cell(flexcpBackColor, .Rows - 2, 0) = lc_Gray Then .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .ColS - 1) = lc_Gray
            .RowOutlineLevel(.Rows - 1) = 0
            .IsSubtotal(.Rows - 1) = True
'---------------------------------------------
'            .Cell(flexcpAlignment, .FixedRows, li_StaticCols, .Rows - 1, .ColS - 1) = flexAlignRightCenter
        End With
    Else
        LblErrMsg = DisplayMsg("4006") 'No data that you want to search !
    End If
    
    If rsGrid.State = adStateOpen Then rsGrid.Close
    Set rsGrid = Nothing
    lb_statusFillGrid = False
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
        If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub griD_RowColChange()
    If grid.Col < li_StaticCols Then
        If grid.Col <> lb_col_C And grid.Col <> lb_col_G Then grid.FocusRect = flexFocusLight
    Else
        If LCase(Split(grid.ColData(grid.Col), ";")(1)) <> "i" And grid.Cell(flexcpBackColor, grid.Row, grid.Col) <> 0 _
        And grid.Cell(flexcpBackColor, grid.Row, 0) <> lc_Gray And grid.Cell(flexcpBackColor, grid.Row, grid.Col) <> vbGrayed _
        And grid.Cell(flexcpChecked, grid.Row, lb_col_G) = flexChecked Then
            grid.FocusRect = flexFocusInset
        Else
            grid.FocusRect = flexFocusLight
        End If
    End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < li_StaticCols Then
        grid.FocusRect = flexFocusLight
        If Col <> lb_col_C And Col <> lb_col_G Then
            Cancel = True
            Exit Sub
        Else
            If grid.TextMatrix(Row, lb_col_Value) <> "" Then
                If Col = lb_col_G And Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_PartReceipt) > 0 Then
                    LblErrMsg = DisplayMsg("1213") 'You can't delete this data ! It's already receipt.
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    Else
        If grid.FocusRect <> flexFocusInset Then
            Cancel = True
            Exit Sub
        Else
            If Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_PartReceipt) > 0 Then
                LblErrMsg = DisplayMsg("1214") 'You can't update this data ! It's already receipt.
                Cancel = True
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col >= li_StaticCols Then
        If Trim$(grid.TextMatrix(Row, Col) & "") <> "" And IsNumeric(grid.TextMatrix(Row, Col)) Then
            If CDbl(grid.TextMatrix(Row, Col)) <= gd_MaxQty Then
                grid.TextMatrix(Row, Col) = Format(CDbl(grid.TextMatrix(Row, Col)), gs_formatQty): grid.TextMatrix(Row, lb_col_Edited) = "1"
                'isi parameter
                grid.TextMatrix(Row, lb_col_Value) = Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_CustCode) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_Po) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_OrderSeq) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_PlanItem) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_PlanSeq) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_PlanCls) & ";" & _
                                                                        CDbl(grid.TextMatrix(Row, Col)) & ";" & _
                                                                        Split(grid.ColData(Col), ";")(0) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_PlanItemLineCode) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_PlanItemUnitCls) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_DailyProd) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_PartReceipt) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_FactoryCode) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_MinStock) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_MaxStock) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_LotQty) & ";" & _
                                                                        Split(grid.TextMatrix(Row, lb_col_Value), ";")(lb_Pos_ReqQty)
                Call S_calcStock(Row, Col, True)
                LblErrMsg = ""
            Else
                If Not lb_statusFillGrid Then LblErrMsg = DisplayMsg("1215") & " " & gd_MaxQty 'Qty must be equal/lower than 9,999,999.99
            End If
        Else
            If Not lb_statusFillGrid Then LblErrMsg = DisplayMsg("1216") 'pls input qty
        End If
    Else
        If Col = lb_col_C Then
            lb_flagCheckGridC = True
            chC.Value = IIf(F_checkC, 1, 0)
            lb_flagCheckGridC = False
        ElseIf Col = lb_col_G Then
        
            If grid.Cell(flexcpChecked, Row, Col) = flexChecked Then
                grid.LeftCol = F_cariKolom(Row)
            End If
            
            lb_flagCheckGridG = True
            chG.Value = IIf(F_checkG, 1, 0)
            lb_flagCheckGridG = False
        End If
    End If
End Sub

Private Sub griD_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = lb_col_C Or Col = lb_col_G Then Cancel = True
End Sub

Private Sub S_hiddenBlankDate()
Dim li_Col As Integer
Dim li_Row As Integer
Dim isValued As Boolean

With grid
    For li_Col = li_StaticCols To .ColS - 1 '((.ColS - li_StaticCols) / 2)
        isValued = False
        For li_Row = .FixedRows To .Rows - 1
            If .RowOutlineLevel(li_Row) = 1 Then
                If Trim$(.TextMatrix(li_Row, li_Col)) <> "" Then
                    isValued = True
                    Exit For
                End If
            End If
        Next li_Row
        If Not isValued Then .ColHidden(li_Col) = True
    Next li_Col
End With
End Sub

Private Sub S_lockControl(Flag As Boolean)
    cboCustomer.Enabled = Not Flag
    DtStart.Enabled = Not Flag
    DtEnd.Enabled = Not Flag
    cboPo.Enabled = Not Flag
    CboItemCode.Enabled = Not Flag
    chDate.Enabled = Not Flag
    chPartMaterial.Enabled = Not Flag
    chItemInformation.Enabled = Not Flag
    cmdSearch.Enabled = Not Flag
    cmdSubmit.Enabled = Not Flag
    cmdGenerate.Enabled = Not Flag
    cmdsubmenu.Enabled = Not Flag
    lstUncalculated.Enabled = Not Flag
    grid.Enabled = Not Flag
End Sub

Private Sub CmdSubmit_Click()
'+++++++++++++++++++++++++++++++++++++++++
'       SUBMIT LOGIC
'1. Loop Grid
'   ~ save CalcData yang ada perubahan [CalcCls = check]
'   ~ delete CalcData [CalcCls = uncheck]
'   End Loop
'2' generate prod sched from table [GenCls = check]
'+++++++++++++++++++++++++++++++++++++++++

Dim li_baris As Integer
Dim lb_Save As Boolean
Dim ls_Que As String
Dim ll_SeqNo As Long
Dim rsDaily As New ADODB.Recordset
Dim li_jumlahRecord As Integer
Dim isDbLock As Boolean
Dim ls_msG As String

If grid.Rows <= grid.FixedRows Then
    LblErrMsg = DisplayMsg("5012")  'There is no data to submit !
    Exit Sub
End If

Me.MousePointer = vbHourglass
S_lockControl (True)
With grid

prgBar.Max = .Rows - .FixedRows
prgBar.Value = 0
prgBar.Visible = True
rsDaily.CursorLocation = adUseClient

Db.BeginTrans
isDbLock = True

    For li_baris = .FixedRows To .Rows - 1
'        If .Cell(flexcpBackColor, li_Baris, lb_col_C) <> lc_Gray Then
            If .RowOutlineLevel(li_baris) = 0 Then
                If .Cell(flexcpChecked, li_baris, lb_col_C) = flexChecked Then lb_Save = True Else lb_Save = False
            Else
                If lb_Save Then
                    If .Cell(flexcpChecked, li_baris, lb_col_G) = flexChecked Then
                    
                        'cek plan qty > 0 and plan qty <= max stock
                        If CDbl(Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanQty)) <= 0 Then
                            'Disable Validasi Qty = 0, di-handle nanti langsung hapus daily production yang qty-nya 0
'                            Db.RollbackTrans
'                            isDbLock = False
'                            prgBar.Visible = False
'                            S_lockControl (False)
'                            lblErrMsg = DisplayMsg(1012)
'                            Me.MousePointer = vbDefault
'                            Exit Sub
                        ElseIf Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_MaxStock) <> 0 Then
                            If CDbl(grid.TextMatrix(li_baris, lb_col_End)) > CDbl(Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_MaxStock)) Then
                                Db.RollbackTrans
                                isDbLock = False
                                prgBar.Visible = False
                                S_lockControl (False)
                                LblErrMsg = DisplayMsg(4045) & " maximum stock !"
                                Me.MousePointer = vbDefault
                                Exit Sub
                            End If
                        End If
                    
                        If Trim$(grid.TextMatrix(li_baris, lb_col_Edited)) = "1" Then
                            'update Calc data
                            ls_Que = "UPDATE productionCalculate_Detail " & vbCrLf & _
                                        " set plan_Qty = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanQty) & "' " & vbCrLf & _
                                        "     ,plan_Date = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanDate) & "' " & vbCrLf & _
                                        " where Cust_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_CustCode) & "' " & vbCrLf & _
                                        "    and Po_No = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_Po) & "' " & vbCrLf & _
                                        "    and seq_No = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_OrderSeq) & "' " & vbCrLf & _
                                        "    and planitem_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanItem) & "' " & vbCrLf & _
                                        "    and plan_SeqNo = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanSeq) & "' "
                            Db.Execute ls_Que
                        End If
                        
                        'if lum ada prodResult
                        li_jumlahRecord = 0
                        ls_Que = " select * " & vbCrLf & _
                                    " from part_receipt " & vbCrLf & _
                                    " where productionResult_Cls = 'P1' " & vbCrLf & _
                                    " and dailyseq_No = ( " & vbCrLf & _
                                    "                                     select seq_no from Daily_Production " & vbCrLf & _
                                    "                                     where isnull(auto_cls,'0') = '1' " & vbCrLf & _
                                    "                                     and item_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanItem) & "' " & vbCrLf & _
                                    "                                     and planCust_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_CustCode) & "' " & vbCrLf & _
                                    "                                     and planPo_No = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_Po) & "' " & vbCrLf & _
                                    "                                     and planPo_Seqno = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_OrderSeq) & "' " & vbCrLf & _
                                    "                                     and Plan_Seqno = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanSeq) & "' " & vbCrLf & _
                                    "                                     ) "
                        If rsDaily.State = adStateOpen Then rsDaily.Close
                        rsDaily.Open ls_Que, Db, adOpenKeyset, adLockReadOnly
                        li_jumlahRecord = rsDaily.RecordCount
                        If rsDaily.State = adStateOpen Then rsDaily.Close
                        If li_jumlahRecord = 0 Then
                            'update DailyProduction ~> if rowAffected = 0 then insert
                            ls_Que = " UPDATE [Daily_Production] " & vbCrLf & _
                                        "    SET [Qty] = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanQty) & "' " & vbCrLf & _
                                        "       ,[Schedule_Date] = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanDate) & "' " & vbCrLf & _
                                        "       ,[Plan_SeqNo] = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanSeq) & "' " & vbCrLf & _
                                        "       ,[Last_Update] = getdate() " & vbCrLf & _
                                        "       ,[Last_User] = '" & userLogin & "' " & vbCrLf & _
                                        "    where isnull(auto_cls,'0') = '1' " & vbCrLf & _
                                        "      and item_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanItem) & "' " & vbCrLf & _
                                        "      and planCust_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_CustCode) & "' " & vbCrLf & _
                                        "      and planPo_No = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_Po) & "' " & vbCrLf & _
                                        "      and planPo_Seqno = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_OrderSeq) & "' " & vbCrLf & _
                                        "      and Plan_Seqno = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanSeq) & "' "
                            Db.Execute ls_Que, li_jumlahRecord
                            If li_jumlahRecord = 0 Then
                                'INSERT DAILY PROD
                                'BARCODE = [Factory_code]+[Line_Code]+[Schedule_Date](yyyymmdd)+[Seq_No]    ~> by trigger
                                ls_Que = " INSERT INTO [Daily_Production] " & vbCrLf & _
                                            "            ([Factory_code] " & vbCrLf & _
                                            "            ,[Line_Code] " & vbCrLf & _
                                            "            ,[Item_code] " & vbCrLf & _
                                            "            ,[Lot_No] " & vbCrLf & _
                                            "            ,[Qty] " & vbCrLf & _
                                            "            ,[Unit_Cls] " & vbCrLf & _
                                            "            ,[Schedule_Date] " & vbCrLf & _
                                            "            ,[Remark] " & vbCrLf & _
                                            "            ,[Request_Cls] " & vbCrLf & _
                                            "            ,[Complete_Cls] " & vbCrLf & _
                                            "            ,[Auto_Cls] " & vbCrLf & _
                                            "            ,[PlanCust_Code] " & vbCrLf & _
                                            "            ,[PlanPO_No] " & vbCrLf & _
                                            "            ,[PlanPO_SeqNo] " & vbCrLf & _
                                            "            ,[Plan_SeqNo] " & vbCrLf & _
                                            "            ,[Prod_Barcode] " & vbCrLf & _
                                            "            ,[Last_Update] " & vbCrLf & _
                                            "            ,[Last_User] " & vbCrLf & _
                                            "            ,[Register_Date]) " & vbCrLf & _
                                            "      VALUES "
                                ls_Que = ls_Que & vbCrLf & _
                                            "            ('" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_FactoryCode) & "' " & vbCrLf & _
                                            "            ,'" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanItemLineCode) & "' " & vbCrLf & _
                                            "            ,'" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanItem) & "' " & vbCrLf & _
                                            "            ,'' " & vbCrLf & _
                                            "            ,'" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanQty) & "' " & vbCrLf & _
                                            "            ,'" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanItemUnitCls) & "' " & vbCrLf & _
                                            "            ,'" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanDate) & "' " & vbCrLf & _
                                            "            ,'' " & vbCrLf & _
                                            "            ,null " & vbCrLf & _
                                            "            ,null " & vbCrLf & _
                                            "            ,'1' " & vbCrLf & _
                                            "            ,'" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_CustCode) & "' " & vbCrLf & _
                                            "            ,'" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_Po) & "' " & vbCrLf & _
                                            "            ,'" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_OrderSeq) & "' " & vbCrLf & _
                                            "            ,'" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanSeq) & "' " & vbCrLf & _
                                            "            ,null " & vbCrLf & _
                                            "            ,null " & vbCrLf & _
                                            "            ,'" & userLogin & "' " & vbCrLf & _
                                            "            ,getdate() ) "
                                Db.Execute ls_Que
                                
                                'update calculation detail
                                ls_Que = "UPDATE productionCalculate_Detail " & vbCrLf & _
                                            " set plan_Qty = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanQty) & "' " & vbCrLf & _
                                            "     ,plan_Date = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanDate) & "' " & vbCrLf & _
                                            " where Cust_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_CustCode) & "' " & vbCrLf & _
                                            "    and Po_No = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_Po) & "' " & vbCrLf & _
                                            "    and seq_No = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_OrderSeq) & "' " & vbCrLf & _
                                            "    and planitem_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanItem) & "' " & vbCrLf & _
                                            "    and plan_SeqNo = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanSeq) & "' "
                            
                                Db.Execute ls_Que
                            
                                'UPDATE GENERATE_CLS di PCD
                                ls_Que = "             UPDATE productioncalculate_detail " & vbCrLf & _
                                            "             SET generate_Cls = '1' " & vbCrLf & _
                                            "             Where Cust_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_CustCode) & "' " & vbCrLf & _
                                            "             and po_no = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_Po) & "' " & vbCrLf & _
                                            "             and seq_no = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_OrderSeq) & "' " & vbCrLf & _
                                            "             and planitem_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanItem) & "' "

                                Db.Execute ls_Que
                                
                            End If
                        Else
                            'brarti udah ada prodResult nya
                        End If
                        
                    Else
                        'Generate di uncheck
                        '~ cek part_receipt_Seqno, 0 -> delete daily_Prod
                        If Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_DailyProd) > 0 _
                        And Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PartReceipt) = "0" Then
                            'delete daily
                            ls_Que = "DELETE daily_Production " & vbCrLf & _
                                        " where seq_No = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_DailyProd) & "' "
                            Db.Execute ls_Que
                            
                            'balikin Plan Qty = Req Qty
                            ls_Que = "UPDATE productionCalculate_Detail " & vbCrLf & _
                                        " set plan_Qty = isnull((" & vbCrLf & _
                                        "    select plan_qty from productionCalculate_Detail a " & vbCrLf & _
                                        "    Where a.cust_code = productionCalculate_Detail.cust_code " & vbCrLf & _
                                        "    and a.po_no = productionCalculate_Detail.po_no " & vbCrLf & _
                                        "    and a.seq_no = productionCalculate_Detail.seq_no " & vbCrLf & _
                                        "    and a.planitem_Code = productionCalculate_Detail.planitem_Code " & vbCrLf & _
                                        "    and a.plan_SeqNo <> productionCalculate_Detail.plan_SeqNo), 0) " & vbCrLf & _
                                        " where Cust_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_CustCode) & "' " & vbCrLf & _
                                        "    and Po_No = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_Po) & "' " & vbCrLf & _
                                        "    and seq_No = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_OrderSeq) & "' " & vbCrLf & _
                                        "    and planitem_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanItem) & "' " & vbCrLf & _
                                        "    and plan_SeqNo = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanSeq) & "' "
                            Db.Execute ls_Que
                        Else
                            'Udah ada data di part_Receipt
                        End If
                    End If
                Else
                    'delete data
                    If Trim$(grid.TextMatrix(li_baris, lb_col_Value) & "") <> "" _
                    And Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_DailyProd) = "0" Then
                        ls_Que = "DELETE productionCalculate_Detail " & vbCrLf & _
                                    " where Cust_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_CustCode) & "' " & vbCrLf & _
                                    "    and Po_No = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_Po) & "' " & vbCrLf & _
                                    "    and seq_No = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_OrderSeq) & "' " & vbCrLf & _
                                    "    and planitem_Code = '" & Split(grid.TextMatrix(li_baris, lb_col_Value), ";")(lb_Pos_PlanItem) & "' "
                        Db.Execute ls_Que
                    End If
                End If
            End If
'        End If
        DoEvents
        prgBar.Value = prgBar.Value + 1
    Next li_baris

End With

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'UPDATE STATUS

    'Update Generate_Cls di ProductionCalculate_Detail
    ls_Que = " UPDATE ProductionCalculate_Detail " & vbCrLf & _
                " SET Generate_Cls = case when " & vbCrLf & _
                "               exists( " & vbCrLf & _
                "                                       select * " & vbCrLf & _
                "                                       from daily_production " & vbCrLf & _
                "                                       where planCust_Code = ProductionCalculate_Detail.Cust_Code " & vbCrLf & _
                "                                       and planPo_No = ProductionCalculate_Detail.Po_No " & vbCrLf & _
                "                                       and planPO_seqNo = ProductionCalculate_Detail.Seq_No " & vbCrLf & _
                "                                       and item_Code = ProductionCalculate_Detail.planitem_Code " & vbCrLf & _
                "                              ) " & vbCrLf & _
                "          then '1' else '0' end "
    Db.Execute ls_Que
    
    'Update Calc_Cls di orderEntry_Detail
    ls_Que = " UPDATE orderEntry_Detail " & vbCrLf & _
                " SET calculate_Cls = case when " & vbCrLf & _
                "               exists( " & vbCrLf & _
                "                                             select * " & vbCrLf & _
                "                                             from productioncalculate_detail " & vbCrLf & _
                "                                             where Cust_Code = orderEntry_Detail.Cust_Code " & vbCrLf & _
                "                                             and Po_No = orderEntry_Detail.Po_No " & vbCrLf & _
                "                                             and seq_No = orderEntry_Detail.Seq_No " & vbCrLf & _
                "                              ) " & vbCrLf & _
                "          then '1' else '0' end " & vbCrLf & _
                " ,generate_Cls = case when " & vbCrLf & _
                "     exists( " & vbCrLf & _
                "                                       select * " & vbCrLf & _
                "                                       from daily_production " & vbCrLf & _
                "                                       where planCust_Code = orderEntry_Detail.Cust_Code " & vbCrLf & _
                "                                       and planPo_No = orderEntry_Detail.Po_No " & vbCrLf & _
                "                                       and planPO_seqNo = orderEntry_Detail.Seq_No " & vbCrLf & _
                "                              ) " & vbCrLf & _
                "          then '1' else '0' end "
    Db.Execute ls_Que
    
    'Delete master yang ga punya detail
    ls_Que = "DELETE productionCalculate_Master " & vbCrLf & _
                " where Cust_Code+Po_No+cast(Seq_No as varchar(18))+PlanItem_Code in ( " & vbCrLf & _
                "     select Cust_Code+Po_No+cast(Seq_No as varchar(18))+PlanItem_Code " & vbCrLf & _
                "     from productioncalculate_master pcm " & vbCrLf & _
                "     where not exists ( " & vbCrLf & _
                "             select * from productioncalculate_detail " & vbCrLf & _
                "             Where Cust_Code = pcm.Cust_Code " & vbCrLf & _
                "             and po_no = pcm.po_no " & vbCrLf & _
                "             and seq_no = pcm.seq_no " & vbCrLf & _
                "                                 ) " & vbCrLf & _
                "    ) "
    Db.Execute ls_Que

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    'Hapus Daily Production yang qty-nya 0
    ls_Que = "delete from daily_production where qty = 0 and auto_cls = 1"
    Db.Execute ls_Que
    
    Db.CommitTrans
    isDbLock = False

    ls_msG = DisplayMsg("1000")
    
normalExit:
    prgBar.Visible = False
    S_lockControl (False)
    cmdSearch.SetFocus
    SendKeys (vbCr)
    DoEvents
    LblErrMsg = ls_msG
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If isDbLock Then Db.RollbackTrans
    ls_msG = err.Description
    err.clear
    Resume normalExit
End Sub

Function F_checkC() As Boolean
Dim li_baris As Integer
With grid
    For li_baris = .FixedRows To .Rows - 1
        If .RowOutlineLevel(li_baris) = 0 And Trim$(.TextMatrix(li_baris, lb_col_PartNumber) & "") <> "" Then
            If .Cell(flexcpChecked, li_baris, lb_col_C) = flexUnchecked Then
                F_checkC = False
                Exit Function
            End If
        Else
        End If
    Next li_baris
End With
F_checkC = True
End Function

Function F_checkG() As Boolean
Dim li_baris As Integer
With grid
    For li_baris = .FixedRows To .Rows - 1
        If .RowOutlineLevel(li_baris) = 1 Then
            If .Cell(flexcpChecked, li_baris, lb_col_G) = flexUnchecked Then
                F_checkG = False
                Exit Function
            End If
        Else
        End If
    Next li_baris
End With
F_checkG = True
End Function

Private Sub chG_Click()
Dim li_baris As Integer
With grid
    If Not lb_flagCheckGridG Then
        For li_baris = .FixedRows To .Rows - 1
            If .RowOutlineLevel(li_baris) = 1 Then
                .Cell(flexcpChecked, li_baris, lb_col_G) = IIf(chG.Value = 1, flexChecked, flexUnchecked)
            End If
        Next li_baris
    End If
End With
End Sub

Private Sub chC_Click()
Dim li_baris As Integer
With grid
    If Not lb_flagCheckGridC Then
        For li_baris = .FixedRows To .Rows - 1
            If .RowOutlineLevel(li_baris) = 0 And Trim$(.TextMatrix(li_baris, lb_col_PartNumber) & "") <> "" Then
                .Cell(flexcpChecked, li_baris, lb_col_C) = IIf(chC.Value = 1, flexChecked, flexUnchecked)
            End If
        Next li_baris
    End If
End With
End Sub

Private Sub S_calcStock(ByVal pRow As Long, ByVal pCol As Long, ByVal pOverride As Boolean)
Dim l_topRow As Integer
Dim l_bottomRow As Integer
Dim l_Col As Integer
Dim l_Row As Integer

Dim l_tempTotalIn As Double
Dim l_tempTotalOut As Double
Dim l_tempPlanQty As Double
Dim l_tempReqQty As Double

Dim l_tempItem As String
Dim l_tempMinStock As Double
Dim l_tempMaxStock As Double
Dim l_tempLotQty As Double
Dim l_tempDate As Date

With grid
    If pRow < .FixedRows Or pRow >= .Rows - 1 Then Exit Sub
    
    l_topRow = pRow
    l_bottomRow = pRow
    
    While .RowOutlineLevel(l_topRow) <> 0 And l_topRow >= .FixedRows
        l_topRow = l_topRow - 1
    Wend
    
    While .RowOutlineLevel(l_bottomRow) <> 0 And l_bottomRow <= .Rows - 1
        l_bottomRow = l_bottomRow + 1
    Wend
    
    If Trim$(.TextMatrix(l_topRow, pCol) & "") = "" Then
        l_tempTotalIn = 0
    Else
        l_tempTotalIn = CDbl(.TextMatrix(l_topRow, pCol))
    End If
        
    For l_Col = pCol To .ColS - 1
        If Trim$(.ColData(l_Col)) <> "" Then
        
            If CDbl(l_tempDate) = 0 Then
                l_tempDate = CDate(Split(.ColData(l_Col), ";")(0))
            ElseIf Month(l_tempDate) <> Month(CDate(Split(.ColData(l_Col), ";")(0))) Then
                l_tempTotalIn = l_tempTotalIn + f_getInventory(l_tempItem, l_tempDate)
                l_tempDate = CDate(Split(.ColData(l_Col), ";")(0))
            End If
            
            If LCase(Split(.ColData(l_Col), ";")(1)) = "i" Then
                .TextMatrix(l_topRow, l_Col) = Format(l_tempTotalIn, gs_formatQty)
            Else
                .TextMatrix(l_topRow, l_Col) = .TextMatrix(l_topRow, l_Col - 1)
            End If
            For l_Row = l_topRow + 1 To l_bottomRow - 1
                l_tempItem = Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_PlanItem)
                l_tempMinStock = CDbl(Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_MinStock))
                l_tempMaxStock = CDbl(Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_MaxStock))
                l_tempLotQty = Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_LotQty)
                l_tempReqQty = Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_ReqQty)
                
                If Trim$(.ColData(l_Col)) <> "" Then
                    If Split(.ColData(l_Col), ";")(1) = "i" Then
                        If Trim$(.TextMatrix(l_Row, l_Col) & "") <> "" Then
                            l_tempTotalIn = l_tempTotalIn - CDbl(.TextMatrix(l_Row, l_Col))
                            l_tempTotalOut = l_tempTotalOut - CDbl(.TextMatrix(l_Row, l_Col))
                        End If
                    Else
                        l_tempPlanQty = 0
                        If Trim$(.TextMatrix(l_Row, l_Col) & "") <> "" Then
                            If pOverride And l_Col = pCol And l_Row = pRow Then
                                l_tempPlanQty = .TextMatrix(l_Row, l_Col)
                                l_tempTotalIn = l_tempTotalIn + l_tempPlanQty
                                l_tempTotalOut = l_tempTotalOut + l_tempReqQty
                            Else
                                If .Cell(flexcpChecked, l_Row, lb_col_G) = flexUnchecked Then
                                    If l_tempTotalIn - l_tempTotalOut - l_tempReqQty >= l_tempMinStock Then
                                        l_tempTotalOut = l_tempTotalOut + l_tempReqQty
                                    Else
                                        If l_tempLotQty = 0 Then
                                            l_tempPlanQty = Abs(l_tempTotalIn - l_tempTotalOut - l_tempReqQty - l_tempMinStock)
                                        Else
                                            l_tempPlanQty = l_tempLotQty * RoundUp(Abs(l_tempTotalIn - l_tempTotalOut - l_tempReqQty - l_tempMinStock) / l_tempLotQty)
                                        End If
                                        l_tempTotalIn = l_tempTotalIn + l_tempPlanQty
                                        l_tempTotalOut = l_tempTotalOut + l_tempReqQty
                                    End If
                                Else
                                    l_tempPlanQty = .TextMatrix(l_Row, l_Col)
                                    l_tempTotalIn = l_tempTotalIn + l_tempPlanQty
                                    l_tempTotalOut = l_tempTotalOut + l_tempReqQty
                                End If
                            End If
                            .TextMatrix(l_Row, l_Col) = Format(l_tempPlanQty, gs_formatQty)
                            .TextMatrix(l_Row, lb_col_Value) = Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_CustCode) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_Po) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_OrderSeq) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_PlanItem) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_PlanSeq) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_PlanCls) & ";" & _
                                                                                    CDbl(.TextMatrix(l_Row, l_Col)) & ";" & _
                                                                                    Split(.ColData(l_Col), ";")(0) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_PlanItemLineCode) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_PlanItemUnitCls) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_DailyProd) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_PartReceipt) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_FactoryCode) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_MinStock) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_MaxStock) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_LotQty) & ";" & _
                                                                                    Split(.TextMatrix(l_Row, lb_col_Value), ";")(lb_Pos_ReqQty) & ";"
                        ElseIf Trim$(.TextMatrix(l_Row, l_Col - 1) & "") <> "" Then
                            If pOverride And l_Col = pCol Then
                                l_tempTotalIn = l_tempTotalIn - CDbl(.TextMatrix(l_Row, l_Col - 1))
                            End If
                        End If
                    End If
                End If
            Next l_Row
            If LCase(Split(.ColData(l_Col), ";")(1)) = "o" Then
                .TextMatrix(l_bottomRow, l_Col) = Format(l_tempTotalIn, gs_formatQty)
                .TextMatrix(l_bottomRow, l_Col - 1) = .TextMatrix(l_bottomRow, l_Col)
            End If
        End If
    Next l_Col
    For l_Row = l_topRow + 1 To l_bottomRow - 1
        .TextMatrix(l_Row, lb_col_End) = l_tempTotalIn - l_tempTotalOut
    Next
End With
End Sub

Private Function f_getInventory(pItemCode As String, pDate As Date) As Double
    Dim adoRs As New ADODB.Recordset
    
    sql = " select isnull(a.inventory, 0) inventory from( " & vbCrLf & _
                "   select sh.item_code, inventory_date = cast(rtrim(sh.stock_year) + '-' + rtrim(sh.stock_month) + '-1' as datetime), isnull(sh.inventory, sh.[current]) - sh.[current] inventory  " & vbCrLf & _
                "   from stock_history sh " & vbCrLf & _
                "   inner join item_master im on sh.item_code = im.item_code and sh.warehouse_code = im.wh_code " & vbCrLf & _
                "   union all " & vbCrLf & _
                "   select st.item_code, inventory_date = ( " & vbCrLf & _
                "       select top 1 cast(cast(inventory_year as varchar) + '-' + cast(inventory_month as varchar) + '-1' as datetime)  " & vbCrLf & _
                "       from inventory_control order by cast(cast(inventory_year as varchar) + '-' + cast(inventory_month as varchar) + '-1' as datetime) desc " & vbCrLf & _
                "   ), isnull(lm_inventory, lm_current) - lm_current inventory " & vbCrLf & _
                "   from stock_master st " & vbCrLf & _
                "   inner join item_master im on st.item_code = im.item_code and st.warehouse_code = im.wh_code " & vbCrLf & _
                " ) a where a.item_code = '" & pItemCode & "' " & vbCrLf & _
                " and year(a.inventory_date) = '" & Year(pDate) & "' " & vbCrLf & _
                " and month(a.inventory_date) = '" & Month(pDate) & "' "

    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then f_getInventory = Val(adoRs.Fields("inventory") & "")
    adoRs.Close
End Function

Private Function F_cariKolom(ByVal pRow As Long) As Long
Dim li_Col As Long
With grid
    For li_Col = li_StaticCols To .ColS - 1
        If .Cell(flexcpBackColor, pRow, li_Col) <> .Cell(flexcpBackColor, pRow, lb_col_C) Then
            Exit For
        End If
    Next li_Col
    F_cariKolom = li_Col
End With
End Function
