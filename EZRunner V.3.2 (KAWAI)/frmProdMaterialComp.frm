VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmProdMaterialComp 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Material Consumption / Loss"
   ClientHeight    =   10980
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   15120
   Icon            =   "frmProdMaterialComp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "&zero"
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
      Index           =   2
      Left            =   11355
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9870
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
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
      Index           =   1
      Left            =   12608
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9870
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Index           =   0
      Left            =   13838
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9870
      Width           =   1140
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Back"
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
      Left            =   233
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9885
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   645
      Left            =   233
      TabIndex        =   19
      Top             =   9120
      Width           =   14745
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
         Height          =   330
         Left            =   105
         TabIndex        =   20
         Top             =   195
         Width           =   14520
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1230
      Left            =   233
      TabIndex        =   6
      Top             =   1110
      Width           =   14745
      Begin VB.Line Line1 
         Index           =   3
         X1              =   7515
         X2              =   8925
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblDailyQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7515
         TabIndex        =   22
         Top             =   750
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan "
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
         Left            =   6990
         TabIndex        =   21
         Top             =   750
         Width           =   420
      End
      Begin VB.Label lblResultQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10575
         TabIndex        =   18
         Top             =   750
         Width           =   1545
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   10575
         X2              =   12105
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result"
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
         Left            =   9735
         TabIndex        =   17
         Top             =   750
         Width           =   525
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1560
         X2              =   3330
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label lblItem 
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
         Left            =   1560
         TabIndex        =   16
         Top             =   300
         Width           =   1155
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
         Index           =   0
         Left            =   255
         TabIndex        =   15
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   3615
         TabIndex        =   14
         Top             =   300
         Width           =   960
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
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
         Left            =   4740
         TabIndex        =   13
         Top             =   300
         Width           =   5040
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4740
         X2              =   10440
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Womin  No"
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
         Left            =   11025
         TabIndex        =   12
         Top             =   300
         Width           =   930
      End
      Begin VB.Label lblFNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
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
         Left            =   12210
         TabIndex        =   11
         Top             =   300
         Width           =   750
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   12210
         X2              =   14520
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot No."
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
         Left            =   255
         TabIndex        =   10
         Top             =   750
         Width           =   600
      End
      Begin VB.Label lblLot 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   9
         Top             =   750
         Width           =   750
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   1560
         X2              =   2910
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result Date"
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
         Left            =   3615
         TabIndex        =   8
         Top             =   750
         Width           =   990
      End
      Begin VB.Label lblDt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4740
         TabIndex        =   7
         Top             =   750
         Width           =   1425
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   4740
         X2              =   6150
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6540
      Left            =   240
      TabIndex        =   0
      Top             =   2490
      Width           =   14745
      _cx             =   26009
      _cy             =   11536
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
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   5
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   3
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
      ScrollTrack     =   -1  'True
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
      Editable        =   2
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
      Begin MSForms.ComboBox CboCls 
         Height          =   315
         Left            =   30
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   330
         Visible         =   0   'False
         Width           =   765
         VariousPropertyBits=   746604571
         MaxLength       =   2
         DisplayStyle    =   7
         Size            =   "1349;556"
         ListRows        =   15
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Material Consumption / Loss"
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
      Height          =   330
      Left            =   5985
      TabIndex        =   5
      Top             =   375
      Width           =   3255
   End
End
Attribute VB_Name = "frmProdMaterialComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Long
Dim HakU As Integer
Public adaErr As Integer

Dim newCls As New clsMRP
Dim ClsProc As New ClsProc
Dim dbTransfer As New ADODB.Connection

Public dailyseqno As Double, KeyProd As Double, tglProd As String
Public factoryCD As String, FactoryDesc As String, lineCD As String
Public ZWHCode As String
Public thnFix As Integer, blnFix As Integer

Dim gridSeqNo As String, childSeqNo As Double
Dim ChildItemCD As String, childWHCode As String, ChildLotNo As String
Dim childQty As String, tampungChildQty As Double, qtyStock As Double
Dim childUnitCls As String, defaultUnit As String
Dim stockItem As String
Dim kondisi As String

Public ResultSeq As Double
Public strChildItemCD As String, schedule_date As String
Public factoryDaily As String, WHDaily As String, completeCls As Byte, seqNoProdReceipt As Double

Dim tampungBrs As Integer, nilKosong As Boolean
Dim matConsumpCls As String
Dim dblQtyTemp As Double, TandaZero As Integer, awal As Double, akhir As Double

Dim bteColSelect As Byte
Dim bteColMatCode As Byte
Dim bteColDesc As Byte
Dim bteColLoc As Byte
Dim bteColLocDesc As Byte
Dim bteColReqDate As Byte
Dim bteColPlanReq As Byte
Dim bteColPlanResult As Byte
Dim bteColResult As Byte
Dim bteColUnitReq As Byte
Dim bteColLotNo As Byte
Dim bteColUnitBom As Byte
Dim bteColJumChild As Byte
Dim bteColBlank1 As Byte
Dim bteColChildReq As Byte
Dim bteColStockCls As Byte
Dim bteColUnit As Byte
Dim bteColDate As Byte
Dim bteColCls As Byte

Private Sub headerGrid()
    
    bteColSelect = 0
    bteColMatCode = 1
    bteColDesc = 2
    bteColLoc = 3
    bteColLocDesc = 4
    bteColReqDate = 5
    bteColPlanReq = 6
    bteColPlanResult = 7
    bteColResult = 8
    bteColUnitReq = 9
    bteColLotNo = 10
    bteColUnitBom = 11
    bteColJumChild = 12
    bteColBlank1 = 13
    bteColChildReq = 14
    bteColStockCls = 15
    bteColUnit = 16
    bteColDate = 17
    bteColCls = 18
    
    With grid
        .clear
        .ColS = 19
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColMatCode) = "Material CD"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColLoc) = "Location"
        .TextMatrix(0, bteColLocDesc) = "Description"
        .TextMatrix(0, bteColReqDate) = "Requirement Date"
        .TextMatrix(0, bteColPlanReq) = "Plan/Req"
        .TextMatrix(0, bteColPlanResult) = "Plan/Result"
        .TextMatrix(0, bteColResult) = "Result"
        .TextMatrix(0, bteColUnitReq) = "Unit"
        .TextMatrix(0, bteColLotNo) = "Lot No"
        .TextMatrix(0, bteColUnitBom) = "UnitBOM"
        .TextMatrix(0, bteColJumChild) = "JumChild"
        .TextMatrix(0, bteColBlank1) = "Blank1"
        .TextMatrix(0, bteColChildReq) = "Child Req Qty"
        .TextMatrix(0, bteColStockCls) = "Stock Cls"
        .TextMatrix(0, bteColUnit) = "Unit Default"
        .TextMatrix(0, bteColDate) = "Scheduledate"
        .TextMatrix(0, bteColCls) = "Cls"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColMatCode) = 1500
        .ColWidth(bteColDesc) = 3000
        .ColWidth(bteColLoc) = 750
        .ColWidth(bteColLocDesc) = 1400
        .ColWidth(bteColReqDate) = 1650
        .ColWidth(bteColPlanReq) = 1250
        .ColWidth(bteColPlanResult) = 1250
        .ColWidth(bteColResult) = 1250
        .ColWidth(bteColUnitReq) = 400
        .ColWidth(bteColLotNo) = 1000
        .ColWidth(bteColCls) = 850
        
        .ColHidden(bteColUnitBom) = True
        .ColHidden(bteColJumChild) = True
        .ColHidden(bteColBlank1) = True
        .ColHidden(bteColChildReq) = True
        .ColHidden(bteColStockCls) = True
        .ColHidden(bteColUnit) = True
        .ColHidden(bteColDate) = True
        

        
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColMatCode) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColLoc) = flexAlignLeftCenter
        .ColAlignment(bteColLocDesc) = flexAlignLeftCenter
        .ColAlignment(bteColReqDate) = flexAlignCenterCenter
        .ColAlignment(bteColPlanReq) = flexAlignRightCenter
        .ColAlignment(bteColPlanResult) = flexAlignRightCenter
        .ColAlignment(bteColResult) = flexAlignRightCenter
        .ColAlignment(bteColUnitReq) = flexAlignLeftCenter
        .ColAlignment(bteColLotNo) = flexAlignLeftCenter
        .ColAlignment(bteColCls) = flexAlignLeftCenter
        
        .MergeCells = flexMergeFixedOnly
    End With
End Sub


Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    HakU = hakUpdate("", "Material Consumption/Loss")
    Call isiCbo(cboCls, "MaterialConsump_Cls", "MaterialConsump_Cls", "Description", 35, 100, "MaterialConsump_Cls", , , , 2)
    TandaZero = 0
End Sub

'************ Isi Grid **********
Sub IsiGrid(seqNo As Double, stat As Boolean, Optional dailyseqno As Double, Optional ItemParent As String)
Dim rsMat As New ADODB.Recordset
Dim rsSupp As New ADODB.Recordset
Dim jmlChild As Integer, qtyParentFormula As Double

If Trim(ItemParent) = "" Then ItemParent = lblitem

With grid
    Call headerGrid
    
    '******** Data Formula Master , Daily (qty en Date) ****************
    'Sql = " Select c.Manufacture_Code as Factory_Code, rtrim(e.trade_name) trade_name, a.Item_Code,rtrim(b.item_name) item_name, a.Qty as qtyAnak,a.Unit_Cls," & _
    '    "   b.WH_Code,b.Address,b.Stockcontrol_Cls as stockItem,d.Stockcontrol_Cls as stockWH, " & _
    '    "   start_date, end_date, " & _
    '    "    a.qty * (Select qty from daily_production where seq_no = '" & dailyseqno & "') TotalPlann," & _
    '    "  (a.qty * '" & Format(lblResultQty.Caption, "#0") & "') PlannResult, " & _
    '    "    (select Isnull(Sum(Consumption_Qty),0) From Part_Supply where ChildItem_Code = a.item_code And DO_No in  " & _
    '    "(Select convert(char,Seq_No) From Part_Receipt Where dailySeq_No = '" & dailyseqno & "')) TotalResult, " & _
    '    "   b.StockControl_Cls StockItem, b.unit_cls unitDefault, '' Lot_no" & _
    '    " from BOM_Master a,Item_Master b, Item_Master c, Warehouse_Master d, trade_master e     " & _
    '    " where a.Item_Code = b.Item_Code             " & _
    '    "   And a.Parent_ItemCode = c.Item_Code  and c.Manufacture_Code = e.trade_code           " & _
    '    "   And b.WH_Code = d.WH_Code             " & _
    '    "   and a.Parent_ItemCode = '" & ItemParent & "'" & _
    '    "   and Start_Date <='" & Format(lblDt, "YYYYMMDD") & "'" & _
    '    "   and End_Date >= '" & Format(lblDt, "YYYYMMDD") & "'"
    
    '*********** Consumption for Dummy WIP **************
    '     Update 20090515 -- For KAWAI
    '**************************************************
    
    Call CreateTableTemp
    
    sql = " Select c.Manufacture_Code as Factory_Code, rtrim(e.trade_name) trade_name,  " & vbCrLf & _
                      "     a.Item_Code,rtrim(b.item_name) item_name, a.QtyBOM as qtyAnak, b.WH_Code, " & vbCrLf & _
                      "     b.Address,b.Stockcontrol_Cls as stockItem,a.unit_cls,d.Stockcontrol_Cls as stockWH,         " & vbCrLf & _
                      "     a.qtyBOM * (Select qty from daily_production where seq_no = '" & dailyseqno & "') TotalPlann,   " & vbCrLf & _
                      "     (a.qtyBOM * '" & Format(lblResultQty.Caption, "#0") & "') PlannResult, " & vbCrLf & _
                      "     (select Isnull(Sum(Consumption_Qty),0) From Part_Supply where  " & vbCrLf & _
                      "         ChildItem_Code = a.item_code And DO_No in   " & vbCrLf & _
                      "             (Select convert(char,Seq_No) From Part_Receipt Where dailySeq_No = '" & dailyseqno & "')) TotalResult, " & vbCrLf & _
                      "     b.StockControl_Cls StockItem, b.unit_cls unitDefault, '' Lot_no  " & vbCrLf & _
                      " from  " & vbCrLf & _
                      "   (Select Utama, Item_Code,sum(QtyBOM) QtyBOM, Control_Cls, Unit_Cls,ItemType From TempStructure "

    sql = sql + "       where itemType='Child' Group By Utama, Item_Code, Control_Cls, Unit_Cls,ItemType) a, " & vbCrLf & _
                      " Item_Master b, Item_Master c, Warehouse_Master d, trade_master e " & vbCrLf & _
                      " where a.Item_Code = b.Item_Code "
    
    sql = sql + "     And a.Utama = c.Item_Code  and c.Manufacture_Code = e.trade_code And b.WH_Code = d.WH_Code " & vbCrLf & _
                      "     and a.Utama = '" & ItemParent & "'  and itemType='Child'  "
    
    
    Set rsMat = Db.Execute(sql)
    
    Do While Not rsMat.EOF
        '****************************** Data Header *************************
        
        'If rsMat("StockItem") = "02" Then Call IsiGrid(seqNo, stat, dailyseqno, rsMat("Item_Code"))
        
        .Rows = .Rows + 1
        i = .Rows - 1
        .TextMatrix(i, bteColSelect) = ""
        .Cell(flexcpBackColor, .Rows - 1, bteColSelect) = vbWhite
        .TextMatrix(i, bteColMatCode) = Trim(rsMat("Item_Code"))
        .TextMatrix(i, bteColDesc) = Trim(rsMat("Item_Name"))
        .TextMatrix(i, bteColLoc) = factoryCD
        .TextMatrix(i, bteColLocDesc) = Trim(rsMat("Trade_Name"))
        .TextMatrix(i, bteColReqDate) = Format(schedule_date, "dd MMM yyyy")
        If IsNull(rsMat("Totalplann")) Then
            .TextMatrix(i, bteColPlanReq) = Format(0, gs_formatQtyBOM)
        Else
            .TextMatrix(i, bteColPlanReq) = Format(rsMat("Totalplann"), gs_formatQtyBOM)
        End If
                
        .TextMatrix(i, bteColPlanResult) = Format(rsMat("PlannResult"), gs_formatQtyBOM)
        .TextMatrix(i, bteColResult) = Format(rsMat("TotalResult"), gs_formatQtyBOM)
        .TextMatrix(i, bteColUnitReq) = uf_GetUnitDescription(Trim(rsMat("unit_Cls")))
        .TextMatrix(i, bteColLotNo) = Trim(rsMat("Lot_No"))
        .TextMatrix(i, bteColUnitBom) = Trim(rsMat("unit_cls"))
        .TextMatrix(i, bteColStockCls) = Trim(rsMat("stockItem")) 'Stock Control Item
        .TextMatrix(i, bteColUnit) = Trim(rsMat("unitDefault")) 'Unit Default
        
        .Cell(flexcpBackColor, .Rows - 1, bteColMatCode, .Rows - 1, .ColS - 1) = &HE0E0E0
        '**** Data Supply berdasarkan Seq No Production Result Dan Item dr Formula Tersebut **********
        sql = "Select Seq_No,ChildItem_Code, " & _
                "ChildRequirement_Qty,Consumption_Qty, ChildUnit_Cls, remarks, " & _
                "MaterialConsump_Cls = isnull(MaterialConsump_Cls,'') " & _
            "From Part_Supply Supp " & _
            "WHere ChildItem_Code ='" & .TextMatrix(.Rows - 1, bteColMatCode) & _
            "' And DO_NO = '" & KeyProd & "'"
        Set rsSupp = Db.Execute(sql)
        
        If Not rsSupp.EOF Then
            jmlChild = 0
            Do While Not rsSupp.EOF
                .Rows = .Rows + 1
                .Cell(flexcpBackColor, .Rows - 1, bteColSelect) = vbWhite 'C
                .TextMatrix(.Rows - 1, bteColSelect) = ""
                .TextMatrix(.Rows - 1, bteColMatCode) = rsSupp("ChildItem_Code") 'Item_Code
                .TextMatrix(.Rows - 1, bteColLoc) = factoryCD  'WH Code
                
                .Cell(flexcpBackColor, .Rows - 1, bteColResult) = vbWhite 'Result
                
                ' Change for Edit Result
                '.TextMatrix(.Rows - 1, bteColResult) = Format(rsSupp("Consumption_Qty"), gs_formatQtyBOM)  'Result
                .TextMatrix(.Rows - 1, bteColResult) = Format(rsMat("PlannResult"), gs_formatQtyBOM)
                ' ----------------------------------
                
                .TextMatrix(.Rows - 1, bteColUnitReq) = uf_GetUnitDescription(Trim(rsSupp("ChildUnit_Cls")))
                
                .Cell(flexcpBackColor, .Rows - 1, bteColLotNo) = vbWhite 'Result
                .TextMatrix(.Rows - 1, bteColLotNo) = Trim(rsSupp("Remarks")) 'Lot No
                .TextMatrix(.Rows - 1, bteColUnitBom) = rsSupp("ChildUnit_Cls") 'Unit Cls
                .TextMatrix(.Rows - 1, bteColJumChild) = rsSupp("Seq_No") 'Update
                .TextMatrix(.Rows - 1, bteColChildReq) = rsSupp("ChildRequirement_Qty") 'Tampung Qty sblm Update
                .TextMatrix(.Rows - 1, bteColStockCls) = Trim(rsMat("stockItem"))
                .TextMatrix(.Rows - 1, bteColUnit) = .TextMatrix(i, bteColUnit) 'unit Default
                .TextMatrix(.Rows - 1, bteColDate) = Format(.TextMatrix(i, bteColReqDate), "yyyy-MM-dd") 'Tgl Prod
                .Cell(flexcpBackColor, .Rows - 1, bteColCls) = vbWhite
                .TextMatrix(.Rows - 1, bteColCls) = Trim(rsSupp("MaterialConsump_Cls"))
                jmlChild = jmlChild + 1
                rsSupp.MoveNext
            Loop
        
        Else
            'Jika tidak ada buat baris kosong dengan data dr Formula kecuali Qty en Lot No
            .AddItem ""
            .Cell(flexcpBackColor, .Rows - 1, bteColSelect) = vbWhite 'C
            .TextMatrix(.Rows - 1, bteColMatCode) = rsMat("Item_Code") 'Item_Code
            .TextMatrix(.Rows - 1, bteColLoc) = factoryCD 'WH Code
            
            .Cell(flexcpBackColor, .Rows - 1, bteColResult) = vbWhite 'Result
            If stat Then
                .TextMatrix(.Rows - 1, bteColResult) = Format(rsMat!plannresult, gs_formatQtyBOM) 'Result Qty
            Else
                .TextMatrix(.Rows - 1, bteColResult) = Format(0, gs_formatQtyBOM)
            End If
            .TextMatrix(.Rows - 1, bteColUnitReq) = uf_GetUnitDescription(Trim(rsMat("unit_Cls")))
            
            .Cell(flexcpBackColor, .Rows - 1, bteColLotNo) = vbWhite 'Lot No
            .TextMatrix(.Rows - 1, bteColLotNo) = "" 'Lot No
            
            .TextMatrix(.Rows - 1, bteColUnitBom) = rsMat("unit_cls") 'Unit Cls
            .TextMatrix(.Rows - 1, bteColJumChild) = "" 'Utk Simpan
            
            .TextMatrix(.Rows - 1, bteColStockCls) = Trim(rsMat("stockItem"))
            .TextMatrix(.Rows - 1, bteColUnit) = .TextMatrix(i, bteColUnit) 'Unit Default
            .TextMatrix(.Rows - 1, bteColDate) = Format(.TextMatrix(i, bteColReqDate), "yyyy-MM-dd") 'Tgl Prod
            
            .Cell(flexcpBackColor, .Rows - 1, bteColCls) = vbWhite
            .TextMatrix(.Rows - 1, bteColCls) = ""
            jmlChild = 1
        End If
        
        .TextMatrix(i, bteColJumChild) = jmlChild 'Utk cek jika dia add bisa dibawahnya langsung
        '*********************************************************************
        
        rsMat.MoveNext
    Loop
    Set rsMat = Nothing
End With
End Sub
Sub IsiGrid1(dailyseqno As Double, Factory As String, Factory_Name As String)
Dim ls_sql As String
Dim rs1 As New ADODB.Recordset

'ls_sql = " Select PM.womIn_No,PD.ChildItem_Code,IM.Item_Name,'" & Factory & "' Location,'" & Factory_Name & "' Description, " & vbCrLf & _
'                  "        Coalesce(ChildRequirement_Qty,0)Plan_Qty,(Coalesce(ChildRequirement_Qty,0)/Coalesce(DP.Qty,0)) * " & lblResultQty & " PlanResult,Result_Qty, " & vbCrLf & _
'                  "        Unit=Un.Description,Lot_No,MaterialConsump_Cls,DP.Schedule_Date,StockControl_Cls,Un.Unit_Cls" & vbCrLf & _
'                  " From PartSupplyRequest_Detail PD Left Join Daily_Production DP On PD.DailySeq_No=DP.Seq_No Left Join PartSupplyRequest_Master PM ON PM.SupplyRec_No=PD.SupplyRec_No " & vbCrLf & _
'                  "      Left Join Item_Master IM On RTRIM(LTRIM(IM.Item_Code))=RTRIM(LTRIM(PD.ChildItem_Code)) " & vbCrLf & _
'                  "      Left Join Unit_Cls Un On UN.Unit_Cls=IM.Unit_Cls    " & vbCrLf & _
'                  "      Left Join( " & vbCrLf & _
'                  "                 Select PS.ChildItem_Code,SUM(Coalesce(ChildRequirement_Qty,0)) Result_Qty,Coalesce(PR.MaterialConsump_Cls,'0') MaterialConsump_Cls From Part_Supply PS  " & vbCrLf & _
'                  "                 Left Join Part_Receipt PR On Ltrim(Rtrim(PS.DO_No))=rtrim(ltrim(Convert(Char(11),PR.Seq_No))) " & vbCrLf & _
'                  "                 Where DailySeq_No=" & dailyseqno & " " & vbCrLf & _
'                  "                 Group by PS.ChildItem_Code,Coalesce(PR.MaterialConsump_Cls,'0') "
'
'ls_sql = ls_sql + "               )A On PD.ChildItem_Code=A.ChildItem_Code " & vbCrLf & _
'                  " where DP.Seq_No=" & dailyseqno & "  "

'optimasi query yang di atas lama bro... 20150826


ls_sql = " Declare @ReceiptSeq_No as varchar(11) " & vbCrLf & _
                  "  " & vbCrLf & _
                  " --Set @ReceiptSeq_No=(Select Seq_No from Part_Receipt PR WHERE   DailySeq_No = " & dailyseqno & " ) " & vbCrLf & _
                  "  " & vbCrLf & _
                  "  SELECT PM.womIn_No , " & vbCrLf & _
                  "         PD.ChildItem_Code , " & vbCrLf & _
                  "         IM.Item_Name , " & vbCrLf & _
                  "         '" & Factory & "' Location , " & vbCrLf & _
                  "         '" & Factory_Name & "' Description , " & vbCrLf & _
                  "         COALESCE(ChildRequirement_Qty, 0) Plan_Qty , " & vbCrLf & _
                  "         ( COALESCE(ChildRequirement_Qty, 0) / COALESCE(DP.Qty, 0) ) * " & CDbl(lblResultQty) & " PlanResult , " & vbCrLf

ls_sql = ls_sql + "         Result_Qty , " & vbCrLf & _
                  "         Unit = Un.Description , " & vbCrLf & _
                  "         Lot_No , " & vbCrLf & _
                  "         '' MaterialConsump_Cls , " & vbCrLf & _
                  "         DP.Schedule_Date , " & vbCrLf & _
                  "         StockControl_Cls , " & vbCrLf & _
                  "         Un.Unit_Cls " & vbCrLf & _
                  "  " & vbCrLf & _
                  "  FROM   (Select * From  PartSupplyRequest_Detail where DailySeq_No=" & dailyseqno & ") PD " & vbCrLf & _
                  "         LEFT JOIN Daily_Production DP ON PD.DailySeq_No = DP.Seq_No " & vbCrLf & _
                  "         LEFT JOIN PartSupplyRequest_Master PM ON PM.SupplyRec_No = PD.SupplyRec_No " & vbCrLf

ls_sql = ls_sql + "         LEFT JOIN Item_Master IM ON IM.Item_Code = PD.ChildItem_Code " & vbCrLf & _
                  "         LEFT JOIN Unit_Cls Un ON UN.Unit_Cls = IM.Unit_Cls " & vbCrLf & _
                  "         LEFT JOIN ( SELECT  PS.ChildItem_Code , " & vbCrLf & _
                  "                             SUM(COALESCE(ChildRequirement_Qty, 0)) Result_Qty  " & vbCrLf & _
                  "                     FROM    Part_Supply PS " & vbCrLf & _
                  "                     Where Do_No IN (Select convert(varchar(11),Seq_No) from Part_Receipt PR WHERE   DailySeq_No = " & dailyseqno & ")" & vbCrLf & _
                  "                     GROUP BY PS.ChildItem_Code " & vbCrLf & _
                  "                   ) A ON PD.ChildItem_Code = A.ChildItem_Code " & vbCrLf & _
                  "  " & vbCrLf


Call headerGrid

If rs1.State <> adStateClosed Then rs1.Close

 Set rs1 = Db.Execute(ls_sql)
With grid

Do While Not rs1.EOF
    .Rows = .Rows + 1
        lblFNo.Caption = Trim(rs1("womIn_No")) 'WOMIN nO
        .Cell(flexcpBackColor, .Rows - 1, bteColSelect) = vbWhite 'C
        .TextMatrix(.Rows - 1, bteColSelect) = ""
        .TextMatrix(.Rows - 1, bteColMatCode) = Trim(rs1("ChildItem_Code")) 'Item_Code
        .TextMatrix(.Rows - 1, bteColDesc) = Trim(rs1("item_Name"))
        
        .TextMatrix(.Rows - 1, bteColLoc) = factoryCD  'WH Code
        .TextMatrix(.Rows - 1, bteColLocDesc) = Trim(rs1("Description"))
        .Cell(flexcpBackColor, .Rows - 1, bteColResult) = vbWhite 'Result
        .TextMatrix(.Rows - 1, bteColReqDate) = Format(rs1("Schedule_Date"), "dd MMM yyyy")
        .TextMatrix(.Rows - 1, bteColPlanReq) = Format(IIf(IsNull(rs1("Plan_Qty")), 0, rs1("Plan_Qty")), gs_formatQtyBOM)
        .TextMatrix(.Rows - 1, bteColPlanResult) = Format(IIf(IsNull(rs1("PlanResult")), 0, rs1("PlanResult")), gs_formatQtyBOM)
        .TextMatrix(.Rows - 1, bteColResult) = Format(IIf(IsNull(rs1("Result_Qty")), 0, rs1("Result_Qty")), gs_formatQtyBOM)
        ' ----------------------------------
        
        .TextMatrix(.Rows - 1, bteColUnitReq) = rs1("unit")
        
        .Cell(flexcpBackColor, .Rows - 1, bteColLotNo) = vbWhite 'Result
        .TextMatrix(.Rows - 1, bteColLotNo) = "" 'Lot No
        .TextMatrix(.Rows - 1, bteColUnitBom) = rs1("unit_Cls") 'Unit Cls
        .TextMatrix(.Rows - 1, bteColJumChild) = 1 'Update
        .TextMatrix(.Rows - 1, bteColChildReq) = "" 'Tampung Qty sblm Update
        .TextMatrix(.Rows - 1, bteColStockCls) = Trim(rs1("StockControl_Cls"))
        .TextMatrix(.Rows - 1, bteColUnit) = rs1("Unit_Cls") 'unit Default
        
        .Cell(flexcpBackColor, .Rows - 1, bteColCls) = vbWhite
        .TextMatrix(.Rows - 1, bteColCls) = Trim(rs1("MaterialConsump_Cls")) & ""
        .Cell(flexcpBackColor, .Rows - 1, bteColMatCode, .Rows - 1, .ColS - 1) = &HE0E0E0
        
        .Rows = .Rows + 1
        .Cell(flexcpBackColor, .Rows - 1, bteColSelect) = vbWhite 'C
        .Cell(flexcpBackColor, .Rows - 1, bteColResult) = vbWhite 'C
        .Cell(flexcpBackColor, .Rows - 1, bteColLotNo) = vbWhite 'C

        .TextMatrix(.Rows - 1, bteColSelect) = ""
        .TextMatrix(.Rows - 1, bteColResult) = Format(IIf(IsNull(rs1("PlanResult")), 0, rs1("PlanResult")), gs_formatQtyBOM)
        .TextMatrix(.Rows - 1, bteColMatCode) = Trim(rs1("ChildItem_Code"))
        .TextMatrix(.Rows - 1, bteColLoc) = factoryCD  'WH Code
        .TextMatrix(.Rows - 1, bteColUnitReq) = rs1("unit")
        .TextMatrix(.Rows - 1, bteColUnitBom) = rs1("unit_Cls") 'Unit Cls
        .TextMatrix(.Rows - 1, bteColUnit) = rs1("Unit_Cls") 'unit Default
        .TextMatrix(.Rows - 1, bteColDate) = rs1("Schedule_Date")
        .TextMatrix(.Rows - 1, bteColStockCls) = Trim(rs1("StockControl_Cls"))
        
        rs1.MoveNext

Loop

End With
End Sub
Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim tambahBrs As Integer
Dim mulai As Long

With grid
    If Row <> 0 Then
        If Col = bteColSelect Then
            If .TextMatrix(Row, bteColSelect) = "C" Then
                tambahBrs = Row + Val(.TextMatrix(Row, bteColJumChild)) + 1
                .AddItem "", tambahBrs
                .Cell(flexcpBackColor, tambahBrs, bteColSelect) = vbWhite 'C
                .TextMatrix(tambahBrs, bteColMatCode) = .TextMatrix(Row, bteColMatCode) 'Item CD
                .TextMatrix(tambahBrs, bteColLoc) = .TextMatrix(Row, bteColLoc) 'WH Code
                
                .Cell(flexcpBackColor, tambahBrs, bteColResult) = vbWhite 'Result
                .TextMatrix(tambahBrs, bteColResult) = Format(0, gs_formatQtyBOM) 'Result Qty
                
                .TextMatrix(tambahBrs, bteColUnitReq) = .TextMatrix(Row, bteColUnitReq) 'Unit
                .TextMatrix(tambahBrs, bteColUnitBom) = .TextMatrix(Row, bteColUnitBom) 'Unit Cls
                
                .Cell(flexcpBackColor, tambahBrs, bteColLotNo) = vbWhite 'Lot No
                .TextMatrix(tambahBrs, bteColLotNo) = "" 'Lot No
                
                .TextMatrix(tambahBrs, bteColStockCls) = .TextMatrix(Row, bteColStockCls) 'Stock Item CLs
                .TextMatrix(Row, bteColJumChild) = .TextMatrix(Row, bteColJumChild) + 1
                
                .TextMatrix(tambahBrs, bteColUnit) = .TextMatrix(Row, bteColUnit) 'Default Unit
                .TextMatrix(tambahBrs, bteColDate) = Format(.TextMatrix(Row, bteColReqDate), "yyyy-MM-dd") 'Tgl Prod
                
                .Cell(flexcpBackColor, tambahBrs, bteColCls) = vbWhite 'Cls
                .TextMatrix(tambahBrs, bteColCls) = "" 'Cls
            End If
        
        ElseIf Col = bteColResult Then 'Result
            If IsNumeric(.TextMatrix(Row, bteColResult)) = False Then .TextMatrix(Row, bteColResult) = Format(0, gs_formatQtyBOM)
            If IsNumeric(.TextMatrix(Row, bteColResult)) = True Then
                If CDbl(.TextMatrix(Row, bteColResult)) > gd_MaxQty Then
                    LblErrMsg = DisplayMsg(4045) & " " & gd_MaxQty & " !"
                    .TextMatrix(Row, bteColResult) = Format(dblQtyTemp, gs_formatQtyBOM)
                    grid.SetFocus
                Else
                    LblErrMsg = ""
                    .TextMatrix(Row, bteColResult) = Format(.TextMatrix(Row, bteColResult), gs_formatQtyBOM)
                End If
            End If
        End If
    End If
    
    If Col = bteColResult Then
       If .TextMatrix(Row, Col) > 0 Then
          TandaZero = 0
       End If
    End If
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid
    If Row <> 0 Then
        If Col = bteColCls Then Cancel = 1: Exit Sub
        If (Col <> bteColSelect Or Col <> bteColResult Or Col <> bteColLotNo) And .Cell(flexcpBackColor, .Row, .Col) <> vbWhite Then Cancel = 1
        If Col = bteColSelect Then 'C / D
            .EditMaxLength = 1
        ElseIf Col = bteColResult Then 'Qty
            .EditMaxLength = 12
            dblQtyTemp = CDbl(grid.TextMatrix(Row, bteColResult))
        ElseIf Col = bteColLotNo Then 'Lot No
            .EditMaxLength = 7
        End If
    End If
End With
End Sub

Private Sub grid_Click()
With grid
    nilKosong = True
    If .Row <> -1 Then
        If .Col = bteColResult Or .Col = bteColLotNo And .Cell(flexcpBackColor, .Row, .Col) = vbWhite Then
            .FocusRect = flexFocusInset
        Else
            .FocusRect = flexFocusNone
        End If
        
        If .Col = bteColCls Then
            If .TextMatrix(.Row, bteColCls) <> "" Then cboCls = .TextMatrix(.Row, bteColCls)
            cboCls.Visible = True
            cboCls.Left = .Cell(flexcpLeft, .Row, bteColCls)
            cboCls.top = .Cell(flexcpTop, .Row, bteColCls)
            cboCls.Width = .Cell(flexcpWidth, .Row, bteColCls)
            cboCls.SetFocus
            tampungBrs = .Row
        Else
            cboCls.Visible = False
        End If
    End If
    nilKosong = False
End With
End Sub

Private Sub Grid_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    cboCls.Visible = False
End Sub

Private Sub cbocls_Change()
If nilKosong Then Exit Sub
With grid
    If tampungBrs > 0 Then .TextMatrix(tampungBrs, bteColCls) = cboCls
End With
End Sub

Private Sub CboCls_LostFocus()
    cboCls.Visible = False
    With grid
        If tampungBrs > 0 Then .TextMatrix(tampungBrs, bteColCls) = cboCls
    End With
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
With grid
    If KeyCode = vbKeyRight Or KeyCode = vbKeyTab Then
        If .Col = bteColSelect Then
            .Col = bteColPlanResult
        ElseIf .Col = bteColResult Then
            .Col = bteColUnitReq
        Else
            .Col = -1
        End If
        .SetFocus
    
    ElseIf KeyCode = vbKeyLeft Then
        If .Col = bteColLotNo Then
            .Col = bteColUnitReq
        ElseIf .Col = bteColResult Then
            .Col = -1
        Else
            .Col = -1
        End If
        .SetFocus
    End If
End With
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

With grid
    If Col = bteColSelect Then
        If KeyAscii = Asc(".") Then KeyAscii = 0: Exit Sub
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If .Cell(flexcpBackColor, Row, bteColResult) = vbWhite Then 'Detail
            If KeyAscii <> Asc("D") And _
                KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And _
                KeyAscii <> vbKeyReturn Then _
                KeyAscii = 0: Exit Sub
        Else 'Header
            If KeyAscii <> Asc("C") And _
                KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And _
                KeyAscii <> vbKeyReturn Then _
                KeyAscii = 0: Exit Sub
        End If
    
    ElseIf Col = bteColResult Then 'Qty Result
        If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack _
            And KeyAscii <> vbKeyDelete And KeyAscii <> Asc(".") And KeyAscii <> vbKeyReturn Then _
            KeyAscii = 0
    End If
End With
End Sub

Private Sub Command1_Click(Index As Integer)
Dim rsCek As New ADODB.Recordset
Dim tampungQty As Double, Zero As Double
Dim pesan As String, X As Double, TempSerial As String

Me.MousePointer = vbHourglass
Select Case Index
    Case 0: 'Submit
        
        If hakUpdate("frmProdMaterialComp") = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
        '************************ Cek Qty Kosong / tdk ************************
        
        If completeCls = 1 Then LblErrMsg = DisplayMsg(1110): Me.MousePointer = vbDefault: Exit Sub 'Daily Completed
        'If seqNoProdReceipt <> 0 Then LblErrMsg = DisplayMsg(1117): Me.MousePointer = vbDefault: Exit Sub  'Already Used By Prod Receipt
        
        With grid
            For i = 1 To .Rows - 1
                If .TextMatrix(i, bteColSelect) <> "D" And .TextMatrix(i, bteColChildReq) <> "" And CDbl(.TextMatrix(i, bteColResult)) = 0 Then
                    LblErrMsg = DisplayMsg(1068) 'Input Qty Result
                    .Col = bteColResult: .Row = i
                    .SetFocus
                    Me.MousePointer = vbDefault: Exit Sub
                End If
            Next i
        End With
        '************************************************************************
        
        dbTransfer.ConnectionTimeout = 0
        dbTransfer.CommandTimeout = 0
        dbTransfer.Open Db.ConnectionString
        dbTransfer.BeginTrans

        Call frmProdResult.simpanUbah(dbTransfer)  'Input Receipt
        Call newCls.HapusDataSupp(dbTransfer, "'" & KeyProd & "'", blnFix, thnFix)
        Call prosesSubmit
        'Call frmProdResult.ErrProsesBag(dbTransfer)   'Input Bag
        Call newCls.UpdateCompleteReq(dbTransfer, Trim(lblitem.Caption), lblLot, schedule_date)
        
        If TandaZero = 1 Then
            sql = "delete Part_Receipt where Seq_No  = " & KeyProd
            dbTransfer.Execute sql
    
            sql = "delete WorkingTime_Master where ProductionSeq_No = " & KeyProd
            dbTransfer.Execute sql
            
            With frmProdResult
            If .TxtSerialFrom.Text <> "" And .TxtSerialTo.Text <> "" Then
                awal = Val(Mid(.TxtSerialFrom, 2, 6))
                akhir = Val(Mid(.TxtSerialTo, 2, 6))

                For X = awal To akhir
                    TempSerial = Left(.TxtSerialFrom, 1) & Mid(1000000 + X, 2, 6)
'                    sql = "Update Serial_Detail " & vbCrLf & _
'                        " Set Result_No=Null, Serial_Status='2' " & vbCrLf & _
'                        " where  item_code='" & lblitem.Caption & "' and " & vbCrLf & _
'                        " Serial_No='" & TempSerial & "'"
'                    dbTransfer.Execute sql
                Next X
                
             End If
          End With
            
        End If

        
        
        dbTransfer.CommitTrans
        dbTransfer.Close
        
        If strChildItemCD <> "" Then Call newCls.UpdateRequirementResult(Db, tglProd, "'" & lblitem & "'", lblLot, Left(strChildItemCD, Len(strChildItemCD) - 1))
        strChildItemCD = ""
        
        'Call IsiGrid(KeyProd, False, dailyseqno)
       ' Call IsiGrid1(dailyseqno, factoryCD, LblDesc)
        LblErrMsg = DisplayMsg(1101) 'Update Success
        
        sql = "Select ProductionSeq_No from WorkingTime_Master where ProductionSeq_No = '" & KeyProd & "'"
        Set rsCek = Db.Execute(sql)
        If rsCek.EOF Then
            With frmProdWorkingTime
                .ProdSeqNo = KeyProd
                .cmdsubmenu.Caption = "&Back"
                .ViewDt (1)
                .Show
            End With
        End If
                
        Set rsCek = Nothing
        
    Case 1: 'Cancel
        LblErrMsg = ""
        Call IsiGrid1(dailyseqno, factoryCD, lbldesc)
        TandaZero = 0
    Case 2: 'Zero
    
        For Zero = 1 To grid.Rows - 1
            If grid.Cell(flexcpBackColor, Zero, bteColResult) = vbWhite Then
               grid.TextMatrix(Zero, bteColResult) = "0"
            End If
        
        Next Zero
        frmProdResult.txtQty = 0
        TandaZero = 1
End Select
Me.MousePointer = vbDefault
End Sub

Sub prosesSubmit()
Dim totResult As Double, stockResult As Double
Dim qtyFormulaProd As Double, convertQtyFormula As Double
    
    '******************** Proses Stock Detail (Pengaruh ke Stock Jika P1) ***************
    adaErr = 0
    
    With grid
        For i = 2 To .Rows - 1
            If .TextMatrix(i, bteColSelect) = "D" Or CDbl(.TextMatrix(i, bteColResult)) = 0 Then 'Jika 0 = Delete Data
            Else
                ChildItemCD = Trim(.TextMatrix(i, bteColMatCode))
                childWHCode = Trim(.TextMatrix(i, bteColLoc))
                childUnitCls = Trim(.TextMatrix(i, bteColUnitBom))
                defaultUnit = Trim(.TextMatrix(i, bteColUnit))
                stockItem = Trim(.TextMatrix(i, bteColStockCls))
                matConsumpCls = Trim(.TextMatrix(i, bteColCls))
                
                childQty = CDbl(.TextMatrix(i, bteColResult))
                qtyStock = ClsProc.nilConvertUnit(CDbl(childQty), childUnitCls, defaultUnit)
                
                
                If .Cell(flexcpBackColor, i, bteColMatCode) <> &HE0E0E0 Then 'Jika Bukan Header baru diproses
                    ChildLotNo = Trim(.TextMatrix(i, bteColLotNo))
                    tglProd = Trim(.TextMatrix(i, bteColDate))
                    childSeqNo = Val(.TextMatrix(i, bteColJumChild))
                                    
                    Call simpanUbah
                    strChildItemCD = strChildItemCD & "'" & ChildItemCD & "',"
                End If
            End If
        Next i
    End With
End Sub

Sub hapusDt()
Dim rsDetailStock As New ADODB.Recordset
Dim rsCek As New ADODB.Recordset
    
    If stockItem = "01" Then
        Call newCls.updateStock(factoryCD, ChildItemCD, qtyStock, "", _
            Format(lblDt, "yyyy-MM-dd"), blnFix, thnFix, dbTransfer, "Supply", 0, 0)
    End If
        
    sql = "delete Part_Supply Where Seq_no = " & childSeqNo
    dbTransfer.Execute sql
End Sub

Sub simpanUbah()
    Call newCls.inputSupply(dbTransfer, 0, Trim(factoryCD), Trim(lineCD), Trim(factoryCD), _
        Format(lblDt, "yyyy-MM-dd"), ChildItemCD, ChildLotNo, "S", CDbl(childQty), CDbl(qtyStock), childUnitCls, _
        Trim(lblitem), Trim(lblLot), CStr(KeyProd), tglProd, matConsumpCls)

    If stockItem = "01" Then
        Call newCls.updateStock(factoryCD, ChildItemCD, qtyStock, _
            "", Format(lblDt, "yyyy-MM-dd"), blnFix, thnFix, dbTransfer, "Supply", 0, 1)
    End If
End Sub

'******************************************************************

'********************* Out ***********************
Private Sub CmdSubMenu_Click()
    With frmProdResult
        Call .kosongBwh
        Call .IsiGrid
        .Command1(0).Enabled = False
        .Show
    End With
    DoEvents
    Unload Me
End Sub
'********************************************

Private Sub CreateTableTemp()
'On Error Resume Next

Dim SqlCreate As String

SqlCreate = " Drop Table TempStructure " & vbCrLf & _
                          "  " & vbCrLf & _
                          " Create Table TempStructure (UTAMA Char(25), Parent_ItemCode Char(25),Item_Code Char(25),QtyBOM Numeric(18,5), Control_cls Char(2), Unit_cls Char(2), ItemType Char(15)) " & vbCrLf
                          
Db.Execute (SqlCreate)

Call GetConsumtionData(lblitem, lblitem, 1)

End Sub
Private Sub GetConsumtionData(Utama As String, ItemParent As String, qtyParent As Double)

Dim SqlSearch As String, Strtype As String, SqlChild As String
Dim RsSearch As New ADODB.Recordset, rsChild As New ADODB.Recordset

SqlSearch = " select Parent_ItemCode, BM.Item_Code, isnull(Qty,0) QtyBOM, Production_Cls, StockControl_Cls, BM.Unit_Cls  " & vbCrLf & _
                          "     From BOM_MASTER BM  " & vbCrLf & _
                          "     Inner Join Item_Master IM on BM.Item_Code=IM.Item_Code " & vbCrLf & _
                          "         Where Parent_ItemCode='" & ItemParent & "' " & vbCrLf & _
                          "   and Start_Date <='" & Format(lblDt, "YYYYMMDD") & "'" & _
                          "   and End_Date >= '" & Format(lblDt, "YYYYMMDD") & "'"
                          

RsSearch.Open SqlSearch, Db, adOpenForwardOnly, adLockReadOnly

Do While Not RsSearch.EOF
    
    SqlChild = "Select * From BOM_Master Where Parent_ItemCode ='" & RsSearch("Item_Code") & "'"
    rsChild.Open SqlChild, Db, adOpenForwardOnly, adLockReadOnly
    
    If Not rsChild.EOF Then
        If RsSearch("StockControl_Cls") = "02" Then
            Strtype = "Parent"
            Call GetConsumtionData(Utama, RsSearch("Item_Code"), qtyParent * RsSearch("QtyBOM"))
        Else
            Strtype = "Child"
        End If
    Else
                Strtype = "Child"
    End If
    
    rsChild.Close
    
    SqlSearch = "Insert Into TempStructure Values ('" & Utama & "','" & Trim(RsSearch("Parent_ItemCode")) & "','" & Trim(RsSearch("Item_Code")) & "'," & _
                    RsSearch("QtyBOM") * qtyParent & "," & vbCrLf & _
                    " '" & Trim(RsSearch("StockControl_Cls")) & "','" & Trim(RsSearch("Unit_Cls")) & "','" & Strtype & "')"
    
    Db.Execute SqlSearch
    
    RsSearch.MoveNext
Loop

End Sub
