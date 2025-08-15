VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm_item_inquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Item Inquiry"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frm_item_inquiry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10545
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13110
      TabIndex        =   17
      Top             =   300
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Copy"
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
      Left            =   11340
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9825
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
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
      Index           =   0
      Left            =   13815
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9825
      Width           =   1155
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Delete"
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
      Left            =   12585
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9825
      Width           =   1155
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
      Left            =   255
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9825
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   255
      TabIndex        =   10
      Top             =   9045
      Width           =   14715
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
         TabIndex        =   11
         Top             =   210
         Width           =   14310
      End
   End
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
      Index           =   0
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8535
      Width           =   1155
   End
   Begin VB.TextBox txtPencarian 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,###"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Left            =   1140
      MaxLength       =   30
      TabIndex        =   0
      Text            =   "AAAAAAA"
      Top             =   8565
      Width           =   3435
   End
   Begin VB.ComboBox cboCari 
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
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   8565
      Width           =   2775
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080FFFF&
      Caption         =   "Filter"
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8535
      Width           =   1155
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080FFFF&
      Caption         =   "Refresh"
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
      Left            =   10695
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8535
      Width           =   1155
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
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
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9825
      Width           =   1155
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   7320
      Left            =   270
      TabIndex        =   16
      Top             =   1035
      Width           =   14655
      _cx             =   25850
      _cy             =   12912
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
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   0
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
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
      Height          =   5550
      Left            =   270
      TabIndex        =   15
      Top             =   1035
      Width           =   11535
      _cx             =   20346
      _cy             =   9790
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
      BackColorFixed  =   12640511
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483624
      BackColorAlternate=   -2147483624
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   55
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
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
      Editable        =   0
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Master Inquiry"
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
      Left            =   6450
      TabIndex        =   14
      Top             =   315
      Width           =   2280
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search :"
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
      TabIndex        =   13
      Top             =   8595
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By :"
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
      Left            =   4770
      TabIndex        =   12
      Top             =   8625
      Width           =   360
   End
End
Attribute VB_Name = "frm_item_inquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rs_item As New ADODB.Recordset
Dim rs_bom_master As New ADODB.Recordset
Dim ib_isNumeric As Boolean
Public is_sql As String
Dim is_SqlFilter As String, sqlfinish As String, sqlpart As String, sqlreserve As String, sqlsuply As String
Dim sqlprovision As String, sqlexplosion As String, sqlmakebuy As String, sqlstockcontrol As String, sqlprod As String, sqlunit As String
Dim sqla As String, sqlB As String
Dim i As Long
Dim cari As String

Dim bteColProdCode As Byte
Dim bteColDesc As Byte
Dim bteColFGCls As Byte
Dim bteColFinishDesc As Byte
Dim bteColDrawNo As Byte
Dim bteColWHCode As Byte
Dim bteColAddress As Byte
Dim bteColSupplier As Byte
Dim bteColDelPlace As Byte
Dim bteColHSCode As Byte
Dim bteColManuCode As Byte
Dim bteColLineCode As Byte
Dim bteColPartCls As Byte
Dim bteColPartDesc As Byte
Dim bteColPartNo As Byte
Dim bteColResvCls As Byte
Dim bteColResvDesc As Byte
Dim bteColSuppCls As Byte
Dim bteColSuppDesc As Byte
Dim bteColProvCls As Byte
Dim bteColProvDesc As Byte
Dim bteColMatCls As Byte
Dim bteColMatDesc As Byte
Dim bteColThickness As Byte
Dim bteColWidth As Byte
Dim bteColWeight As Byte
Dim bteColGrossWeight As Byte
Dim bteColLength As Byte
Dim bteColSheetCoilCls As Byte
Dim bteColSheetCoilDesc As Byte
Dim bteColPitch As Byte
Dim bteColNoProduce As Byte
Dim bteColScrapWeight As Byte
Dim bteColDrawMat As Byte
Dim bteColDrawDesc As Byte
Dim bteColSurfaceCls As Byte
Dim bteColSurfaceDesc As Byte
Dim bteColHeatCls As Byte
Dim bteColHeatDesc As Byte
Dim bteColNoProcess As Byte
Dim bteColMatCoef As Byte
Dim bteColProcCoef As Byte
Dim bteColMinLot As Byte
Dim bteColLotQty As Byte
Dim bteColLotCoef As Byte
Dim bteColProdLead As Byte
Dim bteColYield As Byte
Dim bteColQtyCase As Byte
Dim bteColPackCls As Byte
Dim bteColPackDesc As Byte
Dim bteColPackItem As Byte
Dim bteColGroupCls As Byte
Dim bteColGroupDesc As Byte
Dim bteColProdCls As Byte
Dim bteColprodDesc As Byte
Dim bteColStdStock As Byte
Dim bteColSaveStock As Byte
Dim bteColMaxStock As Byte
Dim bteColMinStock As Byte
Dim bteColAlloDay As Byte
Dim bteColDelLead As Byte
Dim bteColMakeBuyCls As Byte
Dim bteColMakBuyDesc As Byte
Dim bteColControlCls As Byte
Dim bteColControlDesc As Byte
Dim bteColOrderPoint As Byte
Dim bteColUnitCls As Byte
Dim bteColUnitDesc As Byte
Dim bteColQtyBox As Byte
Dim bteColPackMatCls As Byte
Dim bteColPackMatDesc As Byte
Dim bteColAccountCode As Byte
Dim bteColExploCls As Byte
Dim bteColExploDesc As Byte
Dim bteColPICCls As Byte
Dim bteColPICDesc As Byte
Dim bteColStockCls As Byte
Dim bteColStockDesc As Byte
Dim bteColUseEndDay As Byte
Dim bteColLatUpdate As Byte

Private Sub uf_header()
Dim i As Integer

With VSFlexGrid1
    
    bteColProdCode = 0
    bteColPartNo = 1
    bteColDesc = 2
    bteColFGCls = 3
    bteColFinishDesc = 4
    bteColDrawNo = 5
    bteColWHCode = 6
    bteColAddress = 7
    bteColSupplier = 8
    bteColDelPlace = 9
    bteColHSCode = 10
    bteColManuCode = 11
    bteColLineCode = 12
    bteColPartCls = 13
    bteColPartDesc = 14
    bteColResvCls = 15
    bteColResvDesc = 16
    bteColSuppCls = 17
    bteColSuppDesc = 18
    bteColProvCls = 19
    bteColProvDesc = 20
    bteColMatCls = 21
    bteColMatDesc = 22
    bteColThickness = 23
    bteColWidth = 24
    bteColWeight = 25
    bteColGrossWeight = 26
    bteColLength = 27
    bteColSheetCoilCls = 28
    bteColSheetCoilDesc = 29
    bteColPitch = 30
    bteColNoProduce = 31
    bteColScrapWeight = 32
    bteColDrawMat = 33
    bteColDrawDesc = 34
    bteColSurfaceCls = 35
    bteColSurfaceDesc = 36
    bteColHeatCls = 37
    bteColHeatDesc = 38
    bteColNoProcess = 39
    bteColMatCoef = 40
    bteColProcCoef = 41
    bteColMinLot = 42
    bteColLotQty = 43
    bteColLotCoef = 44
    bteColProdLead = 45
    bteColYield = 46
    bteColQtyCase = 47
    bteColPackCls = 48
    bteColPackDesc = 49
    bteColPackItem = 50
    bteColGroupCls = 51
    bteColGroupDesc = 52
    bteColProdCls = 53
    bteColprodDesc = 54
    bteColStdStock = 55
    bteColSaveStock = 56
    bteColMaxStock = 57
    bteColMinStock = 58
    bteColAlloDay = 59
    bteColDelLead = 60
    bteColMakeBuyCls = 61
    bteColMakBuyDesc = 62
    bteColControlCls = 63
    bteColControlDesc = 64
    bteColOrderPoint = 65
    bteColUnitCls = 66
    bteColUnitDesc = 67
    bteColQtyBox = 68
    bteColPackMatCls = 69
    bteColPackMatDesc = 70
    bteColAccountCode = 71
    bteColExploCls = 72
    bteColExploDesc = 73
    bteColPICCls = 74
    bteColPICDesc = 75
    bteColStockCls = 76
    bteColStockDesc = 77
    bteColUseEndDay = 78
    bteColLatUpdate = 79
    
    .clear
    .FixedCols = 1
    .ColS = 80
    .Rows = 1
    
    '#Set Header
    .TextMatrix(0, bteColProdCode) = "Product Code"
    .TextMatrix(0, bteColPartNo) = "Part Number"
    .TextMatrix(0, bteColDesc) = "Description"
    .TextMatrix(0, bteColFGCls) = "Finish Good Part Cls"
    .TextMatrix(0, bteColFinishDesc) = "Finish Desc"
    .TextMatrix(0, bteColDrawNo) = "Drawing Number"
    .TextMatrix(0, bteColWHCode) = "Warehouse Code"
    .TextMatrix(0, bteColAddress) = "Address"
    .TextMatrix(0, bteColSupplier) = "Supplier Code"
    .TextMatrix(0, bteColDelPlace) = "Delivery Place"
    .TextMatrix(0, bteColHSCode) = "HS Code"
    .TextMatrix(0, bteColManuCode) = "Manufacture Code"
    .TextMatrix(0, bteColLineCode) = "Line Code"
    .TextMatrix(0, bteColPartCls) = "Part Cls"
    .TextMatrix(0, bteColPartDesc) = "Part Desc"
    .TextMatrix(0, bteColResvCls) = "Reserve Cls"
    .TextMatrix(0, bteColResvDesc) = "Reserve Desc"
    .TextMatrix(0, bteColSuppCls) = "Supply Cls"
    .TextMatrix(0, bteColSuppDesc) = "Supply Desc"
    .TextMatrix(0, bteColProvCls) = "Provision Cls"
    .TextMatrix(0, bteColProvDesc) = "Provision Desc"
    .TextMatrix(0, bteColMatCls) = "Material Cls"
    .TextMatrix(0, bteColMatDesc) = "Material Desc"
    .TextMatrix(0, bteColThickness) = "Thickness"
    .TextMatrix(0, bteColWidth) = "Width"
    .TextMatrix(0, bteColWeight) = "Weight"
    .TextMatrix(0, bteColGrossWeight) = "Gross Weight"
    .TextMatrix(0, bteColLength) = "Length"
    .TextMatrix(0, bteColSheetCoilCls) = "Sheet Coil Cls"
    .TextMatrix(0, bteColSheetCoilDesc) = "Sheet Coil Desc"
    .TextMatrix(0, bteColPitch) = "Pitch"
    .TextMatrix(0, bteColNoProduce) = "Number Producible"
    .TextMatrix(0, bteColScrapWeight) = "Scrap Weight"
    .TextMatrix(0, bteColDrawMat) = "Drawing Material"
    .TextMatrix(0, bteColDrawDesc) = "Drawing Desc"
    .TextMatrix(0, bteColSurfaceCls) = "Surface Treatment"
    .TextMatrix(0, bteColSurfaceDesc) = "Surface Desc"
    .TextMatrix(0, bteColHeatCls) = "Heat Treatment"
    .TextMatrix(0, bteColHeatDesc) = "Heat Desc"
    .TextMatrix(0, bteColNoProcess) = "Number Process"
    .TextMatrix(0, bteColMatCoef) = "Material Coefficient"
    .TextMatrix(0, bteColProcCoef) = "Process Coefficient"
    .TextMatrix(0, bteColMinLot) = "Min Lot"
    .TextMatrix(0, bteColLotQty) = "Lot Qty"
    .TextMatrix(0, bteColLotCoef) = "Lot Coefficient"
    .TextMatrix(0, bteColProdLead) = "Product Lead Time"
    .TextMatrix(0, bteColYield) = "Yield %"
    .TextMatrix(0, bteColQtyCase) = "Qty/Case"
    .TextMatrix(0, bteColPackCls) = "Packing Style Cls"
    .TextMatrix(0, bteColPackDesc) = "Packing Style Desc"
    .TextMatrix(0, bteColPackItem) = "Packing Item Code"
    .TextMatrix(0, bteColGroupCls) = "Group Cls"
    .TextMatrix(0, bteColGroupDesc) = "Group Desc"
    .TextMatrix(0, bteColProdCls) = "Production Cls"
    .TextMatrix(0, bteColprodDesc) = "Production Desc"
    .TextMatrix(0, bteColStdStock) = "Standard Stock"
    .TextMatrix(0, bteColSaveStock) = "Safety Stock"
    .TextMatrix(0, bteColMaxStock) = "Max Stock"
    .TextMatrix(0, bteColMinStock) = "Min Stock"
    .TextMatrix(0, bteColAlloDay) = "Allowance Day"
    .TextMatrix(0, bteColDelLead) = "Delivery Lead Time"
    .TextMatrix(0, bteColMakeBuyCls) = "Make/Buy"
    .TextMatrix(0, bteColMakBuyDesc) = "Make/Buy Desc"
    .TextMatrix(0, bteColControlCls) = "Control Cls"
    .TextMatrix(0, bteColControlDesc) = "Control Desc"
    .TextMatrix(0, bteColOrderPoint) = "Order Point Qty"
    .TextMatrix(0, bteColUnitCls) = "Unit Cls"
    .TextMatrix(0, bteColUnitDesc) = "Unit Desc"
    .TextMatrix(0, bteColQtyBox) = "Qty/Box"
    .TextMatrix(0, bteColPackMatCls) = "Packing Style Material Cls"
    .TextMatrix(0, bteColPackMatDesc) = "Packing Style Material Desc"
    .TextMatrix(0, bteColAccountCode) = "Accounting Code"
    .TextMatrix(0, bteColExploCls) = "Explosion Cls"
    .TextMatrix(0, bteColExploDesc) = "Explosion Desc"
    .TextMatrix(0, bteColPICCls) = "Person In Charge"
    .TextMatrix(0, bteColPICDesc) = "Person In Charge Desc"
    .TextMatrix(0, bteColStockCls) = "Stock Control Cls"
    .TextMatrix(0, bteColStockDesc) = "Stock Control Desc"
    .TextMatrix(0, bteColUseEndDay) = "Use End Day"
    .TextMatrix(0, bteColLatUpdate) = "Last update"
    
    '#Set Format
    .AddItem ""
    .TextMatrix(1, bteColProdCode) = "text" '"Product Code"
    .TextMatrix(1, bteColPartNo) = "text" '"Part Number"
    .TextMatrix(1, bteColDesc) = "text" '"Description"
    .TextMatrix(1, bteColFGCls) = "text" '"Finish Good Part Cls"
    .TextMatrix(1, bteColFinishDesc) = "text" '"Finish Desc"
    .TextMatrix(1, bteColDrawNo) = "text" '"Drawing Number"
    .TextMatrix(1, bteColWHCode) = "text" '"Warehouse Code"
    .TextMatrix(1, bteColAddress) = "text" '"Address"
    .TextMatrix(1, bteColSupplier) = "text" '"Supplier Code"
    .TextMatrix(1, bteColDelPlace) = "text" '"Delivery Place"
    .TextMatrix(1, bteColHSCode) = "text" 'HS Code'
    .TextMatrix(1, bteColManuCode) = "text" '"Manufacture Code"
    .TextMatrix(1, bteColLineCode) = "text" '"Line Code"
    .TextMatrix(1, bteColPartCls) = "text" '"Part Cls"
    .TextMatrix(1, bteColPartDesc) = "text" '"Part Desc"
    .TextMatrix(1, bteColResvCls) = "text" '"Reserve Cls"
    .TextMatrix(1, bteColResvDesc) = "text" '"Reserve Desc"
    .TextMatrix(1, bteColSuppCls) = "text" '"Supply Cls"
    .TextMatrix(1, bteColSuppDesc) = "text" '"Supply Desc"
    .TextMatrix(1, bteColProvCls) = "text" '"Provision Cls"
    .TextMatrix(1, bteColProvDesc) = "text" '"Provision Desc"
    .TextMatrix(1, bteColMatCls) = "text" '"Material Cls"
    .TextMatrix(1, bteColMatDesc) = "text" '"Material Desc"
    .TextMatrix(1, bteColThickness) = "text" '"Thickness"
    .TextMatrix(1, bteColWidth) = "precision(2)" '"Width"
    .TextMatrix(1, bteColWeight) = "precision(2)" '"Weight"
    .TextMatrix(1, bteColGrossWeight) = "precision(2)" '"Gross Weight"
    .TextMatrix(1, bteColLength) = "precision(3)" '"Length"
    .TextMatrix(1, bteColSheetCoilCls) = "text" '"Sheet Coil Cls"
    .TextMatrix(1, bteColSheetCoilDesc) = "text" '"Sheet Coil Desc"
    .TextMatrix(1, bteColPitch) = "precision(2)" '"Pitch"
    .TextMatrix(1, bteColNoProduce) = "precision(2)" '"Number Producible"
    .TextMatrix(1, bteColScrapWeight) = "precision(2)" '"Scrap Weight"
    .TextMatrix(1, bteColDrawMat) = "text" '"Drawing Material"
    .TextMatrix(1, bteColDrawDesc) = "text" '"Drawing Desc"
    .TextMatrix(1, bteColSurfaceCls) = "text" '"Surface Treatment"
    .TextMatrix(1, bteColSurfaceDesc) = "text" '"Surface Desc"
    .TextMatrix(1, bteColHeatCls) = "text" '"Heat Treatment"
    .TextMatrix(1, bteColHeatDesc) = "text" '"Heat Desc"
    .TextMatrix(1, bteColNoProcess) = "precision(2)" '"Number Process"
    .TextMatrix(1, bteColMatCoef) = "precision(0)" '"Material Coefficient"
    .TextMatrix(1, bteColProcCoef) = "precision(0)" '"Process Coeficient"
    .TextMatrix(1, bteColMinLot) = "precision(0)" '"Min Lot"
    .TextMatrix(1, bteColLotQty) = "precision(0)" '"Lot Qty"
    .TextMatrix(1, bteColLotCoef) = "precision(0)" '"Lot Coefficient"
    .TextMatrix(1, bteColProdLead) = "precision(0)" '"Product Lead Time"
    .TextMatrix(1, bteColYield) = "precision(2)" '"Yield %"
    .TextMatrix(1, bteColQtyCase) = "precision(2)" '"Qty/Case"
    .TextMatrix(1, bteColPackCls) = "text" '"Packing Style Cls"
    .TextMatrix(1, bteColPackDesc) = "text" '"Packing Style Desc"
    .TextMatrix(1, bteColPackItem) = "text" '"Packing Item Code"
    .TextMatrix(1, bteColGroupCls) = "text" '"Group Cls"
    .TextMatrix(1, bteColGroupDesc) = "text" '"Group Desc"
    .TextMatrix(1, bteColProdCls) = "text" '"Production Cls"
    .TextMatrix(1, bteColprodDesc) = "text" '"Production Desc"
    .TextMatrix(1, bteColStdStock) = "precision(0)" '"Standard Stock"
    .TextMatrix(1, bteColSaveStock) = "precision(0)" '"Safety Stock"
    .TextMatrix(1, bteColMaxStock) = "precision(0)" '"Max Stock"
    .TextMatrix(1, bteColMinStock) = "precision(0)" '"Min Stock"
    .TextMatrix(1, bteColAlloDay) = "precision(0)" '"Allowance Day"
    .TextMatrix(1, bteColDelLead) = "precision(0)" '"Delivery Lead Time"
    .TextMatrix(1, bteColMakeBuyCls) = "text" '"Make/Buy"
    .TextMatrix(1, bteColMakBuyDesc) = "text" '"Make/Buy Desc"
    .TextMatrix(1, bteColControlCls) = "text" '"Control Cls"
    .TextMatrix(1, bteColControlDesc) = "text" '"Control Desc"
    .TextMatrix(1, bteColOrderPoint) = "precision(0)" '"Order Point Qty"
    .TextMatrix(1, bteColUnitCls) = "text" '"Unit Cls"
    .TextMatrix(1, bteColUnitDesc) = "text" '"Unit Desc"
    .TextMatrix(1, bteColQtyBox) = "precision(0)" '"Qty/Box"
    .TextMatrix(1, bteColPackMatCls) = "text" '"Packing Style Material Cls"
    .TextMatrix(1, bteColPackMatDesc) = "text" '"Packing Style Material Desc"
    .TextMatrix(1, bteColAccountCode) = "text" '"Accounting Code"
    .TextMatrix(1, bteColExploCls) = "text" '"Explosion Cls"
    .TextMatrix(1, bteColExploDesc) = "text" '"Explosion Desc"
    .TextMatrix(1, bteColPICCls) = "text" '"Person In Charge"
    .TextMatrix(1, bteColPICDesc) = "text" '"Person In Charge Desc"
    .TextMatrix(1, bteColStockCls) = "text" '"Stock Control Cls"
    .TextMatrix(1, bteColStockDesc) = "text" '"Stock Control Desc"
    .TextMatrix(1, bteColUseEndDay) = "text" '"Use End Day"
    .TextMatrix(1, bteColLatUpdate) = "text" '"Last update"

    .RowHidden(1) = True
    
    '#Set Length
    .AddItem ""
    .TextMatrix(2, bteColProdCode) = 18
    .TextMatrix(2, bteColPartNo) = 18
    .TextMatrix(2, bteColDesc) = Len("Description")
    .TextMatrix(2, bteColFGCls) = Len("Finish Good Part Cls")
    .TextMatrix(2, bteColFinishDesc) = Len("Finish Desc")
    .TextMatrix(2, bteColDrawNo) = Len("Drawing Number")
    .TextMatrix(2, bteColWHCode) = Len("Warehouse Code")
    .TextMatrix(2, bteColAddress) = Len("Address")
    .TextMatrix(2, bteColSupplier) = Len("Supplier Code")
    .TextMatrix(2, bteColDelPlace) = Len("Delivery Place")
    .TextMatrix(2, bteColHSCode) = Len("HS Code")
    .TextMatrix(2, bteColManuCode) = Len("Manufacture Code")
    .TextMatrix(2, bteColLineCode) = Len("Line Code")
    .TextMatrix(2, bteColPartCls) = Len("Part Cls")
    .TextMatrix(2, bteColPartDesc) = Len("Part Desc")
    .TextMatrix(2, bteColResvCls) = Len("Reserve Cls")
    .TextMatrix(2, bteColResvDesc) = Len("Reserve Desc")
    .TextMatrix(2, bteColSuppCls) = Len("Supply Cls")
    .TextMatrix(2, bteColSuppDesc) = Len("Supply Desc")
    .TextMatrix(2, bteColProvCls) = Len("Provision Cls")
    .TextMatrix(2, bteColProvDesc) = Len("Provision Desc")
    .TextMatrix(2, bteColMatCls) = Len("Material Cls")
    .TextMatrix(2, bteColMatDesc) = Len("Material Desc")
    .TextMatrix(2, bteColThickness) = Len("Thickness")
    .TextMatrix(2, bteColWidth) = Len("Width")
    .TextMatrix(2, bteColWeight) = Len("Weight")
    .TextMatrix(2, bteColGrossWeight) = Len("Gross Weight")
    .TextMatrix(2, bteColLength) = Len("Length")
    .TextMatrix(2, bteColSheetCoilCls) = Len("Sheet Coil Cls")
    .TextMatrix(2, bteColSheetCoilDesc) = Len("Sheet Coil Desc")
    .TextMatrix(2, bteColPitch) = Len("Pitch")
    .TextMatrix(2, bteColNoProduce) = Len("Number Producible")
    .TextMatrix(2, bteColScrapWeight) = Len("Scrap Weight")
    .TextMatrix(2, bteColDrawMat) = Len("Drawing Material")
    .TextMatrix(2, bteColDrawDesc) = Len("Drawing Desc")
    .TextMatrix(2, bteColSurfaceCls) = Len("Surface Treatment")
    .TextMatrix(2, bteColSurfaceDesc) = Len("Surface Desc")
    .TextMatrix(2, bteColHeatCls) = Len("Heat Treatment")
    .TextMatrix(2, bteColHeatDesc) = Len("Heat Desc")
    .TextMatrix(2, bteColNoProcess) = Len("Number Process")
    .TextMatrix(2, bteColMatCoef) = Len("Material Coefficient")
    .TextMatrix(2, bteColProcCoef) = Len("Process Coeficient")
    .TextMatrix(2, bteColMinLot) = Len("Min Lot")
    .TextMatrix(2, bteColLotQty) = Len("Lot Qty")
    .TextMatrix(2, bteColLotCoef) = Len("Lot Coefficient")
    .TextMatrix(2, bteColProdLead) = Len("Product Lead Time")
    .TextMatrix(2, bteColYield) = Len("Yield %")
    .TextMatrix(2, bteColQtyCase) = Len("Qty/Case")
    .TextMatrix(2, bteColPackCls) = Len("Packing Style Cls")
    .TextMatrix(2, bteColPackDesc) = Len("Packing Style Desc")
    .TextMatrix(2, bteColPackItem) = Len("Packing Item Code")
    .TextMatrix(2, bteColGroupCls) = Len("Group Cls")
    .TextMatrix(2, bteColGroupDesc) = Len("Group Desc")
    .TextMatrix(2, bteColProdCls) = Len("Production Cls")
    .TextMatrix(2, bteColprodDesc) = Len("Production Desc")
    .TextMatrix(2, bteColStdStock) = Len("Standard Stock")
    .TextMatrix(2, bteColSaveStock) = Len("Safety Stock")
    .TextMatrix(2, bteColMaxStock) = Len("Max Stock")
    .TextMatrix(2, bteColMinStock) = Len("Min Stock")
    .TextMatrix(2, bteColAlloDay) = Len("Allowance Day")
    .TextMatrix(2, bteColDelLead) = Len("Delivery Lead Time")
    .TextMatrix(2, bteColMakeBuyCls) = Len("Make/Buy")
    .TextMatrix(2, bteColMakBuyDesc) = Len("Make/Buy Desc")
    .TextMatrix(2, bteColControlCls) = Len("Control Cls")
    .TextMatrix(2, bteColControlDesc) = Len("Control Desc")
    .TextMatrix(2, bteColOrderPoint) = Len("Order Point Qty")
    .TextMatrix(2, bteColUnitCls) = Len("Unit Cls")
    .TextMatrix(2, bteColUnitDesc) = Len("Unit Desc")
    .TextMatrix(2, bteColQtyBox) = Len("Qty/Box")
    .TextMatrix(2, bteColPackMatCls) = Len("Packing Style Material Cls")
    .TextMatrix(2, bteColPackMatDesc) = Len("Packing Style Material Desc")
    .TextMatrix(2, bteColAccountCode) = Len("Accounting Code")
    .TextMatrix(2, bteColExploCls) = Len("Explosion Cls")
    .TextMatrix(2, bteColExploDesc) = Len("Explosion Desc")
    .TextMatrix(2, bteColPICCls) = Len("Person In Charge")
    .TextMatrix(2, bteColPICDesc) = Len("Person In Charge Desc")
    .TextMatrix(2, bteColStockCls) = Len("Stock Control Cls")
    .TextMatrix(2, bteColStockDesc) = Len("Stock Control Desc")
    .TextMatrix(2, bteColUseEndDay) = Len("Use End Day")
    .TextMatrix(2, bteColLatUpdate) = Len("Last update")

    .RowHidden(2) = True
    
   
    For i = 0 To .ColS - 1
        Select Case Trim(.TextMatrix(1, i))
            Case "precision(0)":
                                .ColFormat(i) = "#,##0"
                                .ColAlignment(i) = flexAlignRightCenter
            Case "precision(2)":
                                .ColFormat(i) = "#,##0.00"
                                .ColAlignment(i) = flexAlignRightCenter
            Case "precision(3)":
                                .ColFormat(i) = "#,##0.000"
                                .ColAlignment(i) = flexAlignRightCenter
            Case "text":
                                .ColAlignment(i) = flexAlignLeftCenter
        End Select
        .ColWidth(i) = Trim(.TextMatrix(2, i)) * 120
        If .ColWidth(i) < 1000 Then .ColWidth(i) = 1000
    Next
    
    .ColWidth(bteColDesc) = 4000
    .ColWidth(bteColLatUpdate) = 2500
    .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
    .ColHidden(bteColPackItem) = True '"Packing Item Code"

End With

End Sub

Private Sub uf_clear()
    txtPencarian = ""
    cboCari.ListIndex = 0
    VSFlexGrid1.clear
    LblErrMsg = ""
End Sub

Private Function uf_validateColumnType() As String

ib_isNumeric = IIf(Trim(VSFlexGrid1.TextMatrix(1, cboCari.ListIndex)) = "text", False, True)

If ib_isNumeric = True And IsNumeric(txtPencarian) = False Then
    uf_validateColumnType = "Please input a valid data !"
Else
    uf_validateColumnType = ""
End If

End Function

Private Sub uf_settingCombo()
With cboCari
    For i = 0 To VSFlexGrid1.ColS - 1
        If VSFlexGrid1.ColHidden(i) = False Then
            .AddItem Trim(VSFlexGrid1.TextMatrix(0, i))
        End If
    Next
End With
End Sub

Public Sub uf_settingGrid(ls_sql As String)
Dim j As Integer

If rs_item.State <> adStateClosed Then rs_item.Close
rs_item.CursorLocation = adUseClient
rs_item.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
rs_item.Requery
If rs_item.EOF = False Then
    rs_item.MoveFirst
    Call uf_header
    With VSFlexGrid1
        For i = 1 To rs_item.RecordCount
            .Rows = .Rows + 1
            For j = 0 To .ColS - 1
                 .TextMatrix(.Rows - 1, j) = IIf(IsNull(rs_item(j)), "", rs_item(j))
            Next j
            rs_item.MoveNext
        Next i
        .Row = 2
    End With
Else
    Call uf_header
End If
End Sub

Private Sub old_uf_searchTextInGrid()
Dim li_pos As Integer
Dim ls_GridText As String
Dim ls_SearchText As String

With VSFlexGrid1
    li_pos = 0
    For i = 3 To .Rows - 1
        ls_GridText = LCase(Trim(.TextMatrix(i, (cboCari.ListIndex) + IIf((cboCari.ListIndex) >= 48, 1, 0))))
        ls_SearchText = LCase(Trim(txtPencarian))
        If ls_GridText = ls_SearchText Then .Row = i: .SetFocus: If i <> 1 Then .TopRow = i - 1: Exit Sub
        If ls_SearchText <> "" Then
            If InStr(1, ls_GridText, ls_SearchText) > 0 Then
                .Row = i
                .SetFocus
                If i <> 1 Then .TopRow = i - 1
                Exit Sub
            Else
                li_pos = li_pos + 1
            End If
        Else
            li_pos = li_pos + 1
        End If
    Next i
    If li_pos = .Rows - 3 Then LblErrMsg = DisplayMsg("8012")
End With
End Sub

Private Sub uf_searchTextInGrid()

    Dim lngRow As Long
    Dim strText As String
    Dim booFound As Boolean
    
    With VSFlexGrid1
            
        For lngRow = .Row + 1 To .Rows - 1
            
            strText = Trim(.TextMatrix(lngRow, cboCari.ListIndex))
            If InStr(1, UCase(strText), UCase(txtPencarian.Text)) <> 0 Then
                .Row = lngRow
                .TopRow = lngRow
                booFound = True
                Exit For
            End If
                        
        Next
        .SetFocus
        If Not booFound Then
            .Row = 3
            .TopRow = 3
            LblErrMsg = DisplayMsg("8012")
        End If
        
    End With
    
End Sub

Private Sub cbocari_Click()
Select Case cboCari
    Case "Product Code": cari = "item_code"
    Case "Description": cari = "Item_Name"
    Case "Finish Good Part Cls": cari = "FinishGoodPart_Cls"
    Case "Finish Desc": cari = "Finish_desc"
    Case "Drawing Number": cari = "drawing_number"
    Case "Warehouse Code": cari = "wh_code"
    Case "Address": cari = "address"
    Case "Supplier Code": cari = "supplier_code"
    Case "Delivery Place": cari = "delivery_place"
    Case "HS Code": cari = "HS_Code"
    Case "Manufacture Code": cari = "manufacture_code"
    Case "Line Code": cari = "line_code"
    Case "Part Cls": cari = "part_cls"
    Case "Part Desc": cari = "part_desc"
    Case "Part Number": cari = "makeritem_code"
    Case "Reserve Cls": cari = "reserve_cls"
    Case "Reserve Desc": cari = "reserve_Desc"
    Case "Supply Cls": cari = "suply_cls"
    Case "Supply Desc": cari = "supply_Desc"
    Case "Provision Cls": cari = "provision_cls"
    Case "Provision Desc": cari = "provision_Desc"
    Case "Material Cls": cari = "material_cls"
    Case "Material Desc": cari = "material_desc"
    Case "Thickness": cari = "thickness"
    Case "Width": cari = "width"
    Case "Weight": cari = "weight"
    Case "Gross Weight": cari = "grossweight"
    Case "Length": cari = "length"
    Case "Sheet Coil Cls": cari = "sheetcoil_cls"
    Case "Sheet Coil Desc": cari = "sheet_desc"
    Case "Pitch": cari = "pitch"
    Case "Number Producible": cari = "number_producible"
    Case "Scrap Weight": cari = "scrap_weight"
    Case "Drawing Material": cari = "drawingmaterial_cls"
    Case "Drawing Desc": cari = "draw_desc"
    Case "Surface Treatment": cari = "surfacetreatment_cls"
    Case "Surface Desc": cari = "surface_desc"
    Case "Heat Treatment": cari = "heattreatment_cls"
    Case "Heat Desc": cari = "heat_desc"
    Case "Number Process": cari = "number_process"
    Case "Material Coefficient": cari = "material_coefficient"
    Case "Process Coefficient": cari = "process_coefficient"
    Case "Min Lot": cari = "min_lot"
    Case "Lot Qty": cari = "lot_qty"
    Case "Lot Coefficient": cari = "lot_coefficience"
    Case "Product Lead Time": cari = "product_readtime"
    Case "Yield %": cari = "yield_percentage"
    Case "Qty/Case": cari = "number_entering"
    Case "Packing Style Cls": cari = "packingstyle_cls"
    Case "Packing Style Desc": cari = "packing_desc"
    Case "Group Cls": cari = "group_cls"
    Case "Group Desc": cari = "group_desc"
    Case "Production Cls": cari = "production_cls"
    Case "Production Desc": cari = "production_desc"
    Case "Standard Stock": cari = "standard_stock"
    Case "Safety Stock": cari = "safety_stock"
    Case "Max Stock": cari = "max_stock"
    Case "Min Stock": cari = "min_stock"
    Case "Allowance Day": cari = "alowance_day"
    Case "Delivery Lead Time": cari = "delivery_readtime"
    Case "Make/Buy": cari = "makebuy_cls"
    Case "Make/Buy Desc": cari = "makebuy_desc"
    Case "Control Cls": cari = "control_cls"
    Case "Control Desc": cari = "control_desc"
    Case "Order Point Qty": cari = "orderpoint_qty"
    Case "Unit Cls": cari = "unit_cls"
    Case "Unit Desc": cari = "unit_desc"
    Case "Qty/Box": cari = "number_box"
    Case "Packing Style Material Cls": cari = "packingstylematerial_cls"
    Case "Packing Style Material Desc": cari = " packingmaterial_desc"
    Case "Accounting Code": cari = "accounting_code"
    Case "Explosion Cls": cari = "explosion_cls"
    Case "Explosion Desc": cari = "explosion_desc"
    Case "Person In Charge": cari = "personincharge_cls"
    Case "Person In Charge Desc": cari = "personincharge_cls"
    Case "Stock Control Cls": cari = "stockcontrol_Cls"
    Case "Stock Control Desc": cari = "stockcontrol_desc"
    Case "Use End Day": cari = "use_endday"
    Case "Last update": cari = "last_update"
End Select
End Sub

Private Sub command2_Click()
If VSFlexGrid1.Rows > 1 And VSFlexGrid1.Row <> -1 Then
    With frm_item_master2
        Call .cmd_clear_Click
        .Show
        .txt_item_code = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, bteColProdCode)
        Call .Cmd_Submit_Click
        .k_pertama = False
        .Status = "insert"
        .txt_item_code.Enabled = True
        .txt_item_code.Text = ""
        .txt_item_name.Text = ""
        .txt_maker_item_code = ""
        .lbl_record.Caption = "Record 0 of 0"
        Me.Hide
    End With
    LblErrMsg = ""
Else
    LblErrMsg = DisplayMsg("8011")
End If
End Sub

Private Sub command3_Click()
Dim l_txt_item_code As String, tombol As String, sqlOrder As String, recAff As Double

If VSFlexGrid1.Rows < 1 Or (VSFlexGrid1.Rows = 1 Or VSFlexGrid1.Row = -1) Then
    LblErrMsg.Caption = DisplayMsg("4047"): Exit Sub
Else
    If rs_item.EOF = False Or rs_item.BOF = False Then
        l_txt_item_code = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, bteColProdCode)
        GoTo nexto
    End If
End If
Exit Sub

nexto:

If hakUpdate("frm_item_inquiry") = 0 Then _
     LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
     
Me.MousePointer = vbHourglass
If rs_bom_master.EOF = False Or rs_bom_master.BOF = False Then

'    #######################################################
'    #    Check data if already used in BOM MASTER or Not  #
'    #######################################################

    rs_bom_master.MoveFirst
        rs_bom_master.filter = "parent_itemcode ='" & l_txt_item_code & "' or item_code='" & l_txt_item_code & "'"
        If rs_bom_master.EOF = True Then
            tombol = MsgBox("Are you sure want to delete data with Product Code " & Trim(l_txt_item_code) & " ?", vbQuestion + vbYesNo, "Warning")
            sqlOrder = "select * from orderEntry_detail where item_code='" & l_txt_item_code & "'"
            Db.Execute sqlOrder, recAff
            If recAff = 0 Then
               If tombol = vbYes Then
                    Db.Execute "delete from item_master where  item_code='" & l_txt_item_code & "'"
                    Call uf_settingGrid(is_sql)
                    LblErrMsg.Caption = DisplayMsg(1201) '"Delete data success !"
                    With frm_item_master2
                        Db.Execute "update item_master set item_name=rtrim(item_name), Last_Update = getdate(), Last_User = '" & userLogin & "' where item_code='" & l_txt_item_code & "'", recAff
                        If recAff = 0 Then
                            Call .cmd_clear_Click
                            .lbl_record.Caption = "Record 0 of 0"
                            .rs_item_master.Requery
                        End If
                    End With
                End If
            Else
                LblErrMsg.Caption = DisplayMsg("0038") & " Order Entry !"
            End If
        Else
            LblErrMsg.Caption = DisplayMsg("0038") & " BOM Master !"
        End If
    Set rs_bom_master = Db.Execute("select* from bom_master")

Else

'    ########################################################
'    #    Check data if already used in ORDER ENTRY or Not  #
'    ########################################################

    sqlOrder = "select * from orderEntry_detail where item_code='" & Trim(l_txt_item_code) & "'"
    Db.Execute sqlOrder, recAff
    If recAff = 0 Then
        tombol = MsgBox("Are you sure want to delete this data ?", vbQuestion + vbYesNo, "Warning")
        If tombol = vbYes Then
            Db.Execute "delete from item_master where  item_code='" & l_txt_item_code & "'"
            Call uf_settingGrid(is_sql)
            LblErrMsg.Caption = DisplayMsg(1201)  '"Delete data success !"
            With frm_item_master2
                Db.Execute "update item_master set item_name=rtrim(item_name), Last_Update = getdate(), Last_User = '" & userLogin & "' where  item_code='" & l_txt_item_code & "'", recAff
                If recAff = 0 Then
                    Call .cmd_clear_Click
                    .lbl_record.Caption = "Record 0 of 0"
                    .rs_item_master.Requery
                End If
           End With
        End If
    Else
        LblErrMsg.Caption = DisplayMsg("0038") & " Order Entry !"
    End If
End If
Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
rs_bom_master.Open " select * from bom_master", Db, adOpenKeyset, adLockOptimistic

sqlfinish = "CASE finishgoodpart_cls " & _
                "WHEN '01' THEN 'Finish Goods' " & _
                "WHEN '02' THEN 'Part/WIP/Material' " & _
                "else '' " & _
                "END AS finish_desc "

sqlpart = "CASE part_cls " & _
                "WHEN '01' THEN '' " & _
                "WHEN '02' THEN '' " & _
                "WHEN '03' THEN '' " & _
                "WHEN '04' THEN '' " & _
                "else '' " & _
                "END AS part_desc "

sqlreserve = "CASE reserve_cls " & _
                "WHEN '01' THEN 'Yes' " & _
                "WHEN '02' THEN 'No' " & _
                "else '' " & _
                "END AS reserve_desc "
                
sqlsuply = "CASE suply_cls " & _
                "WHEN '01' THEN 'Yes' " & _
                "WHEN '02' THEN 'No' " & _
                "else '' " & _
                "END AS supply_desc "
                
sqlprovision = "CASE provision_cls " & _
                "WHEN '01' THEN 'Yes' " & _
                "WHEN '02' THEN 'No' " & _
                "else '' " & _
                "END AS provision_desc "
                
sqlexplosion = "CASE explosion_cls " & _
                "WHEN '01' THEN 'All' " & _
                "WHEN '02' THEN '1 Level' " & _
                "else '' " & _
                "END AS explosion_desc "
                
sqlmakebuy = "CASE makebuy_cls " & _
                "WHEN '01' THEN 'Make' " & _
                "WHEN '02' THEN 'Buy' " & _
                "else '' " & _
                "END AS makebuy_desc "
                
sqlstockcontrol = "CASE stockcontrol_cls " & _
                "WHEN '01' THEN 'Yes' " & _
                "WHEN '02' THEN 'No' " & _
                "WHEN '03' THEN '' " & _
                "WHEN '04' THEN '' " & _
                "else '' " & _
                "END AS stockcontrol_desc "
                
sqlprod = "CASE production_cls " & _
                "WHEN '01' THEN 'Yes' " & _
                "WHEN '02' THEN 'No' " & _
                "END AS production_desc "
                             
sqlunit = "(select description from unit_cls uc where uc.unit_cls= item_master.unit_cls )" & _
                " unit_desc "
                                                  

sqla = "select item_master.item_code , item_master.makeritem_code, item_master.Item_Name ,item_master.FinishGoodPart_Cls, " & sqlfinish & " , item_master.drawing_number,item_master.wh_code ,item_master.address ,item_master.supplier_code ,item_master.delivery_place , " & _
        "item_master.hs_code, item_master.manufacture_code ,item_master.line_code ,item_master.part_cls," & sqlpart & " , item_master.reserve_cls," & sqlreserve & " ,item_master.suply_cls," & sqlsuply & " ,item_master.provision_cls," & sqlprovision & " , " & _
        "item_master.material_cls ,material_Cls.description as material_desc,item_master.thickness ,item_master.width,item_master.Weight,item_master.GrossWeight," & _
        "item_master.length ,item_master.sheetcoil_cls ,sheetcoil_cls.description as sheet_desc,item_master.pitch ,item_master.number_producible ,item_master.scrap_weight ,item_master." & _
        "drawingmaterial_cls ,drawingmaterial_cls.description as draw_desc,item_master.surfacetreatment_cls ,surfacetreatment_cls.description as surface_desc , item_master.heattreatment_cls ,heattreatment_cls.description as heat_desc," & _
        "item_master.number_process ,item_master.material_coefficient ,item_master.process_coefficient ,item_master." & _
        "min_lot ,item_master.lot_qty,item_master.lot_coefficience ,item_master.product_readtime ,item_master.yield_percentage ,item_master.number_entering ,item_master." & _
        "packingstyle_cls , packingstyle_cls.description as packing_desc, bag_code='',item_master.group_cls ,group_cls.description as group_desc,item_master.production_cls," & sqlprod & ",item_master.standard_stock ,item_master.safety_stock ,item_master.max_stock ,item_master.min_stock ,item_master." & _
        "alowance_day ,item_master.delivery_readtime ,item_master.makebuy_cls," & sqlmakebuy & " ,item_master.control_cls ,control_Cls.description as control_desc,item_master.orderpoint_qty ,item_master.unit_cls, " & sqlunit & ",item_master." & _
        "number_box,item_master.packingstylematerial_cls , ps.description as packingmaterial_desc ,item_master.accounting_code ,item_master.explosion_cls," & sqlexplosion & " ,item_master.personincharge_cls ,personincharge_cls.description as person_desc,item_master.stockcontrol_Cls," & sqlstockcontrol & ", " & _
        " substring(item_master.use_endday,5,2) + '/' + right(item_master.use_endday,2) + '/' + left(item_master.use_endday,4)  use_endday ,item_master.last_update "
                                                       
sqlB = "From item_Master " & _
        "left join packingstyle_cls on item_master.packingstyle_cls= packingstyle_cls.packingstyle_cls " & _
        "left join packingstyle_cls ps on item_master.packingstylematerial_cls= ps.packingstyle_cls " & _
        "left join drawingmaterial_cls on item_master.drawingmaterial_cls=drawingmaterial_cls.drawingmaterial_cls " & _
        "left join personincharge_cls on item_master.personincharge_cls =personincharge_cls.personincharge_cls " & _
        "left join sheetcoil_cls on item_master.sheetcoil_cls= sheetcoil_cls.sheetcoil_cls " & _
        "left join group_cls on item_master.group_cls = group_cls.group_cls " & _
        "left join surfacetreatment_cls on item_master.surfacetreatment_cls= surfacetreatment_cls.surfacetreatment_cls " & _
        "left join heattreatment_cls on item_master.heattreatment_cls = " & _
        "heattreatment_cls.heattreatment_cls " & _
        "left join material_Cls on item_master.material_Cls=material_Cls.material_Cls " & _
        "left join control_Cls on item_master.control_Cls=control_Cls.control_Cls "

is_sql = "select * from ( " & sqla + sqlB & ") xxx"

Call cmdSearch_Click(2)
Call uf_settingCombo
txtPencarian = ""
cboCari.ListIndex = 0
LblErrMsg = ""
 
End Sub

Private Sub cmdSearch_Click(Index As Integer)

Me.MousePointer = vbHourglass
LblErrMsg = ""
Select Case Index
    Case 0: 'Search
        If txtPencarian = "" Then
            LblErrMsg = DisplayMsg(4007) '"Please input search text !"
        ElseIf cboCari = "" Then
            LblErrMsg = DisplayMsg(4008) '"Please select search category !"
        Else
            Call uf_searchTextInGrid
        End If
    
    Case 1: 'Filter

        If txtPencarian = "" Then
            LblErrMsg = DisplayMsg(5004) '"Please input search text !"
        ElseIf cboCari = "" Then
            LblErrMsg = DisplayMsg(5005) '"Please select search category !"
        Else
            If uf_validateColumnType <> "" Then LblErrMsg = uf_validateColumnType: Me.MousePointer = vbDefault: Exit Sub
            If ib_isNumeric = False Then '#### if non numeric data
                is_SqlFilter = is_sql + " where " & cari & " like '%" & Trim(txtPencarian.Text) & "%'"
            Else '#### if numeric data
                is_SqlFilter = is_sql + " where " & cari & " =" & CDbl(Trim(txtPencarian.Text)) & ""
            End If
            Call uf_settingGrid(is_SqlFilter)
        End If
        
    Case 2: 'Refresh
       Call uf_settingGrid(is_sql)
             
End Select
Me.MousePointer = vbDefault
End Sub

Private Sub Command1_Click(Index As Integer)

Dim rsCek As New ADODB.Recordset
Dim tanya

Select Case Index
    Case 0: 'Submit
        
        If VSFlexGrid1.Rows > 1 And VSFlexGrid1.Row <> -1 Then
            If hakAkses("frm_item_master2") = 0 Then LblErrMsg = DisplayMsg(3007): Exit Sub
            With frm_item_master2
                
                Call .cmd_clear_Click
                .Show
                .txt_item_code = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, bteColProdCode)
                Call .Cmd_Submit_Click
                .k_pertama = False
                Me.Hide
            End With
            LblErrMsg = ""
        Else
            LblErrMsg = DisplayMsg("8011")
        End If
        
End Select
End Sub

Private Sub cmdReport_Click()

  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset
  Dim rssubrpt As New ADODB.Recordset
  Dim Rpt As New FrmRpt3
  LblErrMsg = ""
  Me.MousePointer = vbHourglass
  
  
  sqlprint = "select tb1.* from ( " & _
             "select idx = 'Active', rtrim(im.makeritem_code) makeritem_code, " & _
             "rtrim(im.item_code) item_code, rtrim(im.Item_Name) Item_Name, " & _
             "rtrim(im.FinishGoodPart_Cls) FinishGoodPart_Cls, " & _
             "Case rtrim(im.finishgoodpart_cls) " & _
             "WHEN '01' THEN 'Finish Goods' WHEN '02' THEN 'Part/WIP/Material' else '' " & _
             "END AS Finish_Desc, " & _
             "rtrim(im.wh_code) WH_Code, rtrim(im.supplier_code) Supplier_Code, " & _
             "rtrim(im.manufacture_code) Manufacture_Code, rtrim(im.group_cls) Group_Cls, " & _
             "rtrim(gc.description) as group_desc, " & _
             "rtrim(im.material_cls) material_cls, rtrim(mc.description) material_desc, " & _
             "rtrim(im.hs_code) HS_Code, " & _
             "rtrim(im.suply_cls) Suply_Cls, " & _
             "Case rtrim(im.suply_cls) " & _
             "WHEN '01' THEN 'Yes' " & _
             "WHEN '02' THEN 'No' " & _
             "Else '' END AS supply_desc, " & _
             "rtrim(im.provision_cls) provision_cls, " & _
             "Case rtrim(im.provision_cls) " & _
             "WHEN '01' THEN 'Yes' " & _
             "WHEN '02' THEN 'No' " & _
             "Else '' END AS provision_desc, "
  sqlprint = sqlprint & _
             "rtrim(im.production_cls) production_cls, " & _
             "Case rtrim(im.production_cls) " & _
             "WHEN '01' THEN 'Yes' " & _
             "WHEN '02' THEN 'No' END AS production_desc, " & _
             "rtrim(im.stockcontrol_Cls) stockcontrol_Cls, " & _
             "Case rtrim(im.stockcontrol_cls) " & _
             "WHEN '01' THEN 'Yes' " & _
             "WHEN '02' THEN 'No' " & _
             "WHEN '03' THEN '' " & _
             "WHEN '04' THEN '' else '' END AS stockcontrol_desc, " & _
             "rtrim(im.control_cls) control_cls, rtrim(cc.description) control_desc, " & _
             "im.orderpoint_qty, " & _
             "right(im.use_endday,2) + '/' + substring(im.use_endday,5,2) + '/' + left(im.use_endday,4) as use_endday " & _
             "From item_Master im " & _
             "left join group_cls gc on im.group_cls = gc.group_cls " & _
             "left join material_Cls mc on im.material_Cls = mc.material_Cls " & _
             "left join control_Cls cc on im.control_Cls = cc.control_Cls " & _
             "Where im.Use_EndDay >= convert(varchar, getdate(), 112) "
  sqlprint = sqlprint & _
             "Union All " & _
             "select idx = 'NonActive', rtrim(im.makeritem_code) makeritem_code, " & _
             "rtrim(im.item_code) item_code, rtrim(im.Item_Name) Item_Name, " & _
             "rtrim(im.FinishGoodPart_Cls) FinishGoodPart_Cls, " & _
             "Case rtrim(im.finishgoodpart_cls) " & _
             "WHEN '01' THEN 'Finish Goods' WHEN '02' THEN 'Part/WIP/Material' else '' " & _
             "END AS Finish_Desc, " & _
             "rtrim(im.wh_code) WH_Code, rtrim(im.supplier_code) Supplier_Code, " & _
             "rtrim(im.manufacture_code) Manufacture_Code, rtrim(im.group_cls) Group_Cls, " & _
             "rtrim(gc.description) as group_desc, " & _
             "rtrim(im.material_cls) material_cls, rtrim(mc.description) material_desc, " & _
             "rtrim(im.hs_code) HS_Code, " & _
             "rtrim(im.suply_cls) Suply_Cls, " & _
             "Case rtrim(im.suply_cls) " & _
             "WHEN '01' THEN 'Yes' " & _
             "WHEN '02' THEN 'No' " & _
             "Else '' END AS supply_desc, " & _
             "rtrim(im.provision_cls) provision_cls, " & _
             "Case rtrim(im.provision_cls) " & _
             "WHEN '01' THEN 'Yes' " & _
             "WHEN '02' THEN 'No' " & _
             "Else '' END AS provision_desc, "
  sqlprint = sqlprint & _
             "rtrim(im.production_cls) production_cls, " & _
             "Case rtrim(im.production_cls) " & _
             "WHEN '01' THEN 'Yes' " & _
             "WHEN '02' THEN 'No' END AS production_desc, " & _
             "rtrim(im.stockcontrol_Cls) stockcontrol_Cls, " & _
             "Case rtrim(im.stockcontrol_cls) " & _
             "WHEN '01' THEN 'Yes' " & _
             "WHEN '02' THEN 'No' " & _
             "WHEN '03' THEN '' " & _
             "WHEN '04' THEN '' else '' END AS stockcontrol_desc, " & _
             "rtrim(im.control_cls) control_cls, rtrim(cc.description) control_desc, " & _
             "im.orderpoint_qty, " & _
             "right(im.use_endday,2) + '/' + substring(im.use_endday,5,2) + '/' + left(im.use_endday,4) as use_endday " & _
             "From item_Master im " & _
             "left join group_cls gc on im.group_cls = gc.group_cls " & _
             "left join material_Cls mc on im.material_Cls = mc.material_Cls " & _
             "left join control_Cls cc on im.control_Cls = cc.control_Cls " & _
             "Where im.Use_EndDay < convert(varchar, getdate(), 112) " & _
             ")tb1 " & _
             "order by idx, item_code "

  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic

  If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
  Set report = application.OpenReport(App.path & "\Reports\rpt_itemMaster2.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  
  sqlprint2 = "select idx = 0,  '(' + rtrim(im.material_cls) + ')' material_cls , rtrim(mc.description)  material_desc, " & _
              "count(im.material_cls) TotMaterialCls " & _
              "from Item_Master im " & _
              "left join Material_Cls mc " & _
              "on im.material_cls = mc.material_cls " & _
              "Where im.use_endday >= convert(VarChar, getdate(), 112) " & _
              "group by rtrim(im.material_cls), rtrim(mc.description) " & _
              "Union All " & _
              "select idx = 1,  material_cls = '', material_desc = 'DISCONTINUE', " & _
              "count(im.item_code) TotMaterialCls " & _
              "from item_master im " & _
              "left join Material_Cls mc " & _
              "on im.material_cls = mc.material_cls " & _
              "Where im.use_endday < convert(VarChar, getdate(), 112) " & _
              "order by idx, Material_Cls "
  
   If rssubrpt.State <> adStateClosed Then rssubrpt.Close
   rssubrpt.Open sqlprint2, Db, adOpenKeyset, adLockOptimistic
   report.OpenSubreport("SubRptItemMaster").Database.Tables(1).SetDataSource rssubrpt
    
  printorient = 2
  reportcode = "itemmaster"
   
  report.ReportTitle = "Item Master"

  Rpt.CRViewer1.ReportSource = report
  Rpt.CRViewer1.ViewReport
  Rpt.CRViewer1.Zoom 1

  Rpt.WindowState = 2
  Rpt.Show 1

  Me.MousePointer = vbDefault
  
End Sub

Private Sub CmdSubMenu_Click()
           
If cmdsubmenu.Caption = "&Back" Then
    frm_item_master2.Show
    frm_item_master2.rs_item_master.Requery
     frm_item_master2.k_pertama = False
Else
    frmMainMenu.Show
End If

cmdsubmenu.Caption = "Sub &Menu"

Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rs_item.State <> adStateClosed Then rs_item.Close
    If rs_bom_master.State <> adStateClosed Then rs_bom_master.Close
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub txtPencarian_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cbocari_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cmdSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Command1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CtrlMenu1_KeyPress(KeyAscii As Integer)
Dim stForm As Integer
    LblErrMsg = ""
    If KeyAscii = Asc("'") Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If CtrlMenu1 <> "" Then
            stForm = panggilForm(CtrlMenu1, Me.Name)
            If stForm = 0 Then
                 Unload Me
            ElseIf stForm = 1 Then
                LblErrMsg = DisplayMsg(3006)
            ElseIf stForm = 2 Then
                LblErrMsg = "This Form's Menu ID is " & CtrlMenu1
            End If
        End If
    End If
End Sub

Private Sub vsflexgrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Cancel = True
End Sub




