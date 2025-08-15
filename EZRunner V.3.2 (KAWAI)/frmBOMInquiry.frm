VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBOMInquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BOM Inquiry"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBOMInquiry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10980
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Copy"
      Height          =   375
      Left            =   3075
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9870
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Index           =   0
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9870
      Width           =   1365
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "BOM Master"
      Height          =   375
      Index           =   1
      Left            =   1612
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9870
      Width           =   1365
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   13950
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9870
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   795
      Left            =   150
      TabIndex        =   13
      Top             =   2130
      Width           =   14940
      Begin VB.ComboBox cboExplosion 
         Height          =   315
         ItemData        =   "frmBOMInquiry.frx":0E42
         Left            =   5715
         List            =   "frmBOMInquiry.frx":0E4F
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   705
      End
      Begin MSComCtl2.DTPicker dt 
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   270
         Width           =   1755
         _ExtentX        =   3096
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
         Format          =   130285571
         CurrentDate     =   37860
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2: 1 Level Explosion of BOM"
         Height          =   195
         Index           =   4
         Left            =   10890
         TabIndex        =   20
         Top             =   330
         Width           =   2400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1: Implosion of BOM "
         Height          =   195
         Index           =   3
         Left            =   8790
         TabIndex        =   19
         Top             =   330
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0: Explosion of BOM"
         Height          =   195
         Index           =   2
         Left            =   6780
         TabIndex        =   18
         Top             =   330
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Explosion     :"
         Height          =   195
         Index           =   1
         Left            =   4410
         TabIndex        =   17
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date :"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   330
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   795
      Left            =   150
      TabIndex        =   10
      Top             =   1320
      Width           =   14940
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Left            =   4215
         TabIndex        =   27
         Top             =   307
         Width           =   315
      End
      Begin VB.TextBox txtDocNo 
         Height          =   315
         Left            =   11820
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         Top             =   300
         Width           =   2955
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. Number"
         Height          =   195
         Index           =   1
         Left            =   10455
         TabIndex        =   25
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label lblNm 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   6750
         TabIndex        =   23
         Top             =   360
         Width           =   3495
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   6750
         X2              =   10230
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parent Code   :"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   360
         Width           =   1320
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   4620
         X2              =   6630
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblNm 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   4620
         TabIndex        =   11
         Top             =   360
         Width           =   2025
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   0
         Top             =   300
         Width           =   2535
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "4471;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAAAAAAAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13245
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   510
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid gridBOM 
      Height          =   5820
      Left            =   150
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3330
      Width           =   10725
      _cx             =   18918
      _cy             =   10266
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
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
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
   Begin VSFlex8Ctl.VSFlexGrid gridItem 
      Height          =   5820
      Left            =   10995
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3330
      Width           =   4095
      _cx             =   7223
      _cy             =   10266
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
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   150
      TabIndex        =   15
      Top             =   9165
      Width           =   14940
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
         Height          =   285
         Left            =   105
         TabIndex        =   16
         Top             =   195
         Width           =   14715
      End
   End
   Begin VB.Label lblJudulAtas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item Master"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   10995
      TabIndex        =   22
      Top             =   3030
      Width           =   4095
   End
   Begin VB.Label lblJudulAtas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BOM Master"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   21
      Top             =   3030
      Width           =   10725
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Inquiry"
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
      Left            =   6855
      TabIndex        =   9
      Top             =   510
      Width           =   1530
   End
End
Attribute VB_Name = "frmBOMInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public parent As String
Dim i As Long
Dim sql As String
Dim FinishGood As String, PartsCls As String, YesNo As String
Dim ExplosionCls As String, StockControl As String
Dim MakeBuyCls As String
Dim nilKosong As Boolean

Dim bteColProdCode As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColQtyR As Byte
Dim bteColQty As Byte
Dim bteColUnit As Byte
Dim bteColUnitDesc As Byte
Dim bteColDateStart As Byte
Dim bteColDateEnd As Byte
Dim bteColParent As Byte
Dim bteColRevision As Byte

Private Sub headerGridBOM()
    Dim i As Integer
    
    bteColProdCode = 0
    bteColPartNo = 1
    bteColDesc = 2
    bteColQtyR = 3
    bteColQty = 4
    bteColUnit = 5
    bteColUnitDesc = 6
    bteColDateStart = 7
    bteColDateEnd = 8
    bteColParent = 9
    bteColRevision = 10
    
    With gridBOM
        .clear
        .ColS = 11
        .Rows = 1
        
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColPartNo) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColQtyR) = "R Qty"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColUnitDesc) = "Unit"
        .TextMatrix(0, bteColDateStart) = "Start Date"
        .TextMatrix(0, bteColDateEnd) = "End Date"
        .TextMatrix(0, bteColParent) = "Parent"
        .TextMatrix(0, bteColRevision) = "Revision"
        
        .ColWidth(bteColProdCode) = 2300
        .ColWidth(bteColPartNo) = 2000
        .ColWidth(bteColDesc) = 3000
        .ColWidth(bteColQtyR) = 900
        .ColWidth(bteColQty) = 1000
        .ColWidth(bteColUnit) = 500
        .ColWidth(bteColUnitDesc) = 500
        .ColWidth(bteColDateStart) = 1300
        .ColWidth(bteColDateEnd) = 1300
        .ColWidth(bteColRevision) = 1200
        
        .ColHidden(bteColParent) = True
        .ColHidden(bteColPartNo) = True
        .ColHidden(bteColQtyR) = True
        .ColHidden(bteColUnit) = True
        
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColPartNo) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColQtyR) = flexAlignRightCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColUnit) = flexAlignCenterCenter
        .ColAlignment(bteColUnitDesc) = flexAlignLeftCenter
        .ColAlignment(bteColDateStart) = flexAlignCenterCenter
        .ColAlignment(bteColDateEnd) = flexAlignCenterCenter
        .ColAlignment(bteColRevision) = flexAlignLeftCenter
        
        .EditMaxLength = 1
        .OutlineCol = bteColProdCode
        .OutlineBar = flexOutlineBarSimpleLeaf
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
    End With
End Sub

Private Sub headerGridItem()
    Dim i As Long
    
    With gridItem
        .clear
        .ColS = 2
        .Rows = 67
        
        .TextMatrix(0, 0) = "Description"
        .TextMatrix(0, 1) = "Value"
        
        .ColWidth(0) = 2400
        .ColWidth(1) = 2000
        
        Call judulAtas("Item Info", 1)
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        
        .TextMatrix(2, 0) = "Finish Good Part Cls"
        .TextMatrix(3, 0) = "Part Number"
        
        Call judulAtas("Order & Delivery", 4)
        .TextMatrix(5, 0) = "Warehouse Code"
        .TextMatrix(6, 0) = "Address"
        .TextMatrix(7, 0) = "Supplier Code"
        .TextMatrix(8, 0) = "Delivery Place"
        
        Call judulAtas("Item Classification", 9)
        .TextMatrix(10, 0) = "Parts Cls"
        .TextMatrix(11, 0) = "Reserve Cls"
        .TextMatrix(12, 0) = "Supply Cls"
        .TextMatrix(13, 0) = "Provision Cls"
        .TextMatrix(14, 0) = "Production Cls"
        
        Call judulAtas("Stock", 15)
        .TextMatrix(16, 0) = "Qty/Case (Finish Good)"
        .TextMatrix(17, 0) = "Packing Style"
        .TextMatrix(18, 0) = "Group Cls"
        .TextMatrix(19, 0) = "Standard Stock"
        .TextMatrix(20, 0) = "Safety Stock"
        .TextMatrix(21, 0) = "Max Stock"
        .TextMatrix(22, 0) = "Min Stock"
        .TextMatrix(23, 0) = "Qty/Box (Parts/Material)"
        .TextMatrix(24, 0) = "Accounting"
        .TextMatrix(25, 0) = "Allowance Day"
        .TextMatrix(26, 0) = "Delivery Read Time"
        .TextMatrix(27, 0) = "Make or Buy Cls"
        .TextMatrix(28, 0) = "Control Cls"
        .TextMatrix(29, 0) = "Unit Cls"
        .TextMatrix(30, 0) = "Order Point Qty"
        .TextMatrix(31, 0) = "Explosion Cls"
        .TextMatrix(32, 0) = "Packing Style Part/Material"
        .TextMatrix(33, 0) = "Purchase Person"
        .TextMatrix(34, 0) = "Stock Control"
        .TextMatrix(35, 0) = "Use End Date"
        .TextMatrix(36, 0) = "Last Update"
        
        Call judulAtas("Factory Info", 37)
        .TextMatrix(38, 0) = "Factory Code"
        .TextMatrix(39, 0) = "Line Code"
        
        Call judulAtas("Material Dimension", 40)
        .TextMatrix(41, 0) = "Material Cls"
        .TextMatrix(42, 0) = "Thickness"
        .TextMatrix(43, 0) = "Width"
        .TextMatrix(44, 0) = "Weight"
        .TextMatrix(45, 0) = "Length"
        
        Call judulAtas("Material Clasification", 46)
        .TextMatrix(47, 0) = "Sheet/Coil Cls"
        .TextMatrix(48, 0) = "Pitch"
        .TextMatrix(49, 0) = "Number Preducible"
        .TextMatrix(50, 0) = "Scrap Weight"
        .TextMatrix(51, 0) = "Drawing Material Cls"
        .TextMatrix(52, 0) = "Surface Treatment Cls"
        .TextMatrix(53, 0) = "Surface Order Point Qty"
        .TextMatrix(54, 0) = "Heat Treatment Cls"
        .TextMatrix(55, 0) = "Heat Order Point Qty"
        .TextMatrix(56, 0) = "Sample"
        .TextMatrix(57, 0) = "SW Qty"
        .TextMatrix(58, 0) = "EW Qty"
        .TextMatrix(59, 0) = "Number Of Processes"
        .TextMatrix(60, 0) = "Material Coefficient"
        .TextMatrix(61, 0) = "Process Coefficient"
        .TextMatrix(62, 0) = "Min Lot"
        .TextMatrix(63, 0) = "Lot Qty"
        .TextMatrix(64, 0) = "Lot Coefficienct"
        .TextMatrix(65, 0) = "Product Read"
        .TextMatrix(66, 0) = "Yield"
        
        lblJudulAtas(1) = "Item Master"
    End With
End Sub

Sub Kosong()
nilKosong = True
    cbo(0) = ""
    lblNm(0) = ""
    dt = Now
    cboExplosion.ListIndex = 0
    Call isiCboItem
    Call headerGridBOM
    Call headerGridItem
nilKosong = False
End Sub

Sub isiCboItem()
Dim rscbo As New ADODB.Recordset

    sql = "select Item_Code,MakerItem_Code,ITem_name " & _
         "from Item_Master " & _
         "where use_endday >= convert(char(8), getdate(), 112) " & _
         "order by Item_Code"
    Set rscbo = Db.Execute(sql)
    
    cbo(0).clear
    cbo(0).columnCount = 3
    cbo(0).TextColumn = 1
    
    i = 0
    Do While Not (rscbo.EOF)
        cbo(0).AddItem ""
        cbo(0).List(i, 0) = Trim(rscbo("Item_Code"))
        cbo(0).List(i, 1) = Trim(rscbo("MakerItem_Code"))
        cbo(0).List(i, 2) = Trim(rscbo("Item_Name"))
        i = i + 1
        rscbo.MoveNext
    Loop
    cbo(0).ListWidth = 440
    cbo(0).ColumnWidths = "120 pt;120 pt;200 pt"
    cbo(0).ListRows = 15
    Set rscbo = Nothing
    '************************
End Sub

Sub IsiGrid(ibu As String, lvl As Integer, nmField As String, Optional explosion As Integer)
Dim anak As String
Dim rsAnak As New ADODB.Recordset
Dim rsTglAwal As String, rsTglAkhir As String

If explosion = 1 And lvl = 1 Then Exit Sub

With gridBOM
    sql = "Select a.*, b.Item_Name as nmITem, b.MakerItem_Code from BOM_Master a,Item_Master b where " & _
        "a." & IIf((nmField = "Parent_ItemCode"), "Item_Code", "Parent_ItemCode") & " = b.Item_Code and a." & _
        nmField & "='" & ibu & "' and " & _
        "Start_Date <='" & Format(dt, "yyyyMMdd") & "' and " & _
        "End_Date >='" & Format(dt, "yyyyMMdd") & "'"
    Set rsAnak = Db.Execute(sql)
    
    If Not rsAnak.EOF Then
        Do While Not rsAnak.EOF
            .Rows = .Rows + 1
            
            rsTglAwal = Trim(rsAnak("Start_Date"))
            rsTglAkhir = Trim(rsAnak("End_Date"))
            
            If nmField = "Parent_ItemCode" Then
                anak = Trim(rsAnak("ITem_Code"))
            Else
                anak = Trim(rsAnak("Parent_ITemCode"))
            End If
            
            .TextMatrix(.Rows - 1, bteColProdCode) = anak
            .TextMatrix(.Rows - 1, bteColPartNo) = Trim(rsAnak("MakerITem_Code"))
            If cboExplosion = "1" Then
                .TextMatrix(.Rows - 1, bteColDesc) = Trim(rsAnak("MakerItem_Code")) & " " & Trim(rsAnak("nmITem"))
            Else
                .TextMatrix(.Rows - 1, bteColDesc) = Trim(rsAnak("nmITem"))
            End If
            .TextMatrix(.Rows - 1, bteColQtyR) = Format(rsAnak("R_Qty"), gs_formatQtyBOM)
            .TextMatrix(.Rows - 1, bteColQty) = Format(rsAnak("qty"), gs_formatQtyBOM)
            .TextMatrix(.Rows - 1, bteColUnit) = Trim(rsAnak("Unit_Cls"))
            .TextMatrix(.Rows - 1, bteColUnitDesc) = uf_GetUnitDescription(rsAnak("Unit_Cls"))
            
            .TextMatrix(.Rows - 1, bteColDateStart) = Format(Left(rsTglAwal, 4) & "-" & CInt(Mid(rsTglAwal, 5, 2)) & "-" & _
                                                      Right(rsTglAwal, 2), "dd mmm yyyy")
            
            If rsTglAkhir = "99999999" Then
                .TextMatrix(.Rows - 1, bteColDateEnd) = "99/99/9999"
            Else
                .TextMatrix(.Rows - 1, bteColDateEnd) = Format(Left(rsTglAkhir, 4) & "-" & CInt(Mid(rsTglAkhir, 5, 2)) & "-" & _
                                                        Right(rsTglAkhir, 2), "dd mmm yyyy")
            End If
            
            .TextMatrix(.Rows - 1, bteColParent) = ibu
            .TextMatrix(.Rows - 1, bteColRevision) = Trim(rsAnak("Revision_No") & "")
            txtDocNo = Trim(rsAnak("Doc_No") & "")
            
            .col = 1
            .IsSubtotal(.Rows - 1) = True
            .RowData(.Rows - 1) = .Rows - 1
            .RowOutlineLevel(.Rows - 1) = lvl

            Call IsiGrid(anak, lvl + 1, nmField, explosion)
            If Not (rsAnak.EOF) Then rsAnak.MoveNext
        Loop
    End If
    Set rsAnak = Nothing
End With
End Sub

Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = cbo(0).Text
 frm_BrowseItem.Show 1
 cbo(0).Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub cmdReport_Click()
    
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim Rpt As New FrmRpt3
    Dim rstmax As New ADODB.Recordset
    Dim rsRpt As New ADODB.Recordset
    Dim MaxSeq As Long
    
    On Error GoTo errHandler
    Me.MousePointer = vbHourglass
    LblErrMsg.Caption = ""
    
    sql = "select isnull(max(seqno),0) + 1 no from tempBom "
    Set rstmax = Db.Execute(sql)
    MaxSeq = rstmax!NO
    
    If Not InsertTemp(MaxSeq) Then GoTo ErrExit
    
    sql = " select parentitem_code  as parent_itemCode, " & _
                " case level  " & _
                " when '1' then " & _
                "   rtrim(Childitem_code) " & _
                " when '2' then " & _
                "   '-- '+ rtrim(Childitem_code) " & _
                " when '3' then " & _
                "   '    -- ' + rtrim(Childitem_code) " & _
                " when '4' then " & _
                "   '        -- '+ rtrim(Childitem_code) "
    
    sql = sql + " when '5' then " & _
                "   '            -- ' + rtrim(Childitem_code) " & _
                " when '6' then " & _
                "   '                -- ' + rtrim(Childitem_code) " & _
                " when '7' then " & _
                "   '                    -- '+ rtrim(Childitem_code) " & _
                " end " & _
                " item_Code, bomseq_no, Rqty as R_Qty, Qty, a.Unit_Cls Unit, Start_date, End_date, Revision_No, Doc_No, " & _
                " rtrim(b.Makeritem_Code) MakerParent, rtrim(b.item_name) Description,  " & _
                " rtrim(c.Makeritem_code) Makeritem, rtrim(c.item_name) Item_name, " & _
                " substring(start_Date,5,2) as blnAwal,   "
    
    sql = sql + " substring(end_Date,5,2) as blnEnd, (select description from unit_cls uc where uc.unit_cls=a.unit_cls) unit_desc " & _
                " from tempBom a, item_master b , item_master c " & _
                " where a.parentitem_code = b.item_code " & _
                " and a.Childitem_code = c.item_code and seqno = '" & MaxSeq & "' " & _
                " and start_date <= '" & Format(dt.Value, "yyyyMMdd") & "' " & _
                " and end_date >= '" & Format(dt.Value, "yyyyMMdd") & "' " & _
                " order by bomseq_no"
    
    Set rsRpt = Db.Execute(sql)
    If rsRpt.EOF Then
        LblErrMsg.Caption = DisplayMsg(4002)
        cbo(0).SetFocus
    Else
        sqlprint = sql
        reportcode = "BOM"
        printorient = 2
        
        If cboExplosion.Text = "1" Then
            Set report = application.OpenReport(App.path & "\Reports\BOM_implosion.rpt")
        Else
            Set report = application.OpenReport(App.path & "\Reports\BOM.rpt")
        End If
        report.Database.Tables(1).SetDataSource rsRpt
        
        '#####################################################################
        '# Qty Digit and decimal
        report.FormulaFields(6).Text = "" & gi_decimalDigitQtyBOM & ""
        report.FormulaFields(7).Text = "" & gi_decimalDigitQtyBOM & ""
        Select Case cboExplosion.Text
        Case 0: report.FormulaFields(8).Text = "'Explosion Of BOM'"
        Case 1: report.FormulaFields(8).Text = "'Implosion Of BOM'"
        Case 2: report.FormulaFields(8).Text = "'1 Level Explosion Of BOM'"
        End Select
        '#####################################################################

        Rpt.CRViewer1.ReportSource = report
        Rpt.CRViewer1.ViewReport
        Rpt.CRViewer1.Zoom 1
        
        Rpt.WindowState = 2
        Rpt.Show 1
    End If
    Set rsRpt = Nothing
    
ErrExit:
    sql = "delete from tempBom where seqno = '" & MaxSeq & "'"
    Db.Execute sql
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
    
End Sub

Private Sub command2_Click()
On Error GoTo errHandler
 Dim DbChild As New ADODB.Connection
 Dim rsYourParent As New ADODB.Recordset
 LblErrMsg.Caption = ""
 Me.MousePointer = vbHourglass
   
 If gridBOM.Rows <= 1 Or parent = "" Or cboExplosion.Text = "1" Then  'Jika tdk ada childnya
  GoTo ErrExit
 ElseIf LCase(Trim(cbo(0).Column(0))) = LCase(Trim(parent)) Then 'Jika sama dgn Parentnya
  LblErrMsg = DisplayMsg(1014)
  cbo(0).SetFocus
  GoTo ErrExit
 Else
  '*********Parent Validation**************
  rsYourParent.Open "cek_parent ('" & Trim(cbo(0).Column(0)) & "', '" & Trim(parent) & "')", Db, adOpenDynamic, adLockReadOnly, adCmdStoredProc
  If Not (rsYourParent.EOF) Then
   LblErrMsg.Caption = DisplayMsg(1020)
   GoTo ErrExit
  Else
   DbChild.Open Db.ConnectionString
   sql = "insert into BOM_Master (Parent_ItemCode,Description,Item_Code,Item_name, " & _
         Chr(10) & "Qty,R_Qty, Unit_Cls, Currency_Code, Accounting_Code, Start_Date, End_Date, " & _
         Chr(10) & "Last_Update,Last_User, Doc_No, Revision_No) " & _
         Chr(10) & "select parent_itemcode = '" & Trim(parent) & "', description = (select im.item_name from item_master im where item_code = '" & Trim(parent) & "'), " & _
         Chr(10) & "bm.item_code, bm.item_name, bm.qty, bm.r_qty, bm.unit_cls, bm.currency_code, " & _
         Chr(10) & "bm.accounting_code, bm.start_date, bm.end_date, last_update = getdate(), last_user = '" & userLogin & "', " & _
         Chr(10) & "doc_no,bm.revision_no from bom_master bm " & _
         Chr(10) & "where bm.parent_itemcode = '" & cbo(0).Column(0) & "' "

   DbChild.BeginTrans
   Db.Execute (sql)
   DbChild.CommitTrans
   FrmBOMMaster.Show
   FrmBOMMaster.cbo(0) = parent
   parent = ""
   Unload Me
 End If
End If
ErrExit:
    Set rsYourParent = Nothing
    Set DbChild = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

errHandler:
    DbChild.RollbackTrans
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub

Private Sub dt_change()
    If nilKosong = False Then Call cbo_Click(0)
End Sub

Private Sub gridBOM_Click()
With gridBOM
    If .Rows <> 1 And .row <> -1 Then Call isiGridItem(.TextMatrix(.row, bteColProdCode))
End With
End Sub

Sub isiArr()
    FinishGood = "Finish Goods,Parts/WIP/Material"
    PartsCls = " , , , "
    YesNo = "Yes,No"
    ExplosionCls = "All,1 Level"
    StockControl = "Yes,No, , "
    MakeBuyCls = "Make,Buy"
End Sub

Sub judulAtas(jdl As String, Baris As Integer)
With gridItem
    .TextMatrix(Baris, 0) = jdl
    .TextMatrix(Baris, 1) = jdl
    .MergeRow(Baris) = True
    .MergeCells = flexMergeRestrictRows
    .Cell(flexcpFontBold, Baris, 0) = True
    .Cell(flexcpForeColor, Baris, 0) = &H800000
    .Cell(flexcpBackColor, Baris, 0) = vbWhite
End With
End Sub



Sub isiBaris(kd As Variant, Baris As Integer, Optional nm As String)
With gridItem
    If IsNull(kd) Or Trim(kd) = "" Then
        .TextMatrix(Baris, 1) = ""
    Else
        If nm = "" Then
            .TextMatrix(Baris, 1) = Trim(kd)
        Else
            .TextMatrix(Baris, 1) = Trim(kd) & " (" & Trim(nm) & ")"
        End If
    End If
End With
End Sub

Sub isiBarisArr(kd As Variant, Baris As Integer, nmArr)
Dim nm As String

With gridItem
    If IsNull(kd) Or Trim(kd) = "" Then
        .TextMatrix(Baris, 1) = ""
    Else
        nm = Split(nmArr, ",")(CDbl(kd) - 1)
        .TextMatrix(Baris, 1) = Trim(kd) & " (" & Trim(nm) & ")"
    End If
End With
End Sub


Sub isiGridItem(itemCD As String)
Dim tglEnd As String, tglFormat As String
Dim RsItem As New ADODB.Recordset

Call headerGridItem
    
With gridItem
    '******** Untuk Item Master *********
    sql = "select a.*,isnull(b.WH_Name,'') as nmWarehouse,isnull(c.Trade_Name,'') as nmSupplier," & _
        "isnull(d.Location_Name,'') as nmLocation,isnull(e.Trade_Name,'') as nmManufacture," & _
        "isnull(f.Line_Name,'') as nmLine,isnull(g.description,'') as nmMaterial," & _
        "isnull(h.Description,'') as nmSheet, isnull(i.Description,'') as nmDrawingMaterial," & _
        "isnull(j.Description,'') as nmSurface, isnull(k.Description,'') as nmHeat," & _
        "isnull(l.Description,'') as nmPacking, isnull(m.Description,'') as nmGroup," & _
        "isnull(n.Description,'') as nmControl, isnull(o.Description,'') as nmPerson " & _
        "from Item_Master a " & _
        "left outer Join Warehouse_Master b ON a.WH_Code = b.WH_Code " & _
        "Left Outer Join Trade_Master c ON a.Supplier_Code = c.Trade_Code " & _
        "Left Outer Join Delivery_Place d ON a.Supplier_Code = d.Trade_Code and " & _
            "a.Delivery_Place = D.Location_Code " & _
        "Left Outer Join Trade_Master e ON a.Manufacture_Code = e.Trade_Code " & _
        "Left Outer Join Manufacture_Line f ON a.Manufacture_Code = f.Manufacture_Code and " & _
            "A.Line_Code = F.Line_Code " & _
        "Left Outer Join Material_Cls g On a.Material_Cls = g.Material_Cls " & _
        "Left Outer Join SheetCoil_Cls h On a.SheetCoil_Cls = h.SheetCoil_Cls " & _
        "Left Outer Join DrawingMaterial_Cls i On a.DrawingMaterial_Cls = i.DrawingMaterial_Cls " & _
        "Left Outer Join SurfaceTreatment_Cls j On a.SurfaceTreatment_Cls= j.SurfaceTreatment_Cls " & _
        "Left Outer Join HeatTreatment_Cls k On a.HeatTreatment_Cls= k.HeatTreatment_Cls " & _
        "Left Outer Join PackingStyle_Cls l On a.PackingStyle_Cls= l.PackingStyle_Cls " & _
        "Left Outer Join Group_Cls m On a.Group_Cls = m.Group_Cls " & _
        "Left Outer Join Control_Cls n On a.Control_Cls = n.Control_Cls " & _
        "Left Outer Join PersonInCharge_Cls o On a.PersonInCharge_Cls = o.PersonInCharge_Cls " & _
        "where Item_Code ='" & itemCD & "'"
    Set RsItem = Db.Execute(sql)
    
    If Not RsItem.EOF Then
        
        lblJudulAtas(1) = "Product Code : " & Trim(RsItem("Item_Code")) & " (" & Trim(RsItem("Item_Name")) & ")"
        
        Call isiBarisArr(RsItem("FinishGoodPart_Cls"), 2, FinishGood)
        Call isiBaris(RsItem("MakerItem_Code"), 3)
        
        Call isiBaris(RsItem("WH_Code"), 5, RsItem("nmWarehouse"))
        Call isiBaris(RsItem("Address"), 6)
        Call isiBaris(RsItem("Supplier_Code"), 7, RsItem("nmSupplier"))
        Call isiBaris(RsItem("Delivery_Place"), 8, RsItem("nmLocation"))
        
        Call isiBarisArr(RsItem("Part_Cls"), 10, PartsCls)
        Call isiBarisArr(RsItem("Reserve_Cls"), 11, YesNo)
        Call isiBarisArr(RsItem("Suply_Cls"), 12, YesNo)
        Call isiBarisArr(RsItem("Provision_Cls"), 13, YesNo)
        Call isiBarisArr(RsItem("Production_Cls"), 14, YesNo)

        gridItem.TextMatrix(16, 1) = Format(RsItem!number_entering, gs_formatQtyBOM)
        
        Call isiBaris(RsItem("PackingStyle_Cls"), 17, RsItem("nmPacking"))
        Call isiBaris(RsItem("Group_Cls"), 18, RsItem("nmGroup"))
        
        gridItem.TextMatrix(19, 1) = Format(RsItem!Standard_Stock, gs_formatQtyBOM)
        gridItem.TextMatrix(20, 1) = Format(RsItem!Safety_Stock, gs_formatQtyBOM)
        gridItem.TextMatrix(21, 1) = Format(RsItem!Max_Stock, gs_formatQtyBOM)
        gridItem.TextMatrix(22, 1) = Format(RsItem!Min_Stock, gs_formatQtyBOM)
        gridItem.TextMatrix(23, 1) = Format(RsItem!Number_Box, gs_formatBox)

        Call isiBaris(RsItem("Accounting_Code"), 24)
        
        gridItem.TextMatrix(25, 1) = Format(RsItem!Alowance_Day, gs_formatDay)
        gridItem.TextMatrix(26, 1) = Format(RsItem!Delivery_ReadTime, gs_formatDay)
        
        Call isiBarisArr(RsItem("MakeBuy_Cls"), 27, MakeBuyCls)
        Call isiBaris(RsItem("Control_Cls"), 28, RsItem("nmControl"))
        
        gridItem.TextMatrix(29, 1) = (RsItem!Unit_cls) & " (" & uf_GetUnitDescription(RsItem!Unit_cls) & ")"
        gridItem.TextMatrix(30, 1) = Format(RsItem!OrderPoint_Qty, gs_formatQtyBOM)
        
        Call isiBarisArr(RsItem("Explosion_Cls"), 31, ExplosionCls)
        Call isiBaris(RsItem("PackingStyleMaterial_Cls"), 32, RsItem("nmPacking"))
        Call isiBaris(RsItem("PersonInCharge_Cls"), 33, RsItem("nmPerson"))
        Call isiBarisArr(RsItem("StockControl_Cls"), 34, StockControl)
        
        tglEnd = RsItem("Use_EndDay")
        If Left(tglEnd, 2) = "99" Then
            tglFormat = "99/99/9999"
        Else
            tglFormat = Format(Left(tglEnd, 4) & "-" & CInt(Mid(tglEnd, 5, 2)) & "-" & _
                        Right(tglEnd, 2), "dd mmm yyyy")
        End If
        Call isiBaris(tglFormat, 35)
        Call isiBaris(Format(RsItem("Last_Update"), "dd MMM yyyy HH:MM"), 36)
        Call isiBaris(RsItem("Manufacture_Code"), 38, RsItem("nmManufacture"))
        Call isiBaris(RsItem("Line_Code"), 39, RsItem("nmLine"))
        Call isiBaris(RsItem("Material_Cls"), 41, RsItem("nmMaterial"))
        
        gridItem.TextMatrix(42, 1) = Format(RsItem!Thickness, gs_formatThickness)
        gridItem.TextMatrix(43, 1) = Format(RsItem!Width, gs_formatWidth)
        gridItem.TextMatrix(44, 1) = Format(RsItem!Weight, gs_formatWeight)
        gridItem.TextMatrix(45, 1) = Format(RsItem!Length, gs_formatLength)
        
        Call isiBaris(RsItem("SheetCoil_Cls"), 47, RsItem("nmSheet"))
        
        gridItem.TextMatrix(48, 1) = Format(RsItem!Width, gs_formatWidth)
        gridItem.TextMatrix(49, 1) = Format(RsItem!Weight, gs_formatWeight)
        gridItem.TextMatrix(50, 1) = Format(RsItem!Length, gs_formatLength)
        
        Call isiBaris(RsItem("DrawingMaterial_Cls"), 51, RsItem("nmDrawingMaterial"))
        Call isiBaris(RsItem("SurfaceTreatment_Cls"), 52, RsItem("nmSurface"))
        
        gridItem.TextMatrix(53, 1) = Format(RsItem!Surface_OrderPointQty, gs_formatQtyBOM)
        
        Call isiBaris(RsItem("HeatTreatment_Cls"), 54, RsItem("nmHeat"))
        
        gridItem.TextMatrix(55, 1) = Format(RsItem!Heat_OrderPointQty, gs_formatQtyBOM)
        gridItem.TextMatrix(56, 1) = Format(RsItem!Sample, gs_formatNSample)
        gridItem.TextMatrix(57, 1) = Format(RsItem!SW_Qty, gs_formatSW)
        gridItem.TextMatrix(58, 1) = Format(RsItem!EW_Qty, gs_formatEW)
        gridItem.TextMatrix(59, 1) = RsItem!Number_Process
        gridItem.TextMatrix(60, 1) = Format(RsItem!Material_Coefficient, gs_formatCoefficient)
        gridItem.TextMatrix(61, 1) = Format(RsItem!Process_Coefficient, gs_formatCoefficient)
        gridItem.TextMatrix(62, 1) = Format(RsItem!Min_Lot, gs_formatLot)
        gridItem.TextMatrix(63, 1) = Format(RsItem!Lot_Qty, gs_formatLot)
        gridItem.TextMatrix(64, 1) = Format(RsItem!Lot_Coefficience, gs_formatCoefficient)
        gridItem.TextMatrix(65, 1) = Format(RsItem!Product_ReadTime, gs_formatDay)
        gridItem.TextMatrix(66, 1) = Format(RsItem!Yield_Percentage, gs_formatPercentage)
    
    End If
    Set RsItem = Nothing
End With
End Sub

Private Sub cbo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbo_Click(Index)
End Sub

Private Sub cboExplosion_Click()
    Call cbo_Click(0)
End Sub

Private Sub cbo_Click(Index As Integer)
Dim rsIbu As New ADODB.Recordset

If nilKosong Then Exit Sub
If cbo(0) <> "" Then
    cbo(0) = cbo(0)
    If cbo(0).MatchFound = False Then
        lblNm(0) = ""
        lblNm(1) = ""
        LblErrMsg = DisplayMsg(4002)
        Call headerGridBOM
    Else
        lblNm(0) = cbo(0).Column(1)
        lblNm(1) = cbo(0).Column(2)
        LblErrMsg = ""
        Call headerGridBOM
        Call headerGridItem
        If cboExplosion = 0 Then 'cari anak
            Call IsiGrid(cbo(0), 0, "Parent_ItemCode")
            
        ElseIf cboExplosion = 1 Then 'cari parent

            Call IsiGrid(cbo(0), 0, "Item_Code")
            
            
        ElseIf cboExplosion = 2 Then 'cari anak tp 1 level
            Call IsiGrid(cbo(0), 0, "Parent_ItemCode", 1)
        End If
    End If
Else
    lblNm(0) = ""
    lblNm(1) = ""
    LblErrMsg = ""
End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    Call Kosong
    Call isiArr
End Sub

'************ Unload **********
Private Sub CmdSubMenu_Click(Index As Integer)
If Index = 0 Then 'Sub Menu
    If cmdsubmenu(0).Caption <> "&Back" Then
        DoEvents
        frmMainMenu.Show
        DoEvents
        Unload Me
    Else
        Call CmdSubMenu_Click(1)
    End If

Else 'BOM Master
    If hakAkses("FrmBOMMaster") = 0 Then LblErrMsg = DisplayMsg(3007):  Exit Sub
    With gridBOM
        If .row > 0 And cboExplosion <> 1 Then kirimPar = .TextMatrix(.row, bteColProdCode) & .TextMatrix(.row, bteColDateStart) Else kirimPar = ""
        
        FrmBOMMaster.Show
        
        If .row > 0 Then
            If cboExplosion = 1 Then
                FrmBOMMaster.cbo(0) = .TextMatrix(.row, bteColProdCode)
            Else
                FrmBOMMaster.cbo(0) = .TextMatrix(.row, bteColParent)
            End If
        Else
            FrmBOMMaster.cbo(0) = cbo(0)
        End If
        Unload Me
    End With

End If
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

Private Function InsertTemp(seqNo As Long) As Boolean
    
    Dim cmd As New Command, prm1 As New ADODB.Parameter, prm2 As New ADODB.Parameter, prm3 As New ADODB.Parameter
    
    On Error GoTo errHandler
    
    With cmd
        
        If Trim(cbo(0).Text) = "" Then
            
            prm1.type = adNumeric
            prm1.Precision = 4
            prm1.Value = seqNo
            
            prm2.type = adVarChar
            prm2.Size = 1
            prm2.Value = cboExplosion.Text
            
            .ActiveConnection = Db
            .CommandText = "AllBom"
            .CommandType = adCmdStoredProc
            .Parameters.append prm1
            .Parameters.append prm2
            .Execute
        
        Else
            
            prm1.type = adVarChar
            prm1.Size = 25
            prm1.Value = Trim(cbo(0))
            
            prm2.type = adNumeric
            prm2.Precision = 4
            prm2.Value = seqNo
        
            prm3.type = adVarChar
            prm3.Size = 1
            prm3.Value = cboExplosion.Text
            
            .ActiveConnection = Db
            .CommandType = adCmdStoredProc
            
            If cboExplosion.Text = "1" Then .CommandText = "ChildBOM" Else .CommandText = "ParentBOM"
            
            .Parameters.append prm1
            .Parameters.append prm2
            .Parameters.append prm3
            .Execute
        
        End If
    
    End With
    
    InsertTemp = True

ErrExit:
    Set cmd = Nothing
    Set prm1 = Nothing
    Set prm2 = Nothing
    Set prm3 = Nothing
    Exit Function
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
    
End Function
