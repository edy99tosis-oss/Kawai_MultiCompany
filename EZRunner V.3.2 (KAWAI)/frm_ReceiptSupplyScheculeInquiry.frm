VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ReceiptSupplyScheculeInquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Receipt Supply Schedule Inquiry"
   ClientHeight    =   11010
   ClientLeft      =   45
   ClientTop       =   390
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
   Icon            =   "frm_ReceiptSupplyScheculeInquiry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Tag             =   " "
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   300
      Left            =   4725
      TabIndex        =   19
      Top             =   1350
      Width           =   315
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   420
      TabIndex        =   17
      Top             =   9210
      Width           =   14445
      Begin VB.Label Lblpesan 
         Alignment       =   2  'Center
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
         TabIndex        =   18
         Top             =   195
         Width           =   14220
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   12960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   420
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   767
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Search"
      Height          =   375
      Index           =   9
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1035
   End
   Begin VB.CommandButton cmd_preview 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   13740
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9990
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Sub &Menu"
      Height          =   375
      Index           =   8
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9990
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   1830
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
      Format          =   152633347
      CurrentDate     =   37798
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6795
      Left            =   420
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2325
      Width           =   14445
      _cx             =   25479
      _cy             =   11986
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
      GridColor       =   12632256
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   1
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
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Index           =   1
      Left            =   3540
      TabIndex        =   2
      Top             =   1830
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
      Format          =   152633347
      CurrentDate     =   37798
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      Height          =   195
      Left            =   6300
      TabIndex        =   16
      Top             =   1890
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Line Line2 
      X1              =   12000
      X2              =   12630
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Label Lbl_UnitDesc 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   12000
      TabIndex        =   15
      Top             =   1395
      Width           =   645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      Height          =   195
      Left            =   11415
      TabIndex        =   14
      Top             =   1395
      Width           =   330
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   195
      Left            =   3255
      TabIndex        =   13
      Top             =   1890
      Width           =   195
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Supply Schedule Inquiry"
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
      Left            =   420
      TabIndex        =   12
      Top             =   450
      Width           =   14445
   End
   Begin VB.Line Line1 
      X1              =   6300
      X2              =   11220
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Label LblDesc 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   6300
      TabIndex        =   11
      Top             =   1365
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   0
      Left            =   5160
      TabIndex        =   10
      Top             =   1380
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Period"
      Height          =   195
      Left            =   420
      TabIndex        =   9
      Top             =   1890
      Width           =   540
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
      Height          =   195
      Left            =   420
      TabIndex        =   8
      Top             =   1410
      Width           =   915
   End
   Begin MSForms.ComboBox CboItemCD 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   1350
      Width           =   2985
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "5265;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frm_ReceiptSupplyScheculeInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bteColDate As Byte

Dim bteColOrderQty As Byte
Dim bteColOrderIn As Byte
Dim bteColOrderPONo As Byte
Dim bteColOrderDONo As Byte
Dim bteColOrderCurr As Byte
Dim bteColOrderPrice As Byte
Dim bteColOrderBal As Byte

Dim bteColInCls As Byte
Dim bteColInQty As Byte
Dim bteColInCurr As Byte
Dim bteColInPrice As Byte
Dim bteColInAmount As Byte

Dim bteColOutCls As Byte
Dim bteColOutQty As Byte
Dim bteColOutReq As Byte
Dim bteColOutCurr As Byte
Dim bteColOutPrice As Byte
Dim bteColOutAmount As Byte

Dim bteColBalQty As Byte
Dim bteColBalCurr As Byte
Dim bteColBalPrice As Byte
Dim bteColBalAmount As Byte
Dim bteColBalEffective As Byte

Dim bteHakPrice As Byte
Dim dteMRP As Date

Private Sub AddToComboItem()
    Dim RsItem As New ADODB.Recordset
    
    i = 0
    CboItemCD.columnCount = 3
    CboItemCD.clear
    
    sql = "select item_code, item_name, uc.description unit_desc from item_master im left join unit_cls uc on im.unit_cls = uc.unit_cls where im.use_endday > convert(char(8), getdate(), 112)"
    RsItem.Open sql, Db, adOpenForwardOnly, adLockReadOnly
    Do While Not RsItem.EOF
        CboItemCD.AddItem ""
        CboItemCD.List(i, 0) = Trim(RsItem!Item_Code)
        CboItemCD.List(i, 1) = Trim(RsItem!item_name)
        CboItemCD.List(i, 2) = Trim(RsItem!Unit_Desc)
        i = i + 1
        RsItem.MoveNext
    Loop
    RsItem.Close
    
    CboItemCD.ColumnWidths = "100 pt; 250 pt;0 pt"
    CboItemCD.ListWidth = 350
    CboItemCD.ListRows = 15
End Sub

Private Sub Header()

    bteColDate = 0
    bteColOrderQty = 1
    bteColOrderIn = 2
    bteColOrderPONo = 3
    bteColOrderDONo = 4
    bteColOrderCurr = 5
    bteColOrderPrice = 6
    bteColOrderBal = 7
    
    bteColInCls = 8
    bteColInQty = 9
    bteColInCurr = 10
    bteColInPrice = 11
    bteColInAmount = 12
    
    bteColOutCls = 13
    bteColOutQty = 14
    bteColOutReq = 15
    bteColOutCurr = 16
    bteColOutPrice = 17
    bteColOutAmount = 18
    
    bteColBalQty = 19
    bteColBalCurr = 20
    bteColBalPrice = 21
    bteColBalAmount = 22
    bteColBalEffective = 23
        
    With grid
        .Rows = 2
        .ColS = 24
        
        .FixedCols = 0
        .FixedRows = 2
        .FrozenCols = 1
        
        '#Column Title
        .TextMatrix(0, bteColDate) = "Date"
        
        .TextMatrix(0, bteColOrderQty) = "Order / Production"
        .TextMatrix(0, bteColOrderIn) = "Order / Production"
        .TextMatrix(0, bteColOrderPONo) = "Order / Production"
        .TextMatrix(0, bteColOrderDONo) = "Order / Production"
        .TextMatrix(0, bteColOrderCurr) = "Order / Production"
        .TextMatrix(0, bteColOrderPrice) = "Order / Production"
        .TextMatrix(0, bteColOrderBal) = "Order / Production"
        
        .TextMatrix(0, bteColInCls) = "Incoming"
        .TextMatrix(0, bteColInQty) = "Incoming"
        .TextMatrix(0, bteColInCurr) = "Incoming"
        .TextMatrix(0, bteColInPrice) = "Incoming"
        .TextMatrix(0, bteColInAmount) = "Incoming"
        
        .TextMatrix(0, bteColOutCls) = "Outgoing"
        .TextMatrix(0, bteColOutQty) = "Outgoing"
        .TextMatrix(0, bteColOutReq) = "Outgoing"
        .TextMatrix(0, bteColOutCurr) = "Outgoing"
        .TextMatrix(0, bteColOutPrice) = "Outgoing"
        .TextMatrix(0, bteColOutAmount) = "Outgoing"
        
        .TextMatrix(0, bteColBalQty) = "Balance"
        .TextMatrix(0, bteColBalCurr) = "Balance"
        .TextMatrix(0, bteColBalPrice) = "Balance"
        .TextMatrix(0, bteColBalAmount) = "Balance"
        .TextMatrix(0, bteColBalEffective) = "Effective Balance"
        
        .TextMatrix(1, bteColDate) = "Date"
        
        .TextMatrix(1, bteColOrderQty) = "Quantity"
        .TextMatrix(1, bteColOrderIn) = "Incoming"
        .TextMatrix(1, bteColOrderPONo) = "PO No. / Lot No."
        .TextMatrix(1, bteColOrderDONo) = "DO No. / Lot No."
        .TextMatrix(1, bteColOrderCurr) = "Curr"
        .TextMatrix(1, bteColOrderPrice) = "Unit Price"
        .TextMatrix(1, bteColOrderBal) = "Balance"
        
        .TextMatrix(1, bteColInCls) = "Cls"
        .TextMatrix(1, bteColInQty) = "Quantity"
        .TextMatrix(1, bteColInCurr) = "Curr"
        .TextMatrix(1, bteColInPrice) = "Unit Price"
        .TextMatrix(1, bteColInAmount) = "Amount"
        
        .TextMatrix(1, bteColOutCls) = "Cls"
        .TextMatrix(1, bteColOutQty) = "Quantity"
        .TextMatrix(1, bteColOutReq) = "Requirement"
        .TextMatrix(1, bteColOutCurr) = "Curr"
        .TextMatrix(1, bteColOutPrice) = "Unit Price"
        .TextMatrix(1, bteColOutAmount) = "Amount"
        
        .TextMatrix(1, bteColBalQty) = "Quantity"
        .TextMatrix(1, bteColBalCurr) = "Curr"
        .TextMatrix(1, bteColBalPrice) = "Unit Price"
        .TextMatrix(1, bteColBalAmount) = "Amount"
        .TextMatrix(1, bteColBalEffective) = "Effective Balance"
        
        '#Column Width
        .ColWidth(bteColDate) = 1300 'Date
        
        .ColWidth(bteColOrderQty) = 1250 'Quantity
        .ColWidth(bteColOrderIn) = 1250 'Quantity
        .ColWidth(bteColOrderPONo) = 2600 'PO No.
        .ColWidth(bteColOrderDONo) = 2600 'Do No.
        .ColWidth(bteColOrderCurr) = 500 'Curr
        .ColWidth(bteColOrderPrice) = 1750 'Price.
        .ColWidth(bteColOrderBal) = 1250 'Quantity
        
        .ColWidth(bteColInCls) = 450 'In Cls
        .ColWidth(bteColInQty) = 1250 'Quantity
        .ColWidth(bteColInCurr) = 500 'Curr
        .ColWidth(bteColInPrice) = 1750 'Price.
        .ColWidth(bteColInAmount) = 1750 'Amount.
        
        .ColWidth(bteColOutCls) = 450 'Out Cls
        .ColWidth(bteColOutQty) = 1250 'Quantity
        .ColWidth(bteColOutReq) = 1400 'Requirement
        .ColWidth(bteColOutCurr) = 500 'Curr
        .ColWidth(bteColOutPrice) = 1750 'Price.
        .ColWidth(bteColOutAmount) = 1750 'Amount.
        
        .ColWidth(bteColBalQty) = 1250 'Quantity
        .ColWidth(bteColBalCurr) = 500 'Curr
        .ColWidth(bteColBalPrice) = 1750 'Price.
        .ColWidth(bteColBalAmount) = 1750 'Amount.
        .ColWidth(bteColBalEffective) = 2000 'Effective Balance
        
        'Column Alignment
        .ColAlignment(bteColOrderPONo) = flexAlignLeftCenter
        .ColAlignment(bteColOrderDONo) = flexAlignLeftCenter
        .ColAlignment(bteColOrderCurr) = flexAlignCenterCenter
        .ColAlignment(bteColInCls) = flexAlignCenterCenter
        .ColAlignment(bteColInCurr) = flexAlignCenterCenter
        .ColAlignment(bteColOutCls) = flexAlignCenterCenter
        .ColAlignment(bteColOutCurr) = flexAlignCenterCenter
        .ColAlignment(bteColBalCurr) = flexAlignCenterCenter
        
        '#Merge Column
        .MergeRow(0) = True
        .MergeCol(bteColDate) = True
        .MergeCol(bteColBalEffective) = True
        .MergeCells = flexMergeFixedOnly
        .Editable = flexEDNone
                
        '#Colum Title Alignment
        .Cell(flexcpAlignment, 0, 0, 1, .ColS - 1) = flexAlignCenterCenter
    End With
    
End Sub

Private Sub SettingGrid()
    Dim adoRs As New ADODB.Recordset
    
    Dim dblBalance As Double
    Dim dblBalanceOrder As Double
    Dim dblBalanceEffective As Double
    
    On Error GoTo errHandler
    Me.MousePointer = vbHourglass
    LblPesan.Caption = ""
    
    sql = " declare @item_code char(15) " & vbCrLf & _
                " declare @start_date datetime " & vbCrLf & _
                " declare @end_date datetime " & vbCrLf & _
                " declare @closing_date as datetime " & vbCrLf & _
                "  " & vbCrLf & _
                " set @item_code  = '" & Trim(CboItemCD.Text) & "' " & vbCrLf & _
                " set @start_date = '" & Format(DMonth(0).Value, "yyyy-MM-dd") & "' " & vbCrLf & _
                " set @end_date = '" & Format(DMonth(1).Value, "yyyy-MM-dd") & "' " & vbCrLf & _
                " set @closing_date = (select top 1 cast(inventory_year as nvarchar(4)) + '-' + cast(inventory_month as nvarchar(2)) +  '-1' from inventory_control order by inventory_year desc, inventory_month desc) " & vbCrLf & _
                "  " & vbCrLf
                    
    sql = sql + " select isnull(stock_begin.stock_begin, 0) stock_begin, isnull(stock_in.stock_in, 0) stock_in, isnull(stock_out.stock_out, 0) stock_out,  " & vbCrLf & _
                " isnull(po.po_qty, 0) po_qty, isnull(po_receipt.po_receipt, 0) po_receipt, isnull(po_return.po_return, 0) po_return, isnull(po_cancel.po_cancel, 0) po_cancel,  " & vbCrLf & _
                " isnull(pro.pro_qty, 0) pro_qty, isnull(pro_result.pro_result, 0) pro_result, isnull(pro_cancel.pro_cancel, 0) pro_cancel,  " & vbCrLf & _
                " isnull(req.req_qty, 0) req_qty, isnull(req.req_result, 0) req_result, isnull(req.req_off, 0) req_off, trans.*, isnull(mrp_set, '0') mrp_set,  " & vbCrLf & _
                " rtrim(cp.company_name) company_name, rtrim(cp.address1) address1, rtrim(cp.address2) address2, rtrim(cp.phone1) phone1, rtrim(cp.phone2) phone2, rtrim(cp.fax) fax  " & vbCrLf & _
                " from(  " & vbCrLf
                
    sql = sql + "   select item_code, sum(isnull(bg.premonth, 0)) stock_begin  " & vbCrLf & _
                "   from(  " & vbCrLf & _
                "       select stock_year, stock_month, warehouse_code, item_code, premonth from stock_history  " & vbCrLf & _
                "       union  " & vbCrLf & _
                "       select year(@closing_date) stock_year, month(@closing_date) stock_month, warehouse_code, item_code, lm_premonth from stock_master  " & vbCrLf & _
                "       union  " & vbCrLf & _
                "       select year(dateadd(m, 1, @closing_date)) stock_year, month(dateadd(m, 1, @closing_date)) stock_month, warehouse_code, item_code, tm_premonth from stock_master  " & vbCrLf & _
                "       union  " & vbCrLf & _
                "       select year(dateadd(m, 2, @closing_date)) stock_year, month(dateadd(m, 2, @closing_date)) stock_month, warehouse_code, item_code, nm_premonth from stock_master  " & vbCrLf & _
                "       union  " & vbCrLf & _
                "       select year(@start_date) stock_year, month(@start_date) stock_month, warehouse_code, item_code, nm_current from stock_master where datediff(m, @closing_date, @start_date) > 2  " & vbCrLf & _
                "   )bg where bg.stock_year = year(@start_date)  " & vbCrLf & _
                "   and bg.stock_month = month(@start_date)  " & vbCrLf & _
                "   group by item_code  " & vbCrLf & _
                " )stock_begin  " & vbCrLf & _
                " left join(  " & vbCrLf
                
    sql = sql + "   select item_code, sum(qty) stock_in  " & vbCrLf & _
                "   from(  " & vbCrLf & _
                "       select receipt_date, item_code, qty  " & vbCrLf & _
                "       from part_receipt  " & vbCrLf & _
                "       where receipt_cls not in ('R1')  " & vbCrLf & _
                "       union all  " & vbCrLf & _
                "       select receipt_date, item_code, -qty  " & vbCrLf & _
                "       from part_receipt  " & vbCrLf & _
                "       where receipt_cls in ('R1')  " & vbCrLf & _
                "   )stock_in  " & vbCrLf & _
                "   where receipt_date >= @start_date - (day(@start_date) - 1)  " & vbCrLf & _
                "   and receipt_date < @start_date  " & vbCrLf & _
                "   group by item_code  " & vbCrLf & _
                " )stock_in on stock_begin.item_code = stock_in.item_code  " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select childitem_code, sum(childrequirement_qty) stock_out  " & vbCrLf & _
                "   from part_supply  " & vbCrLf & _
                "   where supply_cls not in ('S1')  " & vbCrLf & _
                "   and childsupply_date >= @start_date - (day(@start_date) - 1)  " & vbCrLf & _
                "   and childsupply_date < @start_date  " & vbCrLf & _
                "   group by childitem_code  " & vbCrLf & _
                " )stock_out on stock_begin.item_code = stock_out.childitem_code  " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select pd.item_code, sum(pd.qty) po_qty  " & vbCrLf & _
                "   from purchaseorder_master pm  " & vbCrLf & _
                "   inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "   where isnull(pm.fix_cls, '0') = '1'  " & vbCrLf & _
                "   and pm.delivery_date < @start_date  " & vbCrLf & _
                "   group by pd.item_code  " & vbCrLf & _
                " )po on stock_begin.item_code = po.item_code  " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select pr.item_code, sum(pr.qty) po_receipt  " & vbCrLf & _
                "   from purchaseorder_master pm  " & vbCrLf & _
                "   inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "   inner join part_receipt pr on pm.supplier_code = pr.supplier_code and pm.po_no = pr.po_no and pd.item_code = pr.item_code  " & vbCrLf & _
                "   where pr.receipt_cls = 'R'  " & vbCrLf & _
                "   and pr.receipt_date < @start_date  " & vbCrLf & _
                "   group by pr.item_code  " & vbCrLf & _
                " )po_receipt on stock_begin.item_code = po_receipt.item_code  " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select pr.item_code, sum(pr.qty) po_return  " & vbCrLf & _
                "   from purchaseorder_master pm  " & vbCrLf & _
                "   inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "   inner join part_receipt pr on pm.supplier_code = pr.supplier_code and pm.po_no = pr.po_no and pd.item_code = pr.item_code  " & vbCrLf & _
                "   where pr.receipt_cls = 'R1'  " & vbCrLf & _
                "   and pr.receipt_date < @start_date  " & vbCrLf & _
                "   group by pr.item_code  " & vbCrLf & _
                " )po_return on stock_begin.item_code = po_return.item_code  " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select item_code, sum(po_cancel) po_cancel  " & vbCrLf & _
                "   from(  " & vbCrLf & _
                "       select pd.item_code, po_cancel = case when isnull(pd.qty, 0) - isnull(rc.qty, 0) > 0 then isnull(pd.qty, 0) - isnull(rc.qty, 0) else 0 end  " & vbCrLf & _
                "       from purchaseorder_master pm  " & vbCrLf & _
                "       inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "       left join(  " & vbCrLf & _
                "           select po_no, supplier_code, item_code, sum(qty) qty  " & vbCrLf & _
                "           from part_receipt  " & vbCrLf & _
                "           where receipt_cls = 'R'  " & vbCrLf & _
                "           group by po_no, supplier_code, item_code  " & vbCrLf & _
                "       )rc on pd.po_no = rc.po_no and pm.supplier_code = rc.supplier_code and pd.item_code = rc.item_code  " & vbCrLf & _
                "       where isnull(pd.complete_cls, '0') = '1'  " & vbCrLf & _
                "       and pm.delivery_date < @start_date  " & vbCrLf & _
                "   )po_cancel group by item_code  " & vbCrLf & _
                " )po_cancel on stock_begin.item_code = po_cancel.item_code  " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select dp.item_code, sum(dp.qty) pro_qty  " & vbCrLf & _
                "   from daily_production dp  " & vbCrLf & _
                "   where dp.schedule_date < @start_date  " & vbCrLf & _
                "   group by dp.item_code  " & vbCrLf & _
                " )pro on stock_begin.item_code = pro.item_code  " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select item_code, sum(qty) pro_result  " & vbCrLf & _
                "   from part_receipt  " & vbCrLf & _
                "   where receipt_cls = 'P1'  " & vbCrLf & _
                "   and receipt_date < @start_date  " & vbCrLf & _
                "   group by item_code  " & vbCrLf & _
                " )pro_result on stock_begin.item_code = pro_result.item_code  " & vbCrLf & _
                " left join(  "
            
    sql = sql + "   select item_code, sum(pro_cancel) pro_cancel  " & vbCrLf & _
                "   from(  " & vbCrLf & _
                "       select dp.item_code, pro_cancel = case when isnull(dp.qty, 0) - isnull(rs.qty, 0) > 0 then isnull(dp.qty, 0) - isnull(rs.qty, 0) else 0 end  " & vbCrLf & _
                "       from daily_production dp  " & vbCrLf & _
                "       left join(  " & vbCrLf & _
                "           select supplier_code, po_no, item_code, suratjalan_no, dailyseq_no, sum(qty) qty  " & vbCrLf & _
                "           from part_receipt  " & vbCrLf & _
                "           where receipt_cls = 'P1'  " & vbCrLf & _
                "           group by supplier_code, po_no, item_code, suratjalan_no, dailyseq_no  " & vbCrLf & _
                "       )rs on dp.factory_code = rs.supplier_code and dp.line_code = rs.po_no and dp.item_code = rs.item_code and dp.lot_no = rs.suratjalan_no and dp.seq_no = rs.dailyseq_no  " & vbCrLf & _
                "       where isnull(dp.complete_cls, '0') = '1'  " & vbCrLf & _
                "       and dp.schedule_date < @start_date  " & vbCrLf & _
                "   )pro_cancel group by item_code  " & vbCrLf & _
                " )pro_cancel on stock_begin.item_code = pro_cancel.item_code  " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select childitem_code, sum(childrequirement_qty) req_qty, sum(childrequirementresult_qty) req_result, sum(offchildrequirement_qty) req_off  " & vbCrLf & _
                "   from requirement  " & vbCrLf & _
                "   where production_date < @start_date  " & vbCrLf & _
                "   group by childitem_code  " & vbCrLf & _
                " )req on stock_begin.item_code = req.childitem_code  " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select 0 idx, 'adj' status, 'ADJ' cls, adj.item_code, adj.date, null od_qty, null od_incoming, null od_pono, null od_dono, null od_curr, null od_price,  " & vbCrLf & _
                "   null in_qty, null in_curr, null in_price, null in_amount, adj.qty_adj out_qty, null out_curr, null out_price, null out_amount  " & vbCrLf & _
                "   from(  " & vbCrLf & _
                "       select dateadd(d, -1, dateadd(m, 1, cast(stock_year + '-' + stock_month + '-1' as datetime))) date, item_code, sum([current]) - sum(inventory) qty_adj  " & vbCrLf & _
                "       from stock_history  " & vbCrLf & _
                "       where [current] <> isnull(inventory, [current])  " & vbCrLf & _
                "       group by stock_year, stock_month, item_code  " & vbCrLf & _
                "       union all  " & vbCrLf & _
                "       select dateadd(d, -1, dateadd(m, 1, @closing_date)) date, item_code, sum(lm_current) - sum(isnull(lm_inventory, lm_current)) qty_ad from stock_master  " & vbCrLf & _
                "       where lm_current <> isnull(lm_inventory, lm_current)  " & vbCrLf & _
                "       group by item_code  " & vbCrLf & _
                "   )adj  "
    
    sql = sql + "   union all  " & vbCrLf & _
                "   select 1 idx, 'po' status, '' cls, pd.item_code, pm.delivery_date date, pd.qty od_qty, null od_incoming, rtrim(pd.po_no) + case when isnull(po_cancel, 0) > 0 then ' (Completed by user)' else '' end od_pono, null od_dono, rtrim(cc.description) od_curr, pd.price + isnull(pd.price_service, 0) od_price,  " & vbCrLf & _
                "   null in_qty, null in_curr, null in_price, null in_amount, isnull(po_cancel, 0) out_qty, null out_curr, null out_price, null out_amount  " & vbCrLf & _
                "   from purchaseorder_master pm  " & vbCrLf & _
                "   inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "   left join curr_cls cc on pd.currency_code = cc.curr_cls  " & vbCrLf & _
                "   left join(  " & vbCrLf & _
                "       select pm.supplier_code, pm.po_no, pd.item_code, pd.qty - isnull(rc.qty, 0) po_cancel  " & vbCrLf & _
                "       from purchaseorder_master pm  " & vbCrLf & _
                "       inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "       left join(  " & vbCrLf & _
                "           select supplier_code, po_no, item_code, sum(qty) qty  " & vbCrLf & _
                "           from part_receipt pr  " & vbCrLf & _
                "           group by supplier_code, po_no, item_code  " & vbCrLf & _
                "       )rc on pm.supplier_code = rc.supplier_code and pm.po_no = rc.po_no and pd.item_code = rc.item_code  " & vbCrLf & _
                "       where pd.complete_cls = '1'  " & vbCrLf & _
                "   )po_cancel on pm.supplier_code = po_cancel.supplier_code and pm.po_no = po_cancel.po_no and pd.item_code = po_cancel.item_code  " & vbCrLf & _
                "   where isnull(pm.fix_cls, '0') = '1'  " & vbCrLf
                
    sql = sql + "   union all  " & vbCrLf & _
                "   select 1 idx, 'pro' status, '' cls, dp.item_code, schedule_date date, dp.qty od_qty, null od_incoming, rtrim(dp.lot_no) + case when isnull(pro_cancel, 0) > 0 then ' (Completed by user)' else '' end od_pono, null od_dono, null od_curr, 0 od_price,  " & vbCrLf & _
                "   null in_qty, null in_curr, null in_price, null in_amount, isnull(pro_cancel, 0) out_qty, null out_curr, null out_price, null out_amount  " & vbCrLf & _
                "   from daily_production dp  " & vbCrLf & _
                "   left join(  " & vbCrLf & _
                "       select dp.factory_code, dp.line_code, dp.item_code, dp.lot_no, dp.seq_no, dp.qty, dp.qty - isnull(rs.qty, 0) pro_cancel  " & vbCrLf & _
                "       from daily_production dp  " & vbCrLf & _
                "       left join(  " & vbCrLf & _
                "           select supplier_code, po_no, item_code, suratjalan_no, dailyseq_no, sum(qty) qty  " & vbCrLf & _
                "           from part_receipt  " & vbCrLf & _
                "           group by supplier_code, po_no, item_code, suratjalan_no, dailyseq_no  " & vbCrLf & _
                "       )rs on dp.factory_code = rs.supplier_code and dp.line_code = rs.po_no and dp.item_code = rs.item_code and dp.lot_no = rs.suratjalan_no and dp.seq_no = rs.dailyseq_no  " & vbCrLf & _
                "       where complete_cls = '1' and dp.qty > rs.qty  " & vbCrLf & _
                "   )pro_cancel on dp.factory_code = pro_cancel.factory_code and dp.line_code = pro_cancel.line_code and dp.item_code = pro_cancel.item_code and dp.lot_no = pro_cancel.lot_no and dp.seq_no = pro_cancel.seq_no  " & vbCrLf
                
    sql = sql + "   union all  " & vbCrLf & _
                "   select 2 idx, 'rec' status, pr.receipt_cls cls, pr.item_code, pr.receipt_date date, null od_qty, case when po.po_no is null then null else pr.qty end od_incoming, isnull(rtrim(po.po_no), 'Receipt Unscheduled') od_pono, rtrim(pr.suratjalan_no) od_dono, rtrim(cc.description) od_curr, pr.price od_price,  " & vbCrLf & _
                "   pr.qty in_qty, rtrim(cc.description) in_curr, pr.price in_price, pr.qty * pr.price in_amount, null out_qty, null out_curr, null out_price, null out_amount  " & vbCrLf & _
                "   from part_receipt pr  " & vbCrLf & _
                "   left join curr_cls cc on pr.currency_code = cc.curr_cls  " & vbCrLf & _
                "   left join(  " & vbCrLf & _
                "       select pm.supplier_code, pm.po_no, pd.item_code  " & vbCrLf & _
                "       from purchaseorder_master pm  " & vbCrLf & _
                "       inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "   )po on pr.supplier_code = po.supplier_code and pr.po_no = po.po_no and pr.item_code = po.item_code  " & vbCrLf & _
                "   where pr.receipt_cls = 'R'  " & vbCrLf
                
    sql = sql + "   union all  " & vbCrLf & _
                "   select 2 idx, 'res' status, pr.receipt_cls cls, pr.item_code, pr.receipt_date date, null od_qty, pr.qty od_incoming, '' od_pono, rtrim(pr.suratjalan_no) od_dono, rtrim(cc.description) od_curr, pr.price od_price,  " & vbCrLf & _
                "   pr.qty in_qty, rtrim(cc.description) in_curr, pr.price in_price, pr.qty * pr.price in_amount, null out_qty, null out_curr, null out_price, null out_amount  " & vbCrLf & _
                "   from part_receipt pr  " & vbCrLf & _
                "   left join curr_cls cc on pr.currency_code = cc.curr_cls  " & vbCrLf & _
                "   where pr.receipt_cls = 'P1'  " & vbCrLf
                
    sql = sql + "   union all  " & vbCrLf & _
                "   select 3 idx, 'ret' status, pr.receipt_cls cls, pr.item_code, pr.receipt_date date, null od_qty, case when po.po_no is null then null else -pr.qty end od_incoming, isnull(rtrim(po.po_no), 'Return Unscheduled') od_pono, rtrim(pr.suratjalan_no) od_dono, rtrim(cc.description) od_curr, pr.price od_price,  " & vbCrLf & _
                "   -pr.qty in_qty, rtrim(cc.description) in_curr, pr.price in_price, -pr.qty * pr.price in_amount, null out_qty, null out_curr, null out_price, null out_amount  " & vbCrLf & _
                "   from part_receipt pr  " & vbCrLf & _
                "   left join curr_cls cc on pr.currency_code = cc.curr_cls  " & vbCrLf & _
                "   left join(  " & vbCrLf & _
                "       select pm.supplier_code, pm.po_no, pd.item_code  " & vbCrLf & _
                "       from purchaseorder_master pm  " & vbCrLf & _
                "       inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "   )po on pr.supplier_code = po.supplier_code and pr.po_no = po.po_no and pr.item_code = po.item_code  " & vbCrLf & _
                "   where pr.receipt_cls = 'R1'  "
    
    sql = sql + "   union all  " & vbCrLf & _
                "   select 3 idx, 'out' status, ps.supply_cls cls, ps.childitem_code item_code, ps.childsupply_date date, null od_qty, null od_incoming, null od_pono, null od_dono, null od_curr, null od_price,  " & vbCrLf & _
                "   null in_qty, null in_curr, null in_price, null in_amount, ps.childrequirement_qty out_qty, rtrim(cc.description) out_curr, ps.price out_price, ps.amount out_amount  " & vbCrLf & _
                "   from part_supply ps  " & vbCrLf & _
                "   left join curr_cls cc on ps.currency_code = cc.curr_cls  " & vbCrLf & _
                "   where ps.supply_cls <> 'S1'  " & vbCrLf
                
    sql = sql + "   union all  " & vbCrLf & _
                "   select 4 idx, 'req' status, 'Req' status, rq.childitem_code item_code, production_date date, null od_qty, null od_incoming, null od_pono, null od_dono, null od_curr, null od_price,  " & vbCrLf & _
                "   null in_qty, null in_curr, null in_price, null in_amount, sum(rq.childrequirement_qty) - sum(rq.childrequirementresult_qty) - sum(rq.offchildrequirement_qty) out_qty, null out_curr, null out_price, null out_amount  " & vbCrLf & _
                "   from requirement rq  " & vbCrLf & _
                "   group by rq.childitem_code, production_date  " & vbCrLf & _
                " ) trans on stock_begin.item_code = trans.item_code  " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select cast(mrp_year + '-' + mrp_month + '-01' as datetime) mrp_date, '1' mrp_set from mrp_setting  " & vbCrLf & _
                " )mrp_set on year(trans.date) = year(mrp_date) and month(trans.date) = month(mrp_date), company_profile cp  " & vbCrLf & _
                " where stock_begin.item_code = @item_code  " & vbCrLf & _
                " and trans.date >= @start_date  " & vbCrLf & _
                " and trans.date <= @end_date  " & vbCrLf & _
                " order by trans.item_code, trans.date, trans.idx, trans.od_pono "
            
    Header
    
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
    
        dblBalanceOrder = CDbl(Format(adoRs.Fields("po_qty") - adoRs.Fields("po_receipt") + adoRs.Fields("po_return") - adoRs.Fields("po_cancel") + _
            adoRs.Fields("pro_qty") - adoRs.Fields("pro_result") - adoRs.Fields("pro_cancel"), gs_formatQty))
        dblBalance = CDbl(Format(adoRs.Fields("stock_begin") + adoRs.Fields("stock_in") - adoRs.Fields("stock_out"), gs_formatQty))
        dblBalanceEffective = dblBalanceOrder + dblBalance
       
        With grid
            .AddItem ""
            .TextMatrix(.Rows - 1, bteColDate) = "Beginning"
            .TextMatrix(.Rows - 1, bteColOrderBal) = Format(dblBalanceOrder, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColBalQty) = Format(dblBalance, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColBalEffective) = Format(dblBalanceEffective, gs_formatQty)
                      
            While Not adoRs.EOF
                dblBalanceOrder = dblBalanceOrder + CDbl(Format(Val(adoRs.Fields("od_qty") & "") - Val(adoRs.Fields("od_incoming") & ""), gs_formatQty))
                If adoRs.Fields("status") = "po" Or adoRs.Fields("status") = "pro" Then dblBalanceOrder = dblBalanceOrder - CDbl(Format(Val(adoRs.Fields("out_qty") & ""), gs_formatQty))
                
                dblBalance = dblBalance + CDbl(Format(Val(adoRs.Fields("in_qty") & ""), gs_formatQty))
                If adoRs.Fields("status") <> "po" And adoRs.Fields("status") <> "pro" And adoRs.Fields("status") <> "req" Then dblBalance = dblBalance - CDbl(Format(Val(adoRs.Fields("out_qty") & ""), gs_formatQty))
                
                dblBalanceEffective = dblBalanceEffective + CDbl(Format(Val(adoRs.Fields("od_qty") & "") - Val(adoRs.Fields("od_incoming") & "") + Val(adoRs.Fields("in_qty") & ""), gs_formatQty))
                If adoRs.Fields("mrp_set") & "" = "0" Then
                    If adoRs.Fields("status") <> "req" Then
                        dblBalanceEffective = dblBalanceEffective - CDbl(Format(Val(adoRs.Fields("out_qty") & ""), gs_formatQty))
                    End If
                Else
                    If adoRs.Fields("status") = "req" Or adoRs.Fields("status") = "po" Or adoRs.Fields("status") = "pro" Then
                        dblBalanceEffective = dblBalanceEffective - CDbl(Format(Val(adoRs.Fields("out_qty") & ""), gs_formatQty))
                    End If
                End If
                
                .AddItem ""
                .TextMatrix(.Rows - 1, bteColDate) = Format(adoRs.Fields("date"), "dd-MMM-yyyy")
                .TextMatrix(.Rows - 1, bteColOrderQty) = IIf(IsNull(adoRs.Fields("od_qty")), "", Format(adoRs.Fields("od_qty"), gs_formatQty))
                .TextMatrix(.Rows - 1, bteColOrderIn) = IIf(IsNull(adoRs.Fields("od_incoming")), "", Format(adoRs.Fields("od_incoming"), gs_formatQty))
                .TextMatrix(.Rows - 1, bteColOrderPONo) = Trim(adoRs.Fields("od_pono") & "")
                .TextMatrix(.Rows - 1, bteColOrderDONo) = Trim(adoRs.Fields("od_dono") & "")
                .TextMatrix(.Rows - 1, bteColOrderCurr) = Trim(adoRs.Fields("od_curr") & "")
                .TextMatrix(.Rows - 1, bteColOrderPrice) = IIf(IsNull(adoRs.Fields("od_price")), "", Format(adoRs.Fields("od_price"), gs_formatPrice))
                .TextMatrix(.Rows - 1, bteColOrderBal) = Format(dblBalanceOrder, gs_formatQty)
                If adoRs.Fields("in_qty") <> 0 Then
                    .TextMatrix(.Rows - 1, bteColInCls) = Trim(adoRs.Fields("cls") & "")
                End If
                .TextMatrix(.Rows - 1, bteColInQty) = IIf(IsNull(adoRs.Fields("in_qty")), "", Format(adoRs.Fields("in_qty"), gs_formatQty))
                .TextMatrix(.Rows - 1, bteColInCurr) = Trim(adoRs.Fields("in_curr") & "")
                .TextMatrix(.Rows - 1, bteColInPrice) = IIf(IsNull(adoRs.Fields("in_price")), "", Format(adoRs.Fields("in_price"), gs_formatPrice))
                .TextMatrix(.Rows - 1, bteColInAmount) = IIf(IsNull(adoRs.Fields("in_amount")), "", Format(adoRs.Fields("in_amount"), gs_formatPrice))
                If adoRs.Fields("out_qty") <> 0 Then
                    .TextMatrix(.Rows - 1, bteColOutCls) = Trim(adoRs.Fields("cls") & "")
                End If
                If adoRs.Fields("status") = "req" Then
                    .TextMatrix(.Rows - 1, bteColOutReq) = IIf(IsNull(adoRs.Fields("out_qty")), "", Format(adoRs.Fields("out_qty"), gs_formatQty))
                ElseIf adoRs.Fields("out_qty") <> 0 Then
                    .TextMatrix(.Rows - 1, bteColOutQty) = IIf(IsNull(adoRs.Fields("out_qty")), "", Format(adoRs.Fields("out_qty"), gs_formatQty))
                End If
                .TextMatrix(.Rows - 1, bteColOutCurr) = Trim(adoRs.Fields("out_curr") & "")
                .TextMatrix(.Rows - 1, bteColOutPrice) = IIf(IsNull(adoRs.Fields("out_price")), "", Format(adoRs.Fields("out_price"), gs_formatPrice))
                .TextMatrix(.Rows - 1, bteColOutAmount) = IIf(IsNull(adoRs.Fields("out_amount")), "", Format(adoRs.Fields("out_amount"), gs_formatPrice))
                .TextMatrix(.Rows - 1, bteColBalQty) = Format(dblBalance, gs_formatQty)
                If Not IsNull(adoRs.Fields("in_curr")) Then
                    .TextMatrix(.Rows - 1, bteColBalCurr) = Trim(adoRs.Fields("in_curr") & "")
                    .TextMatrix(.Rows - 1, bteColBalPrice) = IIf(IsNull(adoRs.Fields("in_price")), "", Format(adoRs.Fields("in_price"), gs_formatPrice))
                    .TextMatrix(.Rows - 1, bteColBalAmount) = IIf(IsNull(adoRs.Fields("in_amount")), "", Format(adoRs.Fields("in_amount"), gs_formatPrice))
                ElseIf Not IsNull(adoRs.Fields("out_curr")) Then
                    .TextMatrix(.Rows - 1, bteColBalCurr) = Trim(adoRs.Fields("out_curr") & "")
                    .TextMatrix(.Rows - 1, bteColBalPrice) = IIf(IsNull(adoRs.Fields("out_price")), "", Format(adoRs.Fields("out_price"), gs_formatPrice))
                    .TextMatrix(.Rows - 1, bteColBalAmount) = IIf(IsNull(adoRs.Fields("out_amount")), "", Format(adoRs.Fields("out_amount"), gs_formatPrice))
                End If
                .TextMatrix(.Rows - 1, bteColBalEffective) = Format(dblBalanceEffective, gs_formatQty)
                    
                adoRs.MoveNext
            Wend
        End With
    End If
    adoRs.Close

ErrExit:
    Me.MousePointer = vbDefault
    Set adoRs = Nothing
    Exit Sub
errHandler:
    LblPesan.Caption = err.Description
    err.clear
    Resume ErrExit
End Sub

Public Sub ClickSearch()
    Call Cmd_Save_Click(9)
End Sub

Private Sub CboItemCD_Change()
    LblPesan = ""
    Call Header
    If CboItemCD.MatchFound Then
        lbldesc = CboItemCD.List(CboItemCD.ListIndex, 1)
        Lbl_UnitDesc = CboItemCD.List(CboItemCD.ListIndex, 2)
    Else
        lbldesc = ""
        Lbl_UnitDesc = ""
        LblPesan = DisplayMsg(4003)
    End If
End Sub

Private Sub CboItemCD_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        LblPesan = ""
        Call Header
        If CboItemCD.MatchFound Then
            lbldesc = CboItemCD.List(CboItemCD.ListIndex, 1)
            Lbl_UnitDesc = CboItemCD.List(CboItemCD.ListIndex, 2)
        Else
            lbldesc = ""
            Lbl_UnitDesc = ""
            LblPesan = DisplayMsg(4003)
        End If
    End If
End Sub

Private Sub CboItemCD_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cmd_preview_Click()
    Dim crApp As New CRAXDDRT.application
    Dim crRpt As New CRAXDDRT.report
    Dim frmRpt As New FrmRpt3
    Dim adoRs As New ADODB.Recordset
    
    On Error GoTo errHandler
    Me.MousePointer = vbHourglass
    LblPesan.Caption = ""
    
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly
    If adoRs.EOF Then
        LblPesan.Caption = DisplayMsg(4006)
        GoTo ErrExit
    End If
    
    sqlprint = sql
    reportcode = "Receiptsupplyschedule"
    printorient = 2
    
    Set crRpt = crApp.OpenReport(App.path & "\Reports\rpt_recSupSchedule.rpt")
    With crRpt
        .Database.Tables(1).SetDataSource adoRs
        .ReportTitle = "Receipt Supply Schedule Inquiry"
        .FormulaFields(2).Text = "'" & Trim(CboItemCD.Text) & "  " & Trim(lbldesc.Caption) & "'"
        .FormulaFields(3).Text = "'" & Trim(Lbl_UnitDesc.Caption) & "'"
        .FormulaFields(4).Text = "'" & Format(DMonth(0).Value, "dd-MMM-yyyy") & " to " & Format(DMonth(1).Value, "dd-MMM-yyyy") & "'"
        .FormulaFields(11).Text = "" & gi_decimalDigitQty & ""
        .FormulaFields(12).Text = "" & gi_decimalDigitPrice & ""
        .FormulaFields(13).Text = "" & gi_decimalDigitAmount & ""
    End With
    
    With frmRpt
        .CRViewer1.ReportSource = crRpt
        .CRViewer1.ViewReport
        .CRViewer1.Zoom 1
        .WindowState = 2
        .Show 1
    End With
    
ErrExit:
    Me.MousePointer = vbDefault
    Set crRpt = Nothing
    Set crApp = Nothing
    Set frmRpt = Nothing
    Set adoRs = Nothing
    Exit Sub
errHandler:
    LblPesan.Caption = err.Description
    err.clear
    Resume ErrExit
End Sub

Private Sub Cmd_Save_Click(Index As Integer)
    Select Case Index
    Case 8
        If Cmd_save(8).Caption = "&Back" And frmPOParts.popanggil = "poparts" Then
            Cmd_save(8).Caption = "Sub &Menu"
            frmPOParts.popanggil = ""
            frmPOParts.Show
            Unload Me
            Exit Sub
        ElseIf Cmd_save(8).Caption = "&Back" And frmPOSteelCoil.popanggil = "posteelcoil" Then
            Cmd_save(8).Caption = "Sub &Menu"
            frmPOSteelCoil.popanggil = ""
            frmPOSteelCoil.Show
            Unload Me
            Exit Sub
        ElseIf Cmd_save(8).Caption = "&Back" And frmPOSubcon.popanggil = "posubcon" Then
            Cmd_save(8).Caption = "Sub &Menu"
            frmPOSubcon.popanggil = ""
            frmPOSubcon.Show
            Unload Me
            Exit Sub
        End If
        frmMainMenu.Show
        Unload Me
    Case 9
        If CboItemCD.Text = "" Then
            LblPesan = DisplayMsg(1009)
        Else
            Call SettingGrid
        End If
    End Select
End Sub

Private Sub Command4_Click()
    Me.MousePointer = vbHourglass
    frm_BrowseItem.getItemCode = CboItemCD.Text
    frm_BrowseItem.Show 1
    CboItemCD.Text = frm_BrowseItem.getItemCode
    CboItemCD.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblPesan.Caption = ErrMsg
    End If
End Sub

Private Sub DMonth_Change(Index As Integer)
    LblPesan.Caption = ""
    If Index = 0 Then
        If DMonth(0).Value > dteMRP Then
            LblPesan.Caption = "[0000] Date start must be lower than MRP setting !"
            DMonth(0).Value = dteMRP
        End If
    End If
    Call Header
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    
    Dim adoRs As New ADODB.Recordset
    dteMRP = Date
    adoRs.Open "select min(cast(mrp_year + '-' + mrp_month + '-1' as datetime)) mrp_date from mrp_setting", Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        dteMRP = IIf(IsNull(adoRs.Fields("mrp_date")), Now, adoRs.Fields("mrp_date"))
    End If
    adoRs.Close
    
    DMonth(0) = Format(dteMRP, "1 MMM yyyy")
    DMonth(1) = Format(Date, "dd MMM yyyy")
    bteHakPrice = hakPrice(Me.Name)
    
    Call AddToComboItem
    Call Header
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub
