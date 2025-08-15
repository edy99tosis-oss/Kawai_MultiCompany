VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAlarmList 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Parts (Material) Alarm List"
   ClientHeight    =   5025
   ClientLeft      =   990
   ClientTop       =   2895
   ClientWidth     =   9525
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAlarmList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   315
      Left            =   3900
      TabIndex        =   6
      Top             =   2535
      Width           =   300
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   3900
      TabIndex        =   4
      Top             =   2115
      Width           =   300
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   330
      TabIndex        =   13
      Top             =   3555
      Width           =   8865
      Begin VB.Label lblPesan 
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
         TabIndex        =   14
         Top             =   195
         Width           =   8640
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xcel"
      Height          =   375
      Left            =   8070
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4245
      Width           =   1125
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00A6D2FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4245
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker dtpDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd MMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   330
      Left            =   1725
      TabIndex        =   0
      Top             =   1275
      Width           =   1500
      _ExtentX        =   2646
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
      Format          =   150994947
      CurrentDate     =   37798
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   7335
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   240
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin MSForms.ComboBox cboSupplier 
      Height          =   330
      Left            =   1725
      TabIndex        =   7
      Top             =   2970
      Width           =   2100
      VariousPropertyBits=   746604571
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "3704;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Code"
      Height          =   195
      Index           =   5
      Left            =   330
      TabIndex        =   26
      Top             =   3045
      Width           =   1215
   End
   Begin VB.Label lblSupplier 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3900
      TabIndex        =   25
      Top             =   3045
      Width           =   5280
   End
   Begin VB.Line Line5 
      X1              =   3900
      X2              =   9135
      Y1              =   3285
      Y2              =   3285
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   1
      Left            =   6150
      TabIndex        =   24
      Top             =   2610
      Width           =   3015
   End
   Begin VB.Label lblProd 
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   1
      Left            =   6150
      TabIndex        =   23
      Top             =   2190
      Width           =   3015
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   6150
      X2              =   9150
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   6150
      X2              =   9150
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Label lblControl 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   6150
      TabIndex        =   22
      Top             =   1350
      Width           =   3000
   End
   Begin VB.Line Line4 
      X1              =   6150
      X2              =   9135
      Y1              =   1590
      Y2              =   1590
   End
   Begin MSForms.ComboBox cboControl 
      Height          =   330
      Left            =   5085
      TabIndex        =   1
      Top             =   1275
      Width           =   960
      VariousPropertyBits=   746604571
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1693;582"
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Control"
      Height          =   195
      Index           =   4
      Left            =   3885
      TabIndex        =   21
      Top             =   1350
      Width           =   630
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   0
      Left            =   4290
      TabIndex        =   20
      Top             =   2610
      Width           =   1740
   End
   Begin VB.Label lblProd 
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   0
      Left            =   4290
      TabIndex        =   19
      Top             =   2190
      Width           =   1740
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   4290
      X2              =   6015
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   4290
      X2              =   6015
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line1 
      X1              =   3900
      X2              =   9135
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Label lblFactory 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3900
      TabIndex        =   18
      Top             =   1770
      Width           =   5280
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Material Code"
      Height          =   195
      Index           =   3
      Left            =   330
      TabIndex        =   17
      Top             =   2610
      Width           =   1185
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   195
      Index           =   2
      Left            =   330
      TabIndex        =   16
      Top             =   2190
      Width           =   1155
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Factory Code"
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   15
      Top             =   1770
      Width           =   1140
   End
   Begin MSForms.ComboBox cboMat 
      Height          =   330
      Left            =   1725
      TabIndex        =   5
      Top             =   2535
      Width           =   2115
      VariousPropertyBits=   746604571
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "3731;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboProd 
      Height          =   330
      Left            =   1725
      TabIndex        =   3
      Top             =   2115
      Width           =   2115
      VariousPropertyBits=   746604571
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "3731;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboFactory 
      Height          =   330
      Left            =   1725
      TabIndex        =   2
      Top             =   1695
      Width           =   2100
      VariousPropertyBits=   746604571
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "3704;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Parts (Material) Alarm List"
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
      Left            =   330
      TabIndex        =   12
      Top             =   450
      Width           =   8865
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Index           =   0
      Left            =   330
      TabIndex        =   11
      Top             =   1350
      Width           =   405
   End
End
Attribute VB_Name = "FrmAlarmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dteMRP As Date

Private Sub AddToComboControl()
    Dim adoRs As New ADODB.Recordset
    
    sql = "Select Control_Cls, Description From Control_Cls Where Control_Cls In ('01', '02', '03')"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With cboControl
        .clear
        .columnCount = 2
        .ColumnWidths = "40 pt;100 pt"
        .ListWidth = 150
        .ListRows = 15
        While Not adoRs.EOF
            .AddItem ""
            .List(.ListCount - 1, 0) = Trim(adoRs!control_cls)
            .List(.ListCount - 1, 1) = Trim(adoRs!Description)
            adoRs.MoveNext
        Wend
        .ListIndex = 0
    End With
End Sub

Private Sub AddToComboFactory()
    Dim adoRs As New ADODB.Recordset
    
    sql = "Select Trade_Code, Trade_Name From Trade_Master Where Trade_Code In (Select Distinct Manufacture_Code From Manufacture_Line) " '& _
            "and trade_code in (" & _
            "       select code from (select username, trade_code code from user_factory union select username, wh_code code from user_warehouse) a where a.username = '" & userLogin & "' " & _
            " )"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With cboFactory
        .clear
        .columnCount = 2
        .ColumnWidths = "70 pt;180 pt"
        .ListWidth = 250
        .ListRows = 15
        .AddItem ""
        .List(.ListCount - 1, 0) = strAll
        .List(.ListCount - 1, 1) = strAll
        While Not adoRs.EOF
            .AddItem ""
            .List(.ListCount - 1, 0) = Trim(adoRs!Trade_Code)
            .List(.ListCount - 1, 1) = Trim(adoRs!trade_name)
            adoRs.MoveNext
        Wend
        .ListIndex = 0
    End With
End Sub

Private Sub AddToComboProdMat()
    Dim adoRs As New ADODB.Recordset
    
    With cboProd
        .clear
        .columnCount = 3
        .ColumnWidths = "80 pt;80 pt;240 pt"
        .ListWidth = 400
        .ListRows = 15
        .AddItem ""
        .List(.ListCount - 1, 0) = strAll
        .List(.ListCount - 1, 1) = strAll
        .List(.ListCount - 1, 2) = strAll
        .ListIndex = 0
    End With
    
    With CboMat
        .clear
        .columnCount = 3
        .ColumnWidths = "80 pt;80 pt;240 pt"
        .ListWidth = 400
        .ListRows = 15
        .AddItem ""
        .List(.ListCount - 1, 0) = strAll
        .List(.ListCount - 1, 1) = strAll
        .List(.ListCount - 1, 2) = strAll
        .ListIndex = 0
    End With
    
    sql = "Select Item_Code, MakerItem_Code, Item_Name, Production_Cls, FinishGoodPart_Cls " & _
        "From Item_Master Where Use_EndDay >= Convert(Char(8), Getdate(), 112) "
    
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    While Not adoRs.EOF
        If adoRs.Fields("FinishGoodPart_Cls") = "01" Or adoRs.Fields("Production_Cls") = "01" Then
            With cboProd
                .AddItem ""
                .List(.ListCount - 1, 0) = Trim(adoRs!Item_Code)
                .List(.ListCount - 1, 1) = Trim(adoRs!MakerItem_Code)
                .List(.ListCount - 1, 2) = Trim(adoRs!item_name)
            End With
        End If
        If adoRs.Fields("FinishGoodPart_Cls") = "02" Then
            With CboMat
                .AddItem ""
                .List(.ListCount - 1, 0) = Trim(adoRs!Item_Code)
                .List(.ListCount - 1, 1) = Trim(adoRs!MakerItem_Code)
                .List(.ListCount - 1, 2) = Trim(adoRs!item_name)
            End With
        End If
        adoRs.MoveNext
    Wend
    adoRs.Close
End Sub

Private Sub AddToComboSupplier()
    Dim adoRs As New ADODB.Recordset
    
    sql = "Select Trade_Code, Trade_Name From Trade_Master Where Trade_Cls In ('1', '2', '3') "
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With cboSupplier
        .clear
        .columnCount = 2
        .ColumnWidths = "70 pt;180 pt"
        .ListWidth = 250
        .ListRows = 15
        .AddItem ""
        .List(.ListCount - 1, 0) = strAll
        .List(.ListCount - 1, 1) = strAll
        While Not adoRs.EOF
            .AddItem ""
            .List(.ListCount - 1, 0) = Trim(adoRs!Trade_Code)
            .List(.ListCount - 1, 1) = Trim(adoRs!trade_name)
            adoRs.MoveNext
        Wend
        .ListIndex = 0
    End With
End Sub

Private Sub PrintMRPExcel()
    Dim adoRs As New ADODB.Recordset
    Dim xlapp As New Excel.application
    Dim xlBook As New Excel.Workbook
    Dim xlSheet As New Excel.Worksheet
    
    Dim strLastCalculate As String
    
    Dim lngRow As Long
    Dim strTempItem As String
    Dim strTempItemPrint As String
    Dim booPrint As Boolean
    
    Dim lngRowBegin As Long
    Dim lngRowEnd As Long
    
    Dim dblStock As Double
    Dim dblOrderLast As Double
    Dim dblOrderCurr As Double
    Dim dblOrderReceipt As Double
    
    Const strColMatCode As String = "A"
    Const strColMatDesc As String = "B"
    Const strColUnit As String = "C"
    Const strColOrderLast As String = "D"
    Const strColOrderCurr As String = "E"
    Const strColOrderReceipt As String = "F"
    Const strColProdQty As String = "G"
    Const strColStock As String = "H"
    Const strColShortDate As String = "I"
    Const strColLeadTime As String = "J"
    Const strColOrderDate As String = "K"
    Const strColControl As String = "L"
    Const strColOrderQty As String = "M"
    Const strColSupplier As String = "N"
    Const strColProdCode As String = "O"
    Const strColProdDesc As String = "P"
    Const strColFactory As String = "Q"
    Const strColLine As String = "R"
    
    On Error GoTo errHandler
    Me.MousePointer = vbHourglass
    
    LblPesan.Caption = ""
    
    sql = "select max(register_date) last_calculate from requirement where production_date <= '" & Format(dteMRP, "yyyy-MM-dd") & "' "
    adoRs.Open sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        strLastCalculate = Format(adoRs.Fields("last_calculate"), "dd-MMM-yyyy hh:mm")
    End If
    adoRs.Close
    
    sql = " declare @start_date datetime " & vbCrLf & _
                " declare @end_date datetime " & vbCrLf & _
                " declare @closing_date as datetime " & vbCrLf & _
                "  " & vbCrLf & _
                " set @start_date = '" & Format(dteMRP, "yyyy-MM-dd") & "' " & vbCrLf & _
                " set @end_date = '" & Format(dtpDate.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
                " set @closing_date = (select top 1 cast(inventory_year as nvarchar(4)) + '-' + cast(inventory_month as nvarchar(2)) +  '-1' from inventory_control order by inventory_year desc, inventory_month desc) " & vbCrLf & _
                "  " & vbCrLf
                
    sql = sql + "select isnull(stock_begin.stock_begin, 0) stock_begin, isnull(stock_in.stock_in, 0) stock_in, isnull(stock_out.stock_out, 0) stock_out,  " & vbCrLf & _
                " isnull(po.po_qty, 0) po_qty, isnull(po_receipt.po_receipt, 0) po_receipt, isnull(po_return.po_return, 0) po_return, isnull(po_cancel.po_cancel, 0) po_cancel,  " & vbCrLf & _
                " isnull(pro.pro_qty, 0) pro_qty, isnull(pro_result.pro_result, 0) pro_result, isnull(pro_cancel.pro_cancel, 0) pro_cancel, " & vbCrLf & _
                " isnull(req.req_qty, 0) req_qty, isnull(req.req_result, 0) req_result, isnull(req.req_off, 0) req_off, trans.*, isnull(mrp_set, '0') mrp_set " & vbCrLf & _
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
                " )stock_begin  " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select item_code, sum(qty) stock_in " & vbCrLf & _
                "   from( " & vbCrLf & _
                "       select receipt_date, item_code, qty " & vbCrLf & _
                "       from part_receipt  " & vbCrLf & _
                "       where receipt_cls not in ('R1')  " & vbCrLf & _
                "       union all " & vbCrLf & _
                "       select receipt_date, item_code, -qty " & vbCrLf & _
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
                "   where pr.receipt_cls = 'R'   " & vbCrLf & _
                "   and pr.receipt_date < @start_date   " & vbCrLf & _
                "   group by pr.item_code  " & vbCrLf & _
                " )po_receipt on stock_begin.item_code = po_receipt.item_code  " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select pr.item_code, sum(pr.qty) po_return  " & vbCrLf & _
                "   from purchaseorder_master pm  " & vbCrLf & _
                "   inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "   inner join part_receipt pr on pm.supplier_code = pr.supplier_code and pm.po_no = pr.po_no and pd.item_code = pr.item_code  " & vbCrLf & _
                "   where pr.receipt_cls = 'R1'   " & vbCrLf & _
                "   and pr.receipt_date < @start_date   " & vbCrLf & _
                "   group by pr.item_code  " & vbCrLf & _
                " )po_return on stock_begin.item_code = po_return.item_code  "
    
    sql = sql + " left join(  " & vbCrLf & _
                "   select item_code, sum(po_cancel) po_cancel  " & vbCrLf & _
                "   from(  " & vbCrLf & _
                "       select pd.item_code, po_cancel = case when isnull(pd.qty, 0) - isnull(rc.qty, 0) > 0 then isnull(pd.qty, 0) - isnull(rc.qty, 0) else 0 end  " & vbCrLf & _
                "       from purchaseorder_master pm  " & vbCrLf & _
                "       inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "       left join(  " & vbCrLf & _
                "           select po_no, supplier_code, item_code, sum(qty) qty  " & vbCrLf & _
                "           from part_receipt   " & vbCrLf & _
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
                " )pro on stock_begin.item_code = pro.item_code  "
    
    sql = sql + " left join(  " & vbCrLf & _
                "   select item_code, sum(qty) pro_result  " & vbCrLf & _
                "   from part_receipt   " & vbCrLf & _
                "   where receipt_cls = 'P1'  " & vbCrLf & _
                "   and receipt_date < @start_date  " & vbCrLf & _
                "   group by item_code  " & vbCrLf & _
                " )pro_result on stock_begin.item_code = pro_result.item_code  " & vbCrLf
                
    sql = sql + " left join(   " & vbCrLf & _
                "   select item_code, sum(pro_cancel) pro_cancel  " & vbCrLf & _
                "   from(  " & vbCrLf & _
                "       select dp.item_code, pro_cancel = case when isnull(dp.qty, 0) - isnull(rs.qty, 0) > 0 then isnull(dp.qty, 0) - isnull(rs.qty, 0) else 0 end  " & vbCrLf & _
                "       from daily_production dp   " & vbCrLf & _
                "       left join(  " & vbCrLf & _
                "           select supplier_code, po_no, item_code, suratjalan_no, dailyseq_no, sum(qty) qty  " & vbCrLf & _
                "           from part_receipt  " & vbCrLf & _
                "           where receipt_cls = 'P1'  " & vbCrLf & _
                "           group by supplier_code, po_no, item_code, suratjalan_no, dailyseq_no  " & vbCrLf & _
                "       )rs on dp.factory_code = rs.supplier_code and dp.line_code = rs.po_no and dp.item_code = rs.item_code and dp.lot_no = rs.suratjalan_no and dp.seq_no = rs.dailyseq_no  " & vbCrLf & _
                "       where isnull(dp.complete_cls, '0') = '1'  " & vbCrLf & _
                "       and dp.schedule_date < @start_date   " & vbCrLf & _
                "   )pro_cancel group by item_code  " & vbCrLf & _
                " )pro_cancel on stock_begin.item_code = pro_cancel.item_code   "
    
    sql = sql + " left join(   " & vbCrLf & _
                "   select childitem_code, sum(childrequirement_qty) req_qty, sum(childrequirementresult_qty) req_result, sum(offchildrequirement_qty) req_off   " & vbCrLf & _
                "   from requirement   " & vbCrLf & _
                "   where production_date < @start_date   " & vbCrLf & _
                "   group by childitem_code   " & vbCrLf & _
                " )req on stock_begin.item_code = req.childitem_code   " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select 0 idx, 'adj' status, adj.item_code, '' partno, '' child_name, adj.date, '' factory_code, '' line_code,  " & vbCrLf & _
                "   '' parentitem_code, '' makeritem_code, '' parent_name, '' control, '' unit_desc, 0 orderpoint_qty, 0 leadtime, '' trade_code, " & vbCrLf & _
                "   0 od_qty, 0 od_incoming, 0 in_qty, adj.qty_adj out_qty " & vbCrLf & _
                "   from(  " & vbCrLf & _
                "       select dateadd(d, -1, dateadd(m, 1, cast(stock_year + '-' + stock_month + '-1' as datetime))) date, item_code, sum([current]) - sum(inventory) qty_adj  " & vbCrLf & _
                "       from stock_history  " & vbCrLf & _
                "       where [current] <> isnull(inventory, [current])  " & vbCrLf & _
                "       group by stock_year, stock_month, item_code  " & vbCrLf & _
                "       union all  " & vbCrLf & _
                "       select dateadd(d, -1, dateadd(m, 1, @closing_date)) date, item_code, sum(lm_current) - sum(isnull(lm_inventory, lm_current)) qty_ad from stock_master   " & vbCrLf & _
                "       where lm_current <> isnull(lm_inventory, lm_current)  " & vbCrLf & _
                "       group by item_code  " & vbCrLf & _
                "   )adj  " & vbCrLf
                
    sql = sql + "   union all  " & vbCrLf & _
                "   select 1 idx, 'po' status, pd.item_code, '' partno, '' child_name, pm.delivery_date date, '' factory_code, '' line_code,  " & vbCrLf & _
                "   '' parentitem_code, '' makeritem_code, '' parent_name, '' control, '' unit_desc, 0 orderpoint_qty, 0 leadtime, '' trade_code, " & vbCrLf & _
                "   pd.qty od_qty, 0 od_incoming, 0 in_qty, isnull(po_cancel, 0) out_qty " & vbCrLf & _
                "   from purchaseorder_master pm  " & vbCrLf & _
                "   inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "   left join curr_cls cc on pd.currency_code = cc.curr_cls  " & vbCrLf & _
                "   left join( " & vbCrLf & _
                "       select pm.supplier_code, pm.po_no, pd.item_code, pd.qty - isnull(rc.qty, 0) po_cancel " & vbCrLf & _
                "       from purchaseorder_master pm  " & vbCrLf & _
                "       inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "       left join( " & vbCrLf & _
                "           select supplier_code, po_no, item_code, sum(qty) qty " & vbCrLf & _
                "           from part_receipt pr " & vbCrLf & _
                "           group by supplier_code, po_no, item_code " & vbCrLf & _
                "       )rc on pm.supplier_code = rc.supplier_code and pm.po_no = rc.po_no and pd.item_code = rc.item_code " & vbCrLf & _
                "       where pd.complete_cls = '1' " & vbCrLf & _
                "   )po_cancel on pm.supplier_code = po_cancel.supplier_code and pm.po_no = po_cancel.po_no and pd.item_code = po_cancel.item_code " & vbCrLf & _
                "   where isnull(pm.fix_cls, '0') = '1'  " & vbCrLf
                
    sql = sql + "   union all  " & vbCrLf & _
                "   select 1 idx, 'pro' status, dp.item_code, '' partno, '' child_name, schedule_date date, '' factory_code, '' line_code,  " & vbCrLf & _
                "   '' parentitem_code, '' makeritem_code, '' parent_name, '' control, '' unit_desc, 0 orderpoint_qty, 0 leadtime, '' trade_code, " & vbCrLf & _
                "   dp.qty od_qty, 0 od_incoming, 0 in_qty, isnull(pro_cancel, 0) out_qty " & vbCrLf & _
                "   from daily_production dp  " & vbCrLf & _
                "   left join( " & vbCrLf & _
                "       select dp.factory_code, dp.line_code, dp.item_code, dp.lot_no, dp.seq_no, dp.qty, dp.qty - isnull(rs.qty, 0) pro_cancel " & vbCrLf & _
                "       from daily_production dp " & vbCrLf & _
                "       left join( " & vbCrLf & _
                "           select supplier_code, po_no, item_code, suratjalan_no, dailyseq_no, sum(qty) qty " & vbCrLf & _
                "           from part_receipt " & vbCrLf & _
                "           group by supplier_code, po_no, item_code, suratjalan_no, dailyseq_no " & vbCrLf & _
                "       )rs on dp.factory_code = rs.supplier_code and dp.line_code = rs.po_no and dp.item_code = rs.item_code and dp.lot_no = rs.suratjalan_no and dp.seq_no = rs.dailyseq_no " & vbCrLf & _
                "       where complete_cls = '1' and dp.qty > rs.qty  " & vbCrLf & _
                "   )pro_cancel on dp.factory_code = pro_cancel.factory_code and dp.line_code = pro_cancel.line_code and dp.item_code = pro_cancel.item_code and dp.lot_no = pro_cancel.lot_no and dp.seq_no = pro_cancel.seq_no " & vbCrLf
                
    sql = sql + "   union all  " & vbCrLf & _
                "   select 2 idx, 'rec' status, pr.item_code, '' partno, '' child_name, pr.receipt_date date, '' factory_code, '' line_code,  " & vbCrLf & _
                "   '' parentitem_code, '' makeritem_code, '' parent_name, '' control, '' unit_desc, 0 orderpoint_qty, 0 leadtime, '' trade_code, " & vbCrLf & _
                "   0 od_qty, case when po.po_no is null then null else pr.qty end od_incoming, pr.qty in_qty, 0 out_qty " & vbCrLf & _
                "   from part_receipt pr  " & vbCrLf & _
                "   left join curr_cls cc on pr.currency_code = cc.curr_cls  " & vbCrLf & _
                "   left join(  " & vbCrLf & _
                "       select pm.supplier_code, pm.po_no, pd.item_code  " & vbCrLf & _
                "       from purchaseorder_master pm  " & vbCrLf & _
                "       inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "   )po on pr.supplier_code = po.supplier_code and pr.po_no = po.po_no and pr.item_code = po.item_code  " & vbCrLf & _
                "   where pr.receipt_cls = 'R'  " & vbCrLf
                
    sql = sql + "   union all  " & vbCrLf & _
                "   select 2 idx, 'res' status, pr.item_code, '' partno, '' child_name, pr.receipt_date date, '' factory_code, '' line_code,  " & vbCrLf & _
                "   '' parentitem_code, '' makeritem_code, '' parent_name, '' control, '' unit_desc, 0 orderpoint_qty, 0 leadtime, '' trade_code, " & vbCrLf & _
                "   0 od_qty, pr.qty od_incoming, pr.qty in_qty, 0 out_qty " & vbCrLf & _
                "   from part_receipt pr  " & vbCrLf & _
                "   left join curr_cls cc on pr.currency_code = cc.curr_cls  " & vbCrLf & _
                "   where pr.receipt_cls = 'P1'  " & vbCrLf
                
    sql = sql + "   union all  " & vbCrLf & _
                "   select 3 idx, 'ret' status, pr.item_code, '' partno, '' child_name, pr.receipt_date date, '' factory_code, '' line_code,  " & vbCrLf & _
                "   '' parentitem_code, '' partno, '' parent_name, '' control, '' unit_desc, 0 orderpoint_qty, 0 leadtime, '' trade_code, " & vbCrLf & _
                "   0 od_qty, case when po.po_no is null then null else -pr.qty end od_incoming, -pr.qty in_qty, 0 out_qty " & vbCrLf & _
                "   from part_receipt pr  " & vbCrLf & _
                "   left join curr_cls cc on pr.currency_code = cc.curr_cls  " & vbCrLf & _
                "   left join(  " & vbCrLf & _
                "       select pm.supplier_code, pm.po_no, pd.item_code  " & vbCrLf & _
                "       from purchaseorder_master pm  " & vbCrLf & _
                "       inner join purchaseorder_detail pd on pm.po_no = pd.po_no  " & vbCrLf & _
                "   )po on pr.supplier_code = po.supplier_code and pr.po_no = po.po_no and pr.item_code = po.item_code  " & vbCrLf & _
                "   where pr.receipt_cls = 'R1'  " & vbCrLf
                
    sql = sql + "   union all  " & vbCrLf & _
                "   select 3 idx, 'out' status, ps.childitem_code item_code, '' partno, '' child_name, ps.childsupply_date date, '' factory_code, '' line_code,  " & vbCrLf & _
                "   '' parentitem_code, '' makeritem_code, '' parent_name, '' control, '' unit_desc, 0 orderpoint_qty, 0 leadtime, '' trade_code, " & vbCrLf & _
                "   0 od_qty, 0 od_incoming, 0 in_qty, ps.childrequirement_qty out_qty " & vbCrLf & _
                "   from part_supply ps   " & vbCrLf & _
                "   left join curr_cls cc on ps.currency_code = cc.curr_cls  " & vbCrLf & _
                "   where ps.supply_cls <> 'S1'  " & vbCrLf
                
    sql = sql + "   union all  " & vbCrLf & _
                "   select 4 idx, 'req' status, rq.childitem_code item_code, ch.makeritem_code partno, ch.item_name child_name, production_date date, rq.factory_code, rq.line_code,  " & vbCrLf & _
                "   rq.parentitem_code, pr.makeritem_code, pr.item_name parent_name, cc.description control, uc.description unit_desc, ch.orderpoint_qty, ch.delivery_readtime leadtime, tm.trade_code, " & vbCrLf & _
                "   0 od_qty, 0 od_incoming, 0 in_qty, sum(rq.childrequirement_qty) - sum(rq.childrequirementresult_qty) - sum(rq.offchildrequirement_qty) out_qty " & vbCrLf & _
                "   from requirement rq " & vbCrLf & _
                "   inner join item_master ch on rq.childitem_code = ch.item_code " & vbCrLf & _
                "   inner join item_master pr on rq.parentitem_code = pr.item_code " & vbCrLf & _
                "   inner join trade_master tm on ch.supplier_code = tm.trade_code " & vbCrLf & _
                "   inner join control_cls cc on ch.control_cls = cc.control_cls " & vbCrLf & _
                "   inner join unit_cls uc on ch.unit_cls = uc.unit_cls " & vbCrLf & _
                "   group by rq.childitem_code, ch.makeritem_code, ch.item_name, production_date, rq.factory_code, rq.line_code,  " & vbCrLf & _
                "   rq.parentitem_code, pr.makeritem_code, pr.item_name, cc.description, uc.description, ch.orderpoint_qty, ch.delivery_readtime, tm.trade_code " & vbCrLf & _
                " ) trans on stock_begin.item_code = trans.item_code " & vbCrLf
                
    sql = sql + " left join(  " & vbCrLf & _
                "   select cast(mrp_year + '-' + mrp_month + '-01' as datetime) mrp_date, '1' mrp_set from mrp_setting   " & vbCrLf & _
                " )mrp_set on year(trans.date) = year(mrp_date) and month(trans.date) = month(mrp_date), company_profile cp   " & vbCrLf & _
                " where stock_begin.item_code in ( " & vbCrLf & _
                "   select distinct childitem_code " & vbCrLf & _
                "   from requirement rq " & vbCrLf & _
                "   where rq.production_date <= @end_date " & vbCrLf
                
    If cboFactory.Text <> strAll Then sql = sql + "   and rq.factory_code = '" & Trim(cboFactory.Text) & "' " & vbCrLf
    If cboProd.Text <> strAll Then sql = sql + "   and rq.parentitem_code = '" & Trim(cboProd.Text) & "' " & vbCrLf
    If CboMat.Text <> strAll Then sql = sql + "   and rq.childitem_code = '" & Trim(CboMat.Text) & "' " & vbCrLf
                
    sql = sql + " ) " & vbCrLf & _
                " and trans.date >= @start_date  " & vbCrLf & _
                " and trans.date <= @end_date   " & vbCrLf
                
    If cboSupplier.Text <> strAll Then sql = sql + " and trans.item_code in (select item_code from item_master where supplier_code = '" & Trim(cboSupplier.Text) & "') " & vbCrLf
    
    sql = sql + " order by trans.item_code, trans.date, trans.idx "
    
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        Set xlapp = CreateObject("Excel.Application")
        Set xlBook = xlapp.Workbooks.Add
        Set xlSheet = xlBook.ActiveSheet
    
        With xlSheet
            lngRow = 1
            .Range(strColMatCode & lngRow & ":" & strColLine & lngRow).Font.Size = 10
            .Range(strColMatCode & lngRow & ":" & strColLine & lngRow).Merge
            .Range(strColMatCode & lngRow & ":" & strColLine & lngRow).Font.Bold = True
            .Range(strColMatCode & lngRow & ":" & strColLine & lngRow).HorizontalAlignment = xlHAlignCenter
            .Range(strColMatCode & lngRow) = "Parts (Material) Alarm List"
            
            lngRow = lngRow + 2
            .Range(strColMatCode & lngRow) = "Date"
            .Range(strColMatDesc & lngRow) = ": " & Format(dtpDate.Value, "dd MMMM yyyy")
            
            .Range(strColShortDate & lngRow) = "Product Code"
            If cboProd.Text = strAll Then
                .Range(strColOrderDate & lngRow) = ": " & cboProd.Text
            Else
                .Range(strColOrderDate & lngRow) = ": " & cboProd.Text & " [" & cboProd.Column(1) & "]"
            End If
            
            lngRow = lngRow + 1
            .Range(strColMatCode & lngRow) = "Control"
            If cboControl.Text = strAll Then
                .Range(strColMatDesc & lngRow) = ": " & cboControl.Text
            Else
                .Range(strColMatDesc & lngRow) = ": " & cboControl.Text & " [" & cboControl.Column(1) & "]"
            End If
            
            .Range(strColShortDate & lngRow) = "Material Code"
            If CboMat.Text = strAll Then
                .Range(strColOrderDate & lngRow) = ": " & CboMat.Text
            Else
                .Range(strColOrderDate & lngRow) = ": " & CboMat.Text & " [" & CboMat.Column(1) & "]"
            End If
            
            lngRow = lngRow + 1
            .Range(strColMatCode & lngRow) = "Factory Code"
            If cboFactory.Text = strAll Then
                .Range(strColMatDesc & lngRow) = ": " & cboFactory.Text
            Else
                .Range(strColMatDesc & lngRow) = ": " & cboFactory.Text & " [" & cboFactory.Column(1) & "]"
            End If
            
            .Range(strColShortDate & lngRow) = "Last Calculate"
            .Range(strColOrderDate & lngRow) = ": " & strLastCalculate
            
            lngRow = lngRow + 2
            .Range(strColMatCode & lngRow & ":" & strColLine & lngRow).HorizontalAlignment = xlHAlignCenter
            .Range(strColMatCode & lngRow & ":" & strColLine & lngRow).VerticalAlignment = xlVAlignCenter
            .Range(strColMatCode & lngRow & ":" & strColLine & lngRow).Borders(xlEdgeTop).LineStyle = xlHairline
            .Range(strColMatCode & lngRow & ":" & strColLine & lngRow).Borders(xlEdgeBottom).LineStyle = xlHairline
            .Range(strColMatCode & lngRow & ":" & strColLine & lngRow).WrapText = True
    
            .Range(strColMatCode & lngRow) = "Material Code"
            .Range(strColMatDesc & lngRow) = "Material Description"
            .Range(strColUnit & lngRow) = "Unit"
            .Range(strColOrderLast & lngRow) = "On Order/Prod (Total)"
            .Range(strColOrderCurr & lngRow) = "Fix Order/Prod (Total)"
            .Range(strColOrderReceipt & lngRow) = "Receipt/Result (Total)"
            .Range(strColProdQty & lngRow) = "Order/  Forecast"
            .Range(strColStock & lngRow) = "Stock/ Requirement"
            .Range(strColShortDate & lngRow) = "Shortage Date"
            .Range(strColLeadTime & lngRow) = "Lead Time (Day)"
            .Range(strColOrderDate & lngRow) = "Estimate Order Date"
            .Range(strColControl & lngRow) = "Control"
            .Range(strColOrderQty & lngRow) = "Order/Prod Qty"
            .Range(strColSupplier & lngRow) = "Supplier"
            .Range(strColProdCode & lngRow) = "Product Code"
            .Range(strColProdDesc & lngRow) = "Product Description"
            .Range(strColFactory & lngRow) = "Factory"
            .Range(strColLine & lngRow) = "Line"
    
            .Range(strColMatCode & lngRow).ColumnWidth = 11
            .Range(strColMatDesc & lngRow).ColumnWidth = 25
            .Range(strColUnit & lngRow).ColumnWidth = 3
            .Range(strColOrderLast & lngRow).ColumnWidth = 11
            .Range(strColOrderCurr & lngRow).ColumnWidth = 11
            .Range(strColOrderReceipt & lngRow).ColumnWidth = 11
            .Range(strColProdQty & lngRow).ColumnWidth = 11
            .Range(strColStock & lngRow).ColumnWidth = 11
            .Range(strColShortDate & lngRow).ColumnWidth = 8
            .Range(strColLeadTime & lngRow).ColumnWidth = 8
            .Range(strColOrderDate & lngRow).ColumnWidth = 8
            .Range(strColControl & lngRow).ColumnWidth = 6
            .Range(strColOrderQty & lngRow).ColumnWidth = 8
            .Range(strColSupplier & lngRow).ColumnWidth = 6
            .Range(strColProdCode & lngRow).ColumnWidth = 11
            .Range(strColProdDesc & lngRow).ColumnWidth = 25
            .Range(strColFactory & lngRow).ColumnWidth = 6
            .Range(strColLine & lngRow).ColumnWidth = 6
            
            lngRow = lngRow + 1
            While adoRs.EOF = False
                If strTempItem <> Trim(adoRs.Fields("item_code") & "") Then
                    strTempItem = Trim(adoRs.Fields("item_code") & "")
                    dblStock = CDbl(Format(adoRs.Fields("po_qty") - adoRs.Fields("po_receipt") + adoRs.Fields("po_return") - adoRs.Fields("po_cancel") + adoRs.Fields("pro_qty") - adoRs.Fields("pro_result") - adoRs.Fields("pro_cancel") + _
                                            adoRs.Fields("stock_begin") + adoRs.Fields("stock_in") - adoRs.Fields("stock_out") - _
                                            adoRs.Fields("req_qty") + adoRs.Fields("req_result") + adoRs.Fields("req_off"), gs_formatQty))
                    dblOrderLast = CDbl(Format(adoRs.Fields("po_qty") - adoRs.Fields("po_receipt") + adoRs.Fields("po_return") - adoRs.Fields("po_cancel") + adoRs.Fields("pro_qty") - adoRs.Fields("pro_result") - adoRs.Fields("pro_cancel"), gs_formatQty))
                    dblOrderCurr = 0
                    dblOrderReceipt = 0
                End If
                
                dblStock = dblStock + CDbl(Format(Val(adoRs.Fields("od_qty") & "") - Val(adoRs.Fields("od_incoming") & "") + Val(adoRs.Fields("in_qty") & ""), gs_formatQty))
                If adoRs.Fields("mrp_set") & "" = "0" Then
                    If adoRs.Fields("status") <> "req" Then
                        dblStock = dblStock - CDbl(Format(Val(adoRs.Fields("out_qty") & ""), gs_formatQty))
                    End If
                Else
                    If adoRs.Fields("status") = "req" Then
                        dblStock = dblStock - CDbl(Format(Val(adoRs.Fields("out_qty") & ""), gs_formatQty))
                    End If
                End If
                
                dblOrderCurr = dblOrderCurr + CDbl(Format(Val(adoRs.Fields("od_qty") & ""), gs_formatQty))
                dblOrderReceipt = dblOrderReceipt + CDbl(Format(Val(adoRs.Fields("od_incoming") & ""), gs_formatQty))
                
                If dblStock < 0 Then
                    If Trim(adoRs.Fields("status") & "") = "req" Then
                        If strTempItemPrint <> Trim(adoRs.Fields("item_code") & "") Then
                            strTempItemPrint = Trim(adoRs.Fields("item_code") & "")
                            lngRow = lngRow + 1
                            lngRowBegin = lngRow
                        End If
                        
                        .Range(strColMatCode & lngRow) = Trim(adoRs.Fields("partno") & "")
                        If Trim(adoRs.Fields("item_code") & "") = Trim(adoRs.Fields("item_code") & "") Then
                            .Range(strColMatDesc & lngRow) = Trim(adoRs.Fields("child_name") & "")
                        Else
                            .Range(strColMatDesc & lngRow) = Trim(adoRs.Fields("partno") & "") & " " & Trim(adoRs.Fields("child_name") & "")
                        End If
                        .Range(strColUnit & lngRow) = Trim(adoRs.Fields("unit_desc") & "")
                        
                        If lngRow = lngRowBegin Then
                            .Range(strColOrderLast & lngRow) = Format(dblOrderLast, gs_formatQty)
                        Else
                            .Range(strColOrderLast & lngRow) = Format(0, gs_formatQty)
                        End If
                        
                        .Range(strColOrderCurr & lngRow) = Format(0, gs_formatQty)
                        .Range(strColOrderReceipt & lngRow) = Format(0, gs_formatQty)
                        .Range(strColProdQty & lngRow) = Format(adoRs.Fields("out_qty"), gs_formatQty)
                        .Range(strColStock & lngRow) = Format(dblStock, gs_formatQty)
                        .Range(strColShortDate & lngRow) = Format(adoRs.Fields("date"), "dd-MMM-yy")
                        .Range(strColLeadTime & lngRow) = Format(adoRs.Fields("leadtime"), gs_formatDay)
                        .Range(strColOrderDate & lngRow) = Format(adoRs.Fields("date"), "dd-MMM-yy")
                        .Range(strColControl & lngRow) = Trim(adoRs.Fields("control") & "")
                        .Range(strColOrderQty & lngRow) = "(                )"
                        .Range(strColSupplier & lngRow) = Trim(adoRs.Fields("trade_code") & "")
                        .Range(strColProdCode & lngRow) = Trim(adoRs.Fields("parentitem_code") & "")
                        If Trim(adoRs.Fields("parentitem_code") & "") = Trim(adoRs.Fields("makeritem_code") & "") Then
                            .Range(strColProdDesc & lngRow) = Trim(adoRs.Fields("parent_name") & "")
                        Else
                            .Range(strColProdDesc & lngRow) = Trim(adoRs.Fields("makeritem_code") & "") & " " & Trim(adoRs.Fields("parent_name") & "")
                        End If
                        .Range(strColFactory & lngRow) = Trim(adoRs.Fields("factory_code") & "")
                        .Range(strColLine & lngRow) = Trim(adoRs.Fields("line_code") & "")
                                            
                        lngRow = lngRow + 1
                    End If
                End If
                
                adoRs.MoveNext
                If Not adoRs.EOF Then
                    If strTempItemPrint = strTempItem Then
                        If strTempItemPrint <> Trim(adoRs.Fields("item_code") & "") Then
                            If dblStock < 0 Then
                                .Range(strColOrderCurr & lngRow - 1) = Format(dblOrderCurr, gs_formatQty)
                                .Range(strColOrderReceipt & lngRow - 1) = Format(dblOrderReceipt, gs_formatQty)
                                .Range(strColStock & lngRow - 1) = Format(dblStock, gs_formatQty)
                            Else
                                .Range(strColMatCode & lngRowBegin & ":" & strColLine & lngRow).delete Shift:=xlUp
                                lngRow = lngRow - (lngRow - (lngRowBegin - 1))
                            End If
                        End If
                    End If
                Else
                    If strTempItemPrint = strTempItem Then
                        If dblStock < 0 Then
                            .Range(strColOrderCurr & lngRow - 1) = Format(dblOrderCurr, gs_formatQty)
                            .Range(strColOrderReceipt & lngRow - 1) = Format(dblOrderReceipt, gs_formatQty)
                            .Range(strColStock & lngRow - 1) = Format(dblStock, gs_formatQty)
                        Else
                            .Range(strColMatCode & lngRowBegin & ":" & strColLine & lngRow).delete Shift:=xlUp
                            lngRow = lngRow - (lngRow - (lngRowBegin - 1))
                        End If
                    End If
                End If
            Wend
            .Range(strColMatCode & "2:" & strColLine & lngRow).Font.Size = 8
            .Range(strColMatCode & lngRow & ":" & strColLine & lngRow).Borders(xlEdgeBottom).LineStyle = xlHairline
        End With
        
        If lngRow > 8 Then
            xlapp.Visible = True
            xlSheet.Range("A8").Select
            xlapp.ActiveWindow.FreezePanes = True
        Else
            xlBook.Close False
            LblPesan.Caption = DisplayMsg("0013")
        End If
    Else
        LblPesan.Caption = DisplayMsg("0013")
    End If
    adoRs.Close
        
ErrExit:
    Me.MousePointer = vbDefault
    Set adoRs = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlapp = Nothing
    Exit Sub
errHandler:
    LblPesan.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub

Private Sub cboControl_Change()
    LblPesan.Caption = ""
    If cboControl.MatchFound Then
        lblControl.Caption = cboControl.Column(1)
    End If
End Sub

Private Sub cboFactory_Change()
    LblPesan.Caption = ""
    If cboFactory.MatchFound Then
        lblFactory.Caption = cboFactory.Column(1)
    End If
End Sub

Private Sub CboMat_Change()
    LblPesan.Caption = ""
    If CboMat.MatchFound Then
        LblMat(0).Caption = CboMat.Column(1)
        LblMat(1).Caption = CboMat.Column(2)
    End If
End Sub

Private Sub cboProd_Change()
    LblPesan.Caption = ""
    If cboProd.MatchFound Then
        lblProd(0).Caption = cboProd.Column(1)
        lblProd(1).Caption = cboProd.Column(2)
    End If
End Sub

Private Sub CboSupplier_Change()
    LblPesan.Caption = ""
    If cboSupplier.MatchFound Then
        lblSupplier.Caption = cboSupplier.Column(1)
    End If
End Sub

Private Sub cmdPrint_Click()
    If cboFactory.MatchFound = False Then
        cboFactory.SetFocus
        LblPesan = DisplayMsg(1060)
        Exit Sub
    End If
    
    If cboProd.MatchFound = False Then
        cboProd.SetFocus
        LblPesan = DisplayMsg(1024)
        Exit Sub
    End If
    
    If CboMat.MatchFound = False Then
        CboMat.SetFocus
        LblPesan = DisplayMsg(8019)
        Exit Sub
    End If
    
    If cboFactory.Text = strAll And cboProd.Text = strAll And CboMat.Text = strAll And cboSupplier.Text = strAll Then
        LblPesan.Caption = DisplayMsg("8113")
        cboFactory.SetFocus
        Exit Sub
    End If
    
    If cboControl.Text = "03" Then
        'PrintOPExcel
    Else
        PrintMRPExcel
    End If
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub Command1_Click()
    Me.MousePointer = vbHourglass
    frm_BrowseItem.getItemCode = cboProd.Text
    frm_BrowseItem.Show 1
    cboProd.Text = frm_BrowseItem.getItemCode
    cboProd.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub command2_Click()
    Me.MousePointer = vbHourglass
    frm_BrowseItem.getItemCode = CboMat.Text
    frm_BrowseItem.Show 1
    CboMat.Text = frm_BrowseItem.getItemCode
    CboMat.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblPesan.Caption = ErrMsg
    End If
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    dtpDate = Date
    AddToComboControl
    AddToComboFactory
    AddToComboProdMat
    AddToComboSupplier

    Dim adoRs As New ADODB.Recordset
    dteMRP = Date
    adoRs.Open "select min(cast(mrp_year + '-' + mrp_month + '-1' as datetime)) mrp_date from mrp_setting", Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        dteMRP = adoRs.Fields("mrp_date")
    End If
    adoRs.Close
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = True
End Sub

