VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSalesReport 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Delivery Note Detail List"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8700
   Icon            =   "FrmSalesReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   435
      TabIndex        =   12
      Top             =   1245
      Width           =   7860
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Left            =   3120
         TabIndex        =   21
         Top             =   1597
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtakhir 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   3405
         TabIndex        =   2
         Top             =   750
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   334954499
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtawal 
         Height          =   315
         Left            =   1530
         TabIndex        =   1
         Top             =   750
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   334954499
         CurrentDate     =   37798
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         TabIndex        =   20
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
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
         TabIndex        =   19
         Top             =   1680
         Width           =   1005
      End
      Begin MSForms.ComboBox cbo_group 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Top             =   1170
         Width           =   1545
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2725;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_item 
         Height          =   315
         Left            =   1530
         TabIndex        =   4
         Top             =   1590
         Width           =   1545
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2725;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         X1              =   3150
         X2              =   6060
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line3 
         X1              =   3540
         X2              =   6540
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Label lbl_group 
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_group"
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
         Left            =   3150
         TabIndex        =   18
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label lbl_item 
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_item"
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
         Left            =   3540
         TabIndex        =   17
         Top             =   1620
         Width           =   3015
      End
      Begin VB.Label lbl_trade 
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_trade"
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
         Left            =   3150
         TabIndex        =   16
         Top             =   360
         Width           =   3225
      End
      Begin VB.Line Line1 
         X1              =   3150
         X2              =   6360
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label LblCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         TabIndex        =   15
         Top             =   390
         Width           =   1575
      End
      Begin MSForms.ComboBox cbo_trade 
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   330
         Width           =   1545
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2725;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   3150
         TabIndex        =   14
         Top             =   780
         Width           =   300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DN Date"
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
         TabIndex        =   13
         Top             =   780
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xcel"
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
      Left            =   5805
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4350
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   435
      TabIndex        =   8
      Top             =   3630
      Width           =   7860
      Begin VB.Label LblErrMsg 
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
         Height          =   270
         Left            =   105
         TabIndex        =   9
         Top             =   195
         Width           =   7635
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
      Left            =   405
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4350
      Width           =   1230
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   0
      Left            =   7095
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4350
      Width           =   1200
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   6450
      TabIndex        =   11
      Top             =   465
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FDDFE3&
      Caption         =   "Delivery Note Detail List"
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
      Left            =   435
      TabIndex        =   10
      Top             =   480
      Width           =   7860
   End
End
Attribute VB_Name = "FrmSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim rsRpt As New ADODB.Recordset
Sub adtocombo()
'********Trade*************
Call up_FillCombo(cbo_trade, "Trade_Master", "trade_code, trade_name", , True)
cbo_trade.ListWidth = 280
cbo_trade.ColumnWidths = "60 pt;220 pt"
cbo_trade.ListIndex = 0
'*******Group Cls**********
Call up_FillCombo(cbo_group, "Group_Cls", , , True)
cbo_group.ListWidth = 150
cbo_group.ColumnWidths = "30 pt;120 pt"
cbo_group.ListIndex = 0
'******Item Code***********
Call up_FillCombo(cbo_item, "Item_Master", "item_code, item_name", IIf(Trim(cbo_group) <> strAll, " where group_cls= '" & Trim(cbo_group) & "' ", ""), True)
cbo_item.ListWidth = 350
cbo_item.ColumnWidths = "130 pt;160 pt"
cbo_item.ListIndex = 0
End Sub
Sub Kosong()
lbl_trade = ""
lbl_group = ""
lbl_item = ""
End Sub
Private Sub cmdAction_Click(Index As Integer)

If cbo_trade = "" Then
   LblErrMsg = DisplayMsg(1017)
   cbo_trade.SetFocus
ElseIf cbo_group = "" Then
   LblErrMsg = DisplayMsg(8081)
   cbo_group.SetFocus
ElseIf cbo_item = "" Then
   LblErrMsg = DisplayMsg(8082)
   cbo_item.SetFocus
Else
   cbo_trade = cbo_trade
   cbo_group = cbo_group
   cbo_item = cbo_item
               
   If cbo_trade.MatchFound = False Then
     LblErrMsg = DisplayMsg(4009) '
     cbo_trade.SetFocus
   ElseIf Format(dtAwal, "yyyy-MM-dd") > Format(CDate(dtAkhir), "yyyy-MM-dd") Then
     LblErrMsg = DisplayMsg(4068)
     dtAwal.SetFocus
   ElseIf cbo_group.MatchFound = False Then
     LblErrMsg = DisplayMsg(8083) '
     cbo_group.SetFocus
   ElseIf cbo_item.MatchFound = False Then
     LblErrMsg = DisplayMsg(8084) '
     cbo_item.SetFocus
   Else
     LblErrMsg = ""
     MousePointer = vbHourglass
            
     '******Trade Master********
     Dim sqltm1 As String, sqltm2 As String
     sqltm1 = vbLf & " '" & Trim(cbo_trade.Column(0)) & "' as toptm, '" & Trim(cbo_trade.Column(1)) & "' as toptmname, "
     If cbo_trade.Text = "ALL" Then
      sqltm2 = ""
     Else
      sqltm2 = vbLf & " and tm.trade_code='" & Trim(cbo_trade.Text) & "' "
     End If
            
     '******Group Cls********
     Dim sqlgr1 As String, sqlgr2 As String
     sqlgr1 = vbLf & " '" & Trim(cbo_group.Column(0)) & "' as topgr, '" & Trim(cbo_group.Column(1)) & "' as topgrdesc, "
     If cbo_group.Text = "ALL" Then
      sqlgr2 = ""
     Else
      sqlgr2 = vbLf & " and im.group_cls='" & Trim(cbo_group.Text) & "' "
     End If
           
     '******Item Master********
     Dim sqlim1 As String, sqlim2 As String
     sqlim1 = vbLf & " '" & Trim(cbo_item.Column(0)) & "' as topim, '" & Trim(cbo_item.Column(1)) & "' as topimname "
     If cbo_item.Text = "ALL" Then
      sqlim2 = ""
     Else
      sqlim2 = vbLf & " and im.item_code='" & Trim(cbo_item.Text) & "' "
     End If
     
     
     sql = "select cp.company_name, cp.address1, cp.address2, cp.city, cp.province, cp.postal_code, cp.phone1, cp.phone2, cp.fax, " & _
           vbLf & "tm.trade_code cust_cd, tm.trade_name cust_name, isnull(rc.description, '') region, tm.affiliate_cls, im.makeritem_code, " & _
           vbLf & "(case tm.country_cls when '0' then 'Domestic' else 'Overseas' end) country_cls,dm.BC40_No, dm.BC40_Date, dm.BC_Type, " & _
           vbLf & "dm.do_no, dm.do_date, im.item_code product_code, im.item_name product_name, invd.qty as size, invd.unit_cls inv_cls, " & _
           vbLf & "uc1.description inv_desc, do.qty do_qty, do.unit_cls do_cls, uc2.description do_desc, om.po_no, om.po_date, " & _
           vbLf & "invm.invoice_date, invd.invoice_no, invd.packing_no, invd.packingseq_no, im.group_cls, gr.description group_desc, " & _
           sqltm1 & sqlgr1 & sqlim1 & _
           vbLf & "from do_master dm " & _
           vbLf & "inner join delivery_order do on dm.do_no = do.do_no " & _
           vbLf & "inner join orderentry_master om on dm.cust_code = om.cust_code and do.po_no = om.po_no " & _
           vbLf & "left  join invoice_detail invd on do.do_no = invd.do_no and do.seq_no = invd.seq_no " & _
           vbLf & "left join invoice_master invm on invd.invoice_no = invm.invoice_no and dm.cust_code = invm.cust_code " & _
           vbLf & "inner join trade_master tm on dm.cust_code = tm.trade_code " & _
           vbLf & "inner join item_master im on im.item_code = do.item_code " & _
           vbLf & "--Left join  Part_Supply pr on pr.TOWarehouse_Code = dm.Cust_Code and pr.ChildItem_Code=do.Item_Code AND do.DO_No=pr.DO_No" & _
           vbLf & "left join unit_cls uc1 on invd.unit_cls = uc1.unit_cls " & _
           vbLf & "left join unit_cls uc2 on do.unit_cls = uc2.unit_cls " & _
           vbLf & "left join region_cls rc on tm.region_cls = rc.region_cls " & _
           vbLf & "left join group_cls gr on im.group_cls = gr.group_cls, company_profile cp " & _
           vbLf & "where dm.do_date >= '" & Format(dtAwal, "yyyy-MM-dd") & "' " & _
           vbLf & "and dm.do_date <= '" & Format(dtAkhir, "yyyy-MM-dd") & "' " & _
           sqltm2 & sqlgr2 & sqlim2
            
     sql = sql & vbLf & "Union All " & _
           vbLf & "select cp.company_name, cp.address1, cp.address2, cp.city, cp.province, cp.postal_code, cp.phone1, cp.phone2, cp.fax, " & _
           vbLf & "tm.trade_code cust_cd, tm.trade_name cust_name, isnull(rc.description, '') region, tm.affiliate_cls, im.makeritem_code, " & _
           vbLf & "(case tm.country_cls when '0' then 'Domestic' else 'Overseas' end) country_cls,dm.BC40_No, dm.BC40_Date, dm.BC_Type, " & _
           vbLf & "pm.packing_no do_no, pm.packing_date do_date, im.item_code product_code, im.item_name product_name, invd.qty size, invd.unit_cls inv_cls, " & _
           vbLf & "uc1.description inv_desc, pd.qty do_qty, pd.unit_cls do_cls, uc2.description do_desc, om.po_no, om.po_date, " & _
           vbLf & "invm.invoice_date, invd.invoice_no, pd.packing_no, pd.packingseq_no, im.group_cls, gr.description group_desc, " & _
           sqltm1 & sqlgr1 & sqlim1 & _
           vbLf & "from packing_master pm " & _
           vbLf & "inner join packing_detail pd on pm.packing_no = pd.packing_no " & _
           vbLf & "inner join orderentry_master om on pm.cust_code = om.cust_code and pd.order_no = om.po_no " & _
           vbLf & "inner join invoice_detail invd on pd.packing_no = invd.packing_no and pd.packingseq_no = invd.packingseq_no " & _
           vbLf & "inner join invoice_master invm on invd.invoice_no = invm.invoice_no and pm.cust_code = invm.cust_code " & _
           vbLf & "inner join trade_master tm on pm.cust_code = tm.trade_code " & _
           vbLf & "inner join item_master im on im.item_code = pd.item_code " & _
           vbLf & "inner join delivery_order do on Pd.Do_No = do.do_no  AND pd.DoSeq_No=do.Seq_No " & _
           vbLf & "inner join do_Master dm on dm.Do_No = do.do_no " & _
           vbLf & "left join unit_cls uc1 on invd.unit_cls = uc1.unit_cls " & _
           vbLf & "left join unit_cls uc2 on pd.unit_cls = uc2.unit_cls " & _
           vbLf & "left join region_cls rc on tm.region_cls = rc.region_cls " & _
           vbLf & "left join group_cls gr on im.group_cls = gr.group_cls, company_profile cp " & _
           vbLf & "where pm.packing_date >= '" & Format(dtAwal, "yyyy-MM-dd") & "' " & _
           vbLf & "and pm.packing_date <= '" & Format(dtAkhir, "yyyy-MM-dd") & "' " & _
           sqltm2 & sqlgr2 & sqlim2 & _
           vbLf & "order by country_cls, tm.trade_code, invm.invoice_date, invd.invoice_no, dm.do_date, dm.do_no, im.group_cls, im.item_code "
            
     sqlprint = sql
     
     If rsRpt.State <> adStateClosed Then rsRpt.Close
     
     rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
     If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
            
     Select Case Index
      Case 0: Call Preview
      Case 1: Call toExcel
     End Select
     
     Me.MousePointer = vbDefault
   End If
End If
End Sub

Sub Preview()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report

Dim Rpt As New FrmRpt3
          
Set report = application.OpenReport(App.path & "\Reports\SalesReport.rpt")
report.Database.Tables(1).SetDataSource rsRpt
report.FormulaFields(1).Text = "'" & Format(dtAwal.Value, "dd mmmm yyyy") & "'"
report.FormulaFields(2).Text = "'" & Format(dtAkhir.Value, "dd mmmm yyyy") & "'"
report.FormulaFields(3).Text = "'" & cbo_trade.Column(0) & "'"
report.FormulaFields(4).Text = "'" & cbo_trade.Column(1) & "'"
report.FormulaFields(5).Text = "'" & cbo_group.Column(0) & "'"
report.FormulaFields(8).Text = "'" & cbo_group.Column(1) & "'"
report.FormulaFields(6).Text = "'" & cbo_item.Column(0) & "'"
report.FormulaFields(7).Text = "'" & cbo_item.Column(1) & "'"
report.FormulaFields(14).Text = gi_decimalDigitQty
         
report.ReportTitle = Trim(Label21.Caption)
reportcode = "salesreport"
printorient = 2
            
Rpt.CRViewer1.ReportSource = report
Rpt.CRViewer1.ViewReport
Rpt.CRViewer1.Zoom (75)
        
Rpt.WindowState = 2
Rpt.Show 1
End Sub
                
Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = cbo_item.Text
 frm_BrowseItem.Show 1
 cbo_item.Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub CmdSubMenu_Click()
DoEvents
frmMainMenu.Show
DoEvents
Unload Me
End Sub

Sub toExcel()
Dim xlapp As New Excel.application
Dim Idx As Long, tempi As String, tempcust As String
Dim bolcust As Boolean, bolinv As Boolean
Dim sql_plus As String
Screen.MousePointer = vbHourglass

With xlapp

    .Workbooks.Add
    
    .Range("a2", "o2").Merge
    .Range("a2") = Trim(rsRpt!company_name)
    .Range("a3", "o3").Merge
    .Range("a3", "o3") = Trim(rsRpt!address1) & " " & Trim(rsRpt!address2) & " " & Trim(rsRpt!City) & " " & Trim(rsRpt!Province) & " " & Trim(rsRpt!postal_code)
    .Range("a4", "o4").Merge
    .Range("a4") = "Phone: " & Trim(rsRpt!phone1) & " " & Trim(rsRpt!phone2) & " Fax: " & Trim(rsRpt!fax)
    
    .Range("a6", "o6").Merge
    .Range("a6") = Trim(Label21.Caption)
    .Range("b6") = ""
    .Range("a6").HorizontalAlignment = xlLeft
    
    '.Range("a7") = rsRpt.Fields(46).Name
    .Range("a7") = "Invoice Date"
    .Range("b7", "o7").Merge
    .Range("b7") = ": " & Format(dtAwal.Value, "[$-409]dd-mmm-yyyy;@") & " to " & Format(dtAkhir.Value, "[$-409]d-mmm-yyyy;@")
    .Range("b7").HorizontalAlignment = xlLeft
    .Range("a8") = "Trade Code"
    .Range("b8", "o8").Merge
    .Range("b8") = ": " & cbo_trade.Column(0) & " / " & cbo_trade.Column(1)
    .Range("a9") = "Group Cls"
    .Range("b9", "o9").Merge
    .Range("b9") = ": " & cbo_group.Column(0) & " / " & cbo_group.Column(1)
    .Range("a10") = "Item Code"
    .Range("b10", "o10").Merge
    .Range("b10") = ": " & cbo_item.Column(0) & " / " & cbo_item.Column(1)
    
    Idx = 12
    
    Dim ls_OverseasCls As String
    ls_OverseasCls = ""
    
    Do While Not rsRpt.EOF
        If Idx = 12 Then
            .Range("a" & Idx) = "Cust Cd"
            .Range("b" & Idx) = "Cust Name"
            .Range("c" & Idx) = "Region"
            .Range("d" & Idx) = "Affiliation"
            .Range("e" & Idx) = "DO NO"
            .Range("f" & Idx) = "DN Date"
            .Range("g" & Idx) = "Product Code"
            .Range("h" & Idx) = "Model"
            .Range("i" & Idx) = "Product Name"
'            .Range("j" & Idx) = "Size"
'            .Range("k" & Idx) = "Unit Cls"
            .Range("j" & Idx) = "DO Qty"
            .Range("k" & Idx) = "Unit Cls"
            .Range("l" & Idx) = "SI/PO No"
            .Range("m" & Idx) = "SI/PO Date"
            '=========== Calips 2013-02-11 =================================
            .Range("n" & Idx) = "BC Type"
            .Range("o" & Idx) = "BC No"
            .Range("p" & Idx) = "BC Date"
            '===============================================================
            .Range("q" & Idx) = "Invoice No"
            .Range("r" & Idx) = "Invoice Date"
            .Range("a" & Idx, "r" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range("a" & Idx, "r" & Idx).Borders(xlEdgeBottom).LineStyle = xlDouble
            Idx = Idx + 1
        End If
        
        Idx = Idx
        'Content
        
        If ls_OverseasCls <> Trim(rsRpt!country_cls) Then
            ls_OverseasCls = Trim(rsRpt!country_cls)
            .Range("a" & Idx) = Trim(rsRpt!country_cls)
            .Range("a" & Idx).Columns.Font.Bold = True
            Idx = Idx + 1
        End If
                        
        .Range("a" & Idx) = Trim(rsRpt!cust_cd)
        .Range("b" & Idx) = Trim(rsRpt!Cust_Name)
        .Range("c" & Idx) = Trim(rsRpt!Region)
        .Range("d" & Idx) = IIf(rsRpt!affiliate_cls = 1, "Yes", "No")
        .Range("e" & Idx) = Trim(rsRpt!do_no)
        .Range("f" & Idx) = rsRpt!do_date
        .Range("g" & Idx) = "'" & Trim(rsRpt!product_code)
        .Range("h" & Idx) = "'" & Trim(rsRpt!MakerItem_Code)
        .Range("i" & Idx) = Trim(rsRpt!Product_Name)
'        .Range("j" & Idx) = rsrpt!Size
'        .Range("k" & Idx) = Trim(rsrpt!InvdDesc)
        .Range("j" & Idx) = rsRpt!do_qty
        .Range("k" & Idx) = Trim(rsRpt!DO_Desc)
        .Range("l" & Idx) = "'" & Trim(rsRpt!po_no)
        .Range("m" & Idx) = rsRpt!po_date
        '=========== Calips 2013-02-11 =================================
        .Range("n" & Idx) = Trim(rsRpt!BC_Type)
        .Range("o" & Idx) = Trim(rsRpt!BC40_No)
        .Range("p" & Idx) = rsRpt!BC40_Date
        '===============================================================
        .Range("q" & Idx) = Trim(rsRpt!Invoice_No)
        .Range("r" & Idx) = rsRpt!Invoice_Date
         
        Idx = Idx + 1
        rsRpt.MoveNext
    Loop
       
    .Range("a1", "r" & Idx).Columns.Font.Name = "Arial"
    .Range("a1", "r" & Idx).Columns.Font.Size = 8
    
       
    .Range("a2", "r2").Columns.Font.Name = "Arial"
    .Range("a2", "r2").Columns.Font.Size = "10"
    .Range("a2", "r2").Columns.Font.Bold = True
    .Range("a2", "r4").HorizontalAlignment = xlCenter
    .Range("a6", "b6").Columns.Font.Bold = True
   
    .Range("a12:a" & Idx).HorizontalAlignment = xlLeft
    .Range("f12:f" & Idx).HorizontalAlignment = xlLeft
    .Range("m12:m" & Idx).HorizontalAlignment = xlLeft
    .Range("n12:m" & Idx).HorizontalAlignment = xlLeft
    .Range("o12:o" & Idx).HorizontalAlignment = xlLeft
    .Range("j12:j" & Idx).NumberFormat = gs_formatQty
    .Range("f12:f" & Idx).NumberFormat = "[$-409]dd-mmm-yyyy;@"
    .Range("m12:m" & Idx).NumberFormat = "[$-409]dd-mmm-yyyy;@"
    '=========== Calips 2013-02-11 =============================
    .Range("p12:p" & Idx).NumberFormat = "[$-409]dd-mmm-yyyy;@"
    '===========================================================
    .Range("r12:r" & Idx).NumberFormat = "[$-409]dd-mmm-yyyy;@"
       
'    .ActiveSheet.PageSetup.PaperSize = xlPaperA4
'    .ActiveSheet.PageSetup.Orientation = 2
'    .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
'    .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.4)
    .Range("a:r").Columns.AutoFit
    .WindowState = xlMaximized
    .Visible = True
End With
Screen.MousePointer = vbDefault
End Sub


Private Sub dtAwal_Change()
    LblErrMsg = ""
    If Format(dtAwal, "yyyy-MM-dd") > _
        Format(CDate(dtAkhir), "yyyy-MM-dd") Then _
    LblErrMsg = DisplayMsg(4068): Exit Sub
End Sub

Private Sub dtAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub dtAkhir_Change()
    LblErrMsg = ""
    If Format(dtAwal, "yyyy-MM-dd") > _
        Format(CDate(dtAkhir), "yyyy-MM-dd") Then _
        LblErrMsg = DisplayMsg(4066): Exit Sub
End Sub

Private Sub dtAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub
Private Sub cbo_group_Change()
lbl_group = ""
LblErrMsg = ""
End Sub

Private Sub cbo_group_Click()
cbo_group = cbo_group
    If cbo_group.MatchFound Then
        lbl_group = cbo_group.Column(1)
        LblErrMsg = ""
    Else
        lbl_group = ""
        LblErrMsg = DisplayMsg(8083)
        cbo_group.SetFocus
    End If
Call up_FillCombo(cbo_item, "Item_Master", "item_code, item_name", IIf(Trim(cbo_group) <> strAll, " where group_cls= '" & Trim(cbo_group) & "' ", ""), True)
cbo_item.ListWidth = 250
cbo_item.ColumnWidths = "90 pt;160 pt"
cbo_item.ListIndex = 0
End Sub
Private Sub cbo_group_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call cbo_group_Click
End Sub
Private Sub cbo_trade_Change()
lbl_trade = ""
LblErrMsg = ""
End Sub
Private Sub cbo_trade_Click()
cbo_trade = cbo_trade
    If cbo_trade.MatchFound Then
        lbl_trade = cbo_trade.Column(1)
        LblErrMsg = ""
    Else
        lbl_trade = ""
        LblErrMsg = DisplayMsg(4009)
    End If
End Sub

Private Sub cbo_trade_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call cbo_trade_Click
End Sub
Private Sub cbo_item_Change()
lbl_item = ""
LblErrMsg = ""
End Sub

Private Sub cbo_item_Click()
cbo_item = cbo_item
    If cbo_item.MatchFound Then
        lbl_item = cbo_item.Column(1)
        LblErrMsg = ""
    Else
        lbl_item = ""
        LblErrMsg = DisplayMsg(8084)
    End If
End Sub

Private Sub cbo_item_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call cbo_item_Click
End Sub
Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

Call Kosong
Call adtocombo

dtAwal = Now
dtAkhir = Now
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
'**************
