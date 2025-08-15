VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmValuationPriceReportCls 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Valuation Price Report Per Classification"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmValuationPriceReportCls.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   593
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5010
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2790
      Left            =   593
      TabIndex        =   11
      Top             =   1410
      Width           =   8865
      Begin MSComCtl2.DTPicker dtAwal 
         Height          =   330
         Left            =   2520
         TabIndex        =   4
         Top             =   2130
         Width           =   1665
         _ExtentX        =   2937
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
         Format          =   141230083
         CurrentDate     =   38808
      End
      Begin MSComCtl2.DTPicker dtAkhir 
         Height          =   330
         Left            =   4560
         TabIndex        =   5
         Top             =   2130
         Width           =   1665
         _ExtentX        =   2937
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
         Format          =   141230083
         CurrentDate     =   39174
      End
      Begin VB.Label lbl_supply 
         BackColor       =   &H00FDDFE3&
         Caption         =   "lbl_supply"
         Height          =   255
         Left            =   4575
         TabIndex        =   20
         Top             =   1710
         Width           =   3495
      End
      Begin VB.Label lbl_towarehouse 
         BackColor       =   &H00FDDFE3&
         Caption         =   "lbl_towarehouse"
         Height          =   255
         Left            =   4575
         TabIndex        =   19
         Top             =   1260
         Width           =   3495
      End
      Begin VB.Line Line8 
         Index           =   2
         X1              =   4575
         X2              =   8145
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   4575
         X2              =   8145
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label lbl_frwarehouse 
         BackColor       =   &H00FDDFE3&
         Caption         =   "lbl_frwarehouse"
         Height          =   255
         Left            =   4575
         TabIndex        =   18
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Date"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   17
         Top             =   2205
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   195
         Index           =   5
         Left            =   4290
         TabIndex        =   16
         Top             =   2190
         Width           =   165
      End
      Begin MSForms.ComboBox cbo_supply 
         Height          =   345
         Left            =   2520
         TabIndex        =   3
         Top             =   1680
         Width           =   1125
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "1984;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_towarehouse 
         Height          =   345
         Left            =   2520
         TabIndex        =   2
         Top             =   1230
         Width           =   1845
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3254;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_frwarehouse 
         Height          =   345
         Left            =   2520
         TabIndex        =   1
         Top             =   810
         Width           =   1845
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3254;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_finish_good 
         Height          =   345
         Left            =   2520
         TabIndex        =   0
         Top             =   330
         Width           =   2475
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4366;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Cls"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   15
         Top             =   1755
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Warehouse Code"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   14
         Top             =   1305
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finish Good Part Cls"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   13
         Top             =   405
         Width           =   1725
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   4575
         X2              =   8145
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Warehouse Code"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   12
         Top             =   885
         Width           =   1965
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   593
      TabIndex        =   9
      Top             =   4275
      Width           =   8865
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
         TabIndex        =   10
         Top             =   195
         Width           =   8640
      End
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   8273
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5010
      Width           =   1185
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   7650
      TabIndex        =   21
      Top             =   225
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valuation Price Report Per Classification"
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
      Left            =   2760
      TabIndex        =   8
      Top             =   675
      Width           =   4545
   End
End
Attribute VB_Name = "FrmValuationPriceReportCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim sql_warehouse As String
Dim rs_warehouse As New ADODB.Recordset

Sub adtocombo()
    
'**********Finish Good Part Cls****************
    cbo_finish_good.clear
    cbo_finish_good.columnCount = 2

    cbo_finish_good.AddItem
    cbo_finish_good.List(0, 0) = strAll
    cbo_finish_good.List(0, 1) = strAll
    cbo_finish_good.AddItem
    cbo_finish_good.List(1, 0) = "01"
    cbo_finish_good.List(1, 1) = "Finish Goods"
    cbo_finish_good.AddItem
    cbo_finish_good.List(2, 0) = "02"
    cbo_finish_good.List(2, 1) = "Parts/WIP/Material"

    cbo_finish_good.ListWidth = 120
    cbo_finish_good.ColumnWidths = "30 pt ; 90 pt "
    cbo_finish_good.ListIndex = 0
    cbo_finish_good.Text = cbo_finish_good.List(0, 0)
    
'**********WAREHOUSE**************
 'From
    If Not (rs_warehouse.BOF And rs_warehouse.EOF) Then
    Dim i As Long
     With cbo_frwarehouse
      .clear
      .columnCount = 2
      .ColumnWidths = "80pt;200pt"
      .ListWidth = 280
              
      .AddItem
      .List(0, 0) = strAll
      .List(0, 1) = strAll
      
      i = 1
      Do While Not rs_warehouse.EOF
       .AddItem
       .List(i, 0) = Trim(rs_warehouse!wh_code)
       .List(i, 1) = Trim(rs_warehouse!WH_Name)
       i = i + 1
       rs_warehouse.MoveNext
      Loop
     End With
  cbo_frwarehouse.Text = cbo_frwarehouse.List(0, 0)
    
 'To
     With cbo_towarehouse
      .clear
      .columnCount = 2
      .ColumnWidths = "80pt;200pt"
      .ListWidth = 280
              
       .AddItem
       .List(0, 0) = strAll
       .List(0, 1) = strAll
      i = 1
      
      rs_warehouse.Requery
      Do While Not rs_warehouse.EOF
       .AddItem
       .List(i, 0) = Trim(rs_warehouse!wh_code)
       .List(i, 1) = Trim(rs_warehouse!WH_Name)
       i = i + 1
       rs_warehouse.MoveNext
      Loop
     End With
    
    End If
    cbo_towarehouse.Text = cbo_towarehouse.List(0, 0)
    
'*************SUPPLY CLS*****************************
    cbo_supply.clear
    cbo_supply.columnCount = 2
    'cbo_supply.TextColumn = 1
    cbo_supply.AddItem
    cbo_supply.List(0, 0) = strAll
    cbo_supply.List(0, 1) = strAll
    cbo_supply.AddItem
    cbo_supply.List(1, 0) = "S1"
    cbo_supply.List(1, 1) = "Supply"
    cbo_supply.AddItem
    cbo_supply.List(2, 0) = "S"
    cbo_supply.List(2, 1) = "Consumption"
    cbo_supply.AddItem
    cbo_supply.List(3, 0) = "L"
    cbo_supply.List(3, 1) = "Loss"
    cbo_supply.AddItem
    cbo_supply.List(4, 0) = "RJ"
    cbo_supply.List(4, 1) = "Reject"
    cbo_supply.ColumnWidths = "25 pt; 75 pt"
    cbo_supply.ListWidth = 100
    cbo_supply.Text = cbo_supply.List(0, 0)
 
 
'==============================================================
    
End Sub

Private Sub cbo_frwarehouse_Change()
lbl_frwarehouse = ""
LblErrMsg = ""
End Sub

Private Sub cbo_frwarehouse_Click()
cbo_frwarehouse = cbo_frwarehouse
    If cbo_frwarehouse.MatchFound Then
        lbl_frwarehouse = cbo_frwarehouse.Column(1)
        LblErrMsg = ""
    Else
        lbl_frwarehouse = ""
        LblErrMsg = DisplayMsg(4018)
    End If
End Sub

Private Sub cbo_frwarehouse_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call cbo_frwarehouse_Click
End Sub

Private Sub cbo_towarehouse_Change()
lbl_towarehouse = ""
LblErrMsg = ""
End Sub

Private Sub cbo_towarehouse_Click()
cbo_towarehouse = cbo_towarehouse
    If cbo_towarehouse.MatchFound Then
        lbl_towarehouse = cbo_towarehouse.Column(1)
        LblErrMsg = ""
    Else
        lbl_towarehouse = ""
        LblErrMsg = DisplayMsg(4018)
    End If
End Sub

Private Sub cbo_towarehouse_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call cbo_towarehouse_Click
End Sub
Private Sub cbo_supply_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbo_supply_Click
End Sub

Public Sub cbo_supply_Click()
    cbo_supply = cbo_supply
    If cbo_supply.MatchFound Then
        lbl_supply = cbo_supply.Column(1)
        LblErrMsg = ""
    Else
        lbl_supply = ""
        LblErrMsg = DisplayMsg(8070)
    End If
End Sub
Private Sub cbo_supply_Change()
    lbl_supply = ""
    LblErrMsg = ""
End Sub
Sub Kosong()

  lbl_frwarehouse.Caption = ""
  lbl_towarehouse.Caption = ""
  lbl_supply.Caption = ""
  
End Sub
Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    sql_warehouse = "select WH_Code, WH_Name from Warehouse_Master " & _
                    "Union " & _
                    "select Trade_Code as WH_Code, Trade_Name as WH_Name from Trade_Master " & _
                    "where trade_code in (select distinct manufacture_code from manufacture_line) " & _
                    "order by WH_Code "
    If rs_warehouse.State <> adStateClosed Then rs_warehouse.Close
    rs_warehouse.Open sql_warehouse, Db, adOpenKeyset, adLockOptimistic

    
    dtAwal = Now
    dtAkhir = Now
    Call Kosong
    Call adtocombo
End Sub
Private Sub dtAwal_Change()
    LblErrMsg = ""
    If Format(dtAwal, "yyyy-MM-dd") > Format(CDate(dtAkhir), "yyyy-MM-dd") Then
        LblErrMsg = DisplayMsg(8055)
        Exit Sub
    End If
End Sub

Private Sub dtAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub dtAkhir_Change()
    LblErrMsg = ""
    If Format(dtAwal, "yyyy-MM-dd") > Format(CDate(dtAkhir), "yyyy-MM-dd") Then
        LblErrMsg = DisplayMsg(4066) '"End Date can't be smaller than " & Format(dtawal, "dd MMM yyyy")
        Exit Sub
    End If
End Sub

Private Sub dtAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub cmdReport_Click()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3

       
    If cbo_finish_good = "" Then
      LblErrMsg = DisplayMsg(8077)
      cbo_finish_good.SetFocus
    ElseIf cbo_frwarehouse = "" Then
        LblErrMsg = DisplayMsg(8067)
        cbo_frwarehouse.SetFocus
    ElseIf cbo_towarehouse = "" Then
        LblErrMsg = DisplayMsg(8069)
        cbo_towarehouse.SetFocus
    ElseIf cbo_supply = "" Then
        LblErrMsg = DisplayMsg(4052)
        cbo_supply.SetFocus
    Else
        cbo_finish_good = cbo_finish_good
        cbo_frwarehouse = cbo_frwarehouse
        cbo_towarehouse = cbo_towarehouse
        cbo_supply = cbo_supply
        
        If cbo_finish_good.MatchFound = False Then
            LblErrMsg = DisplayMsg(8078)
            cbo_finish_good.SetFocus
        ElseIf cbo_frwarehouse.MatchFound = False Then
            LblErrMsg = DisplayMsg(4018)
            cbo_frwarehouse.SetFocus
        ElseIf cbo_towarehouse.MatchFound = False Then
            LblErrMsg = DisplayMsg(4018)
            cbo_towarehouse.SetFocus
        ElseIf cbo_supply.MatchFound = False Then
            LblErrMsg = DisplayMsg(8070)
            cbo_supply.SetFocus
        Else
            LblErrMsg = ""
            MousePointer = vbHourglass
            
            '******Finish Good*********
            Dim sqlfg1 As String, sqlfg2 As String
            If Trim(cbo_finish_good.Text) = strAll Then
             sqlfg1 = " 'ALL' as topfg, "
             sqlfg2 = ""
            Else
             sqlfg1 = " im.finishgoodpart_cls as topfg, "
             sqlfg2 = " and im.finishgoodpart_cls='" & Trim(cbo_finish_good.Text) & "' "
            End If
            
            '******From Warehouse*********
            Dim sqlfrwh1 As String, sqlfrwh2 As String
            sqlfrwh1 = " '" & Trim(cbo_frwarehouse.Column(0)) & "' as topfrwarehouse_code, '" & Trim(cbo_frwarehouse.Column(1)) & "' as topfrwarehouse_name, "
            If Trim(cbo_frwarehouse.Text) = strAll Then
             sqlfrwh2 = ""
            Else
             sqlfrwh2 = " and ps.fromwarehouse_code='" & Trim(cbo_frwarehouse.Text) & "' "
            End If
            
            '******To Warehouse*********
            Dim sqltowh1 As String, sqltowh2 As String
            sqltowh1 = " '" & Trim(cbo_towarehouse.Column(0)) & "' as toptowarehouse_code, '" & Trim(cbo_towarehouse.Column(1)) & "' as toptowarehouse_name, "
            If Trim(cbo_towarehouse.Text) = strAll Then
             sqltowh2 = ""
            Else
             sqltowh2 = " and ps.towarehouse_code='" & Trim(cbo_towarehouse.Text) & "' "
            End If
            
            '******Supply Cls*************
            Dim sqlsupplycls1 As String, sqlsupplycls2 As String
            sqlsupplycls1 = " '" & Trim(cbo_supply.Column(0)) & "' as topsupply_cls, '" & Trim(cbo_supply.Column(1)) & "' as topsupply_clsname, "
            If Trim(cbo_supply.Text) = strAll Then
             sqlsupplycls2 = ""
            Else
             sqlsupplycls2 = " and ps.supply_cls='" & Trim(cbo_supply.Text) & "' "
            End If
            
          sql = "select im.finishgoodpart_cls, " & sqlfg1 & _
                sqlfrwh1 & _
                "rtrim(ps.fromwarehouse_code) fromwarehouse_code, " & _
                "(case when wm.wh_name is null then rtrim(tm.trade_name) else rtrim(wm.wh_name) end) fromwarehouse_name, " & _
                sqltowh1 & _
                "rtrim(ps.towarehouse_code) towarehouse_code, " & _
                "(case when wm2.wh_name is null then rtrim(tm2.trade_name) else rtrim(wm2.wh_name) end) towarehouse_name, " & _
                sqlsupplycls1 & " ps.supply_cls, " & _
               "ps.childsupply_date, rtrim(ps.childitem_code) childitem_code, rtrim(im.MakerItem_Code) part_number, " & _
               "rtrim(im.item_name) childitem_name, " & _
               "rtrim(ps.lot_no) lot_no, isnull(ps.childrequirement_qty,0) childqty, ps.childunit_cls, rtrim(uc.description) description, " & _
               "rtrim(ps.supplyrec_no) supplyrec_no, isnull(ps.do_no,'') do_no, rtrim(isnull(ps.remarks,'')) remarks, " & _
               "price_inventory = (select inventory_price from inventory_price where item_code = ps.childitem_code and inventory_year = year(ps.childsupply_date) and inventory_month = month(ps.childsupply_date)) " & _
               "From part_supply ps " & _
               "inner join item_master im on im.item_code=ps.childitem_code " & _
               "left outer join warehouse_master wm on wm.wh_code=ps.fromwarehouse_code " & _
               "left outer join trade_master tm on tm.trade_code=ps.fromwarehouse_code " & _
               "left outer join warehouse_master wm2 on wm2.wh_code=ps.towarehouse_code " & _
               "left outer join trade_master tm2 on tm2.trade_code=ps.towarehouse_code " & _
               "left outer join unit_cls uc on ps.childunit_cls = uc.unit_cls " & _
               "where ps.supply_cls in ('S','S1','L','RJ') " & _
               "and (childsupply_date>='" & Format(dtAwal.Value, "yyyy-mm-dd") & "' and childsupply_date<='" & Format(dtAkhir.Value, "yyyy-mm-dd") & "') " & _
               sqlfg2 & sqlfrwh2 & sqltowh2 & sqlsupplycls2 & _
               "order by fromwarehouse_code, childsupply_date, childitem_code, towarehouse_code, supply_cls "
            
          sqlprint = sql
          If rsRpt.State <> adStateClosed Then rsRpt.Close
          rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
         
         If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
    
         Set report = application.OpenReport(App.path & "\Reports\rptSupplyListValue.rpt")
         report.Database.Tables(1).SetDataSource rsRpt
         report.FormulaFields(2).Text = "'" & Format(dtAwal.Value, "dd MMM yyyy") & "'"
         report.FormulaFields(3).Text = "'" & Format(dtAkhir.Value, "dd MMM yyyy") & "'"
         report.FormulaFields(5).Text = "'" & cbo_towarehouse.Column(0) & "'"
         report.FormulaFields(6).Text = "'" & cbo_towarehouse.Column(1) & "'"
         report.FormulaFields(7).Text = "'" & cbo_supply.Column(0) & "'"
         report.FormulaFields(8).Text = "'" & cbo_supply.Column(1) & "'"
         report.FormulaFields(9).Text = gi_decimalDigitQty
         report.FormulaFields(10).Text = "'" & cbo_frwarehouse.Column(0) & "'"
         report.FormulaFields(11).Text = "'" & cbo_frwarehouse.Column(1) & "'"
    
    report.ReportTitle = "Valuation Price Report Per Classification"
    printorient = 2 'Landscape
    reportcode = "SupplyListValue"
        
    Rpt.CRViewer1.ReportSource = report
    Rpt.CRViewer1.ViewReport
    Rpt.CRViewer1.Zoom (75)
        
    Rpt.WindowState = 2
    Rpt.Show 1
    Set rsRpt = Nothing
    Me.MousePointer = vbDefault
    End If
    
  End If
End Sub

'************ Unload **********
Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
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


Private Sub Form_Unload(Cancel As Integer)
    If rs_warehouse.State <> adStateClosed Then
        rs_warehouse.Close
        Set rs_warehouse = Nothing
    End If
End Sub
