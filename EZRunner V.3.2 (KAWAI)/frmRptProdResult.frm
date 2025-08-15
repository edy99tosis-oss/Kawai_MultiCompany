VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptProdResult 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Production Result Report"
   ClientHeight    =   5040
   ClientLeft      =   1020
   ClientTop       =   3990
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRptProdResult.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xcel"
      Height          =   375
      Left            =   6818
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4155
      Width           =   1185
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   353
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4155
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2040
      Left            =   353
      TabIndex        =   11
      Top             =   1290
      Width           =   8895
      Begin MSComCtl2.DTPicker dtAwal 
         Height          =   330
         Left            =   1860
         TabIndex        =   2
         Top             =   1335
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   293535747
         CurrentDate     =   37860
      End
      Begin MSComCtl2.DTPicker dtAkhir 
         Height          =   330
         Left            =   3795
         TabIndex        =   3
         Top             =   1335
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   293535747
         CurrentDate     =   37799
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   3495
         X2              =   8670
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To WH Code"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   17
         Top             =   915
         Width           =   1065
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   3495
         TabIndex        =   16
         Top             =   915
         Width           =   960
      End
      Begin MSForms.ComboBox cbo 
         Height          =   330
         Index           =   1
         Left            =   1860
         TabIndex        =   1
         Top             =   855
         Width           =   1530
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2699;582"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo 
         Height          =   330
         Index           =   0
         Left            =   1860
         TabIndex        =   0
         Top             =   375
         Width           =   1530
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2699;582"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   195
         Index           =   2
         Left            =   3495
         TabIndex        =   15
         Top             =   1410
         Width           =   165
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   3495
         TabIndex        =   14
         Top             =   435
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory Code"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   13
         Top             =   435
         Width           =   1140
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   3495
         X2              =   8655
         Y1              =   675
         Y2              =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Production Date"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   12
         Top             =   1410
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   353
      TabIndex        =   9
      Top             =   3390
      Width           =   8895
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
         Width           =   8670
      End
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   8063
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4155
      Width           =   1185
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   7410
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   450
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Production Result Report"
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
      Left            =   3375
      TabIndex        =   8
      Top             =   450
      Width           =   2865
   End
End
Attribute VB_Name = "frmRptProdResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String, i As Integer
Dim QtyResult As Double, GrandResult As Double
Dim QtyLoss As Double, GrandLoss As Double

'******** Combo Factory Code **********
Sub isiCboCust() 'Isi Combo Dealer CD dr Customer Master
Dim RsCust As New ADODB.Recordset 'Data Customer

With Cbo(0)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    '******** Ambil dr Customer Master utk Combo Dealer CD
    sql = "select Trade_Code,Trade_Name from Trade_Master " & _
        "where trade_code in (select distinct manufacture_code from manufacture_line) " & _
        "order by Trade_Code"
    Set RsCust = Db.Execute(sql)
    
    i = 0
    Do While Not (RsCust.EOF)
        .AddItem ""
        .List(i, 0) = Trim(RsCust(0))
        .List(i, 1) = Trim(RsCust(1))
        i = i + 1
        RsCust.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 350
    .ColumnWidths = "50 pt;300 pt"
    
    Set RsCust = Nothing
End With
End Sub

Sub isiCboWH()
Dim rscbo As New ADODB.Recordset

With Cbo(1)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    sql = "select * from (select wh_code,wh_name,stockControl_cls from warehouse_master " & _
        "union all " & _
        "select distinct(manufacture_line.manufacture_code)wh_code, trade_name wh_name, stockControl_Cls = '01'from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code)tbJ order by wh_code"
    Set rscbo = Db.Execute(sql)
    
    .AddItem ""
    .List(0, 0) = strAll
    .List(0, 1) = strAll
    i = 1
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo(0))
        .List(i, 1) = Trim(rscbo(1))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 250
    .ColumnWidths = "50 pt;200 pt"
    .ListIndex = 0
    
    Set rscbo = Nothing
End With
End Sub

Sub TotalCur(xl As Excel.application, Row As Long, Col As String, Col2 As String, coltitle As String)
    With xl
        .Range("a" & Row, "l" & Row).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(coltitle & Row) = "Sub Total"
        .Range(Col & Row) = Format(QtyResult, gs_formatQty)
        .Range(Col2 & Row) = Format(QtyLoss, gs_formatQty)
    End With
End Sub

Sub GrandTotal(xl As Excel.application, Row As Long, Col As String, Col2 As String, coltitle As String)
    With xl
        .Range("a" & Row, "l" & Row).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(coltitle & Row) = "Grand Total"
        .Range(Col & Row) = Format(GrandResult, gs_formatQty)
        .Range(Col2 & Row) = Format(GrandLoss, gs_formatQty)
    End With
End Sub

Private Sub CmdExcel_Click()
    Dim xlapp As New Excel.application
    Dim rsCek As New Recordset, Idx As Long, tempi As String, tempcls As String
    Dim bolcls As Boolean, bolcur As Boolean
    Dim rsCompany As New Recordset
    
    Dim tempLine As String
    
    MousePointer = vbHourglass

    
        sql = "select * from " & _
                vbLf & " (select rtrim(a.supplier_code) supplier_code,rtrim(a.warehouse_code) wh_code,rtrim(c.Trade_Name) Trade_Name, rtrim(po_no)po_no,rtrim(a.item_code) item_code, rtrim(b.makeritem_code) makeritem_code, rtrim(b.Item_Name) Item_Name, receipt_Date, " & _
                vbLf & " isnull(sum(a.qty),0) as Result , 0 LossReject, rtrim(suratjalan_no)suratjalan_no,rtrim(remarks) remarks, d.schedule_date , isnull(d.qty,0) as plann,a.unit_cls,(select rtrim(description) from unit_cls uc   where uc.unit_cls=a.unit_cls) unit_desc, d.seq_No, " & _
                vbLf & " rtrim(d.SerialNoFrom) PSerialFrom,rtrim(d.SerialNoTo) PSerialTo," & _
                vbLf & " rtrim(b.wh_code) item_wh " & _
                vbLf & " from part_Receipt a,item_master b, Trade_Master c, daily_production d " & _
                vbLf & " Where A.Item_Code = B.Item_Code and a.Supplier_Code = c.Trade_Code and receipt_cls = 'P1' " & _
                vbLf & " and ProductionResult_Cls = '1' And a.DailySeq_No = d.Seq_No " & _
                vbLf & " group by a.supplier_code,a.warehouse_code,po_no,a.item_code,b.makeritem_code,b.Item_Name,c.Trade_name, receipt_Date,suratjalan_no,remarks, d.schedule_date,a.unit_cls, d.qty, d.seq_No, b.wh_code, d.serialNoFrom,d.SerialNoTo " & _
                vbLf & " Union " & _
                vbLf & " select rtrim(a.supplier_Code)supplier_Code,rtrim(a.warehouse_code) wh_code,rtrim(c.Trade_Name)Trade_Name,rtrim(po_no)po_no,rtrim(a.item_code)item_code, rtrim(b.makeritem_code)makeritem_code, rtrim(b.Item_Name)Item_Name, receipt_Date, " & _
                vbLf & " 0 result, -(isnull(sum(a.qty),0))  as LossReject, rtrim(suratjalan_no)suratjalan_no,rtrim(remarks)remarks, d.schedule_date , isnull(d.qty,0) as plann,a.unit_cls ,(select rtrim(description) from unit_cls uc   where uc.unit_cls=a.unit_cls) unit_desc, d.seq_No, " & _
                vbLf & " rtrim(d.SerialNoFrom) PSerialFrom,rtrim(d.SerialNoTo) PSerialTo," & _
                vbLf & " rtrim(b.wh_code) item_wh " & _
                vbLf & " from part_Receipt a,item_master b, Trade_Master c, daily_production d " & _
                vbLf & " Where A.Item_Code = B.Item_Code and a.Supplier_Code = c.Trade_Code and receipt_cls <> 'P1' " & _
                vbLf & " and ProductionResult_Cls = '1' And a.DailySeq_No = d.Seq_No " & _
                vbLf & " group by a.supplier_code,a.warehouse_code,po_no,a.item_code,b.makeritem_code,b.Item_Name,c.Trade_Name,receipt_Date,suratjalan_no,remarks, d.schedule_date,a.unit_cls,d.qty, d.seq_No, b.wh_code, d.serialNoFrom,d.SerialNoTo) dt " & _
                vbLf & " where supplier_code= '" & Cbo(0) & "' " & _
                vbLf & " and receipt_Date > = '" & Format(dtAwal, "yyyy-MM-dd") & "'" & _
                vbLf & " and receipt_Date < = '" & Format(dtAkhir, "yyyy-MM-dd") & "'"

            If Cbo(1) <> strAll Then sql = sql & " and wh_code = '" & Cbo(1) & "' "
            sql = sql & " order by supplier_code, po_no, receipt_date, item_wh, item_code "
    
    If rsCek.State <> adStateClosed Then rsCek.Close
    
    rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic
    
    If rsCek.EOF Then
        LblErrMsg.Caption = DisplayMsg(4006)
    Else
            
        With xlapp
            
            .Workbooks.Add
            
            .Range("a2", "k2").Merge
            .Range("a2") = "Production Result Report"

            .Range("a4") = "Factory Code"
            .Range("b4") = ": " & Trim(rsCek!Supplier_Code)
            .Range("c4", "g4").Merge
            .Range("c4") = "Factory Name :  " & rsCek!trade_name
            .Range("b5", "d5").Merge
            .Range("a5") = "Product Date From"
            .Range("b5") = ": " & Format(dtAwal, "dd MMMM YYYY") & " to " & Format(dtAkhir, "dd MMMM YYYY")

            Idx = 7
            tempcls = ""
            tempLine = ""
            QtyResult = 0
            QtyLoss = 0
            GrandResult = 0
            GrandLoss = 0
            
            Do While Not rsCek.EOF
                If Idx <> 7 And tempcls <> Trim(rsCek!Item_Code) Then
                    Call TotalCur(xlapp, Idx, "j", "k", "i")
                    QtyResult = 0
                    QtyLoss = 0
                    Idx = Idx + 2
                End If
                
                If Idx = 7 Then
                    .Range("a" & Idx) = "Production Date"
                    .Range("b" & Idx) = "Plan Date"
                    .Range("c" & Idx) = "Product Code"
                    .Range("d" & Idx) = "Part Number"
                    .Range("e" & Idx) = "Description"
                    .Range("f" & Idx) = "Lot No."
                ' Add Plan Serial Number - 20090212
                    .Range("g" & Idx) = "Serial From"
                    .Range("h" & Idx) = "Serial To"
                ' --
                    .Range("i" & Idx) = "Plan"
                    .Range("j" & Idx) = "Result"
                    .Range("k" & Idx) = "Loss / Reject"
                    .Range("l" & Idx) = "Unit"
                    .Range("m" & Idx) = "Remarks"
                    .Range("a" & Idx, "m" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Range("a" & Idx, "m" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Idx = Idx + 1
                End If
                
                If tempLine <> rsCek!po_no Then
                    Idx = Idx + 1
                    .Range("a" & Idx) = "Machine No"
                    .Range("b" & Idx) = ": " & Trim(rsCek!po_no)
                    Idx = Idx + 1
                End If
                tempLine = rsCek!po_no
                
                'Content
                .Range("a" & Idx).HorizontalAlignment = xlCenter
                .Range("a" & Idx) = Format(rsCek!Receipt_Date, "DD-MMM-YYYY")
                .Range("b" & Idx).HorizontalAlignment = xlCenter
                .Range("b" & Idx) = Format(rsCek!schedule_date, "DD-MMM-YYYY")
                .Range("c" & Idx) = Trim(rsCek!Item_Code)
                .Range("d" & Idx) = Trim(rsCek!MakerItem_Code)
                .Range("e" & Idx) = Trim(rsCek!item_name)
                .Range("f" & Idx) = "'" & Trim(rsCek!SuratJalan_No)
                
                .Range("g" & Idx) = Trim(rsCek!PSerialFrom)
                .Range("h" & Idx) = Trim(rsCek!PSerialTo)
                
                .Range("i" & Idx) = Format(rsCek!plann, gs_formatQty)
                .Range("j" & Idx) = Format(rsCek!result, gs_formatQty)
                .Range("k" & Idx) = Format(rsCek!LossReject, gs_formatQty)
                .Range("l" & Idx) = Trim(rsCek!Unit_Desc)
                .Range("m" & Idx) = Trim(rsCek!Remarks)
                
                Idx = Idx + 1
                tempcls = Trim(rsCek!Item_Code)

                QtyResult = QtyResult + rsCek!result
                QtyLoss = QtyLoss + rsCek!LossReject
                GrandResult = GrandResult + rsCek!result
                GrandLoss = GrandLoss + rsCek!LossReject
                rsCek.MoveNext
            Loop

            Call TotalCur(xlapp, Idx, "j", "k", "i")
            Call GrandTotal(xlapp, Idx + 2, "j", "k", "i")
            Idx = Idx + 3
            
            .Range("a1", "m" & Idx + 3).Columns.Font.Name = "Arial"
            .Range("a1", "m" & Idx + 3).Columns.Font.Size = 8

            .Range("a2", "m2").Columns.Font.Name = "Arial"
            .Range("a2", "m2").Columns.Font.Size = "10"
            .Range("a2", "m2").Columns.Font.Bold = True
            .Range("a2", "m2").HorizontalAlignment = xlCenter
            .Range("c8", "c8").HorizontalAlignment = xlCenter

            .ActiveSheet.PageSetup.Orientation = 2
            .WindowState = xlMaximized
            .Range("a1", "m" & Idx + 3).Columns.AutoFit
            
            .Visible = True
        
        End With
    End If
    MousePointer = vbDefault

End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    lblNm(0) = ""
    lblNm(1) = ""
    dtAwal = Now
    dtAkhir = Now
    Call isiCboCust
    Call isiCboWH
End Sub

Private Sub cbo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbo_Click(Index)
End Sub

Public Sub cbo_Click(Index As Integer)
    Cbo(Index) = Cbo(Index)
    If Cbo(Index).MatchFound Then
        lblNm(Index) = Cbo(Index).Column(1)
        LblErrMsg = ""
    Else
        lblNm(Index) = ""
        LblErrMsg = DisplayMsg(4016)
    End If
End Sub

Private Sub cbo_Change(Index As Integer)
    lblNm(Index) = ""
    LblErrMsg = ""
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

Private Sub cmdReport_Click()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3

    Me.MousePointer = vbHourglass
    
    If Cbo(0) = "" Then
        LblErrMsg = DisplayMsg(1040)
        Cbo(0).SetFocus
    Else
        Cbo(0) = Cbo(0)
        If Cbo(0).MatchFound = False Then
            LblErrMsg = DisplayMsg(4016)
            Cbo(0).SetFocus
        Else
            LblErrMsg = ""
            
            sql = "select * from " & _
                    vbLf & " (select rtrim(a.supplier_code) supplier_code,rtrim(a.warehouse_code) wh_code,rtrim(c.Trade_Name) Trade_Name, rtrim(po_no)po_no,rtrim(a.item_code) item_code, rtrim(b.makeritem_code) makeritem_code, rtrim(b.Item_Name) Item_Name, receipt_Date, " & _
                    vbLf & " isnull(sum(a.qty),0) as Result , 0 LossReject, rtrim(suratjalan_no)suratjalan_no,rtrim(remarks) remarks, d.schedule_date , isnull(d.qty,0) as plann,a.unit_cls,(select rtrim(description) from unit_cls uc   where uc.unit_cls=a.unit_cls) unit_desc, d.seq_No, " & _
                    vbLf & " rtrim(d.SerialNoFrom) PSerialFrom,rtrim(d.SerialNoTo) PSerialTo," & _
                    vbLf & " rtrim(b.wh_code) item_wh " & _
                    vbLf & " from part_Receipt a,item_master b, Trade_Master c, daily_production d " & _
                    vbLf & " Where A.Item_Code = B.Item_Code and a.Supplier_Code = c.Trade_Code and receipt_cls = 'P1' " & _
                    vbLf & " and ProductionResult_Cls = '1' And a.DailySeq_No = d.Seq_No " & _
                    vbLf & " group by a.supplier_code,a.warehouse_code,po_no,a.item_code,b.makeritem_code,b.Item_Name,c.Trade_name, receipt_Date,suratjalan_no,remarks, d.schedule_date,a.unit_cls, d.qty, d.seq_No, b.wh_code, d.serialNoFrom,d.SerialNoTo " & _
                    vbLf & " Union " & _
                    vbLf & " select rtrim(a.supplier_Code)supplier_Code,rtrim(a.warehouse_code) wh_code,rtrim(c.Trade_Name)Trade_Name,rtrim(po_no)po_no,rtrim(a.item_code)item_code, rtrim(b.makeritem_code)makeritem_code, rtrim(b.Item_Name)Item_Name, receipt_Date, " & _
                    vbLf & " 0 result, -(isnull(sum(a.qty),0))  as LossReject, rtrim(suratjalan_no)suratjalan_no,rtrim(remarks)remarks, d.schedule_date , isnull(d.qty,0) as plann,a.unit_cls ,(select rtrim(description) from unit_cls uc   where uc.unit_cls=a.unit_cls) unit_desc, d.seq_No, " & _
                    vbLf & " rtrim(d.SerialNoFrom) PSerialFrom,rtrim(d.SerialNoTo) PSerialTo," & _
                    vbLf & " rtrim(b.wh_code) item_wh " & _
                    vbLf & " from part_Receipt a,item_master b, Trade_Master c, daily_production d " & _
                    vbLf & " Where A.Item_Code = B.Item_Code and a.Supplier_Code = c.Trade_Code and receipt_cls <> 'P1' " & _
                    vbLf & " and ProductionResult_Cls = '1' And a.DailySeq_No = d.Seq_No " & _
                    vbLf & " group by a.supplier_code,a.warehouse_code,po_no,a.item_code,b.makeritem_code,b.Item_Name,c.Trade_Name,receipt_Date,suratjalan_no,remarks, d.schedule_date,a.unit_cls,d.qty, d.seq_No, b.wh_code, d.serialNoFrom,d.SerialNoTo) dt " & _
                    vbLf & " where supplier_code= '" & Cbo(0) & "' " & _
                    vbLf & " and receipt_Date > = '" & Format(dtAwal, "yyyy-MM-dd") & "'" & _
                    vbLf & " and receipt_Date < = '" & Format(dtAkhir, "yyyy-MM-dd") & "'"
            
            If Cbo(1) <> strAll Then sql = sql & " and wh_code = '" & Cbo(1) & "' "
            sql = sql & " order by supplier_code, po_no, receipt_date, item_wh, item_code "

            
            Set rsRpt = Db.Execute(sql)
            
            If rsRpt.EOF Then
                LblErrMsg.Caption = DisplayMsg(4006)
            Else
                sqlprint = sql
                reportcode = "ProdResultByFactory"
                printorient = 2
                
                Set report = application.OpenReport(App.path & "\Reports\rptProdResult.rpt")
                report.Database.Tables(1).SetDataSource rsRpt
                tglAwalRptPrint = Format(dtAwal, "dd MMM yyyy")
                tglAkhirRptPrint = Format(dtAkhir, "dd MMM yyyy")
                
                report.FormulaFields(1).Text = "'" & Format(dtAwal, "dd MMM yyyy") & "'"
                report.FormulaFields(2).Text = "'" & Format(dtAkhir, "dd MMM yyyy") & "'"
                
                '#####################################################################
                '# Qty Digit and decimal
                report.FormulaFields(7).Text = "" & gi_decimalDigitQty & ""
                report.FormulaFields(8).Text = "" & gi_decimalDigitQty & ""
                '#####################################################################
                
                Rpt.CRViewer1.ReportSource = report
                Rpt.CRViewer1.ViewReport
                Rpt.CRViewer1.Zoom 1
                
                Rpt.WindowState = 2
                Rpt.Show 1
            End If
            Set rsRpt = Nothing
        End If
    End If
    
    Me.MousePointer = vbDefault
    
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


