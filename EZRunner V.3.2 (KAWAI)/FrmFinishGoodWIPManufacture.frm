VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFinishGoodWIPManufactureStandard 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Finish Good / WIP Manufacture"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9525
   Icon            =   "FrmFinishGoodWIPManufacture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   9525
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
      Height          =   1770
      Left            =   180
      TabIndex        =   8
      Top             =   1305
      Width           =   9135
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Left            =   3780
         TabIndex        =   16
         Top             =   787
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtawal 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1200
         Width           =   1620
         _ExtentX        =   2858
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
         CurrentDate     =   39173
      End
      Begin VB.Line Line3 
         X1              =   6570
         X2              =   8820
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblitem 
         BackStyle       =   0  'Transparent
         Caption         =   "lblitem"
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
         Left            =   6555
         TabIndex        =   15
         Top             =   840
         Width           =   2310
      End
      Begin VB.Line Line1 
         X1              =   4125
         X2              =   8055
         Y1              =   645
         Y2              =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   405
         Width           =   525
      End
      Begin MSForms.ComboBox cbogroup 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   345
         Width           =   2100
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3704;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboitem 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   780
         Width           =   2550
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4498;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   195
         TabIndex        =   12
         Top             =   1260
         Width           =   405
      End
      Begin VB.Line Line2 
         X1              =   4140
         X2              =   6315
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblgroup 
         BackStyle       =   0  'Transparent
         Caption         =   "lblgroup"
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
         Left            =   4140
         TabIndex        =   11
         Top             =   405
         Width           =   3885
      End
      Begin VB.Label lblitem 
         BackStyle       =   0  'Transparent
         Caption         =   "lblitem"
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
         Left            =   4140
         TabIndex        =   10
         Top             =   840
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   840
         Width           =   915
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
      Left            =   8115
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3975
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   180
      TabIndex        =   5
      Top             =   3210
      Width           =   9135
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
         Left            =   150
         TabIndex        =   6
         Top             =   210
         Width           =   8850
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
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3945
      Width           =   1230
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   7470
      TabIndex        =   13
      Top             =   255
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FDDFE3&
      Caption         =   "Finish Good / WIP Manufacture"
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
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   9135
   End
End
Attribute VB_Name = "FrmFinishGoodWIPManufactureStandard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim SqlExchRate As String
Dim sqlDuty As String
Dim Idx As Long ' index utk row of excel
Dim MainProcess() As String
Dim Curr(3) As String
Dim ERT(3) As Double 'ExchangeRate
Dim Index As Integer 'index utk main process
Dim countchild As Integer
Dim countrow As Integer
Dim main As Boolean
 
Dim xlColLevel As String 'a
Dim xlColPartCode As String 'b
Dim xlColDescription As String 'c
Dim xlColQty As String 'd
Dim xlColUnitCls As String 'e
Dim xlColJPYPartCost As String 'f
Dim xlColUSDPartCost As String 'g
Dim xlColIDRPartCost As String 'h
Dim xlColEURPartCost As String 'i
Dim xlColSupCode As String 'j
Dim xlColUSDProcessCost As String 'k
Dim xlColIDRProcessCost As String 'l
Dim xlColJPYDutyPrice As String 'm
Dim xlColUSDDutyPrice As String 'n
Dim xlColIDRDutyPrice As String 'o
Dim xlColAmount As String 'p
Dim xlColAmountUSD As String 'q

Private Sub cboGroup_Change()
    If cboGroup.MatchFound Then
        lblgroup = cboGroup.Column(1)
        LblErrMsg = ""
    Else
        lblgroup = ""
        LblErrMsg = ""
        If cboGroup.Text <> "" Then LblErrMsg = DisplayMsg(8083)
    End If
    adtocomboitem
End Sub
Private Sub CboItem_Change()
    If cboitem.MatchFound Then
        lblitem(0) = cboitem.Column(1)
        lblitem(1) = cboitem.Column(2)
        LblErrMsg = ""
    Else
        lblitem(0) = ""
        lblitem(1) = ""
        LblErrMsg = ""
        If cboitem.Text <> "" Then LblErrMsg = DisplayMsg(8084)
    End If
End Sub

Sub adtocombo()
'*******Group Cls**********
Call up_FillCombo(cboGroup, "Group_Cls", , , True)
cboGroup.ListWidth = 150
cboGroup.ColumnWidths = "30 pt;120 pt"
cboGroup.ListIndex = 0
End Sub

Private Sub adtocomboitem()
Dim adoRs As New ADODB.Recordset
With cboitem
    
    .clear
    .columnCount = 3
    
    sql = "select item_code, makeritem_code, item_name from item_master " & _
        "where production_cls = '01' and finishgoodpart_cls = '01' and use_endday >= convert(char(8), getdate(), 112) "
    If cboGroup.ListIndex <> 0 Then sql = sql & "and group_cls = '" & Trim(cboGroup.Text) & "'"

    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    While Not adoRs.EOF
        .AddItem ""
        .List(.ListCount - 1, 0) = Trim(adoRs.Fields("item_code"))
        .List(.ListCount - 1, 1) = Trim(adoRs.Fields("makeritem_code"))
        .List(.ListCount - 1, 2) = Trim(adoRs.Fields("item_name"))
        adoRs.MoveNext
    Wend
    adoRs.Close
    
    .ListWidth = 410
    .ColumnWidths = "130 pt;130 pt;150 pt"
    
End With
Set adoRs = Nothing
End Sub

Private Sub cmdAction_Click()
If cboGroup = "" Then
   LblErrMsg = DisplayMsg(8081)
   cboGroup.SetFocus
ElseIf cboitem = "" Then
   LblErrMsg = DisplayMsg(8082)
   cboitem.SetFocus
Else
   If cboGroup.MatchFound = False Then
     LblErrMsg = DisplayMsg(8083)
     cboGroup.SetFocus
   ElseIf cboitem.MatchFound = False Then
     LblErrMsg = DisplayMsg(8084)
     cboitem.SetFocus
   Else
     On Error GoTo errHandler
     LblErrMsg = ""
     MousePointer = vbHourglass
     
     '*********************Export = Ngga Kena Pajak / Local = Kena Pajak********************
     If Left((UCase(Trim(cboitem.Text))), 1) = "L" Then
        sqlDuty = "tax =  (Case (select count(bm2.item_code) from bom_master bm2 where bm2.parent_itemcode = bm.Item_Code) " & _
        Chr(10) & "when 0 then Isnull(hm.tax,0) else 0 end) "
     Else
        sqlDuty = "tax = 0 "
     End If
     
     Dim xlapp As Excel.application
     Dim xlBook As Excel.Workbook
     Dim xlSheet As Excel.Worksheet
     
     '*********************Currency Description utk header & ExchRateTabel In Rupiah********************
     Dim rr As Integer
     Dim SqlExchRateTbl As String
     Dim rsExchRateTbl As New ADODB.Recordset
          
     SqlExchRateTbl = "Select ber.currency_code as CurrCode, cc.description as CurrDesc, isnull(ber.Exch0" & dtAwal.Month & ",0) as ERT " & _
                      "from Book_ExchangeRate ber, Company_Profile cp, Curr_Cls cc " & _
                      "Where ber.Term_Cls = cp.ValuationPrice_ExchTerm " & _
                      "and ber.Currency_Code = cc.Curr_Cls " & _
                      "and ber.Exch_Year = '" & dtAwal.Year & "' " & _
                      "and ber.currency_code <> '05' " & _
                      "order by CurrCode asc "
     
     Set rsExchRateTbl = Db.Execute(SqlExchRateTbl)
     
     'Currency Description utk header ditampung ke array
     rr = 0
     Do While Not (rsExchRateTbl.EOF)
      Curr(rr) = Trim(rsExchRateTbl!CurrDesc)
      ERT(rr) = Format((rsExchRateTbl!ERT), gs_formatExchangeRate)
      rsExchRateTbl.MoveNext
      rr = rr + 1
     '0 -> JPY
     '1 -> USD
     '2 -> IDR
     '3 -> EUR
     Loop
  
     Set xlapp = CreateObject("Excel.Application")
     Set xlBook = xlapp.Workbooks.Add
     Set xlSheet = xlBook.Worksheets("Sheet1")
        
     xlColLevel = "a"
     xlColPartCode = "b"
     xlColDescription = "c"
     xlColQty = "d"
     xlColUnitCls = "e"
     xlColJPYPartCost = "f"
     xlColUSDPartCost = "g"
     xlColIDRPartCost = "h"
     xlColEURPartCost = "i"
     xlColSupCode = "j"
     xlColUSDProcessCost = "k"
     xlColIDRProcessCost = "l"
     xlColJPYDutyPrice = "m"
     xlColUSDDutyPrice = "n"
     xlColIDRDutyPrice = "o"
     xlColAmount = "p"
     xlColAmountUSD = "q"
        
     With xlSheet
      '*******************Header***************************
      .Range(xlColLevel & "2", xlColAmountUSD & "2").Merge
      .Range(xlColLevel & "2") = "FINISH GOOD / WIP MANUFACTURE"
      .Range(xlColLevel & "3", xlColAmountUSD & "3").Merge
      .Range(xlColLevel & "3", xlColAmountUSD & "3") = "(STANDARD)"
      
      Dim rsmakergroup As New ADODB.Recordset
      Dim sqlmakergroup As String
      
      sqlmakergroup = "select im.makeritem_code, im.item_name, gc.description " & _
                      "from item_master im Left Join group_cls gc on im.group_cls = gc.group_cls" & _
                      "where 'A'='A' " & _
                      "and im.item_code = '" & cboitem.Text & "'"
                      
      Set rsmakergroup = Db.Execute(sqlmakergroup)
                      
      .Range(xlColLevel & "5", xlColPartCode & "5").Merge
      .Range(xlColLevel & "5") = "Product Code"
      .Range(xlColDescription & "5", xlColAmountUSD & "5").Merge
      .Range(xlColDescription & "5") = ": " & cboitem.Text & " / " & Trim(rsmakergroup!item_name)
      .Range(xlColDescription & "5").horizontalAlignment = xlLeft
      .Range(xlColLevel & "6", xlColPartCode & "6").Merge
      .Range(xlColLevel & "6") = "Model"
      .Range(xlColDescription & "6", xlColAmountUSD & "6").Merge
      .Range(xlColDescription & "6") = ": " & IIf(IsNull(rsmakergroup!MakerItem_Code), "", Trim(rsmakergroup!MakerItem_Code))
      .Range(xlColDescription & "6").horizontalAlignment = xlLeft
      .Range(xlColLevel & "7", xlColPartCode & "7").Merge
      .Range(xlColLevel & "7") = "Category"
      .Range(xlColDescription & "7", xlColAmountUSD & "7").Merge
      .Range(xlColDescription & "7") = ": " & IIf(IsNull(rsmakergroup!Description), "", Trim(rsmakergroup!Description))
      .Range(xlColDescription & "7").horizontalAlignment = xlLeft
      
      .Range(xlColLevel & "2", xlColAmountUSD & "3").Columns.Font.Size = "10"
      .Range(xlColLevel & "2", xlColAmountUSD & "3").Columns.Font.Bold = True
      .Range(xlColLevel & "2", xlColAmountUSD & "3").horizontalAlignment = xlCenter
     '******************Fieldname************************
     Idx = 9
     'Idx ke-9 dan 10
     .Range(xlColLevel & Idx, xlColLevel & Idx + 1).Merge
     .Range(xlColLevel & Idx) = "LEVEL"
     .Range(xlColPartCode & Idx, xlColPartCode & Idx + 1).Merge
     .Range(xlColPartCode & Idx) = "PART CODE"
     .Range(xlColDescription & Idx, xlColDescription & Idx + 1).Merge
     .Range(xlColDescription & Idx) = "DESCRIPTION"
     .Range(xlColQty & Idx, xlColQty & Idx + 1).Merge
     .Range(xlColQty & Idx) = "QTY"
     
     .Range(xlColUnitCls & Idx, xlColUnitCls & Idx + 1).Merge
     .Range(xlColUnitCls & Idx) = "UNIT" & Chr(10) & "CLS"
     
     .Range(xlColJPYPartCost & Idx, xlColEURPartCost & Idx).Merge
     .Range(xlColJPYPartCost & Idx) = "PART COST"
     .Range(xlColJPYPartCost & Idx + 1) = Curr(0)
     .Range(xlColUSDPartCost & Idx + 1) = Curr(1)
     .Range(xlColIDRPartCost & Idx + 1) = Curr(2)
     .Range(xlColEURPartCost & Idx + 1) = Curr(3)
          
     .Range(xlColSupCode & Idx, xlColSupCode & Idx + 1).Merge
     .Range(xlColSupCode & Idx) = "SUP." & Chr(10) & "CODE"
     
     .Range(xlColUSDProcessCost & Idx, xlColIDRProcessCost & Idx).Merge
     .Range(xlColUSDProcessCost & Idx) = "PROCESS COST"
     .Range(xlColUSDProcessCost & Idx + 1) = Curr(1)
     .Range(xlColIDRProcessCost & Idx + 1) = Curr(2)
          
     .Range(xlColJPYDutyPrice & Idx, xlColIDRDutyPrice & Idx).Merge
     .Range(xlColJPYDutyPrice & Idx) = "DUTY PRICE"
     .Range(xlColJPYDutyPrice & Idx + 1) = Curr(0)
     .Range(xlColUSDDutyPrice & Idx + 1) = Curr(1)
     .Range(xlColIDRDutyPrice & Idx + 1) = Curr(2)
          
     .Range(xlColAmount & Idx, xlColAmount & Idx + 1).Merge
     .Range(xlColAmount & Idx) = "AMOUNT" & Chr(10) & "(Rp)"
     
     .Range(xlColAmountUSD & Idx, xlColAmountUSD & Idx + 1).Merge
     .Range(xlColAmountUSD & Idx) = "AMOUNT" & Chr(10) & "(USD)"
     
     .Range(xlColLevel & Idx & ":" & xlColAmountUSD & Idx + 1).Font.Bold = True
     
     '********************** border ******************************
     .Range(xlColLevel & Idx, xlColAmountUSD & Idx + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
     .Range(xlColLevel & Idx, xlColAmountUSD & Idx + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
     .Range(xlColLevel & Idx, xlColAmountUSD & Idx + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
     .Range(xlColLevel & Idx, xlColAmountUSD & Idx + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
     .Range(xlColLevel & Idx, xlColAmountUSD & Idx + 1).Borders(xlInsideVertical).LineStyle = xlContinuous
     .Range(xlColLevel & Idx, xlColAmountUSD & Idx + 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
     '********************** alignment a-n *************************
     .Range(xlColLevel & Idx, xlColAmountUSD & Idx + 1).horizontalAlignment = xlCenter
     .Range(xlColLevel & Idx, xlColAmountUSD & Idx + 1).verticalAlignment = xlCenter
     '**************************************************************
     Idx = Idx + 2
     Index = 0
     '**********************IsiExcel***************************
     Call IsiExcel(cboitem.Text, 0, 1, xlapp, xlBook, xlSheet)
     '*********************************************************
    
     '*********************Border utk IsiExcel**************************
     If Idx > 11 Then
      With .Range(xlColLevel & "11", xlColAmountUSD & Idx - 1)
           
       With .Borders(xlEdgeLeft)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
       End With
       With .Borders(xlEdgeTop)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
       End With
       With .Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
       End With
       With .Borders(xlEdgeRight)
         .LineStyle = xlContinuous
         .Weight = xlThin
        .ColorIndex = xlAutomatic
       End With
       With .Borders(xlInsideVertical)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
       End With
       If Idx > 12 Then
        With .Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
        End With
       End If
      End With
     End If
    '***********************************************************
    
     '***************alignment utk UnitCLS**************************
     .Range(xlColUnitCls & "11:" & xlColUnitCls & Idx - 1).horizontalAlignment = xlCenter
     '**************************************************************
     
     '*********************Orderby countchild*******************
     If .Range("r11").Value <> "" Then
      .Range(xlColLevel & "11:r" & Idx - 1).Sort Key1:=.Range("r11"), Order1:=xlAscending, Header:=xlGuess, _
      OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
      'DataOption1:=xlSortNormal
     End If
     xlSheet.Columns("r:r").ClearContents

     '*****************Qty/StandardTime*****************
     .Range(xlColQty & "11:" & xlColQty & Idx).NumberFormat = gs_formatQtyBOM
     '*******************PartCost***********************
     .Range(xlColJPYPartCost & "11:" & xlColJPYPartCost & Idx - 1).NumberFormat = gs_formatPrice
     .Range(xlColUSDPartCost & "11:" & xlColUSDPartCost & Idx - 1).NumberFormat = gs_formatPrice
     .Range(xlColIDRPartCost & "11:" & xlColIDRPartCost & Idx - 1).NumberFormat = gs_formatPriceIDR
     .Range(xlColEURPartCost & "11:" & xlColEURPartCost & Idx - 1).NumberFormat = gs_formatPrice
     'Amount of PartCost
     .Range(xlColJPYPartCost & Idx & ":" & xlColUSDPartCost & Idx).NumberFormat = gs_formatAmount
     .Range(xlColIDRPartCost & Idx).NumberFormat = gs_formatAmountIDR
     .Range(xlColEURPartCost & Idx).NumberFormat = gs_formatAmount
     '*****************ProcessCost**********************
     .Range(xlColUSDProcessCost & "11:" & xlColUSDProcessCost & Idx - 1).NumberFormat = gs_formatPrice
     .Range(xlColIDRProcessCost & "11:" & xlColIDRProcessCost & Idx - 1).NumberFormat = gs_formatPriceIDR
     'Amount of ProcessCost
     .Range(xlColUSDProcessCost & Idx).NumberFormat = gs_formatAmount
     .Range(xlColIDRProcessCost & Idx).NumberFormat = gs_formatAmountIDR
     '*****************DutyPrice************************
     .Range(xlColJPYDutyPrice & "11:" & xlColJPYDutyPrice & Idx - 1).NumberFormat = gs_formatPrice
     .Range(xlColUSDDutyPrice & "11:" & xlColUSDDutyPrice & Idx - 1).NumberFormat = gs_formatPrice
     .Range(xlColIDRDutyPrice & "11:" & xlColIDRDutyPrice & Idx - 1).NumberFormat = gs_formatPriceIDR
     'Amount of DutyPrice
     .Range(xlColJPYDutyPrice & Idx & ":" & xlColUSDDutyPrice & Idx).NumberFormat = gs_formatAmount
     .Range(xlColIDRDutyPrice & Idx).NumberFormat = gs_formatAmountIDR
     '********************Amount*************************
     .Range(xlColAmount & "11:" & xlColAmount & Idx).NumberFormat = gs_formatAmountIDR
     .Range(xlColAmountUSD & "11:" & xlColAmountUSD & Idx).NumberFormat = gs_formatAmount
     '***************************Sub Total*******************************
     Dim IdxIsi As Integer
     IdxIsi = 11
     
     .Range(xlColUnitCls & Idx) = "TOTAL"
     .Range(xlColUnitCls & Idx).Font.Bold = True
     
     Do While IdxIsi < Idx
      If .Range(xlColJPYPartCost & IdxIsi).Value <> Empty Then
       .Range(xlColJPYPartCost & Idx).Value = .Range(xlColJPYPartCost & Idx).Value + Format((.Range(xlColJPYPartCost & IdxIsi).Value * .Range(xlColQty & IdxIsi).Value), gs_formatAmount)
      End If
      
      If .Range(xlColUSDPartCost & IdxIsi).Value <> Empty Then
       .Range(xlColUSDPartCost & Idx).Value = .Range(xlColUSDPartCost & Idx).Value + Format((.Range(xlColUSDPartCost & IdxIsi).Value * .Range(xlColQty & IdxIsi).Value), gs_formatAmount)
      End If
      
      If .Range(xlColIDRPartCost & IdxIsi).Value <> Empty Then
       .Range(xlColIDRPartCost & Idx).Value = .Range(xlColIDRPartCost & Idx).Value + Format((.Range(xlColIDRPartCost & IdxIsi).Value * .Range(xlColQty & IdxIsi).Value), gs_formatAmountIDR)
      End If
      
      If .Range(xlColEURPartCost & IdxIsi).Value <> Empty Then
       .Range(xlColEURPartCost & Idx).Value = .Range(xlColEURPartCost & Idx).Value + Format((.Range(xlColEURPartCost & IdxIsi).Value * .Range(xlColQty & IdxIsi).Value), gs_formatAmount)
      End If
      
      If .Range(xlColUSDProcessCost & IdxIsi).Value <> Empty Then
       .Range(xlColUSDProcessCost & Idx).Value = .Range(xlColUSDProcessCost & Idx).Value + Format((.Range(xlColUSDProcessCost & IdxIsi).Value * .Range(xlColQty & IdxIsi).Value), gs_formatAmount)
      End If
      
      If .Range(xlColIDRProcessCost & IdxIsi).Value <> Empty Then
       .Range(xlColIDRProcessCost & Idx).Value = .Range(xlColIDRProcessCost & Idx).Value + Format((.Range(xlColIDRProcessCost & IdxIsi).Value * .Range(xlColQty & IdxIsi).Value), gs_formatAmountIDR)
      End If
      
      If .Range(xlColJPYDutyPrice & IdxIsi).Value <> Empty Then
       .Range(xlColJPYDutyPrice & Idx).Value = .Range(xlColJPYDutyPrice & Idx).Value + Format((.Range(xlColJPYDutyPrice & IdxIsi).Value * .Range(xlColQty & IdxIsi).Value), gs_formatAmount)
      End If
      
      If .Range(xlColUSDDutyPrice & IdxIsi).Value <> Empty Then
       .Range(xlColUSDDutyPrice & Idx).Value = .Range(xlColUSDDutyPrice & Idx).Value + Format((.Range(xlColUSDDutyPrice & IdxIsi).Value * .Range(xlColQty & IdxIsi).Value), gs_formatAmount)
      End If
      
      If .Range(xlColIDRDutyPrice & IdxIsi).Value <> Empty Then
       .Range(xlColIDRDutyPrice & Idx).Value = .Range(xlColIDRDutyPrice & Idx).Value + Format((.Range(xlColIDRDutyPrice & IdxIsi).Value * .Range(xlColQty & IdxIsi).Value), gs_formatAmountIDR)
      End If
      
      If .Range(xlColAmount & IdxIsi).Value <> Empty Then
       .Range(xlColAmount & Idx).Value = .Range(xlColAmount & Idx).Value + .Range(xlColAmount & IdxIsi).Value
      End If
       
      If .Range(xlColAmountUSD & IdxIsi).Value <> Empty Then
       .Range(xlColAmountUSD & Idx).Value = .Range(xlColAmountUSD & Idx).Value + .Range(xlColAmountUSD & IdxIsi).Value
      End If
            
      IdxIsi = IdxIsi + 1
     
     Loop
     
     'SubTotal Borders
     With .Range(xlColJPYPartCost & Idx & ":" & xlColEURPartCost & Idx & "," & xlColUSDProcessCost & Idx & ":" & xlColAmountUSD & Idx)
      With .Borders(xlEdgeLeft)
       .LineStyle = xlContinuous
       .Weight = xlThin
       .ColorIndex = xlAutomatic
      End With
    
      With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
        
      With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
    
      With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
    End With
    
     '***********************ExchangeRate (In Rupiah)**********************
     .Range(xlColPartCode & Idx + 3 & ":" & xlColDescription & Idx + 3).Merge
     .Range(xlColPartCode & Idx + 3) = "Exchange Rate (In Rupiah)"
     .Range(xlColPartCode & Idx + 3).horizontalAlignment = xlCenter
     .Range(xlColPartCode & Idx + 3).Font.Bold = True
     
     For rr = 0 To 3
      .Range(xlColPartCode & Idx + 4 + rr) = Curr(rr)
      .Range(xlColDescription & Idx + 4 + rr) = ERT(rr)
      .Range(xlColDescription & Idx + 4 + rr).NumberFormat = gs_formatExchangeRate
     Next rr
     
     'ExchangeRate (In Rupiah) Borders
     With .Range(xlColPartCode & Idx + 3 & ":" & xlColDescription & Idx + 3 + rr)
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
      
     End With
    
     '******************************NOTE/ SUMMARY***********************************
     .Range(xlColJPYPartCost & Idx + 3 & ":" & xlColEURPartCost & Idx + 3).Merge
     .Range(xlColJPYPartCost & Idx + 3).horizontalAlignment = xlCenter
     .Range(xlColJPYPartCost & Idx + 3) = "NOTE/ SUMMARY"
     .Range(xlColJPYPartCost & Idx + 3).Font.Bold = True
     
     '************************NOTE/SUMMARY (PARTS COST)*****************************
     .Range(xlColJPYPartCost & Idx + 5) = "PARTS COST:"
     .Range(xlColJPYPartCost & Idx + 5).Font.Bold = True
     
     .Range(xlColUSDPartCost & Idx + 6, xlColEURPartCost & Idx + 10).horizontalAlignment = xlRight
     .Range(xlColIDRPartCost & Idx + 6, xlColIDRPartCost & Idx + 10) = "Rp."
     .Range(xlColUSDPartCost & Idx + 6).NumberFormat = gs_formatAmountIDR
     .Range(xlColUSDPartCost & Idx + 7, xlColUSDPartCost & Idx + 10).NumberFormat = gs_formatAmount
     .Range(xlColEURPartCost & Idx + 6, xlColEURPartCost & Idx + 10).NumberFormat = gs_formatAmountIDR
         
     .Range(xlColJPYPartCost & Idx + 6) = Curr(2) & "      :"
     If .Range(xlColIDRPartCost & Idx).Value <> Empty Then
      .Range(xlColUSDPartCost & Idx + 6) = .Range(xlColIDRPartCost & Idx).Value
      .Range(xlColEURPartCost & Idx + 6) = Format((.Range(xlColIDRPartCost & Idx).Value * ERT(2)), gs_formatAmountIDR)
     End If
     
     .Range(xlColJPYPartCost & Idx + 7) = Curr(1) & "     :"
     If .Range(xlColUSDPartCost & Idx).Value <> Empty Then
      .Range(xlColUSDPartCost & Idx + 7) = .Range(xlColUSDPartCost & Idx).Value
      .Range(xlColEURPartCost & Idx + 7) = Format((.Range(xlColUSDPartCost & Idx).Value * ERT(1)), gs_formatAmountIDR)
     End If
     
     .Range(xlColJPYPartCost & Idx + 8) = Curr(0) & "       :"
     If .Range(xlColJPYPartCost & Idx).Value <> Empty Then
      .Range(xlColUSDPartCost & Idx + 8) = .Range(xlColJPYPartCost & Idx).Value
      .Range(xlColEURPartCost & Idx + 8) = Format((.Range(xlColJPYPartCost & Idx).Value * ERT(0)), gs_formatAmountIDR)
     End If
     
     .Range(xlColJPYPartCost & Idx + 9) = Curr(3) & "       :"
     If .Range(xlColEURPartCost & Idx).Value <> Empty Then
      .Range(xlColUSDPartCost & Idx + 9) = .Range(xlColEURPartCost & Idx).Value
      .Range(xlColEURPartCost & Idx + 9) = Format((.Range(xlColEURPartCost & Idx).Value * ERT(3)), gs_formatAmountIDR)
     End If
     
     'Total Part Cost
     .Range(xlColJPYPartCost & Idx + 10) = "TOTAL"
     .Range(xlColEURPartCost & Idx + 10) = .Range(xlColEURPartCost & Idx + 6) + .Range(xlColEURPartCost & Idx + 7) + .Range(xlColEURPartCost & Idx + 8) + .Range(xlColEURPartCost & Idx + 9)
     'Border
     .Range(xlColUSDPartCost & Idx + 10 & ":" & xlColEURPartCost & Idx + 10).Borders(xlEdgeTop).LineStyle = xlContinuous
 
     '*****************************NOTE/SUMMARY (PROCESS COST)*****************************
     .Range(xlColJPYPartCost & Idx + 12) = "PROCESS COST:"
     .Range(xlColJPYPartCost & Idx + 12).Font.Bold = True
     
     .Range(xlColUSDPartCost & Idx + 13, xlColEURPartCost & Idx + 15).horizontalAlignment = xlRight
     .Range(xlColIDRPartCost & Idx + 13, xlColIDRPartCost & Idx + 15) = "Rp."
     .Range(xlColUSDPartCost & Idx + 13).NumberFormat = gs_formatAmountIDR
     .Range(xlColUSDPartCost & Idx + 14).NumberFormat = gs_formatAmount
     .Range(xlColEURPartCost & Idx + 13, xlColEURPartCost & Idx + 15).NumberFormat = gs_formatAmountIDR
         
     .Range(xlColJPYPartCost & Idx + 13) = Curr(2) & "      :"
     If .Range(xlColIDRProcessCost & Idx).Value <> Empty Then
      .Range(xlColUSDPartCost & Idx + 13) = .Range(xlColIDRProcessCost & Idx).Value
      .Range(xlColEURPartCost & Idx + 13) = Format((.Range(xlColIDRProcessCost & Idx).Value * ERT(2)), gs_formatAmountIDR)
     End If
     
     .Range(xlColJPYPartCost & Idx + 14) = Curr(1) & "     :"
     If .Range(xlColUSDProcessCost & Idx).Value <> Empty Then
      .Range(xlColUSDPartCost & Idx + 14) = .Range(xlColUSDProcessCost & Idx).Value
      .Range(xlColEURPartCost & Idx + 14) = Format((.Range(xlColUSDProcessCost & Idx).Value * ERT(1)), gs_formatAmountIDR)
     End If
     
     'Total Process Cost
     .Range(xlColJPYPartCost & Idx + 15) = "TOTAL"
     .Range(xlColEURPartCost & Idx + 15) = .Range(xlColEURPartCost & Idx + 13) + .Range(xlColEURPartCost & Idx + 14)
     'Border
     .Range(xlColUSDPartCost & Idx + 15 & ":" & xlColEURPartCost & Idx + 15).Borders(xlEdgeTop).LineStyle = xlContinuous
     
     '******************************ProcessMain************************************
     Dim ii As Integer
     Dim Total_CostProcessMain As Double
     Total_CostProcessMain = 0
     For ii = 0 To Index - 1
     .Range(xlColJPYPartCost & Idx + ii + 17) = MainProcess(0, ii) 'Process
     .Range(xlColIDRPartCost & Idx + ii + 17) = "Rp."
     .Range(xlColEURPartCost & Idx + ii + 17) = MainProcess(1, ii) 'AmountCost @Process
     .Range(xlColEURPartCost & Idx + ii + 17).NumberFormat = gs_formatAmountIDR
     Total_CostProcessMain = Total_CostProcessMain + MainProcess(1, ii)
     Next ii
     
     .Range(xlColIDRPartCost & Idx + 17, xlColEURPartCost & Idx + ii + 17).horizontalAlignment = xlRight
     .Range(xlColJPYPartCost & Idx + 17, xlColJPYPartCost & Idx + ii + 17).Font.Bold = True
     '***************************DUTY/ TAX******************************
     If ii = 0 Then ii = ii - 1 'kl process main tdk ada ii jadi -1
     .Range(xlColJPYPartCost & Idx + ii + 18) = "DUTY / TAX"
     .Range(xlColJPYPartCost & Idx + ii + 18).Font.Bold = True
     
     .Range(xlColUSDPartCost & Idx + ii + 19, xlColEURPartCost & Idx + ii + 22).horizontalAlignment = xlRight
     .Range(xlColIDRPartCost & Idx + ii + 19, xlColIDRPartCost & Idx + ii + 22) = "Rp."
     .Range(xlColUSDPartCost & Idx + ii + 19).NumberFormat = gs_formatAmountIDR
     .Range(xlColUSDPartCost & Idx + ii + 20, xlColUSDPartCost & Idx + ii + 21).NumberFormat = gs_formatAmount
     .Range(xlColEURPartCost & Idx + ii + 19, xlColEURPartCost & Idx + ii + 22).NumberFormat = gs_formatAmountIDR
         
     .Range(xlColJPYPartCost & Idx + ii + 19) = Curr(2) & "     :"
     If .Range(xlColIDRDutyPrice & Idx).Value <> Empty Then
      .Range(xlColUSDPartCost & Idx + ii + 19) = .Range(xlColIDRDutyPrice & Idx).Value
      .Range(xlColEURPartCost & Idx + ii + 19) = Format((.Range(xlColIDRDutyPrice & Idx).Value * ERT(2)), gs_formatAmountIDR)
     End If
     
     .Range(xlColJPYPartCost & Idx + ii + 20) = Curr(1) & "       :"
     If .Range(xlColUSDDutyPrice & Idx).Value <> Empty Then
      .Range(xlColUSDPartCost & Idx + ii + 20) = .Range(xlColUSDDutyPrice & Idx).Value
      .Range(xlColEURPartCost & Idx + ii + 20) = Format((.Range(xlColUSDDutyPrice & Idx).Value * ERT(1)), gs_formatAmountIDR)
     End If
     
     .Range(xlColJPYPartCost & Idx + ii + 21) = Curr(0) & "       :"
     If .Range(xlColJPYDutyPrice & Idx).Value <> Empty Then
      .Range(xlColUSDPartCost & Idx + ii + 21) = .Range(xlColJPYDutyPrice & Idx).Value
      .Range(xlColEURPartCost & Idx + ii + 21) = Format((.Range(xlColJPYDutyPrice & Idx).Value * ERT(0)), gs_formatAmountIDR)
     End If
     
     'Total Duty
     .Range(xlColJPYPartCost & Idx + ii + 22) = "TOTAL"
     .Range(xlColEURPartCost & Idx + ii + 22) = .Range(xlColEURPartCost & Idx + ii + 19) + .Range(xlColEURPartCost & Idx + ii + 20) + .Range(xlColEURPartCost & Idx + ii + 21)
     'Border
     .Range(xlColUSDPartCost & Idx + ii + 22 & ":" & xlColEURPartCost & Idx + ii + 22).Borders(xlEdgeTop).LineStyle = xlContinuous
     
     '****************TOTAL AKHIR (Total Cost)****************
     .Range(xlColJPYPartCost & Idx + ii + 24) = "TOTAL COST"
     .Range(xlColJPYPartCost & Idx + ii + 24, xlColEURPartCost & Idx + ii + 24).Font.Bold = True
     .Range(xlColIDRPartCost & Idx + ii + 24) = "Rp."
     .Range(xlColIDRPartCost & Idx + ii + 24).horizontalAlignment = xlRight
     .Range(xlColEURPartCost & Idx + ii + 24) = (.Range(xlColEURPartCost & Idx + 10).Value + .Range(xlColEURPartCost & Idx + 15).Value + .Range(xlColEURPartCost & Idx + ii + 22).Value) + Total_CostProcessMain '--> cr lambat (utk buktiin benar ato ga)
     .Range(xlColEURPartCost & Idx + ii + 24).NumberFormat = gs_formatAmountIDR
     .Range(xlColEURPartCost & Idx + ii + 24).Borders(xlEdgeBottom).LineStyle = xlContinuous
     .Range(xlColEURPartCost & Idx + ii + 24).Borders(xlEdgeTop).LineStyle = xlContinuous
     .Range(xlColEURPartCost & Idx + ii + 24).Borders(xlEdgeLeft).LineStyle = xlContinuous
     
     'All Summary Borders
     With .Range(xlColJPYPartCost & Idx + 3 & ":" & xlColEURPartCost & Idx + ii + 25)
      With .Borders(xlEdgeLeft)
       .LineStyle = xlContinuous
       .Weight = xlMedium
       .ColorIndex = xlAutomatic
      End With
      
      With .Borders(xlEdgeTop)
       .LineStyle = xlContinuous
       .Weight = xlMedium
       .ColorIndex = xlAutomatic
      End With
    
      With .Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .Weight = xlMedium
       .ColorIndex = xlAutomatic
      End With
    
      With .Borders(xlEdgeRight)
       .LineStyle = xlContinuous
       .Weight = xlMedium
       .ColorIndex = xlAutomatic
      End With
     End With
    
     With .Range(xlColJPYPartCost & Idx + 3 & ":" & xlColEURPartCost & Idx + 3)
      With .Borders(xlEdgeLeft)
       .LineStyle = xlContinuous
       .Weight = xlMedium
       .ColorIndex = xlAutomatic
      End With
      
      With .Borders(xlEdgeTop)
       .LineStyle = xlContinuous
       .Weight = xlMedium
       .ColorIndex = xlAutomatic
      End With
    
      With .Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .Weight = xlMedium
       .ColorIndex = xlAutomatic
      End With
     End With
     
     
     '*******************Font********************************
     .Range(xlColLevel & "1", xlColAmountUSD & Idx + ii + 25).Columns.Font.Name = "Arial"
     .Range(xlColLevel & "9", xlColAmountUSD & Idx + ii + 25).Columns.Font.Size = 8
     '*******************************************************
      
     '************ width ******************
     .Range(xlColLevel & ":" & xlColAmountUSD).Columns.AutoFit
     .Range(xlColJPYPartCost & "9:" & xlColEURPartCost & "9," & xlColUSDProcessCost & "9:" & xlColAmountUSD & "9").columnWidth = 15
     '*************************************
     
     With .PageSetup
      .PaperSize = xlPaperA4
      .Orientation = 2
'      .LeftMargin = application.InchesToPoints(0.4)
'      .RightMargin = application.InchesToPoints(0.4)
'      .PrintArea = "$" & xlColLevel & "$1:$" & xlColAmountUSD & "$" & Idx + ii + 25
'      .Zoom = 75
'      .CenterHorizontally = True
     End With

'     'Page Breaks
'     .DisplayPageBreaks = True

     xlapp.WindowState = xlMaximized
     xlapp.Visible = True
     
     End With
     
ErrExit:
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlapp = Nothing
    Set rsExchRateTbl = Nothing
    Set rsmakergroup = Nothing
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
   
   End If

End If
End Sub

Sub IsiExcel(ibu As String, lvl As Integer, qtyinduk As Double, xlapp As Excel.application, xlBook As Excel.Workbook, xlSheet As Excel.Worksheet)
Dim anak As String
Dim rsAnak As New ADODB.Recordset

'********VarPenampung********
Dim qtyawal As Double
Dim qtyakhir As Double
Dim cost_minute As Double
Dim part_cost As Double
Dim process_cost As Double
Dim tax As Double
Dim duty_price As Double
Dim ExchRate As Double
Dim Amount As Double
'*****************************

'******VarPembantu*********
Dim amount1 As Double
Dim amount2 As Double
Dim formatprice As String
Dim FormatAmount As String
'**************************

Screen.MousePointer = vbHourglass

With xlSheet
    
    '******************************ExchangeRate*************************/*********
    SqlExchRate = Chr(10) & "IsNull((Select ber.Exch0" & dtAwal.Month & " as ExchRate " & _
                  Chr(10) & "from Book_ExchangeRate ber,Company_Profile cp " & _
                  Chr(10) & "Where ber.Term_Cls = cp.ValuationPrice_ExchTerm " & _
                  Chr(10) & "and ber.Exch_Year = '" & dtAwal.Year & "' " & _
                  Chr(10) & "and ber.Currency_Code = z.Curr_Code),0) "
    '********************************Material**************************************
    sql = "Select *, ExchRate = " & SqlExchRate & "from ( "
    sql = sql & _
          Chr(10) & "--Material" & _
          Chr(10) & "Select '0' idx, bm.parent_itemcode, bm.Item_Code,im.Item_Name, " & _
          Chr(10) & "bm.qty, bm.unit_cls, " & _
          Chr(10) & "trade_code = (Case (select count(bm2.item_code) " & _
          Chr(10) & "from bom_master bm2 where bm2.parent_itemcode = bm.Item_Code) " & _
          Chr(10) & "when 0 then  im.supplier_code Else '' end), " & _
          Chr(10) & "trade_name = '', cost_minute=0, process_cls='',process_cost = 0, "
    sql = sql & _
          Chr(10) & "part_cost = (Case (select count(bm2.item_code) from bom_master bm2 where bm2.parent_itemcode = bm.Item_Code) " & _
          Chr(10) & "when 0 then " & _
          Chr(10) & "(isnull((select top 1 prim.price " & _
          Chr(10) & "from price_master prim " & _
          Chr(10) & "Where prim.item_code = bm.item_code " & _
          Chr(10) & "and prim.price_cls = '01' " & _
          Chr(10) & "and prim.trade_code = im.Supplier_Code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' " & _
          Chr(10) & "and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc),0) ) " & _
          Chr(10) & "else 0 end ), "
    sql = sql & _
          Chr(10) & "curr_code = (Case (select count(bm2.item_code) from bom_master bm2 where bm2.parent_itemcode = bm.Item_Code) " & _
          Chr(10) & "when 0 then (select top 1 prim.currency_code from price_master prim " & _
          Chr(10) & "where prim.item_code = bm.item_code and prim.price_cls = '01' " & _
          Chr(10) & "and prim.trade_code = im.Supplier_Code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' " & _
          Chr(10) & "and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc) else '' end), "
    sql = sql & _
          Chr(10) & sqlDuty & _
          Chr(10) & "from BOM_Master bm,Item_Master im, HS_Master hm where " & _
          Chr(10) & "im.HS_Code *= hm.HS_Code " & _
          Chr(10) & "and bm.Item_Code = im.Item_Code and bm.parent_itemcode = '" & ibu & "' " & _
          Chr(10) & "and (bm.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' " & _
          Chr(10) & "and bm.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') "
    '********************************* WIP **************************************************
    sql = sql & _
          Chr(10) & _
          Chr(10) & "UNION ALL " & _
          Chr(10)
    sql = sql & _
          Chr(10) & "--WIP" & _
          Chr(10) & "select '1' idx, prom.item_code, item_code = '', pros.description process_name, " & _
          Chr(10) & "qty = prom.standard_time, " & _
          Chr(10) & "unit_cls = (case len(isnull(prom.trade_code,'')) when 0 then '0' else '1' end) , " & _
          Chr(10) & "prom.trade_code, tm.trade_name, isnull(prom.cost_minute,0) cost_minute, prom.process_cls, "
    sql = sql & _
          Chr(10) & "process_cost = " & _
          Chr(10) & "--purchase " & _
          Chr(10) & "(case len(isnull(prom.trade_code,'')) when 0 then prom.cost_minute " & _
          Chr(10) & "else ( " & _
          Chr(10) & "isnull((select top 1 prim.price from price_master prim where prim.trade_code = prom.trade_code and " & _
          Chr(10) & "prim.price_cls = '01' and prim.item_code=prom.item_code and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and " & _
          Chr(10) & "prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') order by prim.priority_cls desc),0) " & _
          Chr(10) & "* " & _
          Chr(10) & "(case " & _
          Chr(10) & "(isnull((select top 1 prim.price from price_master prim where prim.trade_code = prom.trade_code and " & _
          Chr(10) & "prim.price_cls = '05' and prim.item_code=prom.item_code and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and " & _
          Chr(10) & "prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') order by prim.priority_cls desc),0) " & _
          Chr(10) & ") when 0 Then 1 "
    sql = sql & _
          Chr(10) & "else ( " & _
          Chr(10) & "-- jika tdk ada price di service pake curr sendr, sebalikny pake curr di service (1) jika sama currnya " & _
          Chr(10) & "-- jika curr di purchase beda dengan service, dianggap (0) " & _
          Chr(10) & "case " & _
          Chr(10) & "(select top 1 prim.currency_code from price_master prim " & _
          Chr(10) & "where prim.trade_code = prom.trade_code and prim.price_cls = '01' and prim.item_code=prom.item_code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc) " & _
          Chr(10) & "when ( " & _
          Chr(10) & "--cek sama ato ngga dg curr yg di service " & _
          Chr(10) & "select top 1 prim.currency_code from price_master prim " & _
          Chr(10) & "where prim.trade_code = prom.trade_code and prim.price_cls = '05' and prim.item_code=prom.item_code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc " & _
          Chr(10) & ") then 1 else 0 end)  end) " & _
          Chr(10) & "+ "
    sql = sql & _
          Chr(10) & "--service " & _
          Chr(10) & "(isnull((select top 1 prim.price from price_master prim where prim.trade_code = prom.trade_code and " & _
          Chr(10) & "prim.price_cls = '05' and prim.item_code=prom.item_code and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and " & _
          Chr(10) & "prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') order by prim.priority_cls desc),0)) " & _
          Chr(10) & ") end), "
    sql = sql & _
          Chr(10) & "part_cost = 0, "
    sql = sql & _
          Chr(10) & "curr_code = " & _
          Chr(10) & "(case len(isnull(prom.trade_code,'')) when 0 then prom.Currency_Code " & _
          Chr(10) & "else ( " & _
          Chr(10) & "case " & _
          Chr(10) & "(select count (Y.currency_code) from " & _
          Chr(10) & "(select top 1 prim.currency_code from price_master prim " & _
          Chr(10) & "where prim.trade_code = prom.trade_code and prim.price_cls = '05' and prim.item_code=prom.item_code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc) Y) "
    sql = sql & _
          Chr(10) & "when  0 then " & _
          Chr(10) & "( " & _
          Chr(10) & "select top 1 prim.currency_code from price_master prim " & _
          Chr(10) & "where prim.trade_code = prom.trade_code and prim.price_cls = '01' and prim.item_code=prom.item_code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc) " & _
          Chr(10) & "Else " & _
          Chr(10) & "( " & _
          Chr(10) & "select top 1 prim.currency_code from price_master prim " & _
          Chr(10) & "where prim.trade_code = prom.trade_code and prim.price_cls = '05' and prim.item_code=prom.item_code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc) end) " & _
          Chr(10) & "end), "
    sql = sql & _
          Chr(10) & "tax = 0 " & _
          Chr(10) & "from process_master prom, process_cls pros, trade_master tm where " & _
          Chr(10) & "prom.process_cls = pros.process_cls and " & _
          Chr(10) & "prom.trade_code *= tm.trade_code and " & _
          Chr(10) & "prom.item_code = '" & ibu & "' "
    '********************************************************************************
    sql = sql & _
          Chr(10) & ")Z order by idx, item_code "
          
    Set rsAnak = Db.Execute(sql)
       
    
  If Not rsAnak.EOF Then
    
    Do While Not rsAnak.EOF
      
      If Trim(rsAnak!curr_code) = "03" Then
        FormatAmount = gs_formatAmountIDR
        formatprice = gs_formatPriceIDR
      Else
        FormatAmount = gs_formatAmount
        formatprice = gs_formatPrice
      End If
      
      anak = Trim(rsAnak!Item_Code)
      qtyawal = Format(rsAnak!Qty, gs_formatQtyBOM)
      qtyakhir = Format((qtyinduk * qtyawal), gs_formatQtyBOM)
      cost_minute = Format(rsAnak!cost_minute, formatprice)
      part_cost = Format(rsAnak!part_cost, formatprice)
      process_cost = Format(rsAnak!process_cost, formatprice)
      tax = Format(((rsAnak!tax) / 100), gs_formatPercentage)
      duty_price = Format((tax * part_cost), formatprice)
      ExchRate = Format(rsAnak!ExchRate, gs_formatExchangeRate)
     
      main = False
     
     If rsAnak!part_cost <> 0 Then
      amount1 = Format(((Format((qtyakhir * part_cost), FormatAmount)) * ExchRate), gs_formatAmountIDR)
      amount2 = Format(((Format((qtyakhir * duty_price), FormatAmount)) * ExchRate), gs_formatAmountIDR)
      Amount = amount1 + amount2
     End If
     
     If (rsAnak!process_cost <> 0) Then
      'subcon
      If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) <> "") Then
       'processcost utk 1 minutenya
       process_cost = Format((process_cost / qtyawal), formatprice)
      End If
      'perlu qtyakhir krn processcostnya hanya utk 1 minute
      Amount = Format(((Format((qtyakhir * process_cost), FormatAmount)) * ExchRate), gs_formatAmountIDR)
     End If
     
     If rsAnak!part_cost = 0 And rsAnak!process_cost = 0 Then
      Amount = Format(0, FormatAmount)
     End If
     
     '*************************MainProcesses**************************
     If anak = "" And Trim(rsAnak!parent_itemcode) = cboitem.Text Then
        ReDim Preserve MainProcess(2, Index + 1)
        MainProcess(0, Index) = Trim(rsAnak!item_name)
        MainProcess(1, Index) = Amount
        Index = Index + 1
     End If
     '************************Main processes tdk ditampilkan************************
     If anak = "" And Trim(rsAnak!parent_itemcode) = cboitem.Text Then
      main = True
      GoTo Rekursif
     End If
     '*************************Content**************************************
     If anak <> "" Then
      .Range(xlColLevel & Idx) = lvl + 1
     Else
      .Range(xlColLevel & Idx) = lvl
     End If
     
     If lvl = 0 Then
      .Range(xlColPartCode & Idx) = anak
      countrow = 0
      countchild = 0
'      .Range(xlColPartCode & Idx & ":" & xlColDescription & Idx).Font.ColorIndex = 3
     Else
      countrow = countrow + 1
      If anak <> "" Then countchild = countchild + 1
      .Range(xlColPartCode & Idx) = Space(lvl) & anak
'      .Range(xlColPartCode & Idx & ":" & xlColDescription & Idx).Font.ColorIndex = 5
     End If
     
     .Range(xlColDescription & Idx) = Trim(rsAnak!item_name)
     .Range(xlColQty & Idx) = IIf((qtyakhir = 0), "", qtyakhir)
     
     If Trim(rsAnak("Unit_Cls")) = "0" Then 'process nonsubcon (WIP)
        .Range(xlColUnitCls & Idx) = "minute"
     ElseIf Trim(rsAnak("Unit_Cls")) = "1" Then 'process subcon (WIP)
        .Range(xlColUnitCls & Idx) = "pcs"
     Else 'material
        .Range(xlColUnitCls & Idx) = Trim(uf_GetUnitDescription(rsAnak("Unit_Cls")))
     End If
     
     Select Case Trim(rsAnak!curr_code)
      Case "01"
       .Range(xlColJPYPartCost & Idx) = IIf(part_cost = 0, "", part_cost)
       .Range(xlColJPYDutyPrice & Idx) = IIf(tax = 0, "", duty_price)
      Case "02"
       .Range(xlColUSDPartCost & Idx) = IIf(part_cost = 0, "", part_cost)
       .Range(xlColUSDDutyPrice & Idx) = IIf(tax = 0, "", duty_price)
      Case "03"
       .Range(xlColIDRPartCost & Idx) = IIf(part_cost = 0, "", part_cost)
       .Range(xlColIDRDutyPrice & Idx) = IIf(tax = 0, "", duty_price)
      Case "04"
       .Range(xlColEURPartCost & Idx) = IIf(part_cost = 0, "", part_cost)
     End Select
     
     .Range(xlColSupCode & Idx) = "'" & Trim(rsAnak!Trade_Code)
     
     Select Case Trim(rsAnak!curr_code)
      Case "02": .Range(xlColUSDProcessCost & Idx) = IIf(process_cost = 0, "", process_cost)
      Case "03": .Range(xlColIDRProcessCost & Idx) = IIf(process_cost = 0, "", process_cost)
     End Select
     
     If Amount <> 0 Then
      .Range(xlColAmount & Idx) = Amount
      .Range(xlColAmountUSD & Idx) = Format((Amount / ERT(1)), gs_formatAmount)
     End If
     
     Idx = Idx + 1
         
Rekursif:
     Call IsiExcel(anak, lvl + 1, rsAnak!Qty * qtyinduk, xlapp, xlBook, xlSheet)
     
     If Not (rsAnak.EOF) Then rsAnak.MoveNext
            
     Loop
     
  Else
     If rsAnak.EOF Then
      If main = False Then
       Dim CC As Integer
       For CC = 0 To countrow
        .Range("r" & Idx - (1 + CC)) = countchild
'        If countchild = 0 Then
'         '*****Color utk bedain*******
'         If countrow > 0 Then
'          .Range(xlColPartCode & Idx - (1 + CC) & ":" & xlColDescription & Idx - (1 + CC)).Font.ColorIndex = 10
'         Else
'          .Range(xlColPartCode & Idx - (1 + CC) & ":" & xlColDescription & Idx - (1 + CC)).Font.ColorIndex = 0
'         End If
'        End If
       Next CC
      End If
     End If
  
  End If
  Set rsAnak = Nothing
  End With
End Sub

Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = cboitem.Text
 frm_BrowseItem.Show 1
 cboitem.Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub CmdSubMenu_Click()
DoEvents
frmMainMenu.Show
DoEvents
Unload Me
End Sub

Sub Kosong()
lblgroup = ""
lblitem(0) = ""
lblitem(1) = ""
End Sub

Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
Call Kosong
Call adtocombo
dtAwal = Now
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

Function Space(lvl As Integer) As String
Dim jlhspace As Integer
Dim chrspace As String
Dim i As Long
 
jlhspace = 5
chrspace = ""

For i = 1 To jlhspace * lvl
 chrspace = chrspace + " "
Next i

Space = chrspace
End Function
