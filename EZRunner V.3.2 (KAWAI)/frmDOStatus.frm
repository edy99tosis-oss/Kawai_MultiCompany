VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDOStatus 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Delivery Note Status"
   ClientHeight    =   10320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frmDOStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1245
      Left            =   540
      TabIndex        =   10
      Top             =   1350
      Width           =   14175
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
         Left            =   5670
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker dtAwal 
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         Top             =   735
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
         Format          =   136577027
         CurrentDate     =   37860
      End
      Begin MSComCtl2.DTPicker dtAkhir 
         Height          =   330
         Left            =   3840
         TabIndex        =   2
         Top             =   735
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
         Format          =   136577027
         CurrentDate     =   37799
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   14
         Top             =   330
         Width           =   840
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   270
         Width           =   1665
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2937;556"
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
         Left            =   3495
         TabIndex        =   13
         Top             =   810
         Width           =   165
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   1
         Left            =   285
         TabIndex        =   12
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         Left            =   3480
         TabIndex        =   11
         Top             =   330
         Width           =   1395
      End
      Begin VB.Line Line1 
         X1              =   3480
         X2              =   9570
         Y1              =   570
         Y2              =   570
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   540
      TabIndex        =   8
      Top             =   9090
      Width           =   14175
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
         Left            =   7050
         TabIndex        =   9
         Top             =   195
         Width           =   75
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
      Left            =   540
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9765
      Width           =   1140
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
      Left            =   13575
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9765
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
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
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9765
      Width           =   1140
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6255
      Left            =   540
      TabIndex        =   15
      Top             =   2745
      Width           =   14175
      _cx             =   25003
      _cy             =   11033
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
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
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12870
      TabIndex        =   16
      Top             =   495
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Note Status"
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
      Left            =   540
      TabIndex        =   7
      Top             =   495
      Width           =   14175
   End
End
Attribute VB_Name = "frmDOStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim dbTransfer As New ADODB.Connection
Dim HakU As Integer
Dim dealerCD As String, Dono As String
Dim blnFix As Integer, thnFix As Integer
Dim statusKlik As Integer
Dim newCls As New clsMRP

Dim bteColTradeCls As Byte
Dim bteColSJNo As Byte
Dim bteColPONo As Byte
Dim bteColCustCode As Byte
Dim bteColCustName As Byte
Dim bteColSJDate As Byte
Dim bteColAmount As Byte
Dim bteColIssue As Byte
Dim bteColFix As Byte
Dim bteColFixCls As Byte
Dim bteColInvCheck As Byte

Dim bteHakPrice As Byte

Private Sub headerGrid()
    Dim i As Integer
    
    bteColTradeCls = 0
    bteColSJNo = 1
    bteColPONo = 2
    bteColCustCode = 3
    bteColCustName = 4
    bteColSJDate = 5
    bteColAmount = 6
    bteColIssue = 7
    bteColFix = 8
    bteColFixCls = 9
    bteColInvCheck = 10
    
    With grid
        .clear
        .ColS = 11
        .Rows = 1
        
        .TextMatrix(0, bteColTradeCls) = "Trade Cls"
        .TextMatrix(0, bteColSJNo) = "DN No."
        .TextMatrix(0, bteColPONo) = "PO NO."
        .TextMatrix(0, bteColCustCode) = "Cust. Code"
        .TextMatrix(0, bteColCustName) = "Cust. Name"
        .TextMatrix(0, bteColSJDate) = "DN Date"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColIssue) = "Issued"
        .TextMatrix(0, bteColFix) = "Fix"
        
        .ColWidth(bteColTradeCls) = 1500
        .ColWidth(bteColSJNo) = 1500
        .ColWidth(bteColPONo) = 1500
        .ColWidth(bteColCustCode) = 1600
        .ColWidth(bteColCustName) = 4000
        .ColWidth(bteColSJDate) = 1500
        .ColWidth(bteColAmount) = 2500
        .ColWidth(bteColIssue) = 800
        
        .ColHidden(bteColTradeCls) = True
        .ColHidden(bteColFixCls) = True
        .ColHidden(bteColInvCheck) = True
        .ColHidden(bteColFixCls) = True
        .ColHidden(bteColAmount) = (bteHakPrice = 0)
        
        .ColAlignment(bteColSJNo) = flexAlignCenterCenter
        .ColAlignment(bteColPONo) = flexAlignLeftCenter
        .ColAlignment(bteColSJDate) = flexAlignLeftCenter
        .ColAlignment(bteColSJDate) = flexAlignCenterCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
        .ColAlignment(bteColIssue) = flexAlignCenterCenter
        .ColAlignment(bteColFix) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, bteColTradeCls, 0, bteColFix) = flexAlignCenterCenter
        
        .EditMaxLength = 1
    End With
End Sub

'******** Combo **********
Sub isiCboCust() 'Isi Combo Dealer CD dr Customer Master
Dim RsCust As New ADODB.Recordset 'Data Customer

With cbo
    .clear
    .columnCount = 3
    .TextColumn = 1
    
    '******** Ambil dr Customer Master utk Combo Dealer CD
    sql = "select Trade_Code,Trade_Name,Trade_Cls from Trade_Master " & _
        "where (Trade_Cls = 2  or Trade_Cls = 4 or Trade_Cls = 3)  order by Trade_Code"
    Set RsCust = Db.Execute(sql)
    
    .AddItem ""
    .List(0, 0) = strAll
    .List(0, 1) = strAll
    .List(0, 2) = strAll
    
    i = 1
    Do While Not (RsCust.EOF)
        .AddItem ""
        .List(i, 0) = Trim(RsCust(0))
        .List(i, 1) = Trim(RsCust(1))
        .List(i, 2) = Trim(RsCust(2))
        i = i + 1
        RsCust.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 350
    .ColumnWidths = "50 pt;300 pt; 0 pt"
    .ListIndex = 0
    
    Set RsCust = Nothing
End With
End Sub

Private Sub cbo_Change()
    cbo = cbo
    dealerCD = cbo
    If cbo.MatchFound Then
        lblNm(0) = cbo.Column(1)
        LblErrMsg = ""
    Else
        lblNm(0) = ""
        LblErrMsg = DisplayMsg(4011)
    End If
    Call headerGrid
End Sub

'******************
Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
HakU = hakUpdate(Me.Name)
bteHakPrice = hakPrice(Me.Name)
Call isiCboCust
dtAwal = Date - (Day(Date) - 1)
dtAkhir = Date
End Sub

Function stInvoice(noDO As String) As Integer
Dim sql As String
Dim rsSt As New ADODB.Recordset
    sql = "select DO_No from Invoice_Detail where  DO_NO = '" & noDO & "' "
    Set rsSt = Db.Execute(sql)
    If rsSt.EOF Then
        stInvoice = 0
    Else
        stInvoice = 1
    End If
End Function

Sub IsiGrid()
Dim rsDOMaster As New ADODB.Recordset

Call headerGrid
With grid
    
    sql = "select distinct tm.Trade_Cls, tm.Trade_Name, dm.DO_NO, dm.Cust_Code, dm.DO_Date, dm.Amount, dm.Reissue_Cls, dm.Fix_Cls,List_PO " & _
        "from Do_master dm " & _
        "inner join trade_master tm on dm.cust_code = tm.trade_code " & _
        "where dm.do_date >= '" & Format(dtAwal, "yyyy-MM-dd") & "' " & _
        "and dm.do_date <= '" & Format(dtAkhir, "yyyy-MM-dd") & "' "
    
    If cbo <> strAll Then sql = sql & "and dm.Cust_Code = '" & cbo & "' "
    sql = sql & "order by dm.Do_NO"
    
    Set rsDOMaster = Db.Execute(sql)
    
    If rsDOMaster.EOF Then

        LblErrMsg = DisplayMsg(4006)
        Exit Sub
    End If

    i = 1
    Do While Not rsDOMaster.EOF
        .Rows = .Rows + 1
        .TextMatrix(i, bteColTradeCls) = Trim(rsDOMaster("Trade_Cls"))
        .TextMatrix(i, bteColSJNo) = Trim(rsDOMaster("DO_NO"))
        .TextMatrix(i, bteColPONo) = Trim(rsDOMaster("List_PO"))
        .TextMatrix(i, bteColCustCode) = Trim(rsDOMaster("Cust_Code"))
        .TextMatrix(i, bteColCustName) = Trim(rsDOMaster("Trade_Name"))
        .TextMatrix(i, bteColSJDate) = Format(Trim(rsDOMaster("do_date")), "dd MMM yyyy")
        .TextMatrix(i, bteColAmount) = Trim(Format(rsDOMaster("Amount"), gs_formatAmountIDR))
        .Cell(flexcpChecked, i, bteColIssue) = IIf(rsDOMaster("Reissue_Cls") = 1, flexChecked, flexUnchecked)
        .Cell(flexcpChecked, i, bteColFix) = IIf(rsDOMaster("Fix_Cls") = 1, flexChecked, flexUnchecked)
        .Cell(flexcpBackColor, i, bteColFix) = vbWhite
        .TextMatrix(i, bteColFixCls) = rsDOMaster("Fix_Cls")
        .TextMatrix(i, bteColInvCheck) = stInvoice(rsDOMaster("DO_NO"))
        i = i + 1
        rsDOMaster.MoveNext
    Loop
    Set rsDOMaster = Nothing
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim pesandtAwal As String, pesandtAkhir As String

With grid
    LblErrMsg = ""
    If Col < bteColFix Then
        Cancel = 1
    Else
        pesandtAwal = up_ValidateDateRange(Format(.TextMatrix(Row, bteColSJDate), "yyyy-MM-dd"), True)
        pesandtAkhir = up_ValidateDateRange(Format(.TextMatrix(Row, bteColSJDate), "yyyy-MM-dd"), True)
        'tidak perlu cek invoice
'        If Trim(.TextMatrix(Row, bteColInvCheck)) = 1 And .Cell(flexChecked, Row, bteColFix) = 0 Then
'            lblErrMsg = DisplayMsg(1106)
'            Cancel = 1
'        Else
            If pesandtAwal <> "" Or pesandtAkhir <> "" Then
                LblErrMsg = IIf(pesandtAwal = "", pesandtAkhir, pesandtAwal)
                Cancel = 1
            End If
'        End If
    End If
End With
End Sub

Private Sub dtAwal_Change()
    LblErrMsg = ""
    If Format(dtAwal, "yyyy-MM-dd") > Format(CDate(dtAkhir), "yyyy-MM-dd") Then LblErrMsg = DisplayMsg("4068"): Exit Sub
    Call headerGrid
End Sub

Private Sub dtAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub dtAkhir_Change()
    LblErrMsg = ""
    If Format(dtAwal, "yyyy-MM-dd") > Format(CDate(dtAkhir), "yyyy-MM-dd") Then LblErrMsg = DisplayMsg("4066"): Exit Sub
    Call headerGrid
End Sub

Private Sub dtAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub cmdSearch_Click()
    MousePointer = vbHourglass
    Call IsiGrid
    MousePointer = vbDefault
End Sub

Sub inputSupply(Dono As String)
Dim RsDo As New ADODB.Recordset
Dim rsAnak As New ADODB.Recordset
Dim fromWHCode As String, fromAddress As String, toWHCode As String, toAddress As String
Dim itemDO As String, tglDO As String
Dim qtyDo As Double, UnitCls As String, currCD As String, Price As Double, Amount As Double
Dim itemAnak As String, nilPrice As String, currAnak As String, qtyAnak As Double
Dim stockWH As String, stockItem As String, LNo As String
Dim Sn As Long

Dim TSeqNo As Integer
Dim TSerialFrom As String, TSerialTo As String
       
    sql = "select DM.Cust_Code, DM.Do_Date, DO.Delivery_Date, DO.Item_Code, DO.Qty,DO.Unit_Cls, DO.Currency_Code, DO.Price,Do.Lot_No, " & vbCrLf & _
        "       case when rtrim(isnull(DM.WHCode, '')) = '' then I.WH_Code else DM.WHCode end WH_Code, I.Address, I.StockCOntrol_Cls as stockItem, WH.StockControl_Cls as stockWH, " & vbCrLf & _
        "       Do.DO_No, Do.Seq_No SeqNo, COALESCE(SerialNoFrom,'') SerialNoFrom, COALESCE(SerialNoTo,'') SerialNoTo " & vbCrLf & _
        "   from Delivery_Order DO, DO_Master DM, Item_Master I, Warehouse_master WH " & vbCrLf & _
        "       where DO.DO_No = DM.DO_No and DO.Item_Code = I.Item_Code and case when rtrim(isnull(DM.WHCode, '')) = '' then I.WH_Code else DM.WHCode end = WH.WH_Code " & vbCrLf & _
        "           and DO.DO_NO = '" & Dono & "'" & vbCrLf
        
    Set RsDo = Db.Execute(sql)
    
    Dim rsmax As New ADODB.Recordset
    Dim strSQL As String

    strSQL = "Select isnull(Max(seq_No),0)+1 SeqNo  from part_supply"

    If Not (RsDo.EOF) Then
        Do While Not RsDo.EOF
            TSeqNo = RsDo("SeqNo")
            TSerialFrom = Trim(RsDo("SerialNoFrom"))
            TSerialTo = Trim(RsDo("SerialNoTo"))
            
            itemDO = Trim(RsDo("Item_Code"))
            currCD = RsDo("Currency_Code")
            qtyDo = CDbl(RsDo("qty"))
            tglDO = Format(RsDo("DO_Date"), "yyyy-MM-dd")
            stockWH = RsDo("StockWH")
            stockItem = RsDo("StockItem")
            fromWHCode = RsDo("Wh_Code")
            LNo = RsDo("Lot_No")
            
            'Set rsmax = Db.Execute(StrSql)
            'Sn = rsmax(0)
            'rsmax.Close
             
            '********* Jika Status Fix = 1 baru insert ke Supply *********
            fromAddress = IIf(IsNull(RsDo("Address")), "", RsDo("Address"))
            toWHCode = RsDo("Cust_Code")
            UnitCls = RsDo("Unit_cls")
            Price = RsDo("Price")
            Amount = qtyDo * Price
        
            sql = "insert into Part_Supply(FromWarehouse_Code,From_Address,ToWarehouse_Code,ChildSupply_date,ChildItem_Code,Supply_Cls," & _
                "ChildRequirement_Qty,ChildUnit_Cls,Currency_Code,Price,Amount,ParentItem_Code,Lot_No,Production_Date,Do_No,Remarks, Last_Update, Last_User) " & _
                "values ('" & Trim(fromWHCode) & "','" & Trim(fromAddress) & "','" & Trim(toWHCode) & "','" & tglDO & "','" & Trim(itemDO) & "','D'," & _
                CDbl(qtyDo) & ",'" & UnitCls & "','" & currCD & "'," & Price & "," & Amount & ",'','" & LNo & "','" & tglDO & "','" & Dono & "','',getdate(),'" & userLogin & "')"
            dbTransfer.Execute sql
            
            sql = " UPDATE Serial_Detail " & vbCrLf & _
                        " SET DO_No='" & Dono & "', DO_SeqNo=" & TSeqNo & " " & vbCrLf & _
                        "   WHERE Item_Code='" & itemDO & "'         " & vbCrLf & _
                        "       AND Serial_No >='" & TSerialFrom & "' AND Serial_No <= '" & TSerialTo & "'    " & vbCrLf
            
            dbTransfer.Execute sql
            
            If stockWH = "01" And stockItem = "01" Then _
                Call newCls.updateStock(fromWHCode, itemDO, qtyDo, "", tglDO, blnFix, thnFix, dbTransfer, "Supply", 0, 1)
            '*********************
            RsDo.MoveNext
        Loop
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim tanya
Dim tampungNoDO As String, tampungBln As String
Dim strDO As String
Dim StatInsertUpdate As Boolean
Dim SetSerial As String

On Error GoTo ErrorMesage
MousePointer = vbHourglass

If HakU = 0 Then LblErrMsg = DisplayMsg(3008): Exit Sub

Select Case Index
    Case 0: 'Submit
        StatInsertUpdate = False
        
        tanya = vbYes 'MsgBox("Do you really want to Process Surat Jalan Status?", vbQuestion & vbYesNo, "Confirmation")
        If tanya = vbYes Then
            Me.MousePointer = vbHourglass
            
            If cbo = "" Then
                LblErrMsg = DisplayMsg(1033)
                cbo.SetFocus
            ElseIf cbo <> dealerCD Then
                LblErrMsg = DisplayMsg(1034)
                cbo.SetFocus
            Else
                cbo = cbo
                If cbo.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4011)
                    cbo.SetFocus
                Else
                    tampungBln = newCls.blnAkhir()
                    blnFix = Split(tampungBln, ",")(0)
                    thnFix = Split(tampungBln, ",")(1)
                    
                    'Jika belum ada Data Stock Inventory Closing
                    If blnFix = 0 Then LblErrMsg = DisplayMsg(4019): Me.MousePointer = vbDefault: Exit Sub
                                
                    LblErrMsg.Caption = ""
                    
                    With grid
                        strDO = ""
                        For i = 1 To .Rows - 1
                            statusKlik = IIf(.Cell(flexcpChecked, i, bteColFix) = flexChecked, 1, 0)
                            If .TextMatrix(i, bteColFixCls) <> statusKlik Then
                                If statusKlik = 0 Then strDO = strDO & "'" & .TextMatrix(i, bteColSJNo) & "',"
                            End If
                        Next i
                    End With
                    
                    dbTransfer.ConnectionTimeout = 0
                    dbTransfer.CommandTimeout = 0
                    dbTransfer.Open Db.ConnectionString
                    dbTransfer.BeginTrans
                    
                    If strDO <> "" Then Call newCls.HapusDataSupp(dbTransfer, Left(strDO, Len(strDO) - 1), blnFix, thnFix)
                    With grid
                        For i = 1 To .Rows - 1
                            
                            'Melakukan perubahan atau tidak
                            statusKlik = IIf(.Cell(flexcpChecked, i, bteColFix) = flexChecked, 1, 0)
                            If .TextMatrix(i, bteColFixCls) <> statusKlik Then
                                Dono = .TextMatrix(i, bteColSJNo)
                                
                                'If cbo.Column(2) <> 3 Then 'Jika Trade Cls bukan 3 (Contractor) Request Pak Toha 20250117
'                                If CInt(Trim(.TextMatrix(i, bteColTradeCls))) <> 3 Then 'Jika Trade Cls bukan 3 (Contractor)
'                                    If statusKlik = 1 Then Call inputSupply(Dono)
'                                End If
                                
                                 If statusKlik = 1 Then Call inputSupply(Dono)
                                
                                '**** Update DO Fix ****
                                sql = "update DO_Master set fix_cls = '" & statusKlik & "', Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                    "where DO_NO ='" & Dono & "'"
                                 
                                 dbTransfer.Execute sql
                                 
                                 '**** Update Serial Detail Status ****
                                 If statusKlik <> 1 Then
                                 
                                    sql = " UPDATE Serial_Detail " & vbCrLf & _
                                                " SET DO_No=NULL , DO_SeqNo=NULL " & vbCrLf & _
                                                "   WHERE DO_No='" & Dono & "' " & vbCrLf
                                    
                                    dbTransfer.Execute sql
                                    
                                End If
                                
                                StatInsertUpdate = True
                            End If
                        Next i
                    End With
                    
                    dbTransfer.CommitTrans
                    dbTransfer.Close
                    
                    If StatInsertUpdate = True Then
                        IsiGrid
                    End If
                    
                    LblErrMsg = DisplayMsg(1101)
                End If
            End If
            Me.MousePointer = vbDefault
        End If
        
    Case 1:
        cbo.ListIndex = 0
        dtAwal = Date - (Day(Date) - 1)
        dtAkhir = Date
End Select

MousePointer = vbDefault
Exit Sub
ErrorMesage:
LblErrMsg = err.number & " " & err.Description
MousePointer = vbDefault
Exit Sub

End Sub

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

Private Function SSeqNo()
Dim rsmax As New ADODB.Recordset
Dim strSQL As String

strSQL = "Select isnull(Max(seq_No),0)+1 SeqNo  from part_supply"

Set rsmax = dbTransfer.Execute(strSQL)

SSeqNo = rsmax(0)

End Function


Sub UpdateSerial(Dono As String, Updatesetting As String)
    
    Dim dbserial As New Connection
    Dim RsDo As New ADODB.Recordset
    Dim rsAnak As New ADODB.Recordset
    Dim fromWHCode As String, fromAddress As String, toWHCode As String, toAddress As String
    Dim itemDO As String, tglDO As String
    Dim qtyDo As Double, UnitCls As String, currCD As String, Price As Double, Amount As Double
    Dim itemAnak As String, nilPrice As String, currAnak As String, qtyAnak As Double
    Dim stockWH As String, stockItem As String, LNo As String
    Dim Sn As Long
    
    Dim TSeqNo As Integer
    Dim TSerialFrom As String, TSerialTo As String
       
       
    dbserial.Open Db.ConnectionString
       
    sql = "select DM.Cust_Code, DM.Do_Date, DO.Delivery_Date, DO.Item_Code, DO.Qty,DO.Unit_Cls, DO.Currency_Code, DO.Price,Do.Lot_No, " & vbCrLf & _
        "       case when rtrim(isnull(DM.WHCode, '')) = '' then I.WH_Code else DM.WHCode end WH_Code, I.Address, I.StockCOntrol_Cls as stockItem, WH.StockControl_Cls as stockWH, " & vbCrLf & _
        "       Do.DO_No, Do.Seq_No SeqNo, COALESCE(SerialNoFrom,'') SerialNoFrom, COALESCE(SerialNoTo,'') SerialNoTo " & vbCrLf & _
        "   from Delivery_Order DO, DO_Master DM, Item_Master I, Warehouse_master WH " & vbCrLf & _
        "       where DO.DO_No = DM.DO_No and DO.Item_Code = I.Item_Code and case when rtrim(isnull(DM.WHCode, '')) = '' then I.WH_Code else DM.WHCode end = WH.WH_Code " & vbCrLf & _
        "           and DO.DO_NO = '" & Dono & "'" & vbCrLf
        
    Set RsDo = dbserial.Execute(sql)
    
    Dim rsmax As New ADODB.Recordset
    Dim strSQL As String

    If Not (RsDo.EOF) Then
        Do While Not RsDo.EOF
            TSeqNo = RsDo("SeqNo")
            TSerialFrom = Trim(RsDo("SerialNoFrom"))
            TSerialTo = Trim(RsDo("SerialNoTo"))
            itemDO = Trim(RsDo("Item_Code"))
            
            If Updatesetting = "1" Then
                sql = " UPDATE Serial_Detail " & vbCrLf & _
                            " SET DO_No='" & Dono & "', DO_SeqNo=" & TSeqNo & " " & vbCrLf & _
                            "   WHERE Item_Code='" & itemDO & "'         " & vbCrLf & _
                            "       AND Serial_No >='" & TSerialFrom & "' AND Serial_No <= '" & TSerialTo & "'    " & vbCrLf
            Else
                sql = " UPDATE Serial_Detail " & vbCrLf & _
                            " SET DO_No=NULL, DO_SeqNo=NULL " & vbCrLf & _
                            "   WHERE Item_Code='" & itemDO & "'         " & vbCrLf & _
                            "       AND Serial_No >='" & TSerialFrom & "' AND Serial_No <= '" & TSerialTo & "'    " & vbCrLf
            End If
            
            dbTransfer.Execute sql
            
        Loop
    End If
End Sub


