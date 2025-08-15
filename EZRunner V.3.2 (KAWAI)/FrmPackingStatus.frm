VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPackingStatus 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Packing List Status"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "FrmPackingStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
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
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9825
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   664
      TabIndex        =   12
      Top             =   9090
      Width           =   13905
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
         TabIndex        =   13
         Top             =   210
         Width           =   13650
      End
   End
   Begin VB.CommandButton CmdSubmit 
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
      Left            =   13436
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9825
      Width           =   1140
   End
   Begin VB.CommandButton CmdClear 
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
      Left            =   12236
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9825
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1140
      Left            =   664
      TabIndex        =   8
      Top             =   1009
      Width           =   13905
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0000FFFF&
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
         Left            =   8085
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   630
         Width           =   1140
      End
      Begin VB.TextBox LblCust 
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
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
         Left            =   3285
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   690
         Width           =   4485
      End
      Begin MSComCtl2.DTPicker DTDel1 
         Height          =   315
         Left            =   1665
         TabIndex        =   0
         Top             =   255
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
         Format          =   334692355
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DtDel2 
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
         Left            =   3615
         TabIndex        =   1
         Top             =   255
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
         Format          =   334692355
         CurrentDate     =   37798
      End
      Begin VB.Label LblCode 
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
         Left            =   315
         TabIndex        =   11
         Top             =   720
         Width           =   840
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3285
         X2              =   7815
         Y1              =   960
         Y2              =   960
      End
      Begin MSForms.ComboBox cbocust 
         Height          =   315
         Left            =   1665
         TabIndex        =   2
         Top             =   660
         Width           =   1500
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label4 
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
         Left            =   3285
         TabIndex        =   10
         Top             =   315
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Date"
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
         Left            =   315
         TabIndex        =   9
         Top             =   315
         Width           =   1125
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   6825
      Left            =   675
      TabIndex        =   7
      Top             =   2235
      Width           =   13905
      _cx             =   95903695
      _cy             =   95891207
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
      FocusRect       =   5
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
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
      WordWrap        =   -1  'True
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
      Height          =   420
      Left            =   12709
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   326
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Packing List Status"
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
      Left            =   671
      TabIndex        =   15
      Top             =   349
      Width           =   13905
   End
End
Attribute VB_Name = "FrmPackingStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SSql As String
Dim KondisiStock As String
Dim ColPac, ColCustCode, ColCustName, ColConsignee, ColPacDate, ColSail, ColQty, ColTtW, ColTTWG, ColTTVol, ColIssue, ColFix, ColFixT, colin As Long
Dim ColSerialFrom As Byte, ColSerialTo As Byte


Private Sub CboCust_Change()
LblErrMsg = ""
Header
If cboCust.MatchFound Then
    lblcust.Text = cboCust.Column(1)
Else
    lblcust.Text = ""
    LblErrMsg.Caption = DisplayMsg(4072)
End If
End Sub

Sub FillGrid()
Dim rsisi As New ADODB.Recordset
Dim rsin As New ADODB.Recordset
If rsisi.State = 1 Then rsisi.Close
Header

SSql = "select packing_no, packing_date, stuffing_date, total_qty, " & _
    "totalweight_Netto, totalweight_gross, total_volume, " & _
    "reissue_cls, fix_cls, cust_code, tm.trade_name, consignee " & _
    "from packing_master pm " & _
    "left outer join trade_master tm on pm.cust_code = tm.trade_code " & _
    "where packing_date >= '" & Format(DTDel1.Value, "yyyy-MM-dd") & "' " & _
    "and packing_date <= '" & Format(DtDel2.Value, "yyyy-MM-dd") & "' "
If cboCust <> strAll Then SSql = SSql & "and cust_code = '" & cboCust.Text & "' "
SSql = SSql & "order by packing_no "

rsisi.Open SSql, Db, adOpenKeyset, adLockOptimistic
If Not rsisi.EOF And Not rsisi.BOF Then
    rsisi.MoveFirst
    While Not rsisi.EOF
        grid.Rows = grid.Rows + 1
        grid.TextMatrix(grid.Rows - 1, ColPac) = rsisi.Fields("packing_no")
        grid.TextMatrix(grid.Rows - 1, ColCustCode) = Trim(rsisi.Fields("cust_code"))
        grid.TextMatrix(grid.Rows - 1, ColCustName) = Trim(rsisi.Fields("trade_name"))
        grid.TextMatrix(grid.Rows - 1, ColConsignee) = Trim(rsisi.Fields("consignee"))
        grid.TextMatrix(grid.Rows - 1, ColPacDate) = Format(rsisi.Fields("packing_date"), "dd MMM yyyy")
        grid.TextMatrix(grid.Rows - 1, ColSail) = ""
        grid.TextMatrix(grid.Rows - 1, ColQty) = Format(rsisi.Fields("total_qty"), gs_formatQty)
        grid.TextMatrix(grid.Rows - 1, ColTtW) = Format(rsisi.Fields("totalweight_Netto"), gs_formatWeight)
        grid.TextMatrix(grid.Rows - 1, ColTTWG) = Format(rsisi.Fields("totalweight_gross"), gs_formatWeight)
        grid.TextMatrix(grid.Rows - 1, ColTTVol) = Format(rsisi.Fields("total_volume"), gs_formatVolume)
        If rsisi.Fields("reissue_cls") = "1" Then
            grid.Cell(flexcpChecked, grid.Rows - 1, ColIssue) = flexChecked
        Else
            grid.Cell(flexcpChecked, grid.Rows - 1, ColIssue) = flexUnchecked
        End If
        If rsisi.Fields("fix_cls") = "1" Then
            grid.Cell(flexcpChecked, grid.Rows - 1, ColFix) = flexChecked
            grid.Cell(flexcpChecked, grid.Rows - 1, ColFixT) = flexChecked
        Else
            grid.Cell(flexcpChecked, grid.Rows - 1, ColFix) = flexUnchecked
            grid.Cell(flexcpChecked, grid.Rows - 1, ColFixT) = flexUnchecked
        End If
        
        If rsin.State = 1 Then rsin.Close
        rsin.Open "select * from invoice_detail where packing_no = '" & rsisi.Fields("packing_no") & "'", Db, adOpenKeyset, adLockOptimistic
        If rsin.EOF And rsin.BOF Then
            grid.TextMatrix(grid.Rows - 1, colin) = "0" 'belum
        Else
            grid.TextMatrix(grid.Rows - 1, colin) = "1" 'sudah
        End If
        
        rsisi.MoveNext
    Wend
    
    If grid.Rows > 0 Then
        grid.Cell(flexcpBackColor, 1, ColFix, grid.Rows - 1, ColFix) = vbWhite
    End If
End If
End Sub

Sub Header()
grid.ColS = 14
grid.Rows = 1

grid.ColWidth(ColPac) = 1300
grid.ColWidth(ColCustCode) = 1200
grid.ColWidth(ColCustName) = 2500
grid.ColWidth(ColConsignee) = 1000
grid.ColWidth(ColPacDate) = 1300
grid.ColWidth(ColQty) = 1400
grid.ColWidth(ColTtW) = 1400
grid.ColWidth(ColTTWG) = 1400
grid.ColWidth(ColTTVol) = 1400
grid.ColWidth(ColIssue) = 550
grid.ColWidth(ColFix) = 350

grid.ColHidden(ColSail) = True
grid.ColHidden(ColFixT) = True
grid.ColHidden(colin) = True

grid.TextMatrix(0, ColPac) = "Packing No"
grid.TextMatrix(0, ColCustCode) = "Cust. Code"
grid.TextMatrix(0, ColCustName) = "Cust. Name"
grid.TextMatrix(0, ColConsignee) = "Consignee"
grid.TextMatrix(0, ColPacDate) = "Packing Date"
grid.TextMatrix(0, ColSail) = "Sailing Date"
grid.TextMatrix(0, ColQty) = "Qty"
grid.TextMatrix(0, ColTtW) = "Total Weight (Netto)"
grid.TextMatrix(0, ColTTWG) = "Total Weight (Gross)"
grid.TextMatrix(0, ColTTVol) = "Total Volume"
grid.TextMatrix(0, ColIssue) = "Issued"
grid.TextMatrix(0, ColFix) = "Fix"
grid.TextMatrix(0, ColFixT) = ""

grid.Cell(flexcpAlignment, 0, ColPac, 0, ColFix) = flexAlignCenterCenter

grid.MergeRow(0) = True
For i = 0 To 9
    grid.MergeCol(i) = True
Next i
grid.MergeCells = flexMergeFixedOnly
End Sub

Private Sub cbocust_LostFocus()
cboCust.Text = UCase(cboCust.Text)
End Sub

Private Sub cmdClear_Click()
IsiCombo
Header
DTDel1.Value = Date - Day(Date) + 1
DtDel2.Value = Now
End Sub

Private Sub cmdSearch_Click()
Header
LblErrMsg = ""
FillGrid
End Sub

Private Sub CmdSubMenu_Click()
 DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub CmdSubmit_Click()
LblErrMsg = ""
Dim LMPMonth, TMPMonth, NMPMonth As Long
Dim LMRec, TMRec, NMRec As Long
Dim LMSupply, TMSupply, NMSupply As Long
Dim LMLossRej, TMLossRej, NMLossRej As Long
Dim LMCur, TMCur, NMCur As Long
Dim RsD As New ADODB.Recordset
Dim RsStockHead As New ADODB.Recordset
Dim Sta As String, SintakSql As String
Dim booUpdate As Boolean

Me.MousePointer = vbHourglass

If cboCust.MatchFound = False Then
    LblErrMsg.Caption = DisplayMsg(4072)
    Me.MousePointer = vbDefault
    Exit Sub
End If

If hakUpdate(Me.Name) = 0 Then _
LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Me.MousePointer = vbDefault: Exit Sub

booUpdate = False

Db.BeginTrans
For i = 1 To grid.Rows - 1

'--- Packing List Status tidak mempengaruhi stock ----
'    Sta = ""
'    If Grid.Cell(flexcpChecked, i, ColFix) <> Grid.Cell(flexcpChecked, i, ColFixT) Then
'        booUpdate = True
'        If Grid.Cell(flexcpChecked, i, ColFix) = flexChecked Then
'            'tambah
'            Sta = "T"
'        ElseIf Grid.Cell(flexcpChecked, i, ColFixT) = flexChecked Then
'            'kurang
'            Sta = "K"
'        End If
'
'        SSql = "select pd.item_code, case when rtrim(isnull(pm.whcode, '')) = '' then im.wh_code else pm.whcode end wh_code, " & _
'            "pd.qty, im.address, pm.packing_date, pm.Stuffing_Date, pd.unit_cls, pd.currency_code,  pd.price, pd.amount, pd.packing_no " & _
'            "from packing_master pm " & _
'            "inner join packing_detail pd on pm.packing_no = pd.packing_no " & _
'            "inner join item_master im on pd.item_code = im.item_code " & _
'            "where pd.packing_no =  '" & Grid.TextMatrix(i, ColPac) & "'"
'
'        If RsD.State = 1 Then RsD.Close
'        RsD.Open SSql, Db, adOpenKeyset, adLockOptimistic
'
'        If Not RsD.EOF And Not RsD.BOF Then
'                While Not RsD.EOF
'                    'UPDATE STOCK MASTER
'                    LMPMonth = 0
'                    TMPMonth = 0
'                    NMPMonth = 0
'                    LMRec = 0
'                    TMRec = 0
'                    NMRec = 0
'                    LMSupply = 0
'                    TMSupply = 0
'                    NMSupply = 0
'                    LMLossRej = 0
'                    TMLossRej = 0
'                    NMLossRej = 0
'                    LMCur = 0
'                    TMCur = 0
'                    NMCur = 0
'
'
'                    KondisiStock = DateDiff("m", uf_GetLastClosing("fulldate"), RsD.Fields("Packing_Date"))
'
'                    If CekClsStokDanWarehouse(RsD.Fields("item_code"), RsD.Fields("wh_code")) = True Then
'                    'BILA STOCK
'                    ' Update Sql Baru 20090202
'                                    If RsStockHead.State = 1 Then RsStockHead.Close
'                                    SintakSql = "SELECT " & _
'                                                "isnull(LM_PreMonth,0) LM_PreMonth, " & _
'                                                "isnull(LM_Receipt,0) LM_Receipt, " & _
'                                                "isnull(LM_Supply,0) LM_Supply, " & _
'                                                "isnull(LM_LossReject,0) LM_LossReject, " & _
'                                                "isnull(LM_Current,0) LM_Current, " & _
'                                                "isnull(TM_PreMonth,0) TM_PreMonth, " & _
'                                                "isnull(TM_Receipt,0) TM_Receipt, " & _
'                                                "isnull(TM_Supply,0) TM_Supply, " & _
'                                                "isnull(TM_LossReject,0) TM_LossReject, " & _
'                                                "isnull(TM_Current,0) TM_Current, " & _
'                                                "isnull(NM_PreMonth,0) NM_PreMonth, " & _
'                                                "isnull(NM_Receipt,0) NM_Receipt, " & _
'                                                "isnull(NM_Supply,0) NM_Supply, " & _
'                                                "isnull(NM_LossReject,0) NM_LossReject, " & _
'                                                "isnull(NM_Current,0) NM_Current " & _
'                                                "From Stock_Master " & _
'                                                "where warehouse_code = '" & RsD.Fields("wh_code") & "' " & _
'                                                "and item_code = '" & RsD.Fields("item_code") & "'"
'
'                                    RsStockHead.Open SintakSql, Db, adOpenKeyset, adLockOptimistic
'                                    If RsStockHead.EOF And RsStockHead.BOF Then
'                                        If Sta = "T" Then
'                                            If KondisiStock = "0" Then
'                                                LMSupply = LMSupply + CDbl(RsD.Fields("qty"))
'                                            ElseIf KondisiStock = "1" Then
'                                                TMSupply = TMSupply + CDbl(RsD.Fields("qty"))
'                                            ElseIf KondisiStock = "2" Then
'                                                NMSupply = NMSupply + CDbl(RsD.Fields("qty"))
'                                            End If
'                                        ElseIf Sta = "K" Then
'                                            If KondisiStock = "0" Then
'                                                LMSupply = LMSupply - CDbl(RsD.Fields("qty"))
'                                            ElseIf KondisiStock = "1" Then
'                                                TMSupply = TMSupply - CDbl(RsD.Fields("qty"))
'                                            ElseIf KondisiStock = "2" Then
'                                                NMSupply = NMSupply - CDbl(RsD.Fields("qty"))
'                                            End If
'                                        End If
'
'                                        LMCur = (((LMPMonth + LMRec) - LMSupply) - LMLossRej)
'                                        If KondisiStock <> "0" Then
'                                            'TMPMonth = LMCur
'                                            TMCur = (((TMPMonth + TMRec) - TMSupply) - TMLossRej)
'                                            NMPMonth = TMCur
'                                            NMCur = (((NMPMonth + NMRec) - NMSupply) - NMLossRej)
'                                        End If
'
'                                        SSql = " INSERT INTO  " & _
'                                                " [Stock_Master] " & _
'                                                " ( " & _
'                                                " [Warehouse_Code],  " & _
'                                                " [Item_Code],  " & _
'                                                " [LM_PreMonth],  " & _
'                                                " [LM_Receipt],  " & _
'                                                " [LM_Supply],  " & _
'                                                " [LM_LossReject],  " & _
'                                                " [LM_Current],  " & _
'                                                " [LM_Inventory],  "
'
'                                        SSql = SSql + " [TM_PreMonth],  " & _
'                                                " [TM_Receipt],  " & _
'                                                " [TM_Supply],  " & _
'                                                " [TM_LossReject],  " & _
'                                                " [TM_Current],  " & _
'                                                " [TM_Inventory],  " & _
'                                                " [NM_PreMonth],  " & _
'                                                " [NM_Receipt],  " & _
'                                                " [NM_Supply],  " & _
'                                                " [NM_LossReject],  " & _
'                                                " [NM_Current],  "
'
'                                        SSql = SSql + " [NM_Inventory] " & _
'                                                " ) " & _
'                                                " VALUES " & _
'                                                " ( " & _
'                                                " '" & RsD.Fields("wh_code") & "', " & _
'                                                " '" & RsD.Fields("item_code") & "', " & _
'                                                " " & LMPMonth & ", " & _
'                                                " " & LMRec & ", " & _
'                                                " " & LMSupply & ", "
'
'                                        SSql = SSql + " " & LMLossRej & ", " & _
'                                                " " & LMCur & ", " & _
'                                                " 0, " & _
'                                                " " & TMPMonth & ", " & _
'                                                " " & TMRec & ", " & _
'                                                " " & TMSupply & ", " & _
'                                                " " & TMLossRej & ", " & _
'                                                " " & TMCur & ", " & _
'                                                " 0, " & _
'                                                " " & NMPMonth & ", " & _
'                                                " " & NMRec & ", " & _
'                                                " " & NMSupply & ", " & _
'                                                " " & NMLossRej & ", "
'
'                                        SSql = SSql + " " & NMCur & ", " & _
'                                                " 0 " & _
'                                                " ) "
'                                        Db.Execute SSql
'                                    Else
'
'                                        LMPMonth = IIf(IsNull(RsStockHead.Fields("LM_PreMonth")), 0, RsStockHead.Fields("LM_PreMonth"))
'                                        LMRec = RsStockHead.Fields("LM_Receipt")
'                                        LMSupply = RsStockHead.Fields("LM_Supply")
'                                        LMLossRej = RsStockHead.Fields("LM_LossReject")
'                                        LMCur = RsStockHead.Fields("LM_Current")
'
'                                        TMPMonth = RsStockHead.Fields("TM_PreMonth")
'                                        TMRec = RsStockHead.Fields("TM_Receipt")
'                                        TMSupply = RsStockHead.Fields("TM_Supply")
'                                        TMLossRej = RsStockHead.Fields("TM_LossReject")
'                                        TMCur = RsStockHead.Fields("TM_Current")
'
'                                        NMPMonth = RsStockHead.Fields("NM_PreMonth")
'                                        NMRec = RsStockHead.Fields("NM_Receipt")
'                                        NMSupply = RsStockHead.Fields("NM_Supply")
'                                        NMLossRej = RsStockHead.Fields("NM_LossReject")
'                                        NMCur = RsStockHead.Fields("NM_Current")
'
'                                        If Sta = "T" Then
'                                            If KondisiStock = "0" Then
'                                                LMSupply = LMSupply + CDbl(RsD.Fields("qty"))
'                                            ElseIf KondisiStock = "1" Then
'                                                TMSupply = TMSupply + CDbl(RsD.Fields("qty"))
'                                            ElseIf KondisiStock = "2" Then
'                                                NMSupply = NMSupply + CDbl(RsD.Fields("qty"))
'                                            End If
'                                        ElseIf Sta = "K" Then
'                                            If KondisiStock = "0" Then
'                                                LMSupply = LMSupply - CDbl(RsD.Fields("qty"))
'                                            ElseIf KondisiStock = "1" Then
'                                                TMSupply = TMSupply - CDbl(RsD.Fields("qty"))
'                                            ElseIf KondisiStock = "2" Then
'                                                NMSupply = NMSupply - CDbl(RsD.Fields("qty"))
'                                            End If
'                                        End If
'
'                                        LMCur = (((LMPMonth + LMRec) - LMSupply) - LMLossRej)
'                                        If KondisiStock <> "0" Then
'                                            'TMPMonth = LMCur
'                                            TMCur = (((TMPMonth + TMRec) - TMSupply) - TMLossRej)
'                                            NMPMonth = TMCur
'                                            NMCur = (((NMPMonth + NMRec) - NMSupply) - NMLossRej)
'                                        End If
'
'                                        SintakSql = "UPDATE Stock_Master Set " & _
'                                                    "LM_Supply = " & LMSupply & ", " & _
'                                                    "TM_Supply = " & TMSupply & ", " & _
'                                                    "NM_Supply = " & NMSupply & ", " & _
'                                                    "LM_Current = " & LMCur & ", " & _
'                                                    "TM_PreMonth = " & TMPMonth & ", " & _
'                                                    "TM_Current =  " & TMCur & ", " & _
'                                                    "NM_PreMonth = " & NMPMonth & ", " & _
'                                                    "NM_Current = " & NMCur & " " & _
'                                                    "where warehouse_code = '" & RsD.Fields("wh_code") & "' " & _
'                                                    "and item_code = '" & RsD.Fields("item_code") & "'"
'                                        Db.Execute SintakSql
'
'                                    End If
'                    End If
'
                    
    ' INSERT PART SUPPLY
    '' -----------------------------
    ' Supply sudah dilakukan pada saat DO Status di FIX
    ' Update 20090202
    ' -----------------------------
    '                    If Sta = "T" Then
    '                        Dim Sn As Long
    '                        Sn = SSeqNo
    '                        'tambah
    '                        SSql = " INSERT  " & _
    '                        " INTO  " & _
    '                        " [dbo].[Part_Supply] " & _
    '                        " ( Seq_No, " & _
    '                        " [FromWarehouse_Code],  " & _
    '                        " [From_Address],  " & _
    '                        " [ToWarehouse_Code],  " & _
    '                        " [ChildSupply_date],  " & _
    '                        " [ChildItem_Code],  " & _
    '                        " [Supply_Cls],  "
    '
    '                        SSql = SSql + " [ChildRequirement_Qty],  " & _
    '                        " [ChildUnit_Cls],  " & _
    '                        " [Currency_Code],  " & _
    '                        " [Price],  " & _
    '                        " [Amount],  " & _
    '                        " [DO_No] " & _
    '                        " ) " & _
    '                        " VALUES " & _
    '                        " ( " & Sn & ","
    '
    '                        SSql = SSql + " '" & RsD.Fields("wh_code") & "' , " & _
    '                        " '" & RsD.Fields("address") & "' , " & _
    '                        " '" & Trim(grid.TextMatrix(i, ColConsignee)) & "' , " & _
    '                        " '" & RsD.Fields("Packing_Date") & "' , " & _
    '                        " '" & RsD.Fields("item_code") & "' , " & _
    '                        " 'D' , " & _
    '                        " " & RsD.Fields("qty") & " , " & _
    '                        " '" & RsD.Fields("unit_cls") & "' , " & _
    '                        " '" & RsD.Fields("currency_code") & "' , " & _
    '                        " " & RsD.Fields("price") & " , " & _
    '                        " " & RsD.Fields("amount") & " , " & _
    '                        " '" & RsD.Fields("packing_no") & "'  " & _
    '                        " ) "
    '
    '                        Db.Execute SSql
    '
    '                    ElseIf Sta = "K" Then
    '                        'hapus
    '                        SSql = "delete from part_supply where do_no = '" & RsD.Fields("packing_no") & "' " & _
    '                        "and childitem_code = '" & RsD.Fields("item_code") & "'"
    '
    '                        Db.Execute SSql
    '                    End If
'
'                    RsD.MoveNext
'                Wend
'            End If
            
        If grid.Cell(flexcpChecked, i, ColFix) = flexChecked And grid.Cell(flexcpChecked, i, ColFixT) = flexUnchecked Then
            SSql = "update packing_master set fix_cls = '1' where packing_no = '" & grid.TextMatrix(i, ColPac) & "' " & _
            "and cust_code = '" & Trim(grid.TextMatrix(i, ColCustCode)) & "'"
            booUpdate = True
        ElseIf grid.Cell(flexcpChecked, i, ColFix) = flexUnchecked And grid.Cell(flexcpChecked, i, ColFixT) = flexChecked Then
            SSql = "update packing_master set fix_cls = '0' where packing_no = '" & grid.TextMatrix(i, ColPac) & "' " & _
            "and cust_code = '" & Trim(grid.TextMatrix(i, ColCustCode)) & "'"
            booUpdate = True
        End If
        Db.Execute SSql
'    End If
Next
Db.CommitTrans
If booUpdate Then
    Header
    FillGrid
    LblErrMsg = DisplayMsg(1000)
End If
Me.MousePointer = vbDefault
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub DTDel1_Change()
Header
LblErrMsg = ""
End Sub

Private Sub DtDel2_Change()
Header
LblErrMsg = ""
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

ColPac = 0
ColCustCode = 1
ColCustName = 2
ColConsignee = 3
ColPacDate = 4
ColSail = 5
ColQty = 6
ColTtW = 7
ColTTWG = 8
ColTTVol = 9
ColIssue = 10
ColFix = 11
ColFixT = 12
colin = 13
KondisiStock = ""
'Conec
IsiCombo
Header
DTDel1.Value = Date - Day(Date) + 1
DtDel2.Value = Now
End Sub

Sub IsiCombo()
Dim rsisi As New ADODB.Recordset
Dim i As Long

'CUST
    cboCust.clear
    cboCust.columnCount = 2
    cboCust.ColumnWidths = "100 pt;330 pt"
    cboCust.ListWidth = 440
    cboCust.ListRows = 15
    cboCust.AddItem
    cboCust.List(0, 0) = strAll
    cboCust.List(0, 1) = strAll
    
    If rsisi.State = 1 Then rsisi.Close
    SSql = "select trade_code, trade_name from trade_master where (trade_cls = 2 or trade_cls = 4) --and country_cls='1'"
    rsisi.Open SSql, Db, adOpenKeyset, adLockOptimistic
    If Not rsisi.EOF And Not rsisi.BOF Then
        i = 1
        rsisi.MoveFirst
        While Not rsisi.EOF
            cboCust.AddItem
            cboCust.List(i, 0) = Trim(rsisi.Fields("trade_code"))
            cboCust.List(i, 1) = rsisi.Fields("trade_name")
            rsisi.MoveNext
            i = i + 1
        Wend
    End If
    cboCust.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'tidak perlu cek ada invoice
'lblErrMsg = ""
'If Grid.Col = ColFix Then
'    If Grid.TextMatrix(Row, colin) = "0" Then
'        Exit Sub
'    Else
'        lblErrMsg = DisplayMsg(4110)
'    End If
'End If
'Cancel = True

If Col = ColIssue Then
 Cancel = True
End If

If Col = ColFix Then
    LblErrMsg = up_ValidateDateRange(Format(grid.TextMatrix(Row, ColPacDate), "yyyy-MM-dd"), True)
    If LblErrMsg <> "" Then Cancel = True
End If
End Sub

Private Sub grid_Click()
LblErrMsg = ""
grid.FocusRect = flexFocusInset
If grid.Col = ColFix Then
    Exit Sub
End If
grid.FocusRect = flexFocusNone
End Sub

Private Function SSeqNo()
Dim rsmax As New ADODB.Recordset
Dim strSQL As String

strSQL = "Select Max(seq_No) from part_supply"

Set rsmax = Db.Execute(strSQL)

SSeqNo = IIf(IsNull(rsmax(0)), 1, rsmax(0) + 1)

End Function

