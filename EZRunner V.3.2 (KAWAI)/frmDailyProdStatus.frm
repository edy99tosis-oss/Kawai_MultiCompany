VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDailyProdStatus 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Daily Production Schedule Status"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDailyProdStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Can&cel"
      Height          =   375
      Index           =   2
      Left            =   12780
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10005
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Index           =   1
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10005
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Search"
      Height          =   405
      Left            =   8970
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2460
      Width           =   1065
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10005
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   0
      Left            =   13935
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10005
      Width           =   1065
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10005
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   180
      TabIndex        =   22
      Top             =   9330
      Width           =   14865
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
         Left            =   105
         TabIndex        =   23
         Top             =   195
         Width           =   14640
      End
   End
   Begin VB.ComboBox cboRemaining 
      Height          =   315
      ItemData        =   "frmDailyProdStatus.frx":0E42
      Left            =   7725
      List            =   "frmDailyProdStatus.frx":0E4C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2505
      Width           =   885
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1170
      Left            =   180
      TabIndex        =   14
      Top             =   1125
      Width           =   14865
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   2940
         TabIndex        =   18
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machine No :"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   750
         Width           =   1110
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   1500
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2355;556"
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
         Height          =   315
         Index           =   1
         Left            =   1500
         TabIndex        =   1
         Top             =   690
         Width           =   1335
         VariousPropertyBits=   746604571
         MaxLength       =   3
         DisplayStyle    =   3
         Size            =   "2355;556"
         ListRows        =   15
         ShowDropButtonWhen=   2
         Value           =   "AAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   16
         Top             =   720
         Width           =   960
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   2940
         X2              =   4440
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory CD :"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   300
         Width           =   1095
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   2940
         X2              =   8010
         Y1              =   540
         Y2              =   540
      End
   End
   Begin MSComCtl2.DTPicker dtAwal 
      Height          =   330
      Left            =   1950
      TabIndex        =   2
      Top             =   2490
      Width           =   1785
      _ExtentX        =   3149
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
      CurrentDate     =   37860
   End
   Begin MSComCtl2.DTPicker dtAkhir 
      Height          =   330
      Left            =   4140
      TabIndex        =   3
      Top             =   2490
      Width           =   1785
      _ExtentX        =   3149
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
      CurrentDate     =   37891
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6255
      Left            =   180
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3030
      Width           =   14865
      _cx             =   26220
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
      AllowUserResizing=   3
      SelectionMode   =   1
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
      Left            =   13110
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   330
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Cls"
      Height          =   195
      Index           =   4
      Left            =   6360
      TabIndex        =   21
      Top             =   2565
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   195
      Index           =   3
      Left            =   3840
      TabIndex        =   20
      Top             =   2565
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule Date :"
      Height          =   195
      Index           =   2
      Left            =   450
      TabIndex        =   19
      Top             =   2565
      Width           =   1380
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Production Schedule Status"
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
      Left            =   180
      TabIndex        =   13
      Top             =   345
      Width           =   14865
   End
End
Attribute VB_Name = "frmDailyProdStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Long
Dim nilKosong As Boolean
Dim clsMRP As New clsMRP

Public fromProd As Boolean

Dim bteColDate As Byte
Dim bteColProdCode As Byte
Dim bteColPart As Byte
Dim bteColDesc As Byte
Dim bteColFNo As Byte
Dim bteColSerialNo As Byte
Dim bteColLotNo As Byte
Dim bteColCustCode As Byte
Dim bteColCustName As Byte
Dim bteColQty As Byte
Dim bteColResult As Byte
Dim bteColRemain As Byte
Dim bteColUnitCls As Byte
Dim bteColUnit As Byte
Dim bteColSeqNo As Byte
Dim bteColBlank1 As Byte
Dim bteColBlank2 As Byte
Dim bteColComplete As Byte
Dim bteColCheck As Byte

Private Sub headerGrid()
    Dim i As Integer
    
    bteColDate = 0
    bteColProdCode = 1
    bteColPart = 2
    bteColDesc = 3
    bteColFNo = 4
    bteColSerialNo = 5
    bteColLotNo = 6
    bteColCustCode = 7
    bteColCustName = 8
    bteColQty = 9
    bteColResult = 10
    bteColRemain = 11
    bteColUnitCls = 12
    bteColUnit = 13
    bteColSeqNo = 14
    bteColBlank1 = 15
    bteColBlank2 = 16
    bteColComplete = 17
    bteColCheck = 18
    
    With grid
        .clear
        
        .Rows = 1
        .ColS = 19
        
        .TextMatrix(0, bteColDate) = "Schedule Date"
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColPart) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColFNo) = "F No"
        .TextMatrix(0, bteColSerialNo) = "Serial No"
        .TextMatrix(0, bteColLotNo) = "Lot No"
        .TextMatrix(0, bteColCustCode) = "Cust Code"
        .TextMatrix(0, bteColCustName) = "Customer"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColResult) = "Result"
        .TextMatrix(0, bteColRemain) = "Remaining"
        .TextMatrix(0, bteColUnitCls) = "UnitCls"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColSeqNo) = "SeqNo"
        .TextMatrix(0, bteColBlank1) = ""
        .TextMatrix(0, bteColBlank2) = ""
        .TextMatrix(0, bteColComplete) = "Complete"
        .TextMatrix(0, bteColCheck) = "Check"
        
        .ColWidth(bteColDate) = 1400
        .ColWidth(bteColProdCode) = 1500
        .ColWidth(bteColPart) = 1500
        .ColWidth(bteColDesc) = 3500
        .ColWidth(bteColLotNo) = 1250
        .ColWidth(bteColQty) = 1250
        .ColWidth(bteColResult) = 1250
        .ColWidth(bteColRemain) = 1250
        .ColWidth(bteColUnit) = 600
        .ColWidth(bteColComplete) = 1000
        
        .ColHidden(bteColFNo) = True
        .ColHidden(bteColSerialNo) = True
        .ColHidden(bteColCustCode) = True
        .ColHidden(bteColCustName) = True
        .ColHidden(bteColUnitCls) = True
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColBlank1) = True
        .ColHidden(bteColBlank2) = True
        .ColHidden(bteColCheck) = True
        
        .ColDataType(bteColDate) = flexDTDate
        
        .Cell(flexcpAlignment, 0, 0, 0, 13) = flexAlignCenterCenter
        .ColAlignment(bteColDate) = flexAlignCenterCenter
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColPart) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColLotNo) = flexAlignCenterCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColResult) = flexAlignRightCenter
        .ColAlignment(bteColRemain) = flexAlignRightCenter
        .ColAlignment(bteColUnit) = flexAlignLeftCenter
        .ColAlignment(bteColComplete) = flexAlignCenterCenter
        
        .EditMaxLength = 1
    End With
End Sub

Sub Kosong()
    nilKosong = True
    cbo(0) = ""
    lblNm(0) = ""
    cbo(1) = ""
    lblNm(1) = ""
    dtAwal = Format(Year(Now) & "-" & Format(Month(Now), "#0") & "-01", "dd MMM yyyy")
    dtAkhir = Format(Now, "dd MMM yyyy")
    cboRemaining.ListIndex = 0
    Call headerGrid
    nilKosong = False
End Sub

'******** Combo Factory Code **********
Sub isiCboCust() 'Isi Combo Dealer CD dr Customer Master
Dim RsCust As New ADODB.Recordset 'Data Customer

With cbo(0)
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

'******** Filter Combo Line Code **********
Sub isiCboLine(factoryCD As String)
Dim rscbo As New ADODB.Recordset

With cbo(1)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    sql = "select Line_Code,Line_Name from Manufacture_line " & _
        "where Manufacture_Code = '" & factoryCD & _
        "' order by Line_Code"
    Set rscbo = Db.Execute(sql)
     
    i = 0
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo(0))
        .List(i, 1) = Trim(rscbo(1))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 200
    .ColumnWidths = "50 pt;150 pt"
    
    Set rscbo = Nothing
End With
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    fromProd = True
    nilKosong = True
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
'    frm_part_supply.ls_dataStatus = "invalid"
    Call isiCboCust
    Call Kosong
    dtAwal = Date
    dtAkhir = Date
    nilKosong = False
    
End Sub

Private Sub cbo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
If nilKosong = True Then Exit Sub
    If KeyCode = 13 Then Call cbo_Click(Index)
End Sub

Private Sub cbo_Change(Index As Integer)
If nilKosong = True Then Exit Sub
    lblNm(Index) = ""
    'Hapus Manufacture Line * Desc
    If Index = 0 Then cbo(1).clear: lblNm(1) = "": Call headerGrid
End Sub

Private Sub cbo_LostFocus(Index As Integer)
If nilKosong = True Then Exit Sub
    If lblNm(Index) = "" Then Call cbo_Click(Index)
End Sub

'*********** Tampilkan Data *********
Private Sub cbo_Click(Index As Integer)
If nilKosong = True Then Exit Sub

If cbo(Index) <> "" Then
    cbo(Index) = cbo(Index)
    If cbo(Index).MatchFound = True Then
        lblNm(Index) = cbo(Index).Column(1)
        If Index = 0 Then 'panggil Manufacture Line
            Call isiCboLine(cbo(0)): lblNm(1) = ""
        End If
        LblErrMsg = ""
    Else
        lblNm(Index) = ""
        If Index = 0 Then 'Hapus Manufacture Line & Desc Line
            cbo(1).clear: lblNm(1) = ""
        End If
        LblErrMsg = DisplayMsg(4016 + Index) 'Err Msg en Panggil Grid
    End If
Else
    lblNm(Index) = ""
    If Index = 0 Then 'Hapus Manufacture Line * Desc
        cbo(1).clear: lblNm(1) = ""
    End If
    LblErrMsg = ""
End If
End Sub

Public Sub cmdSearch_Click()
    cbo(0) = cbo(0)
    cbo(1) = cbo(1)
    If cbo(0) = "" Then
        LblErrMsg = DisplayMsg(1040)
        cbo(0).SetFocus
    ElseIf cbo(0).MatchFound = False Then
        LblErrMsg = DisplayMsg(4016)
        cbo(0).SetFocus
    ElseIf cbo(1).MatchFound = False And cbo(1) <> "" Then
        LblErrMsg = DisplayMsg(4017)
        cbo(1).SetFocus
    Else
        LblErrMsg = ""
        Call IsiGrid
    End If
    
End Sub

Sub IsiGrid()
Dim rsGrid As New ADODB.Recordset
Dim sqlGrid As String

If nilKosong = True Then Exit Sub

Call headerGrid
    
    sqlGrid = "select (select isnull(sum(qty),0) Qtyresult from part_receipt " & _
                    "where receipt_cls='P1' and dailySeq_no = a.seq_no) QtyResult, " & _
                    "a.qty - (select isnull(sum(qty),0) Qtyresult from part_receipt " & _
                    "where receipt_cls='P1' and dailySeq_no = a.seq_no) Remaining, a.*, b.item_name, b.makeritem_code, b.finishgoodpart_cls from daily_production a left join item_master b on b.item_code=a.item_code " & _
              "where a.factory_code='" & cbo(0).Text & "' and " & _
                    "a.line_code='" & cbo(1).Text & "' and a.schedule_date>='" & Format(dtAwal.Value, "yyyymmdd") & _
                    "' and a.schedule_date<='" & Format(dtAkhir.Value, "yyyymmdd") & "'"
    
    If cboRemaining = "Yes" Then
        sqlGrid = sqlGrid & " And a.qty - (select isnull(sum(qty),0) Qtyresult from part_receipt where receipt_cls='P1' and dailySeq_no = a.seq_no) > 0 and (complete_cls is null or complete_cls = 0)"
    Else
        sqlGrid = sqlGrid & " And (a.qty - (select isnull(sum(qty),0) Qtyresult from part_receipt where receipt_cls='P1' and dailySeq_no = a.seq_no) <= 0 or complete_cls = 1)"
    End If
    
    sqlGrid = sqlGrid + " order by a.schedule_date, a.item_code, a.lot_no, a.seq_no"
    
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
    
    If rsGrid.EOF Then
        LblErrMsg = DisplayMsg(4006)
        Exit Sub
    End If
    
    i = 1
    If Not (rsGrid.BOF And rsGrid.EOF) Then
    With grid
    Do While Not rsGrid.EOF
        .Rows = .Rows + 1
        .TextMatrix(i, bteColDate) = Format(Trim(rsGrid("schedule_date")), "dd MMM yyyy")
        .TextMatrix(i, bteColProdCode) = Trim(rsGrid("Item_Code"))
        .TextMatrix(i, bteColPart) = Trim(rsGrid("MakerItem_Code"))
        .TextMatrix(i, bteColDesc) = IIf(IsNull(rsGrid("item_name")), " ", Trim(rsGrid("item_name")))
        .TextMatrix(i, bteColLotNo) = IIf(IsNull(rsGrid("lot_no")), " ", Trim(rsGrid("lot_no")))
        .TextMatrix(i, bteColQty) = IIf(IsNull(rsGrid("Qty")), 0, Trim(rsGrid("Qty")))
        If InStr(1, .TextMatrix(i, bteColQty), ".") > 0 Then
            .TextMatrix(i, bteColQty) = Format(.TextMatrix(i, bteColQty), gs_formatQty)
        Else
            .TextMatrix(i, bteColQty) = Format(.TextMatrix(i, bteColQty), gs_formatQty)
        End If
        .TextMatrix(i, bteColResult) = IIf(IsNull(rsGrid("qtyresult")), " ", Trim(rsGrid("qtyresult")))
        If InStr(1, .TextMatrix(i, bteColResult), ".") > 0 Then
            .TextMatrix(i, bteColResult) = Format(.TextMatrix(i, bteColResult), gs_formatQty)
        Else
            .TextMatrix(i, bteColResult) = Format(.TextMatrix(i, bteColResult), gs_formatQty)
        End If
        .TextMatrix(i, bteColRemain) = IIf(IsNull(rsGrid("remaining")), " ", Trim(rsGrid("remaining")))
        If InStr(1, .TextMatrix(i, bteColRemain), ".") > 0 Then
            .TextMatrix(i, bteColRemain) = Format(.TextMatrix(i, bteColRemain), gs_formatQty)
        Else
            .TextMatrix(i, bteColRemain) = Format(.TextMatrix(i, bteColRemain), gs_formatQty)
        End If
        If IsNull(rsGrid("unit_cls")) Then
          .TextMatrix(i, bteColUnitCls) = " "
          .TextMatrix(i, bteColUnit) = " "
        Else
          .TextMatrix(i, bteColUnitCls) = Trim(rsGrid("Unit_cls"))
          .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(Trim(rsGrid("Unit_Cls")))
        End If
        
        .TextMatrix(i, bteColSeqNo) = Val(rsGrid("seq_no"))

        .Cell(flexcpBackColor, i, bteColComplete) = vbWhite
        If (IsNull(rsGrid("complete_cls"))) Or (rsGrid("complete_cls")) = 0 Then
            .Cell(flexcpChecked, i, bteColComplete) = flexUnchecked
        Else
            .Cell(flexcpChecked, i, bteColComplete) = flexChecked
        End If
        .Cell(flexcpChecked, i, bteColCheck) = flexUnchecked
        rsGrid.MoveNext
        i = i + 1
    Loop
    End With
    End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col <> bteColComplete Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim tampung As Long
With grid
    If Row <> 0 And Col = bteColComplete Then
        If .Cell(flexcpChecked, Row, bteColCheck) = flexUnchecked Then
            .Cell(flexcpChecked, Row, bteColCheck) = flexChecked
        Else
            .Cell(flexcpChecked, Row, bteColCheck) = flexUnchecked
        End If
    End If
End With
LblErrMsg = ""
End Sub

Private Sub Command1_Click(Index As Integer)
Dim sqlUpdate   As String
Dim strTemp     As String
Dim rsTemp      As New ADODB.Recordset
Dim updateValue As Integer
Dim dbw As New Connection
    
dbw.ConnectionString = Db.ConnectionString
    
    'SUBMIT
    If Index = 0 Then
        dbw.Open
        dbw.BeginTrans
        For i = 1 To grid.Rows - 1
            If grid.Cell(flexcpChecked, i, bteColCheck) = flexChecked Then
                If grid.Cell(flexcpChecked, i, bteColComplete) = flexChecked Then
                    updateValue = 1
                    
                    'JIKA REMAINING > 0
                    If grid.TextMatrix(i, bteColRemain) > 0 Then
                        'UPDATE OFF_QTY DAN CHILD_OFF_QTY PADA REQUIREMENT
                        sqlUpdate = "UPDATE requirement " & _
                                    "SET Off_Qty = Off_Qty + " & CDbl(grid.TextMatrix(i, bteColRemain)) & ", " & _
                                    "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                    "WHERE ParentItem_Code = '" & grid.TextMatrix(i, bteColProdCode) & "' " & _
                                    "AND Lot_No = '" & grid.TextMatrix(i, bteColLotNo) & "' " & _
                                    "AND Production_Date = '" & Format(grid.TextMatrix(i, bteColDate), "yyyy-MM-dd") & "'"
                        dbw.Execute sqlUpdate
                        sqlUpdate = "UPDATE requirement " & _
                                    "SET OffChildRequirement_Qty = ChildRequirement_Qty * (Off_Qty / Qty), " & _
                                    "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                    "WHERE ParentItem_Code = '" & grid.TextMatrix(i, bteColProdCode) & "' " & _
                                    "AND Lot_No = '" & grid.TextMatrix(i, bteColLotNo) & "' " & _
                                    "AND Production_Date = '" & Format(grid.TextMatrix(i, bteColDate), "yyyy-MM-dd") & "' "
                        dbw.Execute sqlUpdate
                        'AUTO COMPLETE PADA REQUIREMENT JIKA QTY = OFF_QTY
                        sqlUpdate = "UPDATE requirement " & _
                                    "SET Complete_Cls = 1, " & _
                                    "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                    "WHERE ParentItem_Code = '" & grid.TextMatrix(i, bteColProdCode) & "' " & _
                                    "AND Lot_No = '" & grid.TextMatrix(i, bteColLotNo) & "' " & _
                                    "AND Production_Date = '" & Format(grid.TextMatrix(i, bteColDate), "yyyy-MM-dd") & "' " & _
                                    "AND Qty = Off_Qty"
                        dbw.Execute sqlUpdate
                    End If
                Else
                    updateValue = 0
                    If grid.TextMatrix(i, bteColRemain) > 0 Then
                        'UPDATE OFF_QTY DAN CHILD_OFF_QTY PADA REQUIREMENT
                        sqlUpdate = "UPDATE requirement " & _
                                    "SET Off_Qty = Off_Qty - " & CDbl(grid.TextMatrix(i, bteColRemain)) & ", " & _
                                    "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                    "WHERE ParentItem_Code = '" & grid.TextMatrix(i, bteColProdCode) & "' " & _
                                    "AND Lot_No = '" & grid.TextMatrix(i, bteColLotNo) & "' " & _
                                    "AND Production_Date = '" & Format(grid.TextMatrix(i, bteColDate), "yyyy-MM-dd") & "'"
                        dbw.Execute sqlUpdate
                        sqlUpdate = "UPDATE requirement " & _
                                    "SET OffChildRequirement_Qty = ChildRequirement_Qty * (Off_Qty / Qty), " & _
                                    "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                    "WHERE ParentItem_Code = '" & grid.TextMatrix(i, bteColProdCode) & "' " & _
                                    "AND Lot_No = '" & grid.TextMatrix(i, bteColLotNo) & "' " & _
                                    "AND Production_Date = '" & Format(grid.TextMatrix(i, bteColDate), "yyyy-MM-dd") & "' "
                        dbw.Execute sqlUpdate
                    End If
                End If
                ''''''
                strTemp = "SELECT * FROM requirement " & _
                          "WHERE ParentItem_Code = '" & grid.TextMatrix(i, bteColProdCode) & "' AND " & _
                                "Lot_No = '" & grid.TextMatrix(i, bteColLotNo) & "' AND " & _
                                "Production_Date = '" & Format(grid.TextMatrix(i, bteColDate), "yyyy-MM-dd") & "' "
                If rsTemp.State = adStateOpen Then rsTemp.Close
                rsTemp.Open strTemp, dbw, adOpenKeyset, adLockOptimistic
                Do While Not rsTemp.EOF
                    clsMRP.UpdateRequirementResult dbw, Format(grid.TextMatrix(i, bteColDate), "yyyy-MM-dd"), "'" & grid.TextMatrix(i, bteColProdCode) & "'", grid.TextMatrix(i, bteColLotNo), "'" & rsTemp("ChildItem_Code") & "'"
                    rsTemp.MoveNext
                Loop
                ''''''
                sqlUpdate = "UPDATE daily_production " & _
                                "SET complete_cls = " & updateValue & ", Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                "WHERE Seq_no = " & grid.TextMatrix(i, bteColSeqNo)
                dbw.Execute sqlUpdate
                
                'AUTO COMPLETE PADA MRP JIKA SEMUA RENCANA PADA DAILY SUDAH COMPLETE
                sqlUpdate = " UPDATE requirement SET complete_cls = 1, " & _
                            " Last_Update = getdate(), Last_User = '" & userLogin & "'" & _
                            " WHERE ParentItem_Code = '" & grid.TextMatrix(i, bteColProdCode) & "'  " & _
                            "   AND Lot_No = '" & grid.TextMatrix(i, bteColLotNo) & "'  " & _
                            "   AND Production_Date = '" & Format(grid.TextMatrix(i, bteColDate), "yyyy-MM-dd") & "' " & _
                            "   AND NOT EXISTS  ( " & _
                            "           SELECT * FROM daily_production  " & _
                            "           WHERE   item_code = '" & grid.TextMatrix(i, bteColProdCode) & "'  " & _
                            "               AND Lot_No = '" & grid.TextMatrix(i, bteColLotNo) & "'  " & _
                            "               AND schedule_date = '" & Format(grid.TextMatrix(i, bteColDate), "yyyy-MM-dd") & "'  " & _
                            "               AND (complete_cls = NULL OR complete_cls = 0) " & _
                            "           )  "
                dbw.Execute sqlUpdate
                
                grid.Cell(flexcpChecked, i, bteColCheck) = flexUnchecked
            End If
        Next i
        
        dbw.CommitTrans
        dbw.Close
        
        'REFRESH DATA GRID
        If cbo(0) = "" Then
            LblErrMsg = DisplayMsg(1040)
            cbo(0).SetFocus
        ElseIf cbo(0).MatchFound = False Then
            LblErrMsg = DisplayMsg(4016)
            cbo(0).SetFocus
        ElseIf cbo(1).MatchFound = False And cbo(1) <> "" Then
            LblErrMsg = DisplayMsg(4017)
            cbo(1).SetFocus
        Else
            Call IsiGrid
            LblErrMsg = DisplayMsg(1101)
        End If
    'CLEAR
    ElseIf Index = 1 Then
        cbo(0) = ""
        cbo(1) = ""
        dtAwal = Date
        dtAkhir = Date
        LblErrMsg = ""
        Call headerGrid
    'RESET
    ElseIf Index = 2 Then
        With grid
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, bteColCheck) = flexChecked Then
                    If .Cell(flexcpChecked, i, bteColComplete) = flexUnchecked Then
                        .Cell(flexcpChecked, i, bteColComplete) = flexChecked
                    Else
                        .Cell(flexcpChecked, i, bteColComplete) = flexUnchecked
                    End If
                    .Cell(flexcpChecked, i, bteColCheck) = flexUnchecked
                End If
            Next i
        End With
    End If

End Sub

Private Sub cmdReport_Click()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3
Dim sqlResult As String

    Me.MousePointer = vbHourglass
    
    If cbo(0) = "" Then
        LblErrMsg = DisplayMsg(1040)
        cbo(0).SetFocus
    Else
        cbo(0) = cbo(0)
        If cbo(0).MatchFound = False Then
            LblErrMsg = DisplayMsg(4016)
            cbo(0).SetFocus
        Else
            LblErrMsg = ""
            '**** Utk Report
            
'            Sql = "Select z.* from(" & _
'                "select rtrim(Factory_Code) Factory_Code, rtrim(C.Trade_Name) as TradeFactory, a.Line_Code, Schedule_Date, a.Item_Code, rtrim(b.Item_Name) Descr, Lot_No, a.qty plann, a.complete_cls, " & _
'                "(select Isnull(Sum(Qty), 0) from Part_Receipt where DailySeq_No = a.Seq_No And ProductionResult_Cls = 1) as Result, " & _
'                "(Qty - (select Isnull(Sum(Qty),0) from Part_Receipt where DailySeq_No = a.Seq_No And ProductionResult_Cls = 1)  ) as Sisa, WH_Code " & _
'                "from Daily_Production a, Item_Master b,Trade_Master C " & _
'                "Where a.Item_Code = B.Item_Code And a.Factory_Code = C.Trade_Code " & _
'                "And Schedule_Date >= '" & Format(DtAwal, "yyyy-MM-dd") & "' and Schedule_Date <= '" & Format(DtAkhir, "yyyy-MM-dd") & "' " & _
'                "and Factory_Code = '" & Trim(cbo(0)) & "' and a.Line_Code = '" & Trim(cbo(1)) & "') z "
'
'            If cboRemaining.Text = "Yes" Then
'                Sql = Sql & _
'                    "where sisa > 0 and isnull(complete_cls, 0) = 0"
'            Else
'                Sql = Sql & _
'                    "where sisa <= 0 or complete_cls = 1"
'            End If
            
            
        sql = " Select z.* from(  " & vbCrLf & _
                    " select rtrim(a.factory_code) as factory_code, rtrim(a.line_code)  " & vbCrLf & _
                    "   as line_code, a.schedule_date,   a.SerialNoFrom PSerialFrom,a.SerialNoTo PSerialTo,     " & vbCrLf & _
                    "   rtrim(b.makeritem_code) makeritem_code, rtrim(a.item_code) as item_code,  " & vbCrLf & _
                    "   rtrim(a.lot_no) as lot_no, a.seq_no, a.qty, rtrim(a.unit_cls) as unit_cls,  " & vbCrLf & _
                    "   (select description from unit_cls uc where uc.unit_cls=a.unit_cls)  unit_desc,     " & vbCrLf & _
                    "       a.remark, b.item_name, (select trade_name from trade_master where trade_code = '" & Trim(cbo(0)) & "')  " & vbCrLf & _
                    "   factory_name, (select line_name from manufacture_line where manufacture_code = '" & Trim(cbo(0)) & "'  " & vbCrLf & _
                    "   and line_code = '" & Trim(cbo(1)) & "') line_name,     c.receipt_date, b.number_entering,  a.complete_cls,  " & vbCrLf & _
                    "   c.SerialNoFrom RSerialFrom,C.SerialNoTo RSerialTo, isnull(c.Qty,0) QtyResult,   " & vbCrLf & _
                    "   (select rtrim(production_person) from company_profile) Prod_Person,     "
        
        sql = sql + "   (select rtrim(production_position) from company_profile) production_Position,      " & vbCrLf & _
                    "   (select rtrim(QC_person) from company_profile) QC_Person,     " & vbCrLf & _
                    "   (select rtrim(QC_position) from company_profile) QC_Position,       " & vbCrLf & _
                    "   (select rtrim(PPC_person)from company_profile) PPC_Person,      " & vbCrLf & _
                    "   (select rtrim(PPC_position) from company_profile) PPC_Position,      " & vbCrLf & _
                    "   (select isnull(sum(qty),0) from part_receipt where dailyseq_no = a.seq_no  " & vbCrLf & _
                    "   and receipt_cls = 'P1' and receipt_date= c.receipt_date) QtyResult1   " & vbCrLf & _
                    " from daily_production a      " & vbCrLf & _
                    " left join item_master b on b.item_code=a.item_code      " & vbCrLf & _
                    " left join part_receipt c on (c.dailyseq_no = a.seq_no and c.receipt_cls = 'P1')   " & vbCrLf & _
                    " where a.factory_code='" & Trim(cbo(0)) & "' and a.line_code='" & Trim(cbo(1)) & "' and a.schedule_date>='" & Format(dtAwal, "yyyy-MM-dd") & "'  "
        
        sql = sql + "   and a.schedule_date<='" & Format(dtAkhir, "yyyy-MM-dd") & "' ) z  "
        
            If cboRemaining.Text = "Yes" Then
                sql = sql & _
                    " where z.Qty - z.QtyResult > 0 And  " & vbCrLf & _
                    "   (z.Complete_Cls is Null Or z.Complete_Cls = 0)   "
            Else
                sql = sql & _
                    " where z.Qty - z.QtyResult <= 0 or complete_cls = 1"
            End If
        
            sql = sql & _
                "order by Factory_Code, Line_Code, Schedule_Date, Item_Code, Lot_no "
            
            Set rsRpt = Db.Execute(sql)
            
            If rsRpt.EOF Then
                LblErrMsg.Caption = DisplayMsg(4006)
            Else
                sqlprint = sql
                reportcode = "rptProdResultInquiry"
                tglAwalRptPrint = Format(dtAwal, "dd MMM yyyy")
                tglAkhirRptPrint = Format(dtAkhir, "dd MMM yyyy")
                
                Set report = application.OpenReport(App.path & "\Reports\rptProdResultInquiry.rpt")
                report.Database.Tables(1).SetDataSource rsRpt
                report.FormulaFields(1).Text = "'" & Format(dtAwal, "dd MMM yyyy") & "'"
                report.FormulaFields(2).Text = "'" & Format(dtAkhir, "dd MMM yyyy") & "'"
                report.FormulaFields(4).Text = gi_decimalDigitQty
                'report.FormulaFields(5).Text = "'" & cboRemaining & "'"
                            
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
    If cmdsubmenu.Caption = "&Back" Then
        Call Command1_Click(1)
    Else
        Unload frmProdResult
        Unload frm_part_supply
                    
        DoEvents
        frmMainMenu.Show
        DoEvents
        Unload Me
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
'**************

