VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLotTraceInquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Lot Traceability Inquiry"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13950
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLotTraceInquiry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   13950
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   240
      Left            =   45
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   45
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8610
      Width           =   1365
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Left            =   12555
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8685
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1425
      Left            =   150
      TabIndex        =   5
      Top             =   1320
      Width           =   13545
      Begin VB.ComboBox cboExplosion 
         Height          =   315
         ItemData        =   "frmLotTraceInquiry.frx":0E42
         Left            =   6750
         List            =   "frmLotTraceInquiry.frx":0E4F
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   810
         Width           =   705
      End
      Begin VB.CommandButton CmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Search"
         Height          =   375
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   765
         Width           =   1635
      End
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Left            =   3690
         TabIndex        =   10
         Top             =   307
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Explosion     :"
         Height          =   195
         Index           =   1
         Left            =   5445
         TabIndex        =   19
         Top             =   870
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0: Explosion"
         Height          =   195
         Index           =   2
         Left            =   7785
         TabIndex        =   18
         Top             =   855
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1: Implosion"
         Height          =   195
         Index           =   3
         Left            =   9360
         TabIndex        =   17
         Top             =   855
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2: 1st Level Explosion"
         Height          =   195
         Index           =   4
         Left            =   10890
         TabIndex        =   16
         Top             =   855
         Width           =   1875
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   1
         Left            =   1605
         TabIndex        =   12
         Top             =   810
         Width           =   1995
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3519;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAAAAAAAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOT No            :"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   11
         Top             =   870
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code   :"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   360
         Width           =   1410
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   4335
         X2              =   7740
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblNm 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   4335
         TabIndex        =   6
         Top             =   360
         Width           =   3375
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   0
         Top             =   300
         Width           =   1995
         VariousPropertyBits=   746604571
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "3519;556"
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
      Left            =   11745
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   270
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   195
      TabIndex        =   8
      Top             =   7905
      Width           =   13500
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
         TabIndex        =   9
         Top             =   195
         Width           =   13320
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4785
      Left            =   135
      TabIndex        =   13
      Top             =   2925
      Width           =   13575
      _cx             =   23945
      _cy             =   8440
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
      GridColor       =   8421504
      GridColorFixed  =   8421504
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
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
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Traceability Inquiry"
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
      Left            =   5565
      TabIndex        =   4
      Top             =   450
      Width           =   2700
   End
End
Attribute VB_Name = "frmLotTraceInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public parent As String
Dim i As Integer
Dim sql As String

Dim bteColProdCode As Byte
Dim bteColDesc As Byte
Dim bteColMachNo As Byte
Dim bteColDate As Byte
Dim bteColLotNo As Byte
Dim bteColQty As Byte
Dim bteColRemark As Byte
Dim bteColParent As Byte

Private Sub headerGrid()
    Dim i As Integer
    
    bteColProdCode = 0
    bteColDesc = 1
    bteColMachNo = 2
    bteColDate = 3
    bteColLotNo = 4
    bteColQty = 5
    bteColRemark = 6
    bteColParent = 7
    
    With grid
        .clear
        .ColS = 8
        .Rows = 1
        
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColMachNo) = "Machine No"
        .TextMatrix(0, bteColDate) = "D a t e"
        .TextMatrix(0, bteColLotNo) = "Lot No"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColRemark) = "Remark"
        .TextMatrix(0, bteColParent) = "Parent"
        
        .ColWidth(bteColProdCode) = 2500
        .ColWidth(bteColDesc) = 3000
        .ColWidth(bteColMachNo) = 1500
        .ColWidth(bteColDate) = 2000
        .ColWidth(bteColLotNo) = 2000
        .ColWidth(bteColQty) = 1300
        .ColWidth(bteColRemark) = 2500
        .ColWidth(bteColParent) = 2500
        
        .ColHidden(bteColParent) = True
        
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColMachNo) = flexAlignCenterCenter
        .ColAlignment(bteColDate) = flexAlignCenterCenter
        .ColAlignment(bteColLotNo) = flexAlignCenterCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColRemark) = flexAlignLeftCenter
        
        .FrozenCols = bteColMachNo
        .EditMaxLength = 1
        .OutlineCol = bteColProdCode
        .OutlineBar = flexOutlineBarSimpleLeaf
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
    End With
    LblErrMsg = ""
End Sub

Sub Kosong()
    Cbo(0) = ""
    Cbo(1) = ""
    lblNm(0) = ""
    Text1 = ""
    cboExplosion.ListIndex = 0
    Call headerGrid
End Sub

Sub isiCboItem()
Dim rscbo As New ADODB.Recordset
    
    sql = "select Item_Code,ITem_name " & _
         "from Item_Master " & _
         "where use_endday >= convert(char(8), getdate(), 112) " & _
         "order by Item_Code"
    Set rscbo = Db.Execute(sql)
    
    Cbo(0).clear
    Cbo(0).columnCount = 2
    Cbo(0).TextColumn = 1
    
    i = 0
    Do While Not (rscbo.EOF)
        Cbo(0).AddItem ""
        Cbo(0).List(i, 0) = Trim(rscbo("Item_Code"))
        Cbo(0).List(i, 1) = Trim(rscbo("Item_Name"))
        i = i + 1
        rscbo.MoveNext
    Loop
    Cbo(0).ListWidth = 300
    Cbo(0).ColumnWidths = "100 pt;200 pt"
    Cbo(0).ListRows = 15
    Set rscbo = Nothing
    '************************
End Sub

Private Sub cboExplosion_Change()
Call headerGrid
Call isiCboLot
End Sub

Private Sub cboExplosion_Click()
Call headerGrid
Call isiCboLot
End Sub

Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = Cbo(0).Text
 frm_BrowseItem.Show 1
 Cbo(0).Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub cbo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbo_Click(Index)
End Sub

Private Sub cbo_Click(Index As Integer)
Dim rsIbu As New ADODB.Recordset

'If nilkosong Then Exit Sub

Select Case Index

Case Is = 0
If Cbo(0) <> "" Then
    Cbo(0) = Cbo(0)
    If Cbo(0).MatchFound = False Then
        lblNm(0) = ""
        lblNm(1) = ""
        LblErrMsg = DisplayMsg(4002)
        Call headerGrid
    Else
        lblNm(0) = Cbo(0).Column(1)
        LblErrMsg = ""
        Call headerGrid
        Call isiCboLot
    End If
Else
    lblNm(0) = ""
    LblErrMsg = ""
End If

Case Else
If Cbo(1) <> "" Then
    Cbo(1) = Cbo(1)
    If Cbo(1).MatchFound = False Then
        Text1 = ""
        LblErrMsg = DisplayMsg(4002)
        Call headerGrid
    Else
        Text1 = Cbo(1).Column(1)
        LblErrMsg = ""
        Call headerGrid
    End If
Else
    Text1 = ""
    LblErrMsg = ""
End If
End Select

End Sub

Private Sub cmdReport_Click()
Call Kosong
End Sub

Private Sub cmdSearch_Click()
Dim dsn As Long

Call headerGrid

dsn = IIf(Trim(Text1) = "", 0, Val(Trim(Text1)))

If cboExplosion = 0 Then 'cari anak
    Call IsiGrid(Cbo(0), 0, "ParentItem_Code", Cbo(1), dsn)
ElseIf cboExplosion = 1 Then 'cari parent
    Call IsiGrid(Cbo(0), 0, "ChildItem_Code", Cbo(1), dsn)
ElseIf cboExplosion = 2 Then 'cari anak tp 1 level
    Call IsiGrid(Cbo(0), 0, "ParentItem_Code", Cbo(1), dsn, 1)
End If

If grid.Rows <= 1 Then
    LblErrMsg = " Can not find data ! "
End If
End Sub

Private Sub CmdSubMenu_Click()
frmMainMenu.Show
Unload Me
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    Call Kosong
    Call isiCboItem
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

Sub isiCboLot()
Dim rscbo As New ADODB.Recordset

If cboExplosion = 1 Then

sql = " select distinct ps.childitem_code,ps.remarks lot_no,pr.dailyseq_no From part_supply ps " & _
        " inner join part_receipt pr " & _
        " on ps.parentitem_code=pr.item_code " & _
        "where ps.childitem_code='" & Cbo(0) & _
        "' and ps.remarks<>'' order by ps.remarks"
Else
sql = "select distinct item_code,lot_no,dailyseq_no from " & _
        "(select item_code, lot_no, seq_no dailyseq_no from daily_production " & _
        " Union All " & _
        " select item_code,suratjalan_no lot_no,dailyseq_no from part_receipt)n " & _
        "where Item_Code='" & Cbo(0) & _
        "' and Lot_No<>'' order by Lot_No"
End If

    Set rscbo = Db.Execute(sql)
    
    Cbo(1).clear
    Cbo(1).columnCount = 2
    Cbo(1).TextColumn = 1
    
    i = 0
    Do While Not (rscbo.EOF)
        Cbo(1).AddItem ""
        Cbo(1).List(i, 0) = Trim(rscbo("Lot_No"))
        Cbo(1).List(i, 1) = Trim(rscbo("dailyseq_no"))
        i = i + 1
        rscbo.MoveNext
    Loop
    Cbo(1).ListWidth = 100
    Cbo(1).ColumnWidths = "100 pt;0 pt"
    Cbo(1).ListRows = 15
    Set rscbo = Nothing

    
End Sub

Sub IsiGrid(ibu As String, lvl As Integer, nmField As String, LotItem As String, Optional dseq As Long, Optional explosion As Integer)

Dim anak As String, VSeqNo As Long, VLOT As String
Dim rsAnak As New ADODB.Recordset
Dim rsTglAwal As String, rsTglAkhir As String

If explosion = 1 And lvl = 1 Then Exit Sub

With grid
       

Dim pilih As String
       
If nmField = "ParentItem_Code" Then

sql = "select distinct ps.childitem_code,im.item_Name nmitem,dp.line_code,ps.childsupply_date,ps.lot_no,ps.remarks lno, " & _
        " ps.consumption_qty qty,pr.dailyseq_no Parent_seqno, pc.dailyseq_no child_seqno,pr.remarks,ps.parentitem_code " & _
        " from part_receipt pr " & _
        " inner join part_supply ps on pr.item_code = ps.parentitem_code and pr.suratjalan_no = ps.lot_no and cast(pr.seq_no as varchar) = ps.do_no " & _
        " inner join item_master im on ps.ChildItem_Code =im.item_code " & _
        " inner join daily_production dp on pr.dailyseq_no=dp.seq_no " & _
        " left join part_receipt pc on ps.ChildItem_Code = pc.item_code and ps.remarks = pc.suratjalan_no " & _
        " where ParentItem_Code='" & ibu & "' and ps.Lot_no ='" & LotItem & "'" & _
        " and pr.dailyseq_no = " & dseq & " "

sql = sql & " order by parentitem_code "

Else
sql = " select ps.childitem_code,im.item_name nmitem,dp.line_code,ps.childsupply_date,ps.lot_no,ps.remarks lno," & _
        " pr.qty,pr.remarks,ps.parentitem_code,pr.dailyseq_No child_seqno,prc.dailyseq_no parent_seqno" & _
        " from part_supply ps inner join part_receipt pr " & _
        " on ps.parentitem_code=pr.item_code " & _
        " and ps.lot_no=pr.suratjalan_no " & _
        " and ps.do_no=cast(pr.seq_no as varchar) " & _
        " left join part_supply psc on ps.parentitem_code=psc.childitem_code and ps.lot_no=psc.remarks " & _
        " left join part_receipt prc on psc.parentitem_code=prc.item_code and psc.lot_no=prc.suratjalan_no and psc.do_no =cast(prc.seq_no as varchar) " & _
        " inner join item_master im on ps.parentitem_code=im.item_code " & _
        " inner join daily_production dp on pr.dailyseq_no=dp.seq_no " & _
        " where ps.childitem_code='" & ibu & "' " & _
        " and ps.remarks='" & LotItem & "' " & _
        " and pr.dailyseq_no='" & dseq & "' "
End If


    Set rsAnak = Db.Execute(sql)
    
    If Not rsAnak.EOF Then
        Do While Not rsAnak.EOF
            .Rows = .Rows + 1
            
            If nmField = "ParentItem_Code" Then
                anak = IIf(IsNull(Trim(rsAnak("Childitem_code"))), "", Trim(rsAnak("Childitem_code")))
            Else
                anak = IIf(IsNull(Trim(rsAnak("ParentItem_Code"))), "", Trim(rsAnak("ParentItem_Code")))
            End If
            
            
            If nmField = "ParentItem_Code" Then
                VSeqNo = IIf(IsNull(Trim(rsAnak("Child_seqno"))), 0, Trim(rsAnak("Child_seqno")))
            Else
                VSeqNo = IIf(IsNull(Trim(rsAnak("Parent_seqno"))), 0, Trim(rsAnak("Parent_seqno")))
            End If
            
            If IsNull(VSeqNo) Then VSeqNo = 0
            
            .TextMatrix(.Rows - 1, bteColProdCode) = anak
            
            .TextMatrix(.Rows - 1, bteColDesc) = Trim(rsAnak("nmITem"))
            .TextMatrix(.Rows - 1, bteColMachNo) = IIf(IsNull(Trim(rsAnak("line_code"))), "", Trim(rsAnak("line_code")))
            .TextMatrix(.Rows - 1, bteColDate) = Format(rsAnak("childsupply_date"), "dd mmm yyyy")
            
            If nmField = "ParentItem_Code" Then
                VLOT = IIf(IsNull(Trim(rsAnak("Lno"))), "", Trim(rsAnak("Lno")))
                .TextMatrix(.Rows - 1, bteColLotNo) = IIf(IsNull(Trim(rsAnak("Lno"))), "", Trim(rsAnak("Lno")))
            Else
                .TextMatrix(.Rows - 1, bteColLotNo) = IIf(IsNull(Trim(rsAnak("Lot_no"))), "", Trim(rsAnak("Lot_no")))
                VLOT = IIf(IsNull(Trim(rsAnak("Lot_no"))), "", Trim(rsAnak("Lot_no")))
            End If
            
            .TextMatrix(.Rows - 1, bteColQty) = Format(rsAnak("qty"), gs_formatQty)
            .TextMatrix(.Rows - 1, bteColRemark) = IIf(IsNull(Trim(rsAnak("remarks"))), "", Trim(rsAnak("remarks")))
            .TextMatrix(.Rows - 1, bteColParent) = ibu
            
            .Col = 1
            .IsSubtotal(.Rows - 1) = True
            .RowData(.Rows - 1) = .Rows - 1
            .RowOutlineLevel(.Rows - 1) = lvl
            
            Call IsiGrid(anak, lvl + 1, nmField, VLOT, VSeqNo, explosion)
            If Not (rsAnak.EOF) Then rsAnak.MoveNext
        Loop
    End If
    Set rsAnak = Nothing
End With
End Sub


