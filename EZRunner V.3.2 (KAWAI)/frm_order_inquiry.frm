VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_order_inquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Order Inquiry"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frm_order_inquiry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      Left            =   12570
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9840
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13050
      TabIndex        =   25
      Top             =   120
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080FFFF&
      Caption         =   "Searc&h"
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
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1155
   End
   Begin VB.CommandButton cmd_last 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last Page"
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9840
      Width           =   1170
   End
   Begin VB.CommandButton cmd_next 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Next Page"
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
      Left            =   7530
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9840
      Width           =   1170
   End
   Begin VB.CommandButton cmd_previous 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Prev Page"
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
      Left            =   6270
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9840
      Width           =   1170
   End
   Begin VB.CommandButton cmd_first 
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Page"
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
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9840
      Width           =   1170
   End
   Begin VB.CommandButton cmd_sub_menu 
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
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9840
      Width           =   1125
   End
   Begin VB.CommandButton cmd_submit 
      BackColor       =   &H0080FFFF&
      Caption         =   "To &Update"
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
      Left            =   13770
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9840
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   750
      Left            =   210
      TabIndex        =   14
      Top             =   915
      Width           =   14700
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer CD"
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
         Left            =   135
         TabIndex        =   17
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label LblArea 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   6540
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   3135
         X2              =   9120
         Y1              =   540
         Y2              =   540
      End
      Begin MSForms.ComboBox ComboBox1 
         Height          =   315
         Left            =   1470
         TabIndex        =   0
         Top             =   270
         Width           =   1575
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2778;556"
         ColumnCount     =   2
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lbl_name 
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_name"
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
         Left            =   3135
         TabIndex        =   15
         Top             =   315
         Width           =   6000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   195
      TabIndex        =   12
      Top             =   9120
      Width           =   14715
      Begin VB.Label lbl_pesan 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_pesan"
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
         Height          =   255
         Left            =   90
         TabIndex        =   13
         Top             =   225
         Width           =   14445
      End
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   345
      Left            =   3645
      TabIndex        =   2
      Top             =   1845
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   609
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
      Format          =   287637507
      CurrentDate     =   37818
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   345
      Left            =   1665
      TabIndex        =   1
      Top             =   1845
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   609
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
      Format          =   287637507
      CurrentDate     =   37818
   End
   Begin VSFlex8Ctl.VSFlexGrid vsflexGrid1 
      Height          =   6615
      Left            =   195
      TabIndex        =   24
      Top             =   2340
      Width           =   14685
      _cx             =   25903
      _cy             =   11668
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
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
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
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   275
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm_order_inquiry.frx":0E42
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
   Begin VB.Label lbl_cls 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_cls"
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
      Left            =   7920
      TabIndex        =   23
      Top             =   1890
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   7920
      X2              =   8550
      Y1              =   2115
      Y2              =   2115
   End
   Begin MSForms.ComboBox ComboBox2 
      Height          =   315
      Left            =   6975
      TabIndex        =   3
      Top             =   1845
      Width           =   885
      VariousPropertyBits=   746604571
      MaxLength       =   7
      DisplayStyle    =   3
      Size            =   "1561;556"
      ColumnCount     =   2
      ListRows        =   15
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Cls "
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
      Left            =   5535
      TabIndex        =   22
      Top             =   1890
      Width           =   1290
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   3255
      TabIndex        =   21
      Top             =   1875
      Width           =   210
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
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
      Left            =   270
      TabIndex        =   20
      Top             =   1875
      Width           =   1275
   End
   Begin VB.Label lbl_record 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_record"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   1815
      TabIndex        =   19
      Top             =   9900
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Order Inquiry"
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
      Index           =   0
      Left            =   180
      TabIndex        =   18
      Top             =   240
      Width           =   14730
   End
End
Attribute VB_Name = "frm_order_inquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim rs_order_master As New ADODB.Recordset
Dim rs_order_detail As New ADODB.Recordset
Dim rs_join As New ADODB.Recordset
Dim rs_trade_master As New ADODB.Recordset
Dim update As String, kanan_pertama As Integer, l_combo1 As String
Dim sql_join As String

Dim bteColSelect As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColPONo As Byte
Dim bteColPlan As Byte
Dim BteColSerialFrom As Byte
Dim BteColSerialTo As Byte
Dim bteColResult As Byte
Dim bteColRemaining As Byte
Dim bteColDelDate As Byte
Dim bteColDelTime As Byte
Dim bteColCurr As Byte
Dim bteColCur1 As Byte
Dim bteColService As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColSeqNo As Byte
Dim bteColItemCode As Byte

Dim bteHakPrice As Byte

Private Sub Header()
    With VSFlexGrid1
        
        bteColSelect = 0
        bteColPartNo = 1
        bteColDesc = 2
        bteColPONo = 3
        bteColPlan = 4
        BteColSerialFrom = 5
        BteColSerialTo = 6
        bteColResult = 5 + 2
        bteColRemaining = 6 + 2
        bteColDelDate = 7 + 2
        bteColDelTime = 8 + 2
        bteColCurr = 9 + 2
        bteColPrice = 10 + 2
        'bteColCur1 = 11+2
        bteColService = 11 + 2
        bteColAmount = 12 + 2
        bteColSeqNo = 13 + 2
        bteColItemCode = 14 + 2
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColPartNo) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColPONo) = "SI/PO No"
        .TextMatrix(0, bteColPlan) = "Plan"
        '---
        .TextMatrix(0, BteColSerialFrom) = "Serial From"
        .TextMatrix(0, BteColSerialTo) = "Serial To"
        '---
        .TextMatrix(0, bteColResult) = "Result"
        .TextMatrix(0, bteColRemaining) = "Remaining"
        .TextMatrix(0, bteColDelDate) = "Delivery Date"
        .TextMatrix(0, bteColDelTime) = "Time"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        '.TextMatrix(0, bteColCur1) = "Curr"
        .TextMatrix(0, bteColService) = "Service"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColSeqNo) = "SeqNo"
        .TextMatrix(0, bteColItemCode) = "itemCode"
        
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColItemCode) = True
        
        .ColAlignment(bteColPartNo) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColPONo) = flexAlignLeftCenter
        .ColAlignment(bteColPlan) = flexAlignRightCenter
        .ColAlignment(bteColResult) = flexAlignRightCenter
        .ColAlignment(bteColRemaining) = flexAlignRightCenter
        .ColAlignment(bteColPrice) = flexAlignRightCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter

        
        .ColWidth(bteColSelect) = 250
        .ColWidth(bteColPartNo) = 2000
        .ColWidth(bteColDesc) = 3600
        .ColWidth(bteColPONo) = 2000
        .ColWidth(bteColPlan) = 850
        ' ---
        .ColWidth(BteColSerialFrom) = 1100
        .ColWidth(BteColSerialTo) = 1100
        ' ---
        .ColWidth(bteColResult) = 850
        .ColWidth(bteColRemaining) = 1000
        .ColWidth(bteColDelDate) = 1300
        .ColWidth(bteColDelTime) = 750
        .ColWidth(bteColCurr) = 750
        '.ColWidth(bteColCur1) = 750
        .ColWidth(bteColPrice) = 1500
        .ColWidth(bteColService) = 1600
        .ColWidth(bteColAmount) = 2000
        
        .ColHidden(bteColCurr) = (bteHakPrice = 0)
        .ColHidden(bteColPrice) = (bteHakPrice = 0)
        .ColHidden(bteColAmount) = (bteHakPrice = 0)
        .ColHidden(bteColService) = True

        .EditMaxLength = 1
    End With
End Sub

Public Sub cmdSearch_Click(Index As Integer)

MousePointer = vbHourglass
If Trim(ComboBox1.Text) = "" Then
    lbl_pesan = DisplayMsg(1033) '"Please select customer code !"
Else
    Call set_cust
    Call setting_grid
    lbl_pesan = ""
    If ComboBox1.MatchFound Then
       lbl_name = ComboBox1.List(ComboBox1.ListIndex, 1)
        If VSFlexGrid1.Rows = 1 Then
            lbl_pesan = DisplayMsg(4006) '"customer Code is not found !"
        Else
        lbl_pesan = ""
        End If
    Else
       lbl_name = ""
       lbl_pesan = DisplayMsg(4006) '"customer Code is not found !"
    End If
End If
MousePointer = vbDefault
End Sub

Private Sub ComboBox2_Change()
Call clearGrid
End Sub

Private Sub Combobox2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
KeyCode = 0
End Sub

Private Sub Combobox2_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
If Trim(ComboBox1.Text) = "" Then
    lbl_pesan = DisplayMsg(1033) '"Please select customer code !"
Else
    If ComboBox1.MatchFound Then
       lbl_name = ComboBox1.List(ComboBox1.ListIndex, 1)
       lbl_pesan = ""
       Call toExcel
    Else
       lbl_name = ""
       lbl_pesan = DisplayMsg(4006) '"customer Code is not found !"
    End If
    
End If
End Sub

Sub toExcel()
Dim xlapp As New Excel.application
Dim rsCek As New Recordset, Idx As Long, tempi As String, tempcust As String
Dim bolcust As Boolean, bolinv As Boolean
Dim rsCompany As New Recordset, sql_plus As String

If Trim(ComboBox2.Text) = "No" Then
    sql_plus = " and (sisa=0 or (fix_cls = '1')) "
Else
    sql_plus = " and sisa>0 and (fix_cls is null or fix_cls = '0') "
End If

sql_join = "select xx.*, OD.SerialNoFrom, Od.SerialNoTo from ( " & vbCrLf & _
    "    select *, isnull(qty-qtydo, 0) as sisa from ( " & vbCrLf & _
    "        select od.cust_code, od.item_code, im.makeritem_code, im.item_name, od.po_no, od.seq_no, od.delivery_date, od.delivery_time, od.currency_code, cc.description CurrDesc, od.qty, " & vbCrLf & _
    "        qtyDo = ( " & vbCrLf & _
    "            select isNull(sum(qty), 0) as qtyDo " & vbCrLf & _
    "            from ( " & vbCrLf & _
    "                select po_no, seq_no, qty from delivery_order " & vbCrLf & _
    "                Union " & vbCrLf & _
    "                select order_no, order_seqno, qty from packing_detail " & vbCrLf & _
    "            ) do " & vbCrLf & _
    "            where do.po_no = od.po_no " & vbCrLf & _
    "            and do.seq_no = od.seq_no " & vbCrLf & _
    "        ), od.price, od.unit_cls, od.amount, om.fix_cls, case WHEN od.service IS NULL then '0' ELSE od.Service END As [Service] " & vbCrLf & _
    "        from orderentry_detail od " & vbCrLf & _
    "        left join item_master im on od.item_code = im.item_code " & vbCrLf & _
    "        left join orderentry_master om on od.po_no = om.po_no " & vbCrLf & _
    "        left join Curr_Cls cc on od.Currency_Code = cc.Curr_Cls " & vbCrLf & _
    "    ) xxx " & vbCrLf & _
    ") xx Left Join orderentry_detail OD on xx.PO_No= OD.PO_No and xx.Seq_No=OD.Seq_No " & vbCrLf & _
    " where xx.delivery_date >= '" & Format(DTPicker3.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
    "and xx.delivery_date <= '" & Format(DTPicker4.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
    "and xx.cust_code = '" & Trim(ComboBox1.Text) & "' " & vbCrLf & _
    sql_plus & vbCrLf & _
    "order by xx.Delivery_Date, xx.makeritem_code"
            
If rsCek.State <> adStateClosed Then rsCek.Close
rsCek.CursorLocation = adUseClient
rsCek.Open sql_join, Db, adOpenDynamic, adLockOptimistic


If Not rsCek.EOF Then
Screen.MousePointer = vbHourglass
With xlapp

    sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax  From company_profile "
    If rsCompany.State <> adStateClosed Then rsCompany.Close
    rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
    If rsCompany.EOF Then Screen.MousePointer = vbDefault: Exit Sub
    .Workbooks.Add
    
    .Range("a2", "m2").Merge
    .Range("a2") = rsCompany!company_name
    .Range("a3", "m3").Merge
    .Range("a3") = rsCompany!address1 & " " & rsCompany!address2 & " " & rsCompany!City & " " & rsCompany!Province & " " & rsCompany!postal_code
    .Range("a4", "m4").Merge
    .Range("a4") = "Phone: " & rsCompany!phone1 & " " & rsCompany!phone2 & " Fax: " & rsCompany!fax
    
    .Range("a6") = "Order Inquiry"
    .Range("b6") = ""
    .Range("a6", "b6").Merge
    .Range("a6").HorizontalAlignment = xlLeft
    .Range("a7") = "Customer Code"
    .Range("b7", "k7").Merge
    .Range("b7") = ": " & Trim(ComboBox1.Text) & " / " & Trim(lbl_name)
    .Range("B7").HorizontalAlignment = xlLeft
    .Range("a8") = "Period"
    .Range("b8") = ": " & Format(DTPicker3, "dd MMMM YYYY") & " to " & Format(DTPicker4, "dd MMMM YYYY")
    .Range("F8") = "Remaining Cls"
    .Range("G8") = ": " & ComboBox2.Text
    
    Idx = 10
    
    Do While Not rsCek.EOF
        If Idx = 10 Then
            .Range("A" & Idx) = "Part Number"
            .Range("B" & Idx) = "Description"
            .Range("C" & Idx) = "SI/PO NO"
            .Range("D" & Idx) = "Plan"
            .Range("E" & Idx) = "Serial From"
            .Range("F" & Idx) = "Serial To"
            .Range("G" & Idx) = "Result"
            .Range("H" & Idx) = "Remaining"
            .Range("I" & Idx) = "Delivery Date"
            .Range("J" & Idx) = "Time"
            .Range("K" & Idx) = "Curr"
            .Range("L" & Idx) = "Price"
            .Range("M" & Idx) = "Curr"
            '.Range("N" & Idx) = "Service"
            .Range("N" & Idx) = "Amount"
            .Range("a" & Idx, "N" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range("a" & Idx, "N" & Idx).Borders(xlEdgeBottom).LineStyle = xlDouble
            Idx = Idx + 1
        End If
        
        Idx = Idx
        'Content
        .Range("a" & Idx) = Trim(rsCek!MakerItem_Code)
        .Range("b" & Idx) = Trim(rsCek!Item_Code) & " " & Trim(rsCek!item_name)
        .Range("c" & Idx) = "'" & Trim(rsCek!po_no) & ""
        .Range("d" & Idx) = Format(rsCek!Qty, gs_formatQty)
        .Range("d" & Idx).NumberFormat = gs_formatQty
        '---
        .Range("e" & Idx) = IIf(IsNull(Trim(rsCek!SerialNoFrom)), "", Trim(rsCek!SerialNoFrom))
        .Range("f" & Idx) = IIf(IsNull(Trim(rsCek!SerialNoTo)), "", Trim(rsCek!SerialNoTo))
        '---
        .Range("g" & Idx) = Format(rsCek!qtyDo, gs_formatQty)
        .Range("g" & Idx).NumberFormat = gs_formatQty
        .Range("h" & Idx) = Format(rsCek!sisa, gs_formatQty)
        .Range("h" & Idx).NumberFormat = gs_formatQty
        .Range("i" & Idx) = "'" & Format(rsCek!delivery_Date, "dd MMM YYYY") & ""
        .Range("i" & Idx).NumberFormat = "dd MMM YYYY"
        .Range("j" & Idx) = rsCek!delivery_time
        .Range("k" & Idx) = rsCek!CurrDesc
        .Range("l" & Idx) = rsCek!CurrDesc
        
        If rsCek!currency_code = "03" Then
            .Range("m" & Idx) = Format(rsCek!Price, gs_formatPriceIDR)
            .Range("m" & Idx).NumberFormat = gs_formatPriceIDR
            '.Range("n" & Idx) = Format(rsCek!service, gs_formatPriceIDR)
            '.Range("n" & Idx).NumberFormat = gs_formatPriceIDR
            
        Else
            .Range("l" & Idx) = Format(rsCek!Price, gs_formatPrice)
            .Range("l" & Idx).NumberFormat = gs_formatPrice
            '.Range("n" & Idx) = Format(rsCek!service, gs_formatPrice)
            '.Range("n" & Idx).NumberFormat = gs_formatPrice
        End If
        '.Range("o" & Idx) = Format(.Range("d" & Idx) * (.Range("l" & Idx) + .Range("n" & Idx)), gs_formatQty)
        .Range("n" & Idx) = Format(.Range("d" & Idx) * (.Range("l" & Idx)), gs_formatQty) ' Non service Price
        
        Idx = Idx + 1
        rsCek.MoveNext
    Loop
       
    .Range("a1", "n" & Idx + 3).Columns.Font.Name = "Arial"
    .Range("a1", "n" & Idx + 3).Columns.Font.Size = 8
    
    .Range("a2", "k2").Columns.Font.Name = "Arial"
    .Range("a2", "k2").Columns.Font.Size = "10"
    .Range("a2", "k2").Columns.Font.Bold = True
    .Range("a2", "k2").HorizontalAlignment = xlCenter
    .Range("a3", "k3").HorizontalAlignment = xlCenter
    .Range("a4", "k4").HorizontalAlignment = xlCenter
    .Range("a6", "e6").Columns.Font.Bold = True
    
    .Range("D11:F" & Idx).Select
    .Selection.NumberFormat = gs_formatQty
       
    .Range("K11:K" & Idx).Select
    .Selection.NumberFormat = gs_formatAmountIDR
    
    If bteHakPrice = 0 Then .Range("I1", "K" & Idx).delete xlShiftToLeft
    .Range("A1").Select
    
'    .ActiveSheet.PageSetup.PaperSize = xlPaperA4
'    .ActiveSheet.PageSetup.Orientation = 2
'    .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
'    .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.4)
    .Range("A:n").Columns.AutoFit
    .WindowState = xlMaximized
    .Visible = True
End With
Else
    lbl_pesan = DisplayMsg(4006)
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    lbl_pesan.Caption = ErrMsg
End If
End Sub

Private Sub cmd_first_Click()

With ComboBox1
    If .ListCount <> 0 Then
        .Text = .List(0, 0)
        Call setting_grid
        lbl_pesan.Caption = DisplayMsg(4020) '"This is the first page !"
    End If
End With

End Sub

Private Sub cmd_last_Click()
With ComboBox1
    If .ListCount <> 0 Then
        .Text = .List(.ListCount - 1, 0)
        Call setting_grid
        lbl_pesan.Caption = DisplayMsg(4021) '"This is the last page !"
    End If
End With

End Sub
Private Sub cmd_previous_Click()

With ComboBox1
    If .ListCount <> 0 Then
        If .ListIndex - 1 >= 0 Then
            .Text = .List(.ListIndex - 1, 0)
            lbl_pesan.Caption = ""
        Else
            lbl_pesan.Caption = DisplayMsg(4020) '"This is the first page !"
        End If
        Call setting_grid
    End If
End With

End Sub
Private Sub cmd_next_Click()

With ComboBox1
    If .ListCount <> 0 Then
        If .ListIndex + 1 <= .ListCount - 1 Then
            .Text = .List(.ListIndex + 1, 0)
            lbl_pesan.Caption = ""
        Else
            lbl_pesan.Caption = DisplayMsg(4021) '"This is the last page !"
        End If
        Call setting_grid
        
    End If
End With

End Sub

Private Sub cmd_sub_menu_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub Cmd_Submit_Click()
Dim j As Integer

If Trim(ComboBox1.Text) = "" Then GoTo Kosong
If rs_trade_master.EOF = False Or rs_trade_master.BOF = False Then
    rs_trade_master.MoveFirst
        rs_trade_master.Find "trade_code='" & Trim(ComboBox1.Text) & "'"
End If

If rs_trade_master.EOF = True Then
    lbl_pesan.Caption = DisplayMsg(4011) '"Please insert a valid customer code !"
    Exit Sub
Else
    lbl_pesan.Caption = ""
End If

If hakAkses("frmOrderEntry") = 0 Then lbl_pesan = DisplayMsg(3007):  Exit Sub
i = 0
With VSFlexGrid1

j = 0
For i = 1 To .Rows - 1
    If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then j = 1
Next

If j < 1 Then
    lbl_pesan.Caption = DisplayMsg(4049) '"Please select data to update !"
    Exit Sub
End If

For i = 1 To .Rows - 1
    If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then Exit For
Next
Call frmOrderEntry.dr_orderInquiry(Trim(ComboBox1.Text), .TextMatrix(i, bteColPONo), .TextMatrix(i, bteColItemCode), DTPicker3.Value, DTPicker4.Value, .TextMatrix(i, bteColDelDate), .TextMatrix(i, bteColSeqNo))

End With
frmOrderEntry.Show
frmOrderEntry.frmpanggil = "orderinquiry"
frmOrderEntry.command3.Caption = "Back"
Me.Hide
Exit Sub

Kosong:
    lbl_pesan.Caption = DisplayMsg(1033) '"Please insert customer code!"
End Sub



Private Sub ComboBox1_Click()
Call set_cust
Call clearGrid
lbl_pesan.Caption = ""
lbl_record.Caption = "Page 0 of 0"
    
End Sub
Sub clearGrid()
VSFlexGrid1.clear
VSFlexGrid1.Rows = 1
Call Header
End Sub

Private Sub set_cust()
If rs_trade_master.EOF = False Or rs_trade_master.BOF = False Then
    rs_trade_master.MoveFirst
        rs_trade_master.Find " trade_code='" & Trim(ComboBox1.Text) & "'"
    If rs_trade_master.EOF = False Then
        ComboBox1 = Trim(rs_trade_master!Trade_Code)
        lbl_name = Trim(rs_trade_master!trade_name)
    Else
        lbl_name = ""
    End If
End If
End Sub



Private Sub Combobox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
VSFlexGrid1.clear
Call clearGrid

If KeyCode = 13 Then
    If rs_trade_master.EOF = False Or rs_trade_master.BOF = False Then
        rs_trade_master.MoveFirst
            rs_trade_master.Find "trade_code='" & Trim(ComboBox1.Text) & "'"
        If rs_trade_master.EOF = False Then
            ComboBox1 = Trim(rs_trade_master!Trade_Code)
            lbl_pesan.Caption = ""
            Call ComboBox1_Click
            lbl_record.Caption = "Page 0 of 0"
        Else
            lbl_pesan.Caption = DisplayMsg(4006) '"Data not found !"
            lbl_record.Caption = "Page 0 of 0"
            ComboBox1.Text = ""
            Call setting_grid
            Call label_clear
        End If
    End If
End If

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then lbl_name = ""
End Sub


Private Sub Combobox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub DTPicker3_Change()

If DTPicker3.Value > DTPicker4.Value Then
    lbl_pesan.Caption = DisplayMsg(1021) '"The first of delivery date must be equal or lower than " & Format(DTPicker4.value, "dd MMM yyyy")
    DTPicker3.Value = Format(DTPicker4.Value, "dd MMM yyyy")
Else
    lbl_pesan.Caption = ""
End If
lbl_record.Caption = "Page 0 of 0"

Call clearGrid

End Sub

Private Sub DTPicker4_Change()
If DTPicker3.Value > DTPicker4.Value Then
    lbl_pesan.Caption = DisplayMsg(1021) '"The last of delivery date must be equal or higher than " & Format(DTPicker3.value, "dd MMM yyyy")
    DTPicker4.Value = Format(DTPicker3.Value, "dd MMM yyyy")
Else
    lbl_pesan.Caption = ""
End If
lbl_record.Caption = "Page 0 of 0"

Call clearGrid
End Sub

Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
ComboBox1.Text = ""
lbl_pesan.Caption = ""
lbl_record.Caption = "Page 0 of 0"
bteHakPrice = hakPrice(Me.Name)
Call koneksi
Call setting
Call Header
End Sub

Private Sub koneksi()

Dim sqlcust As String

sqlcust = "select trade_code, trade_name, address1 from trade_master " & _
    vbLf & " where trade_cls = '2' "


rs_trade_master.Open sqlcust, Db, adOpenKeyset, adLockOptimistic

'-----

'=======================================================================================================
rs_order_master.Open "select * from orderentry_master", Db, adOpenKeyset, adLockOptimistic
rs_order_detail.Open "select * from orderentry_detail", Db, adOpenKeyset, adLockOptimistic
'=======================================================================================================

End Sub

Private Sub label_clear()
lbl_name.Caption = ""

End Sub

Private Sub setting()

'=================setting combobox1=======================
ComboBox1.clear
ComboBox1.columnCount = 2
ComboBox1.TextColumn = 1

i = 0

If rs_trade_master.EOF = False Or rs_trade_master.BOF = False Then
    rs_trade_master.MoveFirst
    While rs_trade_master.EOF = False
        ComboBox1.AddItem ""
        ComboBox1.List(i, 0) = Trim(rs_trade_master!Trade_Code)
        ComboBox1.List(i, 1) = Trim(rs_trade_master!trade_name)
        rs_trade_master.MoveNext
        i = i + 1
    Wend
    ComboBox1.ColumnWidths = "50 pt; 350 pt"
    ComboBox1.ListWidth = 400
End If
'==========================================================

'=================setting combobox2=======================
ComboBox2.clear
ComboBox2.columnCount = 1
ComboBox2.TextColumn = 1

ComboBox2.AddItem ""
'ComboBox2.List(0, 0) = "0"
ComboBox2.List(0, 0) = "Yes"
ComboBox2.AddItem ""
'ComboBox2.List(1, 0) = "1"
ComboBox2.List(1, 0) = "No"

'ComboBox2.ColumnWidths = "50 pt; 50 pt"
ComboBox2.ColumnWidths = "50 pt"
ComboBox2.ListWidth = 50

ComboBox2.Text = ComboBox2.List(0, 0)

'==========================================================

'=================setting label clear============================
Call label_clear
'==========================================================
'=================setting vsflexgrid1============================
Call setting_grid
'==========================================================
'=================setting dtpicker======================

DTPicker3.Value = Now
DTPicker4.Value = Now
'==========================================================

End Sub

Private Sub setting_grid()
Dim l_cur As String, l_price1 As String, sql_plus As String, d As String
Dim l_cur1 As String, L_Service As String
VSFlexGrid1.Rows = 1
                              
If rs_join.State <> adStateClosed Then rs_join.Close

If Trim(ComboBox2.Text) = "No" Then
    sql_plus = " and (sisa=0 or (fix_cls = '1')) "
Else
    sql_plus = " and sisa>0 and (fix_cls is null or fix_cls = '0') "
End If

sql_join = "select xx.*,OD.SerialNoFrom, OD.SerialNoTo from ( " & vbCrLf & _
    "    select *, isnull(qty-qtydo, 0) as sisa from ( " & vbCrLf & _
    "        select od.cust_code, od.item_code, im.makeritem_code, im.item_name, od.po_no, od.seq_no, od.delivery_date, od.delivery_time, od.currency_code, od.qty, " & vbCrLf & _
    "        qtyDo = ( " & vbCrLf & _
    "            select isNull(sum(qty), 0) as qtyDo " & vbCrLf & _
    "            from ( " & vbCrLf & _
    "                select po_no, seq_no, qty from delivery_order " & vbCrLf & _
    "                Union " & vbCrLf & _
    "                select order_no, order_seqno, qty from packing_detail " & vbCrLf & _
    "            ) do " & vbCrLf & _
    "            where do.po_no = od.po_no " & vbCrLf & _
    "            and do.seq_no = od.seq_no " & vbCrLf & _
    "        ), od.price, od.unit_cls, od.amount, om.fix_cls, case WHEN od.service IS NULL then '0' ELSE od.Service END As [Service]  " & vbCrLf & _
    "        from orderentry_detail od " & vbCrLf & _
    "        left join item_master im on od.item_code = im.item_code " & vbCrLf & _
    "        left join orderentry_master om on od.po_no = om.po_no " & vbCrLf & _
    "    ) xxx " & vbCrLf & _
    ") xx Left join orderentry_detail OD on xx.PO_No= OD.PO_No and xx.Seq_No=OD.Seq_No " & vbCrLf & _
    " where xx.delivery_date >= '" & Format(DTPicker3.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
    "and xx.delivery_date <= '" & Format(DTPicker4.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
    "and xx.cust_code = '" & Trim(ComboBox1.Text) & "' " & vbCrLf & _
    sql_plus & vbCrLf & _
    "order by xx.Delivery_Date, xx.makeritem_code"
                
rs_join.Open sql_join, Db, adOpenKeyset, adLockOptimistic

If rs_join.EOF = False Or rs_join.BOF = False Then

    rs_join.MoveFirst

    While Not rs_join.EOF
        l_cur = uf_GetCurrencyDescription(rs_join!currency_code)
        
        d = Trim(rs_join!Item_Code) & " " & uf_GetItemDescription(rs_join!Item_Code)
        If rs_join!currency_code = "03" Then
            l_price1 = Format(Trim(rs_join!Price), gs_formatPriceIDR)
            L_Service = Format(Trim(rs_join!service), gs_formatPriceIDR)
        Else
            l_price1 = Format(Trim(rs_join!Price), gs_formatPrice)
            L_Service = Format(Trim(rs_join!service), gs_formatPrice)
        End If
        With VSFlexGrid1
            .AddItem ""
            .TextMatrix(.Rows - 1, bteColPartNo) = Trim(rs_join!MakerItem_Code)
            .TextMatrix(.Rows - 1, bteColDesc) = Trim(d)
            .TextMatrix(.Rows - 1, bteColPONo) = Trim(rs_join!po_no)
            .TextMatrix(.Rows - 1, bteColPlan) = Format(Trim(rs_join!Qty), gs_formatQty)
            .TextMatrix(.Rows - 1, bteColResult) = Format(Trim(rs_join!qtyDo), gs_formatQty)
            ' ---
            .TextMatrix(.Rows - 1, BteColSerialFrom) = IIf(IsNull(Trim(rs_join!SerialNoFrom)), "", Trim(rs_join!SerialNoFrom))
            .TextMatrix(.Rows - 1, BteColSerialTo) = IIf(IsNull(Trim(rs_join!SerialNoTo)), "", Trim(rs_join!SerialNoTo))
            ' ---
            .TextMatrix(.Rows - 1, bteColRemaining) = Format(Trim(rs_join!sisa), gs_formatQty)
            .TextMatrix(.Rows - 1, bteColDelDate) = Format(Trim(rs_join!delivery_Date), "dd MMM yyyy")
            .TextMatrix(.Rows - 1, bteColDelTime) = Trim(rs_join!delivery_time)
            .TextMatrix(.Rows - 1, bteColCurr) = l_cur
            .TextMatrix(.Rows - 1, bteColPrice) = l_price1
            .TextMatrix(.Rows - 1, bteColService) = L_Service
            .TextMatrix(.Rows - 1, bteColCur1) = l_cur
            .TextMatrix(.Rows - 1, bteColAmount) = Format(Trim(rs_join!Amount), gs_formatAmountIDR)
            .TextMatrix(.Rows - 1, bteColSeqNo) = Trim(rs_join!Seq_no)
            .TextMatrix(.Rows - 1, bteColItemCode) = Trim(rs_join!Item_Code)
        End With
        
        rs_join.MoveNext
    Wend

Else
    VSFlexGrid1.clear
End If

Call Header
With VSFlexGrid1
   
    
    For i = 1 To .Rows - 1
        .Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
    Next
    
    

End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

If rs_order_master.State <> adStateClosed Then rs_order_master.Close
If rs_order_detail.State <> adStateClosed Then rs_order_detail.Close
If rs_join.State <> adStateClosed Then rs_join.Close
If rs_trade_master.State <> adStateClosed Then rs_trade_master.Close
End Sub

Private Sub VSFlexGrid1_Click()
If VSFlexGrid1.Col = bteColSelect And VSFlexGrid1.Row > 0 Then
    If VSFlexGrid1.Cell(flexcpChecked, VSFlexGrid1.Row, bteColSelect) = flexUnchecked Then
        VSFlexGrid1.Cell(flexcpChecked, VSFlexGrid1.Row, bteColSelect) = flexUnchecked
    Else
        For i = 1 To VSFlexGrid1.Rows - 1
            VSFlexGrid1.Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
        Next
        VSFlexGrid1.Cell(flexcpChecked, VSFlexGrid1.Row, bteColSelect) = flexChecked
    End If
End If
End Sub

Private Sub VSFlexGrid1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub VSFlexGrid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
KeyAscii = 0
End Sub
