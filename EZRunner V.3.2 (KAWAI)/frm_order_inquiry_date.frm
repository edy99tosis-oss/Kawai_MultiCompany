VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_order_inquiry_date 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Order Inquiry ( Delivery Date )"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frm_order_inquiry_date.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
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
      Left            =   12660
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9630
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13080
      TabIndex        =   21
      Top             =   240
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   270
      TabIndex        =   13
      Top             =   8910
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
         TabIndex        =   14
         Top             =   240
         Width           =   14490
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   750
      Left            =   285
      TabIndex        =   11
      Top             =   1275
      Width           =   14670
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
         Index           =   0
         Left            =   7965
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   1155
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   345
         Left            =   3555
         TabIndex        =   1
         Top             =   270
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
         Format          =   287571971
         CurrentDate     =   37818
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   345
         Left            =   1485
         TabIndex        =   0
         Top             =   270
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
         Format          =   287571971
         CurrentDate     =   37818
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
         Left            =   180
         TabIndex        =   18
         Top             =   300
         Width           =   1275
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
         Left            =   3165
         TabIndex        =   17
         Top             =   300
         Width           =   210
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
         Left            =   5445
         TabIndex        =   16
         Top             =   315
         Width           =   1290
      End
      Begin MSForms.ComboBox ComboBox2 
         Height          =   315
         Left            =   6885
         TabIndex        =   2
         Top             =   270
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
      Begin VB.Label LblArea 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   6540
         TabIndex        =   12
         Top             =   840
         Width           =   1815
      End
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
      Left            =   13860
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9630
      Width           =   1125
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
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9630
      Width           =   1125
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
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9630
      Visible         =   0   'False
      Width           =   1125
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
      Left            =   6060
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9630
      Visible         =   0   'False
      Width           =   1125
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
      Left            =   7275
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9630
      Visible         =   0   'False
      Width           =   1125
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
      Left            =   8475
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9630
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid vsflexGrid1 
      Height          =   6600
      Left            =   270
      TabIndex        =   19
      Top             =   2205
      Width           =   14655
      _cx             =   25850
      _cy             =   11642
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
      Cols            =   19
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   275
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm_order_inquiry_date.frx":0E42
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Order Inquiry ( Delivery Date )"
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
      Left            =   240
      TabIndex        =   20
      Top             =   480
      Width           =   14670
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
      Left            =   1845
      TabIndex        =   15
      Top             =   9690
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "frm_order_inquiry_date"
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
Dim bteColDelDate As Byte
Dim bteColDelTime As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColCustCode As Byte
Dim bteColPONo As Byte
Dim bteColUnit As Byte
Dim bteColPlan As Byte
Dim BteColSerialFrom As Byte
Dim BteColSerialTo As Byte
Dim bteColResult As Byte
Dim bteColRemaining As Byte
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
        bteColDelDate = 1
        bteColDelTime = 2
        bteColPartNo = 3
        bteColDesc = 4
        bteColCustCode = 5
        bteColPONo = 6
        bteColUnit = 7
        bteColPlan = 8
        BteColSerialFrom = 9
        BteColSerialTo = 10
        bteColResult = 9 + 2
        bteColRemaining = 10 + 2
        bteColCurr = 11 + 2
        bteColPrice = 12 + 2
        bteColService = 13 + 2
        bteColAmount = 14 + 2
        bteColSeqNo = 15 + 2
        bteColItemCode = 16 + 2
        bteColCur1 = 17 + 2
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColDelDate) = "Delivery Date"
        .TextMatrix(0, bteColDelTime) = "Time"
        .TextMatrix(0, bteColPartNo) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColCustCode) = "Cust CD"
        .TextMatrix(0, bteColPONo) = "SI/PO No"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColPlan) = "Plan"
        '---
        .TextMatrix(0, BteColSerialFrom) = "Serial From"
        .TextMatrix(0, BteColSerialTo) = "Serial To"
        '---
        .TextMatrix(0, bteColResult) = "Result"
        .TextMatrix(0, bteColRemaining) = "Remaining"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColService) = "Service"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColSeqNo) = "seqNo"
        .TextMatrix(0, bteColItemCode) = "itemCode"
        
        .ColAlignment(bteColDelDate) = flexAlignLeftCenter
        .ColAlignment(bteColDelTime) = flexAlignLeftCenter
        .ColAlignment(bteColPartNo) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColCustCode) = flexAlignLeftCenter
        .ColAlignment(bteColPONo) = flexAlignLeftCenter
        .ColAlignment(bteColUnit) = flexAlignLeftCenter
        .ColAlignment(bteColPlan) = flexAlignRightCenter
        .ColAlignment(bteColResult) = flexAlignRightCenter
        .ColAlignment(bteColRemaining) = flexAlignRightCenter
        .ColAlignment(bteColPrice) = flexAlignRightCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
        
        .ColWidth(bteColSelect) = 250
        .ColWidth(bteColDelDate) = 1400
        .ColWidth(bteColDelTime) = 750
        .ColWidth(bteColPartNo) = 2000
        .ColWidth(bteColDesc) = 2400
        .ColWidth(bteColCustCode) = 1200
        .ColWidth(bteColPONo) = 1500
        .ColWidth(bteColUnit) = 500
        .ColWidth(bteColPlan) = 850
        '---
        .ColWidth(BteColSerialFrom) = 1100
        .ColWidth(BteColSerialTo) = 1100
        '---
        .ColWidth(bteColResult) = 850
        .ColWidth(bteColRemaining) = 1000
        .ColWidth(bteColCurr) = 650
        '.ColWidth(bteColCur1) = 650
        .ColWidth(bteColPrice) = 1400
        .ColWidth(bteColService) = 1400
        .ColWidth(bteColAmount) = 1400
        
        .ColHidden(bteColUnit) = True
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColItemCode) = True
        
        .ColHidden(bteColCurr) = (bteHakPrice = 0)
        .ColHidden(bteColPrice) = (bteHakPrice = 0)
'       .ColHidden(bteColCur1) = (bteHakPrice = 0)
        .ColHidden(bteColService) = True
        .ColHidden(bteColAmount) = (bteHakPrice = 0)
        
        .EditMaxLength = 1
    End With
End Sub

Public Sub cmdSearch_Click(Index As Integer)
MousePointer = vbHourglass
lbl_pesan.Caption = ""
Call setting_grid
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
Private Sub clearGrid()
VSFlexGrid1.clear
VSFlexGrid1.Rows = 1
Call Header
End Sub

Private Sub Command1_Click()
Dim xlapp As New Excel.application
Dim rsCek As New Recordset, Idx As Long, tempi As String, tempcust As String
Dim bolcust As Boolean, bolinv As Boolean
Dim rsCompany As New Recordset, sql_plus As String

If Trim(ComboBox2.Text) = "No" Then
    sql_plus = " and (sisa=0 or (fix_cls = '1')) "
Else
    sql_plus = " and sisa>0 and (fix_cls is null or fix_cls = '0') "
End If

' Edit For Add Serial Number From  and Serial Number To
' Update 20090205

sql_join = "select xx.*,ODD.SerialNoFrom,ODD.SerialNoTo from ( " & vbCrLf & _
    "    select *, isnull(qty-qtydo, 0) as sisa from ( " & vbCrLf & _
    "        select od.cust_code, od.item_code, im.makeritem_code, im.item_name, od.po_no, od.seq_no, od.delivery_date, od.delivery_time, od.currency_code, cc.description currDesc, od.qty, " & vbCrLf & _
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
    "        left join curr_cls cc on cc.curr_cls = od.Currency_Code " & vbCrLf & _
    "    ) xxx " & vbCrLf & _
    ") xx Left join orderentry_detail ODD on xx.PO_No= ODD.PO_No and xx.Seq_No=ODD.Seq_No" & vbCrLf & _
    " where xx.delivery_date >= '" & Format(DTPicker3.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
    "and xx.delivery_date <= '" & Format(DTPicker4.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
    sql_plus & vbCrLf & _
    "order by xx.Delivery_Date, xx.makeritem_code"

' ---------

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
    
    .Range("a2", "n2").Merge
    .Range("a2") = rsCompany!company_name
    .Range("a3", "n3").Merge
    .Range("a3") = rsCompany!address1 & " " & rsCompany!address2 & " " & rsCompany!City & " " & rsCompany!Province & " " & rsCompany!postal_code
    .Range("a4", "n4").Merge
    .Range("a4") = "Phone: " & rsCompany!phone1 & " " & rsCompany!phone2 & " Fax: " & rsCompany!fax
    
    .Range("a6", "l6").Merge
    .Range("a6") = "Order Inquiry (Delivery Date)"
    .Range("a6").HorizontalAlignment = xlLeft
    .Range("a8") = "Period"
    .Range("b8", "D8").Merge
    .Range("B8") = ": " & Format(DTPicker3, "dd MMMM YYYY") & " to " & Format(DTPicker4, "dd MMMM YYYY")
    .Range("B8").HorizontalAlignment = xlLeft
    .Range("F8") = "Remaining Cls"
    .Range("G8") = ": " & ComboBox2.Text
    
    Idx = 10
    
    Do While Not rsCek.EOF
        If Idx = 10 Then
            .Range("a" & Idx) = "Delivery Date"
            .Range("b" & Idx) = "Time"
            .Range("c" & Idx) = "Product Code"
            .Range("d" & Idx) = "Description"
            .Range("e" & Idx) = "Cust Code"
            .Range("f" & Idx) = "SI/PO NO"
            .Range("g" & Idx) = "Plan"
            '--
            .Range("h" & Idx) = "Serial From"
            .Range("i" & Idx) = "Serial To"
            '--
            .Range("j" & Idx) = "Result"
            .Range("k" & Idx) = "Remaining"
            .Range("l" & Idx) = "Curr"
            .Range("m" & Idx) = "Price"
            '.Range("n" & Idx) = "Service"
            .Range("n" & Idx) = "Amount"
            
            .Range("a" & Idx, "N" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range("a" & Idx, "N" & Idx).Borders(xlEdgeBottom).LineStyle = xlDouble

            Idx = Idx + 1
        End If
        
        Idx = Idx
        'Content
        .Range("a" & Idx) = "'" & Format(rsCek!delivery_Date, "dd MMM YYYY") & ""
        .Range("b" & Idx) = rsCek!delivery_time
        .Range("c" & Idx) = Trim(rsCek!MakerItem_Code)
        .Range("d" & Idx) = Trim(rsCek!Item_Code) & " " & Trim(rsCek!item_name)
        .Range("e" & Idx) = Trim(rsCek!Cust_CodE)
        .Range("f" & Idx) = "'" & Trim(rsCek!po_no) & ""
        .Range("g" & Idx) = Format(rsCek!Qty, gs_formatQty)
        '--
        .Range("h" & Idx) = IIf(IsNull(Trim(rsCek!SerialNoFrom)), "", Trim(rsCek!SerialNoFrom))
        .Range("i" & Idx) = IIf(IsNull(Trim(rsCek!SerialNoTo)), "", Trim(rsCek!SerialNoTo))
        '--
        .Range("j" & Idx) = Format(rsCek!qtyDo, gs_formatQty)
        .Range("k" & Idx) = Format(rsCek!sisa, gs_formatQty)
        .Range("l" & Idx) = Trim(rsCek!CurrDesc)
        .Range("m" & Idx) = Trim(rsCek!CurrDesc)
        If rsCek!currency_code = "03" Then
            .Range("m" & Idx) = Format(rsCek!Price, gs_formatPriceIDR)
            .Range("m" & Idx).NumberFormat = gs_formatPriceIDR
            
            '.Range("n" & Idx) = Format(rsCek!service, gs_formatPriceIDR)
            '.Range("n" & Idx).NumberFormat = gs_formatPriceIDR
            
        Else
            .Range("m" & Idx) = Format(rsCek!Price, gs_formatPrice)
            .Range("m" & Idx).NumberFormat = gs_formatPrice
            
            
            '.Range("n" & Idx) = Format(rsCek!service, gs_formatPrice)
            '.Range("n" & Idx).NumberFormat = gs_formatPrice
        End If
        '.Range("o" & Idx) = Format(.Range("g" & Idx) * (.Range("m" & Idx) + .Range("n" & Idx)), gs_formatAmountIDR)
        .Range("N" & Idx) = Format(.Range("g" & Idx) * (.Range("m" & Idx)), gs_formatAmountIDR) ' Non Service Price
        
        Idx = Idx + 1
        rsCek.MoveNext
    Loop
    

    .Range("a1", "o" & Idx + 3).Columns.Font.Name = "Arial"
    .Range("a1", "o" & Idx + 3).Columns.Font.Size = 8
    
    .Range("a2", "l2").Columns.Font.Name = "Arial"
    .Range("a2", "l2").Columns.Font.Size = "10"
    .Range("a2", "l2").Columns.Font.Bold = True
    .Range("a2").HorizontalAlignment = xlCenter
    .Range("a3").HorizontalAlignment = xlCenter
    .Range("a4").HorizontalAlignment = xlCenter
    .Range("a6", "e6").Columns.Font.Bold = True
    
    .Range("G11:I" & Idx).Select
    .Selection.NumberFormat = gs_formatQty
       
'    .Range("K11:K" & Idx).Select
'    .Selection.NumberFormat = gs_formatPrice
    
    .Range("n11:n" & Idx).Select
    .Selection.NumberFormat = gs_formatAmountIDR
    
    If bteHakPrice = 0 Then .Range("l1", "n" & Idx).delete xlShiftToLeft
    .Range("A1").Select
    
'    .ActiveSheet.PageSetup.PaperSize = xlPaperA4
'    .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
'    .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.4)
    .Range("A:o").Columns.AutoFit
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

Private Sub cmd_sub_menu_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub Cmd_Submit_Click()
Dim j As Integer

If hakAkses("frmOrderEntry") = 0 Then lbl_pesan = DisplayMsg(3007):  Exit Sub
i = 0
With VSFlexGrid1
If .Rows = 1 = True Then
    lbl_pesan.Caption = DisplayMsg(4007) 'Please select data first !"
    Exit Sub
Else
    lbl_pesan.Caption = ""
End If

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
Call frmOrderEntry.dr_orderInquiry(.TextMatrix(i, bteColCustCode), .TextMatrix(i, bteColPONo), .TextMatrix(i, bteColItemCode), DTPicker3.Value, DTPicker4.Value, .TextMatrix(i, bteColDelDate), .TextMatrix(i, bteColSeqNo))

End With
frmOrderEntry.Show
frmOrderEntry.frmpanggil = "orderinquirydate"
frmOrderEntry.command3.Caption = "Back"
Me.Hide

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
Me.Caption = "Order Inquiry ( Delivery Date )"
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
lbl_record.Caption = "Page 0 of 0"
bteHakPrice = hakPrice(Me.Name)
Call koneksi
Call setting
Call Header
lbl_pesan.Caption = ""
End Sub

Private Sub koneksi()

'=======================================================================================================
rs_order_master.Open "select * from orderentry_master", Db, adOpenKeyset, adLockOptimistic
rs_order_detail.Open "select * from orderentry_detail", Db, adOpenKeyset, adLockOptimistic
rs_trade_master.Open "select * from trade_master where trade_cls='4'", Db, adOpenKeyset, adLockOptimistic

'=======================================================================================================

End Sub

Private Sub label_clear()

End Sub

Private Sub setting()


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

ComboBox2.ColumnWidths = "50 pt" '; 50 pt"
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
Dim l_price1 As String, sql_plus As String, d As String
Dim L_Service As String
VSFlexGrid1.Rows = 1
                              
If rs_join.State <> adStateClosed Then rs_join.Close

If Trim(ComboBox2.Text) = "No" Then
    sql_plus = " and (sisa=0 or (fix_cls = '1')) "
Else
    sql_plus = " and sisa>0 and (fix_cls is null or fix_cls = '0') "
End If

' Edit For Add Serial Number From  and Serial Number To
' Update 20090205

sql_join = "select xx.*,ODD.SerialNoFrom,ODD.SerialNoTo from ( " & vbCrLf & _
    "    select *, isnull(qty-qtydo, 0) as sisa from ( " & vbCrLf & _
    "        select od.cust_code, od.item_code, im.makeritem_code, im.item_name, od.po_no, od.seq_no, od.delivery_date, od.delivery_time, od.currency_code, cc.description currDesc, od.qty, " & vbCrLf & _
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
    "        left join curr_cls cc on cc.curr_cls = od.Currency_Code " & vbCrLf & _
    "    ) xxx " & vbCrLf & _
    ") xx Left join orderentry_detail ODD on xx.PO_No= ODD.PO_No and xx.Seq_No=ODD.Seq_No" & vbCrLf & _
    " where xx.delivery_date >= '" & Format(DTPicker3.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
    "and xx.delivery_date <= '" & Format(DTPicker4.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
    sql_plus & vbCrLf & _
    "order by xx.Delivery_Date, xx.makeritem_code"

' ---------

rs_join.Open sql_join, Db, adOpenKeyset, adLockOptimistic

If rs_join.EOF = False Or rs_join.BOF = False Then

    rs_join.MoveFirst

    While Not rs_join.EOF
    Dim l_cur As String
    
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
            .TextMatrix(.Rows - 1, bteColDelDate) = Format(Trim(rs_join!delivery_Date), "dd MMM yyyy")
            .TextMatrix(.Rows - 1, bteColDelTime) = Trim(rs_join!delivery_time)
            .TextMatrix(.Rows - 1, bteColPartNo) = Trim(rs_join!MakerItem_Code)
            .TextMatrix(.Rows - 1, bteColDesc) = Trim(d)
            .TextMatrix(.Rows - 1, bteColCustCode) = Trim(rs_join!Cust_CodE)
            .TextMatrix(.Rows - 1, bteColPONo) = Trim(rs_join!po_no)
            .TextMatrix(.Rows - 1, bteColUnit) = Trim(rs_join!Unit_cls)
            .TextMatrix(.Rows - 1, bteColPlan) = Format(Trim(rs_join!Qty), gs_formatQty)
            '---
            .TextMatrix(.Rows - 1, BteColSerialFrom) = IIf(IsNull(Trim(rs_join!SerialNoFrom)), "", Trim(rs_join!SerialNoFrom))
            .TextMatrix(.Rows - 1, BteColSerialTo) = IIf(IsNull(Trim(rs_join!SerialNoTo)), "", Trim(rs_join!SerialNoTo))
            '---
            .TextMatrix(.Rows - 1, bteColResult) = Format(Trim(rs_join!qtyDo), gs_formatQty)
            .TextMatrix(.Rows - 1, bteColRemaining) = Format(Trim(rs_join!sisa), gs_formatQty)
            .TextMatrix(.Rows - 1, bteColCurr) = l_cur
            .TextMatrix(.Rows - 1, bteColPrice) = l_price1
            .TextMatrix(.Rows - 1, bteColService) = L_Service
'           .TextMatrix(.Rows - 1, bteColCur1) = l_cur
            .TextMatrix(.Rows - 1, bteColAmount) = Format(Trim(rs_join!Amount), gs_formatAmountIDR)
            .TextMatrix(.Rows - 1, bteColSeqNo) = Trim(rs_join!Seq_no)
            .TextMatrix(.Rows - 1, bteColItemCode) = Trim(rs_join!Item_Code)
        End With
        rs_join.MoveNext
    Wend
    
Else
    VSFlexGrid1.clear
    lbl_pesan.Caption = DisplayMsg(4006)
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

'With vsflexGrid1
'    If .Row = 1 And .Col <> 0 Then
'      If .Col = 1 Or .Col = 2 Or .Col = 3 Or .Col = 4 Or .Col = 5 Or .Col = 6 Then
'
'          If .ColSort(.Col) = flexSortStringAscending Then
'             .ColSort(.Col) = flexSortStringDescending
'          Else
'             .ColSort(.Col) = flexSortStringAscending
'          End If
'     Else
'
'          If .ColSort(.Col) = flexSortNumericAscending Then
'             .ColSort(.Col) = flexSortNumericDescending
'          Else
'             .ColSort(.Col) = flexSortNumericAscending
'          End If
'
'      End If
'       .Sort = .ColSort(.Col)
'    End If
'End With
End Sub

Private Sub VSFlexGrid1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub VSFlexGrid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
KeyAscii = 0
End Sub



