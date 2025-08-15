VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPORecInquiryMaterial 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Purchase Order / Result Inquiry (Each Material Code)"
   ClientHeight    =   10560
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
   Icon            =   "FrmPORecInquiryMaterial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdexcel 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xcel"
      Height          =   375
      Left            =   13613
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9960
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   353
      TabIndex        =   17
      Top             =   1500
      Width           =   14520
      Begin VB.TextBox txtPO 
         Height          =   315
         Left            =   10500
         TabIndex        =   27
         Top             =   705
         Width           =   2175
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
         Height          =   375
         Left            =   12960
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   660
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker deldate 
         Height          =   315
         Left            =   1590
         TabIndex        =   2
         Top             =   705
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   151322627
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker deldate1 
         Height          =   315
         Left            =   3930
         TabIndex        =   3
         Top             =   705
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   151322627
         CurrentDate     =   37810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO No."
         Height          =   195
         Left            =   9225
         TabIndex        =   28
         Top             =   765
         Width           =   585
      End
      Begin MSForms.ComboBox cbocountry 
         Height          =   315
         Left            =   10485
         TabIndex        =   5
         Top             =   255
         Visible         =   0   'False
         Width           =   2175
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3836;556"
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
         Caption         =   "Country Cls"
         Height          =   195
         Left            =   9225
         TabIndex        =   25
         Top             =   315
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Material Code"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   315
         Width           =   1185
      End
      Begin MSForms.ComboBox cbosupplier 
         Height          =   315
         Left            =   1590
         TabIndex        =   1
         Top             =   255
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
      Begin VB.Line Line2 
         X1              =   4275
         X2              =   8730
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4275
         TabIndex        =   21
         Top             =   315
         Width           =   4455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   765
         Width           =   1185
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   195
         Left            =   3465
         TabIndex        =   19
         Top             =   765
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining Cls"
         Height          =   195
         Left            =   6300
         TabIndex        =   18
         Top             =   765
         Width           =   1230
      End
      Begin MSForms.ComboBox cboremaincls 
         Height          =   315
         Left            =   7740
         TabIndex        =   4
         Top             =   705
         Width           =   915
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "1614;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Index           =   0
      Left            =   353
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9960
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last Page"
      Height          =   375
      Index           =   4
      Left            =   7703
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Next Page"
      Height          =   375
      Index           =   3
      Left            =   6398
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Prev Page"
      Height          =   375
      Index           =   2
      Left            =   5123
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Page"
      Height          =   375
      Index           =   1
      Left            =   3803
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   353
      TabIndex        =   15
      Top             =   9150
      Width           =   14520
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
         TabIndex        =   16
         Top             =   195
         Width           =   14175
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   6255
      Left            =   353
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2820
      Width           =   14520
      _cx             =   25612
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
      GridColorFixed  =   12582912
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   0
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
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
      Height          =   420
      Left            =   13013
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   270
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.Label lblgroup 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4260
      TabIndex        =   26
      Top             =   1050
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   4260
      X2              =   6120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
      Height          =   195
      Left            =   600
      TabIndex        =   24
      Top             =   1050
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSForms.ComboBox cbogroup 
      Height          =   315
      Left            =   1950
      TabIndex        =   0
      Top             =   1020
      Visible         =   0   'False
      Width           =   2175
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3836;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order / Result Inquiry (Each Material CD)"
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
      Left            =   368
      TabIndex        =   23
      Top             =   300
      Width           =   14490
   End
End
Attribute VB_Name = "FrmPORecInquiryMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim RS As New ADODB.Recordset

Dim bteColMatCode As Byte
Dim bteColMatName As Byte
Dim bteColSupplier As Byte
Dim bteColName As Byte
Dim bteColDate As Byte
Dim bteColPONo As Byte
Dim bteColPlan As Byte
Dim bteColResult As Byte
Dim bteColRemain As Byte
Dim bteColComplete As Byte

Sub Header()

    bteColMatCode = 0
    bteColMatName = 1
    bteColSupplier = 2
    bteColName = 3
    bteColPONo = 4
    bteColDate = 5
    bteColPlan = 6
    bteColResult = 7
    bteColRemain = 8
    bteColComplete = 9
    
    With grid
        .Rows = 1
        .ColS = 10
        
        .TextMatrix(0, bteColMatCode) = "Material"
        .TextMatrix(0, bteColMatName) = "Description"
        .TextMatrix(0, bteColSupplier) = "Supplier"
        .TextMatrix(0, bteColName) = "Name"
        .TextMatrix(0, bteColDate) = "Delivery Date"
        .TextMatrix(0, bteColPONo) = "PO No"
        .TextMatrix(0, bteColPlan) = "Plan"
        .TextMatrix(0, bteColResult) = "Result"
        .TextMatrix(0, bteColRemain) = "Remaining"
        .TextMatrix(0, bteColComplete) = "Complete"
        
        .ColWidth(bteColMatCode) = 1500
        .ColWidth(bteColMatName) = 3500
        .ColWidth(bteColSupplier) = 1500
        .ColWidth(bteColName) = 3000
        .ColWidth(bteColDate) = 1300
        .ColWidth(bteColPONo) = 2600
        .ColWidth(bteColPlan) = 1200
        .ColWidth(bteColResult) = 1200
        .ColWidth(bteColRemain) = 1200
        .ColWidth(bteColComplete) = 1000
        
        .ColDataType(bteColDate) = flexDTDate
        
        .ColAlignment(bteColMatCode) = flexAlignLeftCenter
        .ColAlignment(bteColMatName) = flexAlignLeftCenter
        .ColAlignment(bteColSupplier) = flexAlignCenterCenter
        .ColAlignment(bteColName) = flexAlignLeftCenter
        .ColAlignment(bteColDate) = flexAlignLeftCenter
        .ColAlignment(bteColPONo) = flexAlignCenterCenter
        .ColAlignment(bteColPlan) = flexAlignRightCenter
        .ColAlignment(bteColResult) = flexAlignRightCenter
        .ColAlignment(bteColRemain) = flexAlignRightCenter
        .ColAlignment(bteColComplete) = flexAlignLeftCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
        
        If cboSupplier.Text = strAll Then
            .ColHidden(bteColMatCode) = False
            .ColHidden(bteColMatName) = False
        Else
            .ColHidden(bteColMatCode) = True
            .ColHidden(bteColMatName) = True
        End If
        
    End With
    
    LblErrMsg.Caption = ""

End Sub

Sub adtocombo()
'Sql = "SELECT rtrim(item_code) item_code, case finishGoodPart_cls " & _
'                "when '01' then Rtrim(item_master.Grade_cls)+ '/' + rtrim(item_master.UserColor_Cls)+ '/'+ rtrim(item_master.NIPColor_Cls) " & _
'                "when '02' then item_name end item_name FROM item_master"
'Sql = "select im.item_code, case " & _
'        "when im.sheetcoil_cls is null then rtrim(im.item_name) " & _
'        "Else " & _
'        "rtrim(im.item_name)+ '(' + rtrim(description)+ ',T' + rtrim(convert(char(15),im.thickness)) + ' x W' + " & _
'        "rtrim(convert(char(15),im.Width)) + ' x L'  + rtrim(convert(char(15),im.length)) + ')' " & _
'        "End item_name " & _
'        "from item_master im left join sheetcoil_cls sh on im.sheetcoil_cls=sh.sheetcoil_cls " & _
'        "where use_endday > convert(char(8), getdate(), 112) "

sql = "select im.item_code, rtrim(im.item_name) item_Name " & _
        "from item_master im left join sheetcoil_cls sh on im.sheetcoil_cls=sh.sheetcoil_cls " & _
        "where use_endday > convert(char(8), getdate(), 112) "

Set RS = Db.Execute(sql)

With cboSupplier
    .clear
    .columnCount = 2
    .ColumnWidths = "150 pt;270 pt"
    .ListWidth = 420
    .ListRows = 15
    .AddItem ""
    .List(0, 0) = strAll
    .List(0, 1) = strAll
    
i = 1
Do Until RS.EOF
    .AddItem ""
    .List(i, 0) = Trim(RS!Item_Code)
    .List(i, 1) = Trim(RS!item_name)
    i = i + 1
    RS.MoveNext
Loop
.ListIndex = 0
End With

'**************Group Cls**************************
Call up_FillCombo(cboGroup, "Group_Cls", , , True)
cboGroup.ListWidth = 150
cboGroup.ColumnWidths = "30 pt;120 pt"
cboGroup.ListIndex = 0
'*************************************************

'**************Country Cls************************
With cbocountry
 .clear
 .AddItem "Domestic"
 .AddItem "Overseas"
End With
'*************************************************

End Sub
Sub query()
'****Group_Cls*******
Dim sqlgr As String
'If cboGroup.Text = "ALL" Then
' sqlgr = ""
'Else
' sqlgr = " and im.group_cls = '" & Trim(cboGroup.Text) & "' "
'End If
'
''***Supllier All*******
Dim sqlcc As String
If cboSupplier.Text = strAll Then
 sqlcc = " "
Else
 sqlcc = " and pd.Item_code = '" & Trim(cboSupplier.Text) & "' "
End If

sql = "select  cp.company_name,cp.Address1,cp.Address2,cp.City, " & _
      "cp.Province, cp.Postal_Code, cp.Phone1, cp.Phone2, cp.Fax, " & _
      "pm.supplier_code, rtrim(tm.trade_name) supplier_name, pm.delivery_date, pd.item_code, im.finishgoodpart_cls, tm.country_cls,  " & _
      "case when tm.country_cls='1' then 'Overseas' else 'Domestic' End as country_desc, " & _
      "im.item_name, pm.po_no, isnull(pd.qty,0) as orders, " & _
      "isnull((select sum(qty) as qty from part_receipt where supplier_code=pm.supplier_code and " & _
      "po_no=pm.po_no and item_code=pd.item_code),0) as receipt, " & _
      " case  when  pd.complete_cls ='1' then " & _
      "case when (isnull(pd.qty,0) - isnull((select sum(qty) from part_receipt where supplier_code=pm.supplier_code and po_no=pm.po_no and item_code=pd.item_code),0)) < 0 then " & _
      "(isnull(pd.qty,0) - isnull((select sum(qty) from part_receipt where supplier_code=pm.supplier_code and po_no=pm.po_no and item_code=pd.item_code),0)) " & _
      "Else" & _
      "(isnull(pd.qty,0) - isnull((select sum(qty) from part_receipt where supplier_code=pm.supplier_code and po_no=pm.po_no and item_code=pd.item_code),0)) " & _
      "End " & _
      "Else " & _
      "(isnull(pd.qty,0) - isnull((select sum(qty) from part_receipt where supplier_code=pm.supplier_code and po_no=pm.po_no and item_code=pd.item_code),0)) " & _
      "End     as remaining, isnull(complete_cls,0) complete_Cls " & _
      "from purchaseorder_master pm, purchaseorder_detail pd, item_master im, trade_master tm, company_profile cp " & _
      "Where pm.po_no = pd.po_no And pd.item_code = im.item_code and pm.supplier_code = tm.trade_code " & _
      " " & sqlcc & "" & _
      "and pm.po_no like '%" & txtpo.Text & "%' "
'      sqlgr & sqlcc
      
If cboremaincls.ListIndex = 0 Then
sql = sql & " and pd.delivery_date>='" & Format(DelDate, "yyyy-mm-dd") & "' and pd.delivery_date<='" & _
                    Format(deldate1.Value, "yyyy-mm-dd") & "'"
ElseIf cboremaincls.ListIndex = 1 Then
sql = sql & " and pd.delivery_date>='" & Format(DelDate, "yyyy-mm-dd") & "' and pd.delivery_date<='" & _
      Format(deldate1.Value, "yyyy-mm-dd") & "' and (isnull(pd.qty,0) - isnull((select sum(qty) from part_receipt where supplier_code=pm.supplier_code and " & _
      "po_no=pm.po_no and item_code=pd.item_code),0)) > 0 and (pd.complete_cls is null or pd.complete_Cls <>'1') "
Else
sql = sql & " and pd.delivery_date>='" & Format(DelDate, "yyyy-mm-dd") & "' and pd.delivery_date<='" & _
      Format(deldate1.Value, "yyyy-mm-dd") & "' and ((isnull(pd.qty,0) - isnull((select sum(qty) from part_receipt where supplier_code=pm.supplier_code and " & _
      "po_no=pm.po_no and item_code=pd.item_code),0)) <= 0  or  complete_Cls = '1')"
End If

sql = sql & "Order By pd.Item_Code,pd.Delivery_Date"
End Sub


Private Sub Browse()
Header

Call query

Set RS = Db.Execute(sql)

If (RS.BOF And RS.EOF) Then
    LblErrMsg.Caption = DisplayMsg(4006)
    Exit Sub
Else
i = 1
With grid
    Do While Not RS.EOF
        .Rows = .Rows + 1
        .TextMatrix(i, bteColMatCode) = Trim(RS("Item_Code"))
        .TextMatrix(i, bteColMatName) = Trim(RS("Item_Name"))
        .TextMatrix(i, bteColSupplier) = Trim(RS("supplier_Code"))
        .TextMatrix(i, bteColName) = IIf(IsNull(RS("supplier_name")), "", Trim(RS("supplier_name")))
        .TextMatrix(i, bteColDate) = Format(RS("delivery_date"), "dd mmm yyyy")
        .TextMatrix(i, bteColPONo) = Trim(RS("po_no"))
        .TextMatrix(i, bteColPlan) = Format(RS("orders"), gs_formatQty)
        .TextMatrix(i, bteColResult) = Format(RS("receipt"), gs_formatQty)
        .TextMatrix(i, bteColRemain) = Format(RS("remaining"), gs_formatQty)
        If RS!complete_cls = "1" Then
            .Cell(flexcpChecked, i, bteColComplete) = flexChecked
        Else
            .Cell(flexcpChecked, i, bteColComplete) = flexUnchecked
        End If
        i = i + 1
        RS.MoveNext
    Loop
End With
End If

End Sub

Private Sub cboremaincls_Change()
        grid.Rows = 1
End Sub

Private Sub CmdExcel_Click()
Dim xlapp As New Excel.application
Dim Idx As Long

'If cbogroup = "" Then
'   LblErrMsg = DisplayMsg(8081)
'   cbogroup.SetFocus
If cboSupplier = "" Then
   LblErrMsg = DisplayMsg(1009)
   cboSupplier.SetFocus
ElseIf CDate(deldate1) < CDate(DelDate) Then
   LblErrMsg.Caption = DisplayMsg(4077) & " " & Format(DelDate, "dd MMM yyyy")   '"delivery Date must be higher than "
   deldate1.SetFocus
'ElseIf cbocountry = "" Then
'   LblErrMsg = DisplayMsg(8085)
'   cbocountry.SetFocus
Else
   'cbogroup = cbogroup
   cboSupplier = cboSupplier
   'cbocountry = cbocountry
               
'   If cbogroup.MatchFound = False Then
'    LblErrMsg = DisplayMsg(8083)
'    cbogroup.SetFocus
   If cboSupplier.MatchFound = False Then
    LblErrMsg = DisplayMsg(4003)
    cboSupplier.SetFocus
'   ElseIf cbocountry.MatchFound = False Then
'    LblErrMsg = DisplayMsg(8086)
'    cbocountry.SetFocus
   Else
    LblErrMsg = ""
    MousePointer = vbHourglass
         
    Call query
   
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenDynamic, adLockOptimistic
    If RS.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
    
    Screen.MousePointer = vbHourglass
    With xlapp

     .Workbooks.Add
     .Range("a2", "j2").Merge
     .Range("a2") = RS!company_name
     .Range("a3", "j3").Merge
     .Range("a3", "j3") = Trim(RS!address1) & " " & Trim(RS!address2) & " " & Trim(RS!City) & " " & Trim(RS!Province) & " " & Trim(RS!postal_code)
     .Range("a4", "j4").Merge
     .Range("a4") = "Phone: " & RS!phone1 & " " & RS!phone2 & " Fax: " & RS!fax
     
     .Range("a6") = "Purchase Order / Result Inquiry (Each Material CD)"
     .Range("b6") = ""
     .Range("a6", "j6").Merge
     .Range("a6").HorizontalAlignment = xlLeft
     
     '.Range("a7") = "Group Cls"
     '.Range("b7", "h7").Merge
     '.Range("b7") = ": " & cbogroup.Column(0) & " / " & cbogroup.Column(1)
     '.Range("b7").HorizontalAlignment = xlLeft
     .Range("a7") = "Material Code"
     .Range("b7", "h7").Merge
     .Range("b7") = ": " & cboSupplier.Column(0) & " / " & cboSupplier.Column(1)
     .Range("a8") = "Delivery Date"
     .Range("b8", "h8").Merge
     .Range("b8") = ": " & Format(DelDate.Value, "[$-409]d-mmm-yyyy;@") & " to " & Format(deldate1.Value, "[$-409]d-mmm-yyyy;@")
     '.Range("a10") = "Country Cls"
     '.Range("b10", "h10").Merge
     '.Range("b10") = ": " & cbocountry.Column(0)
     .Range("a9") = "Remaining Cls"
     .Range("b9", "h9").Merge
     .Range("b9") = ": " & cboremaincls.Column(0)
    
    
     Idx = 11
     Do While Not RS.EOF
      If Idx = 11 Then
        .Range("a" & Idx) = "Material"
        .Range("b" & Idx) = "Description"
        .Range("c" & Idx) = "Supplier"
        .Range("d" & Idx) = "Name"
        .Range("e" & Idx) = "PO No"
        .Range("f" & Idx) = "Delivery Date"
        .Range("g" & Idx) = "Plan"
        .Range("h" & Idx) = "Result"
        .Range("i" & Idx) = "Remaining"
        .Range("j" & Idx) = "Complete"
        .Range("a" & Idx, "J" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("a" & Idx, "J" & Idx).Borders(xlEdgeBottom).LineStyle = xlDouble
        Idx = Idx + 1
      End If
        
     'Idx = Idx
     'Content
     .Range("a" & Idx) = Trim(RS("Item_Code"))
     .Range("b" & Idx) = Trim(RS("Item_Name"))
     .Range("c" & Idx) = Trim(RS("supplier_Code"))
     .Range("d" & Idx) = IIf(IsNull(RS("supplier_name")), "", Trim(RS("supplier_name")))
     .Range("e" & Idx) = Trim(RS("po_no"))
     .Range("f" & Idx) = Format(RS("delivery_date"), "dd mmm yyyy")
     .Range("g" & Idx) = Format(RS("orders"), gs_formatQty)
     .Range("h" & Idx) = Format(RS("receipt"), gs_formatQty)
     .Range("i" & Idx) = Format(RS("remaining"), gs_formatQty)
     If RS!complete_cls = "1" Then
      .Range("j" & Idx) = "Yes"
     Else
      .Range("j" & Idx) = "No"
     End If
     RS.MoveNext
      Idx = Idx + 1
     Loop
       
    .Range("a1", "j" & Idx).Columns.Font.Name = "Arial"
    .Range("a1", "j" & Idx).Columns.Font.Size = 8
          
    .Range("a2", "j2").Columns.Font.Name = "Arial"
    .Range("a2", "j2").Columns.Font.Size = "10"
    .Range("a2", "j2").Columns.Font.Bold = True
    .Range("a2", "j4").HorizontalAlignment = xlCenter
    .Range("a6", "j6").Columns.Font.Bold = True
   
    .Range("a11:d" & Idx).HorizontalAlignment = xlLeft
    .Range("f11:f" & Idx).NumberFormat = "[$-409]d-mmm-yyyy;@"
    .Range("g11:g" & Idx).NumberFormat = gs_formatQty
    .Range("h11:h" & Idx).NumberFormat = gs_formatQty
    .Range("i11:i" & Idx).NumberFormat = gs_formatQty
       
    .Visible = True
    .ActiveSheet.PageSetup.PaperSize = xlPaperA4
    .ActiveSheet.PageSetup.Orientation = 2
    .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
    .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.4)
    .Range("a:j").Columns.AutoFit
    .WindowState = xlMaximized
  End With
     
     
  Screen.MousePointer = vbDefault
  MousePointer = vbDefault
 End If
End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
  adtocombo
  Header
    
  With cboremaincls
    .clear
    .AddItem strAll
    .AddItem "Yes"
    .AddItem "No"
    
    .ListIndex = 0
  End With
  
  DelDate.Value = Format(Now, "dd MMM yyyy")
  deldate1.Value = Format(Now, "dd MMM yyyy")
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
End Sub

Private Sub cbosupplier_Click()
  If cboSupplier.ListIndex <> -1 Then
    LblName.Caption = cboSupplier.Column(1)
  End If
End Sub

Private Sub cbosupplier_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then cbosupplier_Click
End Sub

Private Sub deldate_Change()
   If CDate(DelDate) > CDate(deldate1) Then
      LblErrMsg.Caption = DisplayMsg(4076) & " " & Format(deldate1, "dd MMM yyyy")  '"delivery Date must be lower than "
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
End Sub

Private Sub deldate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then deldate_Change
End Sub

Private Sub deldate1_Change()
   If CDate(deldate1) < CDate(DelDate) Then
      LblErrMsg.Caption = DisplayMsg(4077) & " " & Format(DelDate, "dd MMM yyyy")   '"delivery Date must be higher than "
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
End Sub

Private Sub deldate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then deldate1_Change
End Sub

'Private Sub Grid_Click()
'With Grid
'    If .row = 1 Then
'      If .Col = bteColSupplier Or .Col = bteColName Or .Col = bteColDate Or .Col = 3 Then
'        If .ColSort(.Col) = flexSortGenericAscending Then
'           .ColSort(.Col) = flexSortGenericDescending
'        Else
'           .ColSort(.Col) = flexSortGenericAscending
'        End If
'      Else
'        If .ColSort(.Col) = flexSortNumericAscending Then
'           .ColSort(.Col) = flexSortNumericDescending
'        Else
'           .ColSort(.Col) = flexSortNumericAscending
'        End If
'      End If
'       .Sort = .ColSort(.Col)
'    End If
'End With
'End Sub

Private Sub cmdSearch_Click()
'    If cboGroup.Text = "" Then
'     lblErrMsg = DisplayMsg(8081)
'     cboGroup.SetFocus
'     Exit Sub
    If cboSupplier.Text = "" Then
     LblErrMsg = DisplayMsg(1009)
     cboSupplier.SetFocus
     Exit Sub
    ElseIf CDate(deldate1) < CDate(DelDate) Then
     LblErrMsg.Caption = DisplayMsg(4077) & " " & Format(DelDate, "dd MMM yyyy")   '"delivery Date must be higher than "
     deldate1.SetFocus
     Exit Sub
'    ElseIf cbocountry.Text = "" Then
'     lblErrMsg = DisplayMsg(8085)
'     'cbocountry.SetFocus
'     Exit Sub
    End If
    
'    If cboGroup.Text <> "" Then
'      cboGroup.MatchEntry = 1
'      cboGroup.Text = cboGroup.Text
'      If cboGroup.MatchFound = False Then
'          lblErrMsg = DisplayMsg(8083)
'          cboGroup.SetFocus
'          cboGroup.MatchEntry = 2
'          Exit Sub
'      End If
'      cboGroup.MatchEntry = 2
'    End If
        
    If cboSupplier.Text <> "" Then
      cboSupplier.MatchEntry = 1
      cboSupplier.Text = cboSupplier.Text
      If cboSupplier.MatchFound = False Then
          LblErrMsg = DisplayMsg(4003)
          cboSupplier.SetFocus
          cboSupplier.MatchEntry = 2
          Exit Sub
      End If
      cboSupplier.MatchEntry = 2
    End If
    
'    If cbocountry.Text <> "" Then
'     cbocountry.MatchEntry = 1
'     cbocountry.Text = cbocountry.Text
'     If cbocountry.MatchFound = False Then
'      lblErrMsg = DisplayMsg(8086)
'      cbocountry.SetFocus
'      cbocountry.MatchEntry = 2
'      Exit Sub
'     End If
'     cbocountry.MatchEntry = 2
'    End If
    
    Browse
End Sub

Private Sub Command1_Click(Index As Integer)
Dim top As Integer

Select Case Index
Case 0: Unload Me
        frmMainMenu.Show

Case 1: On Error Resume Next
        grid.TopRow = 1
        LblErrMsg.Caption = DisplayMsg(5007)    '"This is first page"

Case 2: On Error Resume Next
        top = grid.TopRow
        grid.TopRow = grid.TopRow - 10
        If top = grid.TopRow Then grid.TopRow = 1
            
Case 3: On Error Resume Next
        grid.TopRow = grid.TopRow + 10

Case 4: On Error Resume Next
        grid.TopRow = grid.Rows
        LblErrMsg.Caption = DisplayMsg(5008)    '"This is last page"

End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set RS = Nothing
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub cboGroup_Change()
lblgroup = ""
LblErrMsg = ""
End Sub

Private Sub CboGroup_Click()
cboGroup = cboGroup
    If cboGroup.MatchFound Then
        lblgroup = cboGroup.Column(1)
        LblErrMsg = ""
    Else
        lblgroup = ""
        LblErrMsg = DisplayMsg(8083)
    End If
End Sub

Private Sub CboGroup_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call CboGroup_Click
End Sub

Private Sub Cbocountry_Change()
LblErrMsg = ""
End Sub

Private Sub cbocountry_Click()
cbocountry = cbocountry
    If cbocountry.MatchFound Then
        LblErrMsg = ""
    Else
        LblErrMsg = DisplayMsg(8086)
    End If
End Sub

Private Sub cbocountry_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call cbocountry_Click
End Sub

