VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmPOResultInquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Purchase Order / Result Inquiry (Each Supplier)"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "FrmPOResultInquiry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdexcel 
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
      Left            =   13770
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9930
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   360
      TabIndex        =   19
      Top             =   9120
      Width           =   14550
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
         TabIndex        =   20
         Top             =   195
         Width           =   14205
      End
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   4223
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Index           =   2
      Left            =   5543
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Index           =   3
      Left            =   6818
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Index           =   4
      Left            =   8138
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9960
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1215
      Left            =   360
      TabIndex        =   14
      Top             =   1320
      Width           =   14520
      Begin VB.TextBox txtPO 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10155
         TabIndex        =   4
         Top             =   720
         Width           =   2340
      End
      Begin VB.CommandButton cmdsearch 
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
         Left            =   13020
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   690
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker deldate 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   720
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
         Format          =   150994947
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker deldate1 
         Height          =   315
         Left            =   3720
         TabIndex        =   2
         Top             =   720
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
         Format          =   150994947
         CurrentDate     =   37810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO No."
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
         Left            =   9375
         TabIndex        =   23
         Top             =   765
         Width           =   585
      End
      Begin MSForms.ComboBox cboremaincls 
         Height          =   315
         Left            =   7905
         TabIndex        =   3
         Top             =   720
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining Cls"
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
         Left            =   6450
         TabIndex        =   21
         Top             =   765
         Width           =   1230
      End
      Begin VB.Label Label7 
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
         Left            =   315
         TabIndex        =   17
         Top             =   765
         Width           =   1185
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
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
         Left            =   3735
         TabIndex        =   16
         Top             =   300
         Width           =   5535
      End
      Begin MSForms.ComboBox cbosupplier 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   255
         Width           =   1950
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3440;556"
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
         Caption         =   "Supplier CD"
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
         TabIndex        =   15
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
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
         Left            =   3375
         TabIndex        =   18
         Top             =   765
         Width           =   195
      End
      Begin VB.Line Line2 
         X1              =   3735
         X2              =   9255
         Y1              =   540
         Y2              =   540
      End
   End
   Begin MSComDlg.CommonDialog cdlReport 
      Left            =   390
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   6315
      Left            =   360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2760
      Width           =   14520
      _cx             =   25612
      _cy             =   11139
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
      Left            =   12960
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   360
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order / Result Inquiry (Each Supplier)"
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
      Left            =   390
      TabIndex        =   13
      Top             =   360
      Width           =   14460
   End
End
Attribute VB_Name = "FrmPOResultInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim RS As New ADODB.Recordset

Dim bteSuppCode As Byte
Dim bteSuppName As Byte
Dim bteDate As Byte
Dim bteMatCode As Byte
Dim bteDesc As Byte
Dim btePONo As Byte
Dim btePlan As Byte
Dim bteResult As Byte
Dim bteRemain As Byte
Dim bteComplete As Byte

Sub Header()
    
    bteSuppCode = 0
    bteSuppName = 1
    bteMatCode = 2
    bteDesc = 3
    btePONo = 4
    bteDate = 5
    btePlan = 6
    bteResult = 7
    bteRemain = 8
    bteComplete = 9
    
    With grid
        .Rows = 1
        .ColS = 10
        
        .TextMatrix(0, bteSuppCode) = "Supplier Code"
        .TextMatrix(0, bteSuppName) = "Supplier Name"
        .TextMatrix(0, bteDate) = "Delivery Date"
        .TextMatrix(0, bteMatCode) = "Material Code"
        .TextMatrix(0, bteDesc) = "Description"
        .TextMatrix(0, btePONo) = "PO No"
        .TextMatrix(0, btePlan) = "Plan"
        .TextMatrix(0, bteResult) = "Result"
        .TextMatrix(0, bteRemain) = "Remaining"
        .TextMatrix(0, bteComplete) = "Complete"
        
        .ColWidth(bteSuppCode) = 1500
        .ColWidth(bteSuppName) = 2000
        .ColWidth(bteDate) = 1300
        .ColWidth(bteMatCode) = 1500
        .ColWidth(bteDesc) = 3500
        .ColWidth(btePONo) = 2700
        .ColWidth(btePlan) = 1000
        .ColWidth(bteResult) = 1000
        .ColWidth(bteRemain) = 1000
        .ColWidth(bteComplete) = 975
        
        .ColDataType(bteDate) = flexDTDate
        
        .ColAlignment(bteSuppCode) = flexAlignLeftCenter
        .ColAlignment(bteSuppName) = flexAlignLeftCenter
        .ColAlignment(bteDate) = flexAlignCenterCenter
        .ColAlignment(bteMatCode) = flexAlignLeftCenter
        .ColAlignment(bteDesc) = flexAlignLeftCenter
        .ColAlignment(btePONo) = flexAlignLeftCenter
        .ColAlignment(btePlan) = flexAlignRightCenter
        .ColAlignment(bteResult) = flexAlignRightCenter
        .ColAlignment(bteRemain) = flexAlignRightCenter
        .ColAlignment(bteComplete) = flexAlignCenterCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
        
        If cboSupplier.Text = strAll Then
            .ColHidden(bteSuppCode) = False
            .ColHidden(bteSuppName) = False
        Else
            .ColHidden(bteSuppCode) = True
            .ColHidden(bteSuppName) = True
        End If
        
        
    End With
    
    LblErrMsg.Caption = ""
 
End Sub

Sub adtocombo()
sql = "SELECT Trade_Code, Trade_Name FROM Trade_Master where trade_cls='2' or trade_cls='3'"
Set RS = Db.Execute(sql)

With cboSupplier
.clear
.columnCount = 2
.ColumnWidths = "80 pt;300 pt"
.ListWidth = 380
.ListRows = 15
.AddItem ""
.List(0, 0) = strAll
.List(0, 1) = strAll
i = 1
Do Until RS.EOF
    .AddItem ""
    .List(i, 0) = Trim(RS!Trade_Code)
    .List(i, 1) = Trim(RS!trade_name)
    i = i + 1
    RS.MoveNext
Loop
.ListIndex = 0
End With
End Sub

Private Sub Browse()

Me.MousePointer = vbHourglass

Header

sql = "select pm.supplier_code, tm.trade_name, pd.delivery_date, pd.item_code, " & vbCrLf & _
      "im.item_name, pm.po_no, pd.complete_cls, isnull(pd.qty,0) as orders, " & vbCrLf & _
      "isnull((select sum(qty) as qty from part_receipt where supplier_code=pm.supplier_code and " & vbCrLf & _
      "po_no=pm.po_no and item_code=pd.item_code),0) as receipt, " & vbCrLf & _
      "(case pd.complete_cls when 1 then 0 else (isnull(pd.qty,0) - isnull((select sum(qty) from part_receipt where supplier_code=pm.supplier_code and " & vbCrLf & _
      "po_no=pm.po_no and item_code=pd.item_code),0)) end) as remaining " & vbCrLf & _
      "from purchaseorder_master pm inner join purchaseorder_detail pd on pm.po_no=pd.Po_No " & vbCrLf & _
      " Inner Join  trade_master tm  on  pm.supplier_code = tm.trade_code " & vbCrLf & _
      " Left Join item_master im on pd.item_code = im.item_code" & vbCrLf & _
      "Where 'A'='A' " & vbCrLf & _
      "and pm.po_no like '%" & txtpo.Text & "%' " & vbCrLf

If cboSupplier <> strAll Then sql = sql & "and pm.supplier_code='" & Trim(cboSupplier.Text) & "' "
If cboremaincls.ListIndex = 0 Then
sql = sql & " and pd.delivery_date>='" & Format(DelDate, "yyyy-mm-dd") & "' and pd.delivery_date<='" & _
      Format(deldate1.Value, "yyyy-mm-dd") & "' and (case pd.complete_cls when 1 then 0 else (isnull(pd.qty,0) - isnull((select sum(qty) from part_receipt where supplier_code=pm.supplier_code and " & _
      "po_no=pm.po_no and item_code=pd.item_code),0)) end) > 0 "
Else
sql = sql & " and pd.delivery_date>='" & Format(DelDate, "yyyy-mm-dd") & "' and pd.delivery_date<='" & _
      Format(deldate1.Value, "yyyy-mm-dd") & "' and (case pd.complete_cls when 1 then 0 else (isnull(pd.qty,0) - isnull((select sum(qty) from part_receipt where supplier_code=pm.supplier_code and " & _
      "po_no=pm.po_no and item_code=pd.item_code),0)) end) <= 0 "
End If

sql = sql & "Order By im.Supplier_Code, PM.Po_No "

Set RS = Db.Execute(sql)

If (RS.BOF And RS.EOF) Then
    LblErrMsg.Caption = DisplayMsg(4006)
    Me.MousePointer = vbDefault
    Exit Sub
Else
i = 1
With grid
    Do While Not RS.EOF
        .Rows = .Rows + 1
        .TextMatrix(i, bteSuppCode) = Trim(RS("supplier_code"))
        .TextMatrix(i, bteSuppName) = Trim(RS("trade_name"))
        .TextMatrix(i, bteDate) = Format(RS("delivery_date"), "dd mmm yyyy")
        .TextMatrix(i, bteMatCode) = Trim(RS("item_code") & "")
        .TextMatrix(i, bteDesc) = IIf(IsNull(RS("item_name")), " ", Trim(RS("item_name")))
        .TextMatrix(i, btePONo) = Trim(RS("po_no"))
        .TextMatrix(i, btePlan) = Format(RS("orders"), gs_formatQty)
        .TextMatrix(i, bteResult) = Format(RS("receipt"), gs_formatQty)
        .TextMatrix(i, bteRemain) = Format(RS("remaining"), gs_formatQty)
        If RS("complete_cls") = 1 Then
            .Cell(flexcpChecked, i, bteComplete) = flexChecked
        Else
            .Cell(flexcpChecked, i, bteComplete) = flexUnchecked
        End If
        
            
        i = i + 1
        RS.MoveNext
    Loop
End With
End If
Me.MousePointer = vbDefault
End Sub
Private Sub toExcel()
        
    Dim adoRs As New ADODB.Recordset
    
    Dim xlapp As New Excel.application
    Dim Idx As Long
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Dim strFileName As String
    
    Me.MousePointer = vbHourglass
    On Error GoTo errHandler
    
    LblErrMsg.Caption = ""
    
    If grid.Rows > 1 Then
        
                
        Me.MousePointer = vbHourglass
                
        sql = "Select Company_Name, Address1, Address2, Province, City, Postal_Code, Phone1, Phone2, Fax From Company_Profile"
        If adoRs.State <> adStateClosed Then adoRs.Close
        adoRs.Open sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
    
    Screen.MousePointer = vbHourglass
    With xlapp

        .Workbooks.Add
            
        .Range("A1:A11").EntireRow.Insert
        .Range("A2:J2").Merge
        .Range("A3:J3").Merge
        .Range("A4:J4").Merge
        .Range("A6:J6").Merge
        .Range("A2:J4").HorizontalAlignment = xlHAlignCenter
                             
        .Range("A2").Font.Bold = True
        .Range("A2").Font.Size = 10
        .Range("A2") = Trim(adoRs.Fields("Company_Name"))
        .Range("A3") = Trim(adoRs.Fields("Address1")) & " " & Trim(adoRs.Fields("Address2")) & " " & Trim(adoRs.Fields("City")) & " " & Trim(adoRs.Fields("Postal_Code")) & ", " & Trim(adoRs.Fields("Province"))
        .Range("A4") = "Phone : " & Trim(adoRs.Fields("Phone1")) & " " & Trim(adoRs.Fields("Phone2")) & " Fax : " & Trim(adoRs.Fields("Fax"))
        
        .Range("A6").Font.Bold = True
        .Range("A6").Font.Size = 8
        .Range("A6") = "Purchase Order / Result Inquiry (Each Supplier)"
        .Range("A7") = "Supplier Code"
        .Range("B7", "J7").Merge
        .Range("B7") = ": " & cboSupplier.Column(0) & " / " & cboSupplier.Column(1)
        .Range("B7").HorizontalAlignment = xlLeft
        .Range("A8") = "Delivery Date"
        .Range("B8", "J8").Merge
        .Range("B8") = ": " & Format(DelDate.Value, "[$-409]d-mmm-yyyy;@") & " to " & Format(deldate1.Value, "[$-409]d-mmm-yyyy;@")
        .Range("A9") = "Remaining Cls"
        .Range("B9", "J9").Merge
        .Range("B9") = ": " & cboremaincls.Column(0)
                   
        '********************Garis Header****************************
        .Range("A11:J11").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("A11:J11").Borders(xlEdgeBottom).LineStyle = xlDouble
        '***********************************************************
        
        '********************Garis Footer************************************************************
        .Range("A" & 11 + grid.Rows - 1 & ":J" & 11 + grid.Rows - 1).Borders(xlEdgeBottom).LineStyle = xlDouble
        '********************************************************************************************
        
        .Range("A11:J11").HorizontalAlignment = xlHAlignCenter
                    
        .Range("A11") = "Supplier Code"
        .Range("B11") = "Supplier Name"
        .Range("C11") = "Material Code"
        .Range("D11") = "Description"
        .Range("E11") = "PO No."
        .Range("F11") = "Delivery Date"
        .Range("G11") = "Plan"
        .Range("H11") = "Result"
        .Range("I11") = "Remaining"
        .Range("J11") = "Completed"
                    
        '******Complete*********
        Dim Row As Integer, i As Integer
        Row = 1
        i = 12
        Do While i <= 10 + grid.Rows
            .Range("A" & i) = grid.TextMatrix(Row, bteSuppCode)
            .Range("B" & i) = grid.TextMatrix(Row, bteSuppName)
            .Range("C" & i) = grid.TextMatrix(Row, bteMatCode)
            .Range("D" & i) = grid.TextMatrix(Row, bteDesc)
            .Range("E" & i) = grid.TextMatrix(Row, btePONo)
            .Range("F" & i) = grid.TextMatrix(Row, bteDate)
            .Range("G" & i) = grid.TextMatrix(Row, btePlan)
            .Range("H" & i) = grid.TextMatrix(Row, bteResult)
            .Range("I" & i) = grid.TextMatrix(Row, bteRemain)
            
          If grid.Cell(flexcpChecked, Row, bteComplete) = flexChecked Then
           .Range("J" & i) = "Yes"
          Else
           .Range("J" & i) = "No"
          End If
        i = i + 1
        Row = Row + 1
        Loop
        '***********************
        
        .Range("A1", "J" & 11 + grid.Rows - 1).Columns.Font.Name = "Arial"
        .Range("A11", "J" & 11 + grid.Rows - 1).Columns.Font.Size = 8
        
        .Range("F12:F" & 11 + grid.Rows - 1).NumberFormat = "[$-409]d-mmm-yyyy;@"
        .Range("A12:E" & 11 + grid.Rows - 1).HorizontalAlignment = xlLeft
        .Range("g12:g" & 11 + grid.Rows - 1).NumberFormat = gs_formatQty
        .Range("h12:h" & 11 + grid.Rows - 1).NumberFormat = gs_formatQty
        .Range("i12:i" & 11 + grid.Rows - 1).NumberFormat = gs_formatQty
        .Range("A:J").Columns.AutoFit

        .Visible = True
        .ActiveSheet.PageSetup.PaperSize = xlPaperA4
        .ActiveSheet.PageSetup.Orientation = 2
        .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
        .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.4)
        .WindowState = xlMaximized
    End With
    adoRs.Close
        
    Screen.MousePointer = vbDefault
    MousePointer = vbDefault
    Else
        LblErrMsg.Caption = DisplayMsg("0013")
    End If

ErrExit:
    Set adoRs = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
    
End Sub

Private Sub cboremaincls_Change()
Header
End Sub
Private Sub CboSupplier_Change()
Header
End Sub
Private Sub CmdExcel_Click()
Call toExcel
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
  Header
  adtocombo
    
  With cboremaincls
    .clear
    
    .AddItem "Yes"
    .AddItem "No"
    
    .ListIndex = 0
  End With
  
  DelDate.Value = Format(Now, "dd MMM yyyy")
  deldate1.Value = Format(Now, "dd MMM yyyy")
    
  'Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
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
Header
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
Header
End Sub

Private Sub deldate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then deldate1_Change
End Sub

Private Sub cmdSearch_Click()
    If cboSupplier.Text = "" Then
        LblErrMsg = DisplayMsg(1054)
        cboSupplier.SetFocus
        Exit Sub
    ElseIf CDate(deldate1) < CDate(DelDate) Then
      LblErrMsg.Caption = DisplayMsg(4077) & " " & Format(DelDate, "dd MMM yyyy")   '"delivery Date must be higher than "
      deldate1.SetFocus
      Exit Sub
    End If
    
    If cboSupplier.Text <> "" Then
      cboSupplier.MatchEntry = 1
      cboSupplier.Text = cboSupplier.Text
      If cboSupplier.MatchFound = False Then
          LblErrMsg = DisplayMsg(4050)
          cboSupplier.SetFocus
          cboSupplier.MatchEntry = 2
          Exit Sub
      End If
      cboSupplier.MatchEntry = 2
    End If
    
    Browse
End Sub

Private Sub Command1_Click(Index As Integer)
Dim top As Integer
LblErrMsg.Caption = ""

Select Case Index
Case 0: Unload Me
        frmMainMenu.Show

Case 1: On Error Resume Next
        grid.TopRow = 1
        LblErrMsg.Caption = DisplayMsg(4020)    '"This is first page"

Case 2: On Error Resume Next
        top = grid.TopRow
        grid.TopRow = grid.TopRow - 10
        If top = grid.TopRow Then grid.TopRow = 1
            
Case 3: On Error Resume Next
        grid.TopRow = grid.TopRow + 10

Case 4: On Error Resume Next
        grid.TopRow = grid.Rows
        LblErrMsg.Caption = DisplayMsg(4021)    '"This is last page"

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

