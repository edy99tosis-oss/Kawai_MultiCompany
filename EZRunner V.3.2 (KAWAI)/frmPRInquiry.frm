VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPRInquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Purchase Request Inquiry"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPRInquiry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Left            =   13845
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10080
      Width           =   1140
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10080
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   240
      TabIndex        =   13
      Top             =   9330
      Width           =   14730
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
         TabIndex        =   14
         Top             =   195
         Width           =   13725
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   2490
      Left            =   270
      TabIndex        =   8
      Top             =   810
      Width           =   14730
      Begin VB.ComboBox CboRem 
         Height          =   315
         ItemData        =   "frmPRInquiry.frx":0E42
         Left            =   1935
         List            =   "frmPRInquiry.frx":0E4F
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1950
         Width           =   885
      End
      Begin VB.CommandButton CmdSearch 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Search"
         Height          =   375
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1935
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker DtAwal 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   240
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
         Format          =   334036995
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker DtAkhir 
         Height          =   315
         Left            =   3990
         TabIndex        =   1
         Top             =   240
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
         Format          =   334036995
         CurrentDate     =   37810
      End
      Begin MSForms.ComboBox CboType 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   675
         Width           =   1875
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   7
         Size            =   "3307;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Request Type"
         Height          =   195
         Index           =   3
         Left            =   510
         TabIndex        =   20
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining Cls"
         Height          =   195
         Left            =   510
         TabIndex        =   19
         Top             =   2010
         Width           =   1335
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   5385
         X2              =   10455
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Label LblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   5385
         TabIndex        =   18
         Top             =   1590
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         Height          =   195
         Index           =   1
         Left            =   510
         TabIndex        =   17
         Top             =   1590
         Width           =   1155
      End
      Begin MSForms.ComboBox CboReqNo 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   1110
         Width           =   1875
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3307;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Request No"
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   12
         Top             =   1170
         Width           =   975
      End
      Begin MSForms.ComboBox CboItemCode 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   1530
         Width           =   3135
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "5530;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Request Date"
         Height          =   195
         Index           =   2
         Left            =   510
         TabIndex        =   11
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   255
         Left            =   3540
         TabIndex        =   10
         Top             =   270
         Width           =   375
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5805
      Left            =   240
      TabIndex        =   15
      Top             =   3465
      Width           =   14730
      _cx             =   25982
      _cy             =   10239
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
      Height          =   435
      Left            =   13035
      TabIndex        =   21
      Top             =   180
      Width           =   1845
      _extentx        =   3254
      _extenty        =   767
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Request Inquiry"
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
      Left            =   6180
      TabIndex        =   16
      Top             =   270
      Width           =   2880
   End
End
Attribute VB_Name = "frmPRInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClsProc As New ClsProc
Dim nilKosong As Boolean, i As Integer

Dim ColRequestor As Integer
Dim ColReqNo As Integer
Dim ColReqDate As Integer
Dim ColDelDate As Integer
Dim ColItemCode As Integer
Dim ColDesc As Integer
Dim ColReqQty As Integer
Dim colpoqty As Integer
Dim colrem As Integer
Dim ColUnit As Integer
Dim ColComplete As Integer

Sub Kosong()
    dtAwal = Format(Year(Now) & "-" & Format(Month(Now), "#0") & "-01", "dd MMM yyyy")
    dtAkhir = Format(Now, "dd MMM yyyy")
    
    CboType.ListIndex = 0
    CboRem.ListIndex = 0
    CboReqNo.clear
    CboReqNo.Text = ""
    cboItemCode.clear
    cboItemCode.Text = ""
    lblitem.Caption = ""
    
    Call isiCboRequest
    Call isiCboItem
    
    CboReqNo.ListIndex = 0
    cboItemCode.ListIndex = 0
    
End Sub

Sub isiCboRequest() 'Filter Request No
Dim rscbo As New ADODB.Recordset 'Data Customer

With CboReqNo
    .clear
    .columnCount = 1
    .TextColumn = 1
    
    sql = "select PORequest_No From PORequest_Master " & _
        "Where PORequest_Date >= '" & Format(dtAwal, "yyyy-MM-dd") & _
            "' And PORequest_Date <= '" & Format(dtAkhir, "yyyy-MM-dd") & "'"
            
    If CboType.ListIndex = 1 Then
        sql = sql + " And Others_Cls = '0' And SheetCoil_Cls = '0' "
    ElseIf CboType.ListIndex = 2 Then
        sql = sql + " And Others_Cls = '0' And SheetCoil_Cls = '1' "
    ElseIf CboType.ListIndex = 3 Then
        sql = sql + " And Others_Cls = '1' "
    End If
    
    Set rscbo = Db.Execute(sql)
    
    .AddItem ""
    .List(0, 0) = strAll
    
    i = 1
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo(0))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .ListWidth = 250
    .ColumnWidths = "250 pt"
    
    Set rscbo = Nothing
End With
End Sub

Sub isiCboItem()
Dim rscbo As New ADODB.Recordset

With cboItemCode
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    sql = "select PD.Item_Code, I.Item_Name From PORequest_Master PM, PORequest_Detail PD, Item_Master I " & _
        "Where PM.PORequest_No = PD.PORequest_No " & _
            "And PD.Item_Code = I.Item_Code "
                
    If CboReqNo.ListIndex = 0 Then
        sql = sql & "And PM.PORequest_Date >= '" & Format(dtAwal, "yyyy-MM-dd") & _
            "' And PM.PORequest_Date <= '" & Format(dtAkhir, "yyyy-MM-dd") & "' "
        If CboType.ListIndex = 1 Then
            sql = sql + " And PM.Others_Cls = '0' and PM.SheetCoil_Cls = '0' "
        ElseIf CboType.ListIndex = 2 Then
            sql = sql + " And PM.Others_Cls = '0' and PM.SheetCoil_Cls = '1' "
        ElseIf CboType.ListIndex = 3 Then
            sql = sql + " And PM.Others_Cls = '1' "
        End If
    Else
        sql = sql & "And PM.PORequest_No = '" & CboReqNo.Text & "' "
    End If
    sql = sql & "Order By PD.Item_Code"
    Set rscbo = Db.Execute(sql)
    
    .AddItem ""
    .List(0, 0) = strAll
    .List(0, 1) = strAll
    
    i = 1
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo("Item_Code"))
        .List(i, 1) = Trim(rscbo("Item_Name"))
        
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .ListWidth = 450
    .ColumnWidths = "150 pt;300 pt"
    Set rscbo = Nothing
End With
End Sub

Private Sub CboItemCode_Change()
cboItemCode.Text = cboItemCode.Text
lblitem.Caption = ""
If cboItemCode.MatchFound = True Then
    lblitem.Caption = cboItemCode.List(cboItemCode.ListIndex, 1)
End If
headerGrid
End Sub

Private Sub cboitemcode_Click()
cboItemCode.Text = cboItemCode.Text
lblitem.Caption = ""
If cboItemCode.MatchFound = True Then
    lblitem.Caption = cboItemCode.List(cboItemCode.ListIndex, 1)
End If
headerGrid
End Sub

Private Sub CboRem_Change()
headerGrid
End Sub

Private Sub CboRem_Click()
headerGrid
End Sub

Private Sub CboReqNo_Change()
Call isiCboItem
headerGrid
End Sub

Private Sub CboReqNo_Click()
Call isiCboItem
headerGrid
End Sub

Private Sub cboType_Change()
Call isiCboRequest
headerGrid
End Sub

Private Sub CboType_Click()
Call isiCboRequest
headerGrid
End Sub

Private Sub cmdClear_Click()
Kosong
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub Form_Load()
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    
    dtAkhir.Value = Now
    dtAwal.Value = Now
    
    ColRequestor = 0
    ColReqNo = 1
    ColReqDate = 2
    ColDelDate = 3
    ColItemCode = 4
    ColDesc = 5
    ColReqQty = 6
    colpoqty = 7
    colrem = 8
    ColUnit = 9
    ColComplete = 10

    CboType.AddItem strAll
    CboType.AddItem "Part/Material"
    CboType.AddItem "Sheet/Coil"
    CboType.AddItem "Other Item"
    
    Call Kosong
    Call headerGrid
End Sub

Public Sub cmdSearch_Click()
If CDate(dtAwal) > CDate(dtAkhir) Then
    LblErrMsg.Caption = "Start Date must be lower than " & dtAkhir.Value & " !!!"
    Exit Sub
ElseIf CDate(dtAkhir) < CDate(dtAwal) Then
    LblErrMsg.Caption = "End Date must be higher than " & dtAwal.Value & " !!!"
    Exit Sub
End If
    
    CboReqNo.Text = CboReqNo.Text
    cboItemCode.Text = cboItemCode.Text
    If Trim(CboReqNo.Text) = "" Then
        LblErrMsg = DisplayMsg(1067) 'Please Input Req No
        CboReqNo.SetFocus
    ElseIf CboReqNo.MatchFound = False Then
        LblErrMsg = DisplayMsg(4081) 'Req No Not Found
        CboReqNo.SetFocus
    ElseIf Trim(cboItemCode.Text) = "" Then
        LblErrMsg = DisplayMsg(1009) 'Please Input Product Code
        cboItemCode.SetFocus
    ElseIf cboItemCode.MatchFound = False Then
        LblErrMsg = DisplayMsg(4003) 'Product Code Not Found
        cboItemCode.SetFocus
    Else
        LblErrMsg = ""
        Call IsiGrid
    End If
End Sub

Private Sub headerGrid()
Dim i As Integer

With grid
    .clear
    .ColS = 11
    .Rows = 1
       
    .ColWidth(ColRequestor) = 1200
    .ColWidth(ColReqNo) = 1250
    .ColWidth(ColReqDate) = 1300
    .ColWidth(ColDelDate) = 1400
    .ColWidth(ColItemCode) = 2500
    .ColWidth(ColDesc) = 3500
    .ColWidth(ColReqQty) = 1200
    .ColWidth(colpoqty) = 1200
    .ColWidth(colrem) = 1100
    .ColWidth(ColUnit) = 500
    .ColWidth(ColComplete) = 900
    
    .TextMatrix(0, ColRequestor) = "Dept. Req"
    .TextMatrix(0, ColReqNo) = "Request No"
    .TextMatrix(0, ColReqDate) = "Request Date"
    .TextMatrix(0, ColDelDate) = "Req Dlvy Date"
    .TextMatrix(0, ColItemCode) = "Product Code"
    .TextMatrix(0, ColDesc) = "Description"
    .TextMatrix(0, ColReqQty) = "Req Qty"
    .TextMatrix(0, colpoqty) = "PO Qty"
    .TextMatrix(0, colrem) = "Remaining"
    .TextMatrix(0, ColUnit) = "Unit"
    .TextMatrix(0, ColComplete) = "Complete"
    
    .ColAlignment(ColRequestor) = flexAlignCenterCenter
    .ColAlignment(ColReqNo) = flexAlignCenterCenter
    .ColAlignment(ColReqDate) = flexAlignLeftCenter
    .ColAlignment(ColDelDate) = flexAlignCenterCenter
    .ColAlignment(ColItemCode) = flexAlignLeftCenter
    .ColAlignment(ColDesc) = flexAlignLeftCenter
    .ColAlignment(ColReqQty) = flexAlignRightCenter
    .ColAlignment(colpoqty) = flexAlignRightCenter
    .ColAlignment(colrem) = flexAlignRightCenter
    .ColAlignment(ColUnit) = flexAlignCenterCenter
    .ColAlignment(ColComplete) = flexAlignCenterCenter
    
    .EditMaxLength = 1
    Call ClsProc.AlignHeader(grid)
End With
End Sub

Sub IsiGrid()
Dim rsGrid As New ADODB.Recordset
Dim sqlResult As String


With grid
    Call headerGrid

    sql = " Select  " & _
                " dt.*,(select description from unit_cls uc where uc.unit_cls= dt.unit_cls ) unit_desc ,  " & _
                " Sisa = Case dt.complete_Cls When 1 Then 0 Else ReqQty - totPO  End  " & _
                " From  " & _
                " (SELECT     department_cls.description, PM.PORequest_No, PM.PORequest_Date, PD.ReqDelivery_Date, ISNULL(PM.Fix_Cls, 0) AS Fix_Cls, ISNULL(PM.complete_Cls, 0) AS complete_Cls, PD.Item_Code, ISNULL(I.Item_Name,  " & _
                "                       PD.Item_Name) AS Item_name, PM.Others_Cls, PM.SheetCoil_Cls, ISNULL(I.Unit_Cls, PD.Unit_Cls) AS Unit_Cls, PD.Qty AS ReqQty, ISNULL " & _
                "                           ((SELECT     SUM(Qty) " & _
                "                               FROM         PurchaseOrder_Master POM, PurchaseOrder_Detail POD " & _
                "                               WHERE     POM.PO_No = POD.PO_No AND POD.PORequest_No = PM.PORequest_No AND POD.POReq_Seqno = PD.PoReq_seqno), 0) AS totPO " & _
                " FROM         dbo.PORequest_Master PM INNER JOIN " & _
                "                       dbo.PORequest_Detail PD ON PM.PORequest_No = PD.PORequest_No LEFT OUTER JOIN " & _
                "                       dbo.Item_Master I ON PD.Item_Code = I.Item_Code " & _
                "              inner join department_Cls on department_cls.department_cls = PM.Department_Cls " & _
                " )dt  Where "
                
    sql = sql & " dt.PORequest_Date >= '" & Format(dtAwal, "yyyy-MM-dd") & _
            "' AND dt.PORequest_Date <= '" & Format(dtAkhir, "yyyy-MM-dd") & "' "
    
    If CboType.ListIndex = 1 Then
        sql = sql + " And dt.Others_Cls = '0' and dt.SheetCoil_Cls = '0' "
    ElseIf CboType.ListIndex = 2 Then
        sql = sql + " And dt.Others_Cls = '0' and dt.SheetCoil_Cls = '1' "
    ElseIf CboType.ListIndex = 3 Then
        sql = sql + " And dt.Others_Cls = '1' "
    End If
    
    If CboReqNo.ListIndex <> 0 Then
        sql = sql & " AND dt.PORequest_No = '" & CboReqNo.Text & "' "
    End If
    
    'Item
    If cboItemCode.ListIndex > 0 Then sql = sql & "And Item_Code = '" & cboItemCode & "' "
    
    If CboRem.ListIndex = 1 Then  'Remaining
        sql = sql & "And (ReqQty > totPO And Fix_Cls = 0) "
    ElseIf CboRem.ListIndex = 2 Then
        sql = sql & "And (ReqQty <= totPO Or Fix_Cls = 1) "
    End If
    
    sql = sql & "Order By PORequest_No, PORequest_Date, Item_Code"
    Set rsGrid = Db.Execute(sql)

    i = 1
    If Not (rsGrid.EOF) Then
        Do While Not rsGrid.EOF
            .Rows = .Rows + 1
            .TextMatrix(i, ColRequestor) = Trim(rsGrid("description"))
            .TextMatrix(i, ColReqNo) = Trim(rsGrid("PORequest_No"))
            .TextMatrix(i, ColReqDate) = Format(rsGrid("PORequest_Date"), "dd MMM yyyy")
            .TextMatrix(i, ColDelDate) = Format(rsGrid("ReqDelivery_Date"), "dd MMM yyyy")
            .TextMatrix(i, ColItemCode) = Trim(rsGrid("Item_Code"))
            .TextMatrix(i, ColDesc) = Trim(rsGrid("Item_Name"))
            .TextMatrix(i, ColReqQty) = ClsProc.formatnilai(rsGrid("ReqQty"))
            .TextMatrix(i, colpoqty) = ClsProc.formatnilai(rsGrid("totPO"))
            .TextMatrix(i, colrem) = ClsProc.formatnilai(rsGrid("Sisa"))
            .TextMatrix(i, ColUnit) = Trim(rsGrid("Unit_desc"))
            .Cell(flexcpChecked, i, ColComplete) = IIf(rsGrid("complete_Cls") = 0, flexUnchecked, flexChecked)
            i = i + 1
            rsGrid.MoveNext
        Loop
    Else
        LblErrMsg = DisplayMsg(4006)
    End If
    Set rsGrid = Nothing
End With
End Sub

Private Sub dtAwal_Change()
If CDate(dtAwal) > CDate(dtAkhir) Then
    LblErrMsg.Caption = "Start Date must be lower than " & dtAkhir.Value & " !!!"
    Exit Sub
ElseIf CDate(dtAkhir) < CDate(dtAwal) Then
    LblErrMsg.Caption = "End Date must be higher than " & dtAwal.Value & " !!!"
    Exit Sub
End If
    
    
    Call headerGrid
    Call isiCboRequest
    If Format(dtAwal, "yyyy-MM-dd") > Format(dtAkhir, "yyyy-MM-dd") Then LblErrMsg = DisplayMsg(4068) & " " & Format(dtAkhir, "dd MMM yyyy") Else LblErrMsg = ""
End Sub

Private Sub dtAkhir_Change()
If CDate(dtAwal) > CDate(dtAkhir) Then
    LblErrMsg.Caption = "Start Date must be lower than " & dtAkhir.Value & " !!!"
    Exit Sub
ElseIf CDate(dtAkhir) < CDate(dtAwal) Then
    LblErrMsg.Caption = "End Date must be higher than " & dtAwal.Value & " !!!"
    Exit Sub
End If


    Call headerGrid
    Call isiCboRequest
    If Format(dtAwal, "yyyy-MM-dd") > Format(dtAkhir, "yyyy-MM-dd") Then LblErrMsg = DisplayMsg(4066) & " " & Format(dtAwal, "dd MMM yyyy") Else LblErrMsg = ""
End Sub

Private Sub cboRemaining_Click()
    Call headerGrid
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
    If ErrMsg = "" Then Unload Me Else LblErrMsg.Caption = ErrMsg
End Sub




