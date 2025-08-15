VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ProdResultAutoRequest 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Parts (Material) Supply Request [Automatic]"
   ClientHeight    =   10980
   ClientLeft      =   105
   ClientTop       =   405
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
   Icon            =   "frm_ProdResultAutoRequest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin EZRunnerv3.Anchor Anchor1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   300
      TabIndex        =   19
      Top             =   9375
      Width           =   14715
      Begin VB.Label LblErrMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   240
         Left            =   60
         TabIndex        =   20
         Top             =   210
         Width           =   14550
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "To Material Request"
      Height          =   375
      Index           =   1
      Left            =   13125
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10050
      Width           =   1860
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Search"
      Height          =   405
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1065
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10050
      Width           =   1140
   End
   Begin VB.ComboBox cboRemaining 
      Height          =   315
      ItemData        =   "frm_ProdResultAutoRequest.frx":0E42
      Left            =   7140
      List            =   "frm_ProdResultAutoRequest.frx":0E4C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2565
      Width           =   885
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1215
      Left            =   300
      TabIndex        =   11
      Top             =   1170
      Width           =   14715
      Begin VB.Label lblItemCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Left            =   10065
         TabIndex        =   25
         Top             =   855
         Width           =   3525
      End
      Begin MSForms.ComboBox cboitemcode 
         Height          =   315
         Left            =   8085
         TabIndex        =   24
         Top             =   795
         Width           =   1890
         VariousPropertyBits=   746604571
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "3334;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         Height          =   195
         Index           =   6
         Left            =   6465
         TabIndex        =   23
         Top             =   855
         Width           =   1155
      End
      Begin VB.Line Line8 
         Index           =   3
         X1              =   10065
         X2              =   13575
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Line Line8 
         Index           =   2
         X1              =   9705
         X2              =   13215
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warehouse Code"
         Height          =   195
         Index           =   5
         Left            =   6465
         TabIndex        =   22
         Top             =   330
         Width           =   1470
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   2
         Left            =   8085
         TabIndex        =   2
         Top             =   270
         Width           =   1530
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2699;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblNm 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   2
         Left            =   9705
         TabIndex        =   21
         Top             =   330
         Width           =   3525
      End
      Begin VB.Label lblNm 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   3285
         TabIndex        =   15
         Top             =   330
         Width           =   3045
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machine No"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   14
         Top             =   780
         Width           =   975
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   1665
         TabIndex        =   0
         Top             =   270
         Width           =   1530
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2699;556"
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
         Left            =   1650
         TabIndex        =   1
         Top             =   720
         Width           =   1530
         VariousPropertyBits=   746604571
         MaxLength       =   3
         DisplayStyle    =   3
         Size            =   "2699;556"
         ListRows        =   15
         ShowDropButtonWhen=   2
         Value           =   "AAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblNm 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   3270
         TabIndex        =   13
         Top             =   780
         Width           =   1515
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   3270
         X2              =   4770
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory Code"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   12
         Top             =   330
         Width           =   1140
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   3285
         X2              =   6315
         Y1              =   570
         Y2              =   570
      End
   End
   Begin MSComCtl2.DTPicker dtAwal 
      Height          =   330
      Left            =   1950
      TabIndex        =   3
      Top             =   2550
      Width           =   1530
      _ExtentX        =   2699
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
      Format          =   128909315
      CurrentDate     =   37860
   End
   Begin MSComCtl2.DTPicker dtAkhir 
      Height          =   330
      Left            =   3840
      TabIndex        =   4
      Top             =   2550
      Width           =   1530
      _ExtentX        =   2699
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
      Format          =   128909315
      CurrentDate     =   37891
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6255
      Left            =   300
      TabIndex        =   9
      Top             =   3045
      Width           =   14715
      _cx             =   25956
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
      GridColor       =   8421504
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
      Left            =   12960
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "FTTF*/"
      Top             =   240
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
      Left            =   5565
      TabIndex        =   18
      Top             =   2625
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   195
      Index           =   3
      Left            =   3585
      TabIndex        =   17
      Top             =   2625
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule Date :"
      Height          =   195
      Index           =   2
      Left            =   465
      TabIndex        =   16
      Top             =   2625
      Width           =   1380
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Parts (Material) Supply Request [Automatic]"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   300
      TabIndex        =   10
      Top             =   330
      Width           =   14715
   End
End
Attribute VB_Name = "frm_ProdResultAutoRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim i As Long
Dim nilKosong As Boolean

Public is_request As String
Public is_groupRequest As String
Public il_selectedRecord As Long
Public fromProd As Boolean

Dim ls_supReqNoAcc As String
Dim ls_group As String
Dim ls_dailySeqNo As String
Dim ls_str As String

Dim arrBOM() As String
Dim IntIndex As Integer
Dim tempRowBom As Integer

Dim bteColSelect As Byte
Dim bteColScheduleDate As Byte
Dim bteColProductCode As Byte
Dim bteColPartCode As Byte
Dim bteColDesc As Byte
Dim bteColLotNo As Byte
Dim bteColPackCode As Byte
Dim bteColPackDesc As Byte
Dim bteColPackSize As Byte
Dim bteColCustName As Byte
Dim bteColPlan As Byte
Dim bteColResult As Byte
Dim bteColRemaining As Byte
Dim bteColWHCode As Byte
Dim bteColSeqNo As Byte
Dim bteColUnitCls As Byte
Dim bteColReqCls As Byte
Dim bteColGroupCls As Byte
Dim bteColRequest As Byte
Dim bteColRequestNo As Byte
Dim bteColControlCls As Byte
Dim bteColSupplyCls As Byte
Dim bteColWominNo As Byte

Private Sub headerGrid()
    
    bteColSelect = 0
    bteColScheduleDate = 1
    bteColProductCode = 2
    bteColPartCode = 3
    bteColDesc = 4
    bteColLotNo = 5
    bteColPackCode = 6
    bteColPackDesc = 7
    bteColPackSize = 8
    bteColCustName = 9
    bteColPlan = 10
    bteColResult = 11
    bteColRemaining = 12
    bteColSeqNo = 13
    bteColWHCode = 14
    bteColUnitCls = 15
    bteColReqCls = 16
    bteColGroupCls = 17
    bteColRequest = 18
    bteColRequestNo = 19
    bteColControlCls = 20
    bteColSupplyCls = 21
    bteColWominNo = 22
    
    With grid
        .clear
        .ColS = 23
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = "S"
        .TextMatrix(0, bteColScheduleDate) = "Schedule Date"
        .TextMatrix(0, bteColProductCode) = "Product CD"
        .TextMatrix(0, bteColPartCode) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColLotNo) = "Lot No"
        .TextMatrix(0, bteColPackCode) = "Packing Code"
        .TextMatrix(0, bteColPackDesc) = "Packing Desc"
        .TextMatrix(0, bteColPackSize) = "P. Size"
        .TextMatrix(0, bteColCustName) = "Customer Name"
        .TextMatrix(0, bteColPlan) = "Plan"
        .TextMatrix(0, bteColResult) = "Result"
        .TextMatrix(0, bteColRemaining) = "Remaining"
        .TextMatrix(0, bteColWHCode) = "WH Code"
        .TextMatrix(0, bteColSeqNo) = "SeqNo"
        .TextMatrix(0, bteColUnitCls) = "Unit Cls"
        .TextMatrix(0, bteColReqCls) = "Request Cls"
        .TextMatrix(0, bteColGroupCls) = "Group Cls"
        .TextMatrix(0, bteColRequest) = "Request"
        .TextMatrix(0, bteColRequestNo) = "Request No"
        .TextMatrix(0, bteColControlCls) = "Control Cls"
        .TextMatrix(0, bteColSupplyCls) = "Supply Cls"
        .TextMatrix(0, bteColWominNo) = "Womin Number"
        
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColScheduleDate) = 1350
        .ColWidth(bteColProductCode) = 1500
        .ColWidth(bteColPartCode) = 1500
        .ColWidth(bteColDesc) = 3900
        .ColWidth(bteColLotNo) = 1000
        .ColWidth(bteColPlan) = 1300
        .ColWidth(bteColResult) = 1300
        .ColWidth(bteColRemaining) = 1300
        .ColWidth(bteColRequest) = 800
        .ColWidth(bteColRequestNo) = 1500
        .ColWidth(bteColWominNo) = 3000

        
        .ColHidden(bteColPackCode) = True
        .ColHidden(bteColPackDesc) = True
        .ColHidden(bteColPackSize) = True
        .ColHidden(bteColCustName) = True
        .ColHidden(bteColWHCode) = True 'WH Code
        .ColHidden(bteColSeqNo) = True 'Seq No
        .ColHidden(bteColUnitCls) = True 'Unit Cls
        .ColHidden(bteColReqCls) = True 'Request Cls
        .ColHidden(bteColGroupCls) = True 'Group Cls
        .ColHidden(bteColControlCls) = True 'Control Cls
        .ColHidden(bteColSupplyCls) = True 'Supply Cls
        .ColHidden(bteColRequestNo) = True 'RequestNo Change to WomIn
        
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColScheduleDate) = flexAlignCenterCenter
        .ColAlignment(bteColProductCode) = flexAlignLeftCenter
        .ColAlignment(bteColPartCode) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColLotNo) = flexAlignLeftCenter
        .ColAlignment(bteColPackCode) = flexAlignLeftCenter
        .ColAlignment(bteColPackDesc) = flexAlignLeftCenter
        .ColAlignment(bteColPackSize) = flexAlignRightCenter
        .ColAlignment(bteColCustName) = flexAlignLeftCenter
        .ColAlignment(bteColPlan) = flexAlignRightCenter
        .ColAlignment(bteColResult) = flexAlignRightCenter
        .ColAlignment(bteColRemaining) = flexAlignRightCenter
        .ColAlignment(bteColReqCls) = flexAlignRightCenter
        .ColAlignment(bteColRequest) = flexAlignCenterCenter
        .ColAlignment(bteColRequestNo) = flexAlignLeftCenter
        
        .EditMaxLength = 1
    End With

End Sub

Private Sub InputRequestMaterial(strItemCode As String, Optional strParentCode As String, Optional StrSeqNo As String, _
    Optional strLotNo As String, Optional strRemark As String, Optional dblQty As Double)
    
    Dim adoRs As New ADODB.Recordset
    Dim booRecurring As Boolean
    
    Dim intCount As Integer
    Dim booExist As Boolean
    
    On Error GoTo ErrHandler
        
    sql = "Select a.Parent_ItemCode, a.Unit_Cls, c.Production_Cls, a.Qty As QtyBOM, c.Item_Code, c.WH_Code, Material_Cls = Case When c.Material_Cls = '01' Then '1' Else '0' End, 'Formula' As Data, " & _
        "Item_Child = (Select Distinct Parent_ItemCode From BOM_Master Where Parent_ItemCode = a.Item_Code) " & _
        "From BOM_Master a " & _
        "Inner Join Item_Master b On a.Parent_ItemCode = b.Item_Code " & _
        "Inner Join Item_Master c On a.Item_Code = c.Item_Code " & _
        "Where b.Control_Cls = '01' And c.Control_Cls = '01' And c.SupplyIssue_Cls = '01' And a.Parent_ItemCode = '" & strItemCode & "' " & _
        "And a.Start_Date <= '" & Format(Date, "YYYYMMDD") & "' And a.End_Date >= '" & Format(Date, "YYYYMMDD") & "'"
    
    adoRs.Open sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
    While Not adoRs.EOF
        
        booRecurring = adoRs.Fields("Production_Cls") <> "01"
        If booRecurring Then booRecurring = Not IsNull(adoRs.Fields("Item_Child"))
        
        If booRecurring Then
            
            InputRequestMaterial Trim(adoRs.Fields("Item_Code")), strParentCode, StrSeqNo, strLotNo, strRemark, dblQty * adoRs.Fields("QtyBOM")
        
        Else
        
            For intCount = 1 To UBound(arrBOM, 2)
                booExist = (StrSeqNo = Trim(arrBOM(2, intCount)))
                If booExist Then booExist = (Trim(adoRs.Fields("Item_Code")) = Trim(arrBOM(5, intCount)))
                If booExist Then Exit For
            Next
            
            If Not booExist Then
                
                ReDim Preserve arrBOM(10, IntIndex) As String
                
                arrBOM(1, IntIndex) = strParentCode
                arrBOM(2, IntIndex) = StrSeqNo
                arrBOM(3, IntIndex) = strLotNo
                arrBOM(4, IntIndex) = strRemark
                arrBOM(5, IntIndex) = Trim(adoRs.Fields("Item_Code"))
                arrBOM(6, IntIndex) = adoRs.Fields("Unit_Cls")
                arrBOM(7, IntIndex) = dblQty * adoRs.Fields("QtyBOM")
                If gb_WarehouseReferToItem_MaterialSupplyRequestAutomatic Then
                    arrBOM(8, IntIndex) = Trim(adoRs.Fields("WH_Code"))
                Else
                    arrBOM(8, IntIndex) = Trim(cbo(2))
                End If
                arrBOM(9, IntIndex) = adoRs.Fields("Material_Cls")
                arrBOM(10, IntIndex) = adoRs.Fields("Data")
                
                IntIndex = IntIndex + 1
                
            Else
                
                arrBOM(7, intCount) = arrBOM(7, intCount) + (dblQty * adoRs.Fields("QtyBOM"))
                
            End If
            
        End If
        
        adoRs.MoveNext
    
    Wend
    adoRs.Close
    
ErrExit:
    Set adoRs = Nothing
    Exit Sub
ErrHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit

End Sub

'Private Sub SaveRequestMaterial()
'
'    Dim lrs_srMaster As New ADODB.Recordset
'    Dim lrs_srDetail As New ADODB.Recordset
'
'    Dim adoRs As New ADODB.Recordset
'    Dim intCount As Integer
'
'    On Error GoTo ErrHandler
'
'    '#Open Rs Part Supply Request Master
'    If lrs_srMaster.State <> adStateClosed Then lrs_srMaster.Close
'    lrs_srMaster.Open "select * from partSupplyRequest_Master ", Db, adOpenKeyset, adLockOptimistic
'
'    '#Open Rs Part Supply Request Detail
'    If lrs_srDetail.State <> adStateClosed Then lrs_srDetail.Close
'    lrs_srDetail.Open "select * from partSupplyRequest_Detail ", Db, adOpenKeyset, adLockOptimistic
'
'    '#Looping Insert To DataBase
'    Dim ll_counter As Long
'    Dim ls_supReqNo As String
'
'    ls_supReqNo = ""
'
'    For intCount = 1 To UBound(arrBOM, 2)
'
'        Sql = "Select Distinct a.SupplyRec_No From PartSupplyRequest_Master a " & _
'            "Inner Join PartSupplyRequest_Detail b On a.SupplyRec_No = b.SupplyRec_No " & _
'            "Where DailySeq_No In (" & ls_dailySeqNo & ") And FromWarehouse_Code = '" & arrBOM(8, intCount) & " '"
'
'        adoRs.Open Sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
'        If adoRs.EOF Then
'
'            '#genSupNo
'            ls_supReqNo = uf_GenerateSupplyRequestNo(Format(Date, "MM"), Format(Date, "yyyy"))
'            ls_supReqNoAcc = ls_supReqNoAcc + IIf(Trim(ls_supReqNoAcc) <> "", ";", "") + ls_supReqNo
'
'            '#Insert Data To Supply Request Master
'            lrs_srMaster.AddNew
'            lrs_srMaster!supplyRec_No = Trim(ls_supReqNo)
'            lrs_srMaster!fromwarehouse_code = arrBOM(8, intCount)
'            lrs_srMaster!Machine_no = Trim(cbo(1))
'            lrs_srMaster!towarehouse_code = Trim(cbo(0))
'            lrs_srMaster!childsupply_date = Format(Date, "yyyy-MM-dd")
'            lrs_srMaster!supply_cls = "S1"
'            lrs_srMaster!Auto_Cls = "1"
'            lrs_srMaster!Request_Cls = Trim(ls_group)
'            lrs_srMaster.update
'
'        Else
'
'            ls_supReqNo = adoRs.Fields("SupplyRec_No")
'
'        End If
'        adoRs.Close
'
'        '#Init Counter
'        Sql = "Select IsNull(Max(Seq_No), 0) + 1 As NewNumber From PartSupplyRequest_Detail Where SupplyRec_No = '" & Trim(ls_supReqNo) & "'"
'        adoRs.Open Sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
'        ll_counter = Val(adoRs.Fields("NewNumber") & "")
'        adoRs.Close
'
'        '#Insert Data To Supply Request Detail
'        lrs_srDetail.AddNew
'        lrs_srDetail!supplyRec_No = Trim(ls_supReqNo)
'        lrs_srDetail!Seq_No = Trim(ll_counter)
'        lrs_srDetail!childitem_code = arrBOM(5, intCount)
'        lrs_srDetail!ChildLot_no = arrBOM(3, intCount)
'        lrs_srDetail!ChildRequirement_qty = IIf(arrBOM(10, intCount) = "Formula", Trim(Format(arrBOM(7, intCount), gs_formatQtyBOM)), Trim(Format(0, gs_formatQtyBOM)))
'        lrs_srDetail!childunit_cls = arrBOM(6, intCount)
'        lrs_srDetail!parentItem_code = arrBOM(1, intCount)
'        lrs_srDetail!last_update = Now
'        lrs_srDetail!last_user = userLogin
'        lrs_srDetail!Remarks = arrBOM(4, intCount)
'        lrs_srDetail!dailyseq_no = arrBOM(2, intCount)
'        lrs_srDetail.update
'
'    Next
'
'ErrExit:
'    Set lrs_srMaster = Nothing
'    Set lrs_srDetail = Nothing
'    Set adoRs = Nothing
'    Exit Sub
'ErrHandler:
'    LblErrMsg.Caption = "[" & Err.number & "] " & Err.Description
'    Err.clear
'    Resume ErrExit
'
'End Sub

Private Sub SaveRequestMaterial_New(parent As String, seqNo As String, lotno As String)
    
    Dim adoCon As New ADODB.Connection
    Dim adoCmd As New ADODB.Command
    Dim rsRequest As New ADODB.Recordset
    
    Dim lngCount As Integer
    Dim ls_supReqNo As String
    Dim strSQL As String
    
       
    On Error GoTo ErrHandler

    
    'Open New Connection
    adoCon.ConnectionString = Db.ConnectionString
    adoCon.CursorLocation = adUseClient
    adoCon.Open
    
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = "PartSupplyRequest_Save"
       
    'Perubahan untuk pengambilan Request, bukan dari Array, tapi dari Temp Table
    
    strSQL = " select Parent_ItemCode,Seqno,LotNo,Remarks,Item_Code,Unit_Cls,WH_Code, " & vbCrLf & _
                      "     Sum(QtyBom) QtyBOM From TempRequest Where ItemType='Child' " & vbCrLf & _
                      "     And Parent_ItemCode='" & Trim(parent) & "' " & _
                                             " And SeqNo='" & Trim(seqNo) & "' " & _
                                             " And LotNo='" & Trim(lotno) & "' " & _
                      "     Group By Parent_ItemCode,Seqno,LotNo,Remarks,Item_Code,Unit_Cls,WH_Code "
                            
                          
    rsRequest.Open strSQL, adoCon, adOpenDynamic, adLockReadOnly
    
    Do While Not rsRequest.EOF
        
        adoCmd.Parameters(1) = ls_dailySeqNo  '@DailySeq_No_List
        adoCmd.Parameters(2) = Trim(rsRequest("WH_Code")) '@FromWarehouse_Code
        adoCmd.Parameters(3) = Trim(cbo(1)) '@Machine_No
        adoCmd.Parameters(4) = Trim(cbo(0)) '@ToWarehouse_Code
        adoCmd.Parameters(5) = "S1" '@Supply_Cls
        adoCmd.Parameters(6) = "1" '@Auto_Cls
        adoCmd.Parameters(7) = Trim(ls_group) '@Request_Cls
        adoCmd.Parameters(8) = Trim(rsRequest("Item_Code")) '@ChildItem_Code
        adoCmd.Parameters(9) = Trim(rsRequest("LotNo")) '@ChildLot_No
        adoCmd.Parameters(10) = Format(rsRequest("QtyBOM"), gs_formatQtyBOM)  '@ChildRequirement_Qty
        adoCmd.Parameters(11) = Trim(rsRequest("Unit_Cls")) '@ChildUnit_Cls
        adoCmd.Parameters(12) = Trim(rsRequest("Parent_ItemCode")) '@ParentItem_Code
        adoCmd.Parameters(13) = Trim(rsRequest("Remarks")) '@Remarks2
        adoCmd.Parameters(14) = Trim(rsRequest("SeqNo")) '@DailySeq_No
        adoCmd.Parameters(15) = userLogin  '@UserLogin
        adoCmd.Parameters(16) = "" '@NewSupplyRec_No
        
        adoCmd.Execute
        While adoCmd.State = adStateExecuting
        Wend
        
        If Trim(adoCmd.Parameters(16)) <> "" Then ls_supReqNoAcc = ls_supReqNoAcc + IIf(Trim(ls_supReqNoAcc) <> "", ";", "") + Trim(adoCmd.Parameters(16))
        
        rsRequest.MoveNext
        
    Loop
                      
'    For lngCount = 1 To UBound(arrBOM, 2)
'
'        adoCmd.Parameters(1) = ls_dailySeqNo  '@DailySeq_No_List
'        adoCmd.Parameters(2) = arrBOM(8, lngCount) '@FromWarehouse_Code
'        adoCmd.Parameters(3) = Trim(cbo(1)) '@Machine_No
'        adoCmd.Parameters(4) = Trim(cbo(0)) '@ToWarehouse_Code
'        adoCmd.Parameters(5) = "S1" '@Supply_Cls
'        adoCmd.Parameters(6) = "1" '@Auto_Cls
'        adoCmd.Parameters(7) = Trim(ls_group) '@Request_Cls
'        adoCmd.Parameters(8) = arrBOM(5, lngCount) '@ChildItem_Code
'        adoCmd.Parameters(9) = arrBOM(3, lngCount) '@ChildLot_No
'        adoCmd.Parameters(10) = IIf(arrBOM(10, lngCount) = "Formula", Trim(Format(arrBOM(7, lngCount), gs_formatQtyBOM)), Trim(Format(0, gs_formatQtyBOM)))    '@ChildRequirement_Qty
'        adoCmd.Parameters(11) = arrBOM(6, lngCount) '@ChildUnit_Cls
'        adoCmd.Parameters(12) = arrBOM(1, lngCount) '@ParentItem_Code
'        adoCmd.Parameters(13) = arrBOM(4, lngCount) '@Remarks2
'        adoCmd.Parameters(14) = arrBOM(2, lngCount) '@DailySeq_No
'        adoCmd.Parameters(15) = userLogin  '@UserLogin
'        adoCmd.Parameters(16) = "" '@NewSupplyRec_No
'
'        adoCmd.Execute
'        While adoCmd.State = adStateExecuting
'        Wend
'
'        If Trim(adoCmd.Parameters(16)) <> "" Then ls_supReqNoAcc = ls_supReqNoAcc + IIf(Trim(ls_supReqNoAcc) <> "", ";", "") + Trim(adoCmd.Parameters(16))
'
'    Next
    
    adoCon.Close
    
'            '#genSupNo
'            ls_supReqNo = uf_GenerateSupplyRequestNo(Format(Date, "MM"), Format(Date, "yyyy"))
'            ls_supReqNoAcc = ls_supReqNoAcc + IIf(Trim(ls_supReqNoAcc) <> "", ";", "") + ls_supReqNo
    
ErrExit:
    Set adoCmd = Nothing
    Set adoCon = Nothing
    Exit Sub
ErrHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
    
End Sub


Sub Kosong()
    nilKosong = True
    cbo(0) = ""
    cbo(1) = ""
    cbo(2) = ""
    CboItemCode = "ALL"
    lblNm(0) = ""
    lblNm(1) = ""
    lblNm(2) = ""
    lblItemCode.Caption = "ALL"
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


    RsCust.Close
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


    rscbo.Close
End With
End Sub

Private Sub SetComboWH()
    
    Dim adoRs As New ADODB.Recordset
    
    sql = "select wh_code, wh_name from warehouse_master " & _
        "where stockcontrol_cls='01' and Adm_Group = '" & Trim(cbo(0)) & "'"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    cbo(2).columnCount = 2
    cbo(2).clear
    i = 0
    Do While Not adoRs.EOF
        cbo(2).AddItem ""
        cbo(2).List(i, 0) = Trim(adoRs("wh_code"))
        cbo(2).List(i, 1) = Trim(adoRs("wh_name"))
        i = i + 1
        adoRs.MoveNext
    Loop
    cbo(2).ColumnWidths = "50 pt; 300 pt"
    cbo(2).ListWidth = 350
    cbo(2).ListRows = 15
    If cbo(2).ListCount > 0 Then
        cbo(2).ListIndex = 0
    End If
    cbo(2).locked = True
    
End Sub

Sub addToCboItemCode()

    Dim sqlLine As String
    Dim RsLine As New Recordset
    Dim i As Long
    
    Me.MousePointer = vbHourglass
    
    sqlLine = "SELECT DISTINCT DP.item_code, makeritem_code, item_name FROM  " & _
        " Daily_Production DP  " & _
        " LEFT JOIN Item_Master IM ON DP.Item_code = IM.Item_Code " & _
        " WHERE production_cls = 01 and use_endday > convert(char(8), getdate(), 112) and FinishGoodPart_Cls='01' "
            
    Set RsLine = Db.Execute(sqlLine)
        
    With CboItemCode
        .clear
        .columnCount = 3
        .ColumnWidths = "80pt;80pt;165pt"
        .ListWidth = 325
        .ListRows = 15
        .AddItem
        .List(0, 0) = strAll
        .List(0, 1) = strAll
        .List(0, 2) = strAll
        i = 1
        Do While Not RsLine.EOF
            .AddItem
            .List(i, 0) = Trim(RsLine("item_code"))
            .List(i, 1) = Trim(RsLine("makeritem_code"))
            .List(i, 2) = Trim(RsLine("item_Name"))
            RsLine.MoveNext
            i = i + 1
        Loop
        .ListIndex = 0
    End With
    
     Me.MousePointer = vbDefault
End Sub

Private Sub CboItemCode_Change()
    LblErrMsg = ""
    If CboItemCode.MatchFound Then
        If CboItemCode.Column(0) = CboItemCode.Column(1) Then
            lblItemCode.Caption = CboItemCode.Column(2)
        Else
            lblItemCode.Caption = CboItemCode.Column(1) & " " & CboItemCode.Column(2)
        End If
    Else
        lblItemCode.Caption = ""
    End If
End Sub

Private Sub cboRemaining_Change()
Call headerGrid
End Sub

Private Sub cboRemaining_Click()
Call headerGrid
End Sub

Private Sub dtAkhir_Change()
Call headerGrid
End Sub

Private Sub dtAwal_Change()
Call headerGrid
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    fromProd = True
    nilKosong = True
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode("frm_part_supplyAuto") & ")"

    Call isiCboCust
    Call Kosong
    nilKosong = False
    Call initGroup
'    addToCboItemCode
    
    Label1(5).Visible = Not gb_WarehouseReferToItem_MaterialSupplyRequestAutomatic
    cbo(2).Visible = Not gb_WarehouseReferToItem_MaterialSupplyRequestAutomatic
    lblNm(2).Visible = Not gb_WarehouseReferToItem_MaterialSupplyRequestAutomatic
    Line8(2).Visible = Not gb_WarehouseReferToItem_MaterialSupplyRequestAutomatic
    
End Sub

Public Sub initGroup()

With Me
 '#Reset Group & request Data
    'If .il_selectedRecord = 0 Then
        .is_request = ""
        .is_groupRequest = ""
         .il_selectedRecord = 0
    'End If

For i = 0 To .grid.Rows - 1

    If .grid.Cell(flexcpChecked, i, 0) = flexChecked Then

        '#Init Grid For The FirstTime
        If Trim(.is_request) = "" Then .is_request = Trim(.grid.TextMatrix(i, bteColReqCls))
        If Trim(.is_groupRequest) = "" Then .is_groupRequest = Trim(.grid.TextMatrix(i, bteColGroupCls))

        '#Init Selected rows
        .il_selectedRecord = .il_selectedRecord + 1

    End If

Next
End With
End Sub

Private Sub cbo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
If nilKosong = True Then Exit Sub
    If KeyCode = 13 Then Call cbo_Click(Index)
End Sub

Private Sub cbo_Change(Index As Integer)

    Call headerGrid
    If nilKosong = True Then Exit Sub
    lblNm(Index) = ""
    'Hapus Manufacture Line * Desc
    If Index = 0 Then cbo(1).clear: lblNm(1) = "": Call headerGrid
'Me.MousePointer = vbDefault
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
            Call SetComboWH
            addToCboItemCode
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
'On Error Resume Next
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

Public Sub IsiGrid()

Dim rsProd As New ADODB.Recordset
Dim sqlResult As String, UWom As String

If nilKosong = True Then Exit Sub
With grid
    Call headerGrid

'    Sql = "select isnull(a.Request_cls,0)Request_cls,Seq_No,Schedule_Date,a.Item_Code, im.Makeritem_Code, " & _
'            uf_GetQueryDescription("ItemDesc", "FinishGoodPart_Cls") & " As Descr,Lot_No," & _
'            " qty "
    
    sql = "select isnull(a.Request_cls,0)Request_cls,Seq_No,Schedule_Date,a.Item_Code, im.Makeritem_Code, " & _
            "im.item_name As Descr,Lot_No," & _
            " qty "


    sqlResult = "(select Isnull(Sum(Qty),0) from " & _
            "Part_Receipt where DailySeq_No = a.Seq_No " & _
            "And Receipt_Cls = 'P1') "

    sql = sql & "," & sqlResult & " as Result " & _
        ",(Qty - " & sqlResult & ") as Sisa,WH_Code, im.Unit_Cls, im.Control_Cls, im.Suply_Cls, " & _
        "isnull((select max(supplyrec_no) from partsupplyrequest_master where request_cls = a.request_cls), '') supplyrec_no, " & _
        "isnull((select max(womin_no) from partsupplyrequest_master where request_cls = a.request_cls), '') womin_no " & _
        "from Daily_Production a,Item_Master im " & _
        "Where a.Item_Code = im.Item_Code " & _
        "And Schedule_Date >= '" & Format(dtAwal, "yyyy-MM-dd") & _
        "' and Schedule_Date <= '" & Format(dtAkhir, "yyyy-MM-dd") & _
        "' and Factory_Code = '" & cbo(0) & _
        "' And a.line_Code = '" & cbo(1) & "' "

    If CboItemCode <> strAll Then
        sql = sql & " and a.item_code  = '" & CboItemCode & "'"
    End If

    If cboRemaining = "Yes" Then
        sql = sql & " And Qty - " & sqlResult & " > 0 "
    Else
        sql = sql & " And Qty - " & sqlResult & " <= 0 "
    End If

    sql = sql & "order by Schedule_Date,a.Item_Code,Lot_no "
    Set rsProd = Db.Execute(sql)
    rsProd.Requery

    '#Init Selected Grid Row
    il_selectedRecord = 0

    i = 1
    If Not (rsProd.EOF) Then
        Do While Not rsProd.EOF
            .Rows = .Rows + 1
            .Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
            .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
            .TextMatrix(i, bteColScheduleDate) = Format(Trim(rsProd("Schedule_Date")), "dd MMM yyyy")
            .TextMatrix(i, bteColProductCode) = Trim(rsProd("Item_Code"))
            .TextMatrix(i, bteColPartCode) = Trim(rsProd("MakerItem_Code"))
            .TextMatrix(i, bteColDesc) = Trim(rsProd("Descr"))
            .TextMatrix(i, bteColLotNo) = Trim(rsProd("Lot_No"))
            '.TextMatrix(i, bteColPackCode) = Trim(rsProd("Bag_code"))
            '.TextMatrix(i, bteColPackDesc) = Trim(rsProd("bag_name"))
            '.TextMatrix(i, bteColPackSize) = Trim(rsProd("PackingSize"))
            .TextMatrix(i, bteColCustName) = "" 'IIf(IsNull(rsProd("Cust_Code")), "COMMON", Trim(rsProd("Trade_Name")))
            .TextMatrix(i, bteColPlan) = Format(rsProd("Qty"), gs_formatQty)
            .TextMatrix(i, bteColResult) = Format(rsProd("Result"), gs_formatQty)
            .TextMatrix(i, bteColRemaining) = Format(rsProd("Sisa"), gs_formatQty)
            .TextMatrix(i, bteColWHCode) = Trim(rsProd("WH_Code"))
            .TextMatrix(i, bteColSeqNo) = rsProd("Seq_No")
            .TextMatrix(i, bteColUnitCls) = rsProd("Unit_Cls")
            .TextMatrix(i, bteColReqCls) = IIf(rsProd("Request_Cls") <> "0", "1", "0")
            .TextMatrix(i, bteColGroupCls) = Trim(rsProd("Request_Cls"))
            .TextMatrix(i, bteColRequest) = IIf(rsProd("Request_Cls") <> "0", "Yes", "No")
            .TextMatrix(i, bteColRequestNo) = Trim(rsProd("supplyrec_no"))
            .TextMatrix(i, bteColControlCls) = Trim(rsProd("Control_Cls"))
            .TextMatrix(i, bteColSupplyCls) = Trim(rsProd("Suply_Cls"))
            ' Penambahan womin No
            If Trim(rsProd("Womin_No")) <> "" Then
                .TextMatrix(i, bteColWominNo) = Trim(rsProd("Womin_No"))
            Else
                If Trim(rsProd("supplyrec_no")) <> "" Then
                    '.TextMatrix(i, bteColWominNo) = Trim(rsProd("supplyrec_no")) & " - " & Trim(rsProd("Descr"))
                    .TextMatrix(i, bteColWominNo) = Trim(rsProd("Descr")) & "/" & Right(Trim(rsProd("supplyrec_no")), 4) & Mid(Trim(rsProd("supplyrec_no")), 7, 2) & Left(Trim(rsProd("supplyrec_no")), 5)
                    
                    UWom = " Update PartSupplyRequest_Master Set Womin_No='" & Trim(rsProd("Descr")) & "/" & Right(Trim(rsProd("supplyrec_no")), 4) & Mid(Trim(rsProd("supplyrec_no")), 7, 2) & Left(Trim(rsProd("supplyrec_no")), 5) & "' " & vbCrLf & _
                                " Where SupplyRec_No='" & Trim(rsProd("supplyrec_no")) & "' "
                     Db.Execute (UWom)
                Else
                    .TextMatrix(i, bteColWominNo) = ""
                End If
            End If
            
            i = i + 1
            rsProd.MoveNext
        Loop
    Else
        LblErrMsg = DisplayMsg(4006)
    End If
    rsProd.Close
End With
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim InRowPos As Integer, InRow As Integer

With grid
    If Row <> 0 And Col = bteColSelect Then
        If .Cell(flexcpChecked, Row, bteColSelect) = flexChecked Then
            InRowPos = Row
            '#Reset Group & request Data
            If il_selectedRecord = 0 Then
                is_request = ""
                is_groupRequest = ""
            End If

            '#Init Grid For The FirstTime
            If Trim(is_request) = "" Then is_request = Trim(.TextMatrix(Row, bteColReqCls))
            If Trim(is_groupRequest) = "" Then is_groupRequest = Trim(.TextMatrix(Row, bteColGroupCls))

            If Trim(is_request) = Trim(.TextMatrix(Row, bteColReqCls)) And _
                Trim(is_groupRequest) = Trim(.TextMatrix(Row, bteColGroupCls)) Then

                '#Init Selected rows
                il_selectedRecord = il_selectedRecord + 1

                .Cell(flexcpChecked, Row, bteColSelect) = flexChecked
                LblErrMsg = ""
                InRow = 1
                Do
                    If InRow > .Rows - 1 Then Exit Do
                    .Cell(flexcpChecked, InRow, bteColSelect) = flexUnchecked
                    InRow = InRow + 1
                Loop
                .Cell(flexcpChecked, InRowPos, bteColSelect) = flexChecked
                
            Else

                .Cell(flexcpChecked, Row, bteColSelect) = flexUnchecked

                LblErrMsg = DisplayMsg(8090) ' "Record has different Group Request!"

            End If
            
        Else
            '#Init Selected rows
            il_selectedRecord = il_selectedRecord - 1

        End If
    End If
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    LblErrMsg = ""
    If Col <> bteColSelect Then Cancel = 1
    If grid.TextMatrix(Row, bteColControlCls) <> "01" Then
        LblErrMsg = DisplayMsg(8091) ' "This item's Control Cls is not 'MRP' !"
        Cancel = 1 'Cek Control Cls (Continue If Control_Cls = 'MRP')
    End If
    If grid.TextMatrix(Row, bteColSupplyCls) = "01" Then
        LblErrMsg = "You cannot select this item ! Supply Cls must be set to 'NO' ": Cancel = 1  'Cek Supply Cls (Continue If Supply_Cls = 'No')"
        Cancel = 1
    End If

End Sub

Private Sub Command1_Click(Index As Integer) '#Submit
    
    Dim adoRs As New ADODB.Recordset
    
    Dim ll_cek As Long
    Dim ls_status As String
    Dim blncex As Boolean
    
    Dim ls_requestCls As String
    Dim ls_sqlGroup As String
    Dim ls_update As String
    Dim ls_sql As String
    
    Dim ls_SNR() As String
    Dim Dbl_Qty As Double
    
    Dim StrParent As String
    Dim StrSeqNo As String
    Dim strLotNo As String
    
    On Error GoTo ErrHandler
    Me.MousePointer = vbHourglass
    
    With grid
        
        ls_dailySeqNo = ""
        ll_cek = 0
        blncex = False
        ls_group = ""
        
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
                blncex = True
                ls_dailySeqNo = ls_dailySeqNo + Trim(grid.TextMatrix(i, bteColSeqNo)) + ","
                
                ls_FG = .TextMatrix(i, bteColDesc)
                
                If .TextMatrix(i, bteColReqCls) = "0" Or .TextMatrix(i, bteColRequestNo) = "" Then
                    .TextMatrix(i, bteColReqCls) = "1"
                    .TextMatrix(i, bteColRequest) = "Yes"
                    ls_status = "insert"
                Else
                    ls_group = Trim(.TextMatrix(i, bteColGroupCls))
                    ls_status = "update"
                End If
        
                ll_cek = ll_cek + 1
        
            End If
        Next i
        
    End With

    If Not blncex Then
        LblErrMsg = DisplayMsg(8011)  '"Please select data first!"
        MousePointer = vbDefault: Exit Sub
    End If

    '#Init Selected Daily SeqNo
    If Len(Trim(ls_dailySeqNo)) > 0 Then ls_dailySeqNo = Left(ls_dailySeqNo, Len(Trim(ls_dailySeqNo)) - 1)

    '#Check Selected Data in Grid
    If ll_cek = 0 Then

        '#There is no Selected Data in Grid
        Me.MousePointer = vbDefault
        Me.Hide
        frm_part_supplyAuto.Show
        frm_part_supplyAuto.cmd_sub_menu.Caption = "&Back"
        LblErrMsg = ""
        Exit Sub

    Else

        '#Check Empty Header Combo
        If cbo(0) = "" Then
            LblErrMsg = DisplayMsg(1040)
            cbo(0).SetFocus
            Me.MousePointer = vbDefault
            frm_part_supplyAuto.Show
            frm_part_supplyAuto.cmd_sub_menu.Caption = "&Back"
            Exit Sub
        End If

        '#Check Valid Header Combo
        cbo(0) = cbo(0)
        cbo(1) = cbo(1)
        If cbo(0).MatchFound = False Then
            LblErrMsg = DisplayMsg(4016)
            cbo(0).SetFocus
            Me.MousePointer = vbDefault
            Exit Sub
        ElseIf cbo(1) <> "" And cbo(1).MatchFound = False Then
            LblErrMsg = DisplayMsg(4017)
            cbo(1).SetFocus
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        If Not gb_WarehouseReferToItem_MaterialSupplyRequestAutomatic Then
            If cbo(2).MatchFound = False Then
                LblErrMsg = DisplayMsg("0031")
                cbo(2).SetFocus
                Me.MousePointer = vbDefault
                Exit Sub
            End If
        End If

    End If

    If hakAkses("frm_part_supplyAuto") = 0 Then
        Me.MousePointer = vbDefault
        LblErrMsg = DisplayMsg(3007)
        Me.MousePointer = vbDefault
        Exit Sub
    End If

    If hakUpdate("frm_part_supplyAuto") = 0 Then
        Me.MousePointer = vbDefault
        LblErrMsg = DisplayMsg(3007)
        Exit Sub
    End If

    ls_supReqNoAcc = ""

    If ls_status = "insert" Then
        
        'Update Status di pindah kan di akhir proses 20210603
        '#Update Status Request_Cls Daily
        sql = "Select Max(Isnull(Request_Cls,0)) + 1 As Request_Cls From Daily_Production"
        adoRs.Open sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
        If Not adoRs.EOF Then ls_requestCls = Trim(adoRs!Request_Cls) Else ls_requestCls = "1"
        adoRs.Close
        ' Error in trigger if choose more than 1 record
        ls_sql = "Update Daily_Production Set Request_Cls = '" & Trim(ls_requestCls) & "', Last_Update = GETDATE(), Last_User = '" & userLogin & "' Where Seq_No In (" + ls_dailySeqNo + ")"
        Db.Execute ls_sql
                
'        For i = 1 To grid.Rows - 1
'            If grid.Cell(flexcpChecked, i, bteColselect) = flexChecked Then
'                ls_sql = "Update Daily_Production Set Request_Cls = '" & Trim(ls_requestCls) & "' Where Seq_No =" & Trim(grid.TextMatrix(i, bteColSeqNo))
'                Db.Execute ls_sql
'            End If
'        Next i
        
        '#Init Group request
        ls_sqlGroup = "Select Distinct Request_Cls  From Daily_Production Where Seq_no In (" + ls_dailySeqNo + ")"
        adoRs.Open ls_sqlGroup, Db, adOpenDynamic, adLockReadOnly, adCmdText
        ls_group = Trim(adoRs!Request_Cls)
        adoRs.Close
        
        '#Create Material Request Data
        IntIndex = 1
        'ReDim arrBOM(10, intIndex) As String
        Call CreateTableTemp ' Mengubah Array menjadi Temporary Table
        
        sql = "Select a.Item_code, a.Seq_No, Lot_No, Remark, Qty " & _
            "From Daily_Production a " & _
            "Inner Join Item_Master b On a.Item_Code = b.Item_Code " & _
            "Where b.Control_Cls = '01' And a.Seq_No In (" & ls_dailySeqNo & ")"
    
        adoRs.Open sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
        While Not adoRs.EOF
        
            StrParent = Trim(adoRs.Fields("Item_code"))
            StrSeqNo = Trim(adoRs.Fields("Seq_No"))
            strLotNo = Trim(adoRs.Fields("Lot_No"))
            
'            InputRequestMaterial Trim(adoRs.Fields("Item_code")), Trim(adoRs.Fields("Item_Code")), Trim(adoRs.Fields("Seq_No")), _
'                Trim(adoRs.Fields("Lot_No")), Trim(adoRs.Fields("Remark")), adoRs.Fields("Qty")

            ' Perubahan metode pencarian Item BOM - 20090518
            sql = " Delete From TempRequest Where Parent_ItemCode='" & Trim(adoRs.Fields("Item_code")) & "' " & _
                    " And SeqNo='" & Trim(adoRs.Fields("Seq_No")) & "' " & _
                    " And LotNo='" & Trim(adoRs.Fields("Lot_No")) & "' "
            Db.Execute (sql)
            
            Call GetConsumtionData(Trim(adoRs.Fields("Item_code")), Trim(adoRs.Fields("Item_code")), adoRs.Fields("Qty"), Trim(adoRs.Fields("Seq_No")), _
                Trim(adoRs.Fields("Lot_No")), Trim(adoRs.Fields("Remark")))
            adoRs.MoveNext
        Wend
        adoRs.Close
        
        ' Cek data pada Temporary Request
        
            sql = " Select * From TempRequest Where Parent_ItemCode='" & StrParent & "' " & _
                    " And SeqNo='" & StrSeqNo & "' " & _
                    " And LotNo='" & strLotNo & "' "
        
        adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly
        
        If adoRs.EOF Then
            LblErrMsg = DisplayMsg(8092) '"Data BOM for this item is not valid"
            Me.MousePointer = vbDefault: Exit Sub
        End If
        
'        If arrBOM(1, 1) = "" Then
'            LblErrMsg = DisplayMsg(8092) '"Data BOM for this item is not valid"
'            Me.MousePointer = vbDefault: Exit Sub
'        End If
        
        SaveRequestMaterial_New StrParent, StrSeqNo, strLotNo
        
        For i = 1 To grid.Rows - 1
            If grid.Cell(flexcpChecked, i, bteColSelect) = flexChecked Then grid.TextMatrix(i, bteColGroupCls) = Trim(ls_group)
        Next
        
    
    ElseIf ls_status = "update" Then
    
        ls_update = "Select Distinct Isnull(SupplyRec_No, '') SupplyRec_No " & _
                "From PartSupplyRequest_Master prm " & _
                "Right Join (" & _
                "Select Request_Cls From Daily_Production " & _
                "Where Request_Cls = '" + Trim(ls_group) + "') dp On prm.Request_Cls = dp.Request_Cls"

        adoRs.Open ls_update, Db, adOpenDynamic, adLockReadOnly, adCmdText
        While adoRs.EOF = False
            ls_supReqNoAcc = ls_supReqNoAcc + IIf(Trim(ls_supReqNoAcc) <> "", ";", "") + Trim(adoRs!supplyRec_No)
            adoRs.MoveNext
        Wend
        adoRs.Close
        
    End If
            
    Load frm_part_supplyAuto
    With frm_part_supplyAuto
        
        .cbo_supplyNo(0).clear
        ls_SNR = Split(ls_supReqNoAcc, ";")
        
        
            .cbo_supplyNo(0).AddItem ls_SNR(0)
            .cbo_supplyNo(0) = ls_SNR(0)
            .LblWomIn = Trim(ls_FG) & "/" & Right(Trim(ls_SNR(0)), 4) & Mid(Trim(ls_SNR(0)), 7, 2) & Left(Trim(ls_SNR(0)), 5)

        
        'For i = 0 To UBound(ls_SNR)
'            .cbo_supplyNo(0).AddItem grid.TextMatrix(grid.Row, bteColRequestNo)
'            .cbo_supplyNo(0) = grid.TextMatrix(grid.Row, bteColRequestNo)
       ' Next
        .txtTemp = tempRowBom
        .is_status = ls_status
        .is_group = Trim(ls_group)
        .cbo_location(0) = Trim(cbo(0))
        .cbo_MachineNo(0) = Trim(cbo(1))
        
        If cmdsubmenu.Caption <> "&Back" Then .cmd_sub_menu.Caption = "&Back"
        .ib_fromProd = True
        
        Call .cmd_update_Click(0)
            
        .is_groupRequest = is_groupRequest
        .il_selectedRecord = il_selectedRecord
        .navigateButton (ls_status = "update")
        .is_DailySeqNo = ls_dailySeqNo
        .ib_fromProd = False
        .cmd_sub_menu.Enabled = True
        .parentItemCode = ls_str
        If ls_status <> "insert" Then .lbl_pesan = ""
        If .cbo_supplyNo(0).ListCount > 1 Then .lbl_pesan = DisplayMsg("0077") & " " & .cbo_supplyNo(0).ListCount
        
        .Show
        
    End With
    
    Me.Hide
    
ErrExit:
    Set adoRs = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
    
End Sub

'#Unload
Private Sub CmdSubMenu_Click()
    If cmdsubmenu.Caption = "&Back" Then
        Call Command1_Click(1)
    Else
        Unload frmProdResult
        Unload frm_part_supplyAuto
        frmMainMenu.Show
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


Private Sub CreateTableTemp()
On Error Resume Next

Dim SqlCreate As String

SqlCreate = " Create Table TempRequest (Parent_ItemCode Char(25), SeqNo Numeric(18),LotNo Char(15),Remarks Char(25), " & _
                 " Item_Code Char(25),Unit_Cls Char(2),QtyBOM Numeric(18,5), WH_Code Char(15), Material_cls Char(2), Data Char(15), ItemType Char(15)) " & vbCrLf
                          
Db.Execute (SqlCreate)

ErrExit:
    Exit Sub
ErrHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit

End Sub
'Consumption harus ada improvement 20250507
'Pada saat create planning seharusnya langsung menyimpan ke table bom master production, seqno dailyproduction harus diambil sebagai key di bom master production

Private Sub GetConsumtionData(Utama As String, ItemParent As String, qtyParent As Double, Optional StrSeqNo As String, _
    Optional strLotNo As String, Optional strRemark As String)

Dim SqlSearch As String, Strtype As String, SqlChild As String
Dim RsSearch As New ADODB.Recordset, rsChild As New ADODB.Recordset
Dim Warehouse As String

On Error GoTo ErrHandler

'SqlSearch = " select Parent_ItemCode, BM.Item_Code, isnull(Qty,0) QtyBOM, Production_Cls, StockControl_Cls, BM.Unit_Cls, IM.WH_Code," & vbCrLf & _
'                          "         Material_Cls = Case When IM.Material_Cls = '01' Then '1' Else '0' End, 'Formula' As Data  " & vbCrLf & _
'                          "     From BOM_MASTER BM  " & vbCrLf & _
'                          "     Inner Join Item_Master IM on BM.Item_Code=IM.Item_Code " & vbCrLf & _
'                          "         Where Parent_ItemCode='" & ItemParent & "' " & vbCrLf & _
'                          "   and Start_Date <='" & Format(Date, "YYYYMMDD") & "'" & _
'                          "   and End_Date >= '" & Format(Date, "YYYYMMDD") & "'"

SqlSearch = "EXEC dbo.sp_SupplyAutoByBom @Seqno = '" & StrSeqNo & "'," & vbCrLf & _
                          "@ParentItemCode = '" & ItemParent & "'," & vbCrLf & _
                          "@StartDate = '" & Format(Date, "YYYYMMDD") & "', -- varchar(8)" & vbCrLf & _
                          "  @EndDate = '" & Format(Date, "YYYYMMDD") & "' -- varchar(8)"
                          
'RsSearch.Open SqlSearch, Db, adOpenForwardOnly, adLockReadOnly
If RsSearch.State <> adStateClosed Then RsSearch.Close
RsSearch.Open SqlSearch, Db, adOpenKeyset, adLockOptimistic
'RsSearch.Open SqlSearch, Db, adOpenStatic, adLockReadOnly

'disini di tampung RecordCount nya 20230118
tempRowBom = RsSearch("RowsData")

Do While Not RsSearch.EOF
    
    SqlChild = "Select * From BOM_Master Where Parent_ItemCode ='" & RsSearch("Item_Code") & "'"
    rsChild.Open SqlChild, Db, adOpenForwardOnly, adLockReadOnly
    
    If Not rsChild.EOF Then
        If RsSearch("StockControl_Cls") = "02" Then
            Strtype = "Parent"
            Call GetConsumtionData(Utama, RsSearch("Item_Code"), qtyParent * RsSearch("QtyBOM"), StrSeqNo, strLotNo, strRemark)
        Else
            Strtype = "Child"
        End If
    Else
                Strtype = "Child"
    End If
    
    rsChild.Close
    
    If gb_WarehouseReferToItem_MaterialSupplyRequestAutomatic Then
        Warehouse = Trim(RsSearch.Fields("WH_Code"))
    Else
        Warehouse = Trim(cbo(2))
    End If
    
    SqlSearch = "Insert Into TempRequest Values ('" & Utama & "'," & StrSeqNo & ",'" & strLotNo & "','" & strRemark & "'," & _
                    " '" & Trim(RsSearch("Item_Code")) & "','" & Trim(RsSearch("Unit_Cls")) & "'," & RsSearch("QtyBOM") * qtyParent & "," & vbCrLf & _
                    " '" & Warehouse & "','" & Trim(RsSearch("Material_Cls")) & "','" & Trim(RsSearch("Data")) & "','" & Strtype & "')"
    
    Db.Execute SqlSearch
    
    RsSearch.MoveNext
Loop

ErrExit:
    Set RsSearch = Nothing
    Set rsChild = Nothing
    Exit Sub
ErrHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit

End Sub


