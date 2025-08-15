VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDNReturn 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivery Note Return"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "frmDNReturn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Cancel"
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
      Left            =   12825
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10020
      Width           =   1125
   End
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clea&r"
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
      Left            =   11610
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10020
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   75
      TabIndex        =   11
      Top             =   9330
      Width           =   15105
      Begin VB.Label lblErrMsg 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   14865
      End
   End
   Begin VB.CommandButton cmdProcess 
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
      Left            =   14055
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10020
      Width           =   1125
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
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10020
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   7110
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   15105
      _cx             =   26644
      _cy             =   12541
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   22
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDNReturn.frx":0E42
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
      Begin MSComCtl2.DTPicker dtReturn 
         Height          =   315
         Left            =   390
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   141230083
         CurrentDate     =   37859
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1215
      Left            =   75
      TabIndex        =   13
      Top             =   660
      Width           =   15105
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Search"
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
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtStart 
         Height          =   330
         Left            =   1860
         TabIndex        =   1
         Top             =   630
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
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtEnd 
         Height          =   330
         Left            =   4350
         TabIndex        =   2
         Top             =   630
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
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin MSForms.ComboBox CboCust 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   180
         Width           =   1605
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2831;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboDO 
         Height          =   333
         Left            =   7410
         TabIndex        =   3
         Top             =   660
         Width           =   2430
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "4286;587"
         TextColumn      =   1
         ListRows        =   7
         cColumnInfo     =   1
         ShowDropButtonWhen=   2
         Value           =   "001/AAAAAA/07/2003"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN No."
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
         Index           =   3
         Left            =   6750
         TabIndex        =   19
         Top             =   705
         Width           =   600
      End
      Begin VB.Label Label1 
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
         Left            =   3870
         TabIndex        =   17
         Top             =   705
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN Date from"
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
         Left            =   270
         TabIndex        =   16
         Top             =   705
         Width           =   1185
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3540
         X2              =   9540
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
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
         Left            =   270
         TabIndex        =   15
         Top             =   285
         Width           =   1350
      End
      Begin VB.Label lblCust 
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
         Left            =   3540
         TabIndex        =   14
         Top             =   240
         Width           =   6015
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13290
      TabIndex        =   20
      Top             =   90
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Note Return"
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
      Left            =   5707
      TabIndex        =   10
      Top             =   180
      Width           =   3840
   End
End
Attribute VB_Name = "frmDNReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created BY Dudi S, November akhir dan desember Awal 2008

Option Explicit
Dim dbTrans As New ADODB.Connection
Dim ClsProc As New ClsProc, clsMRP As New clsMRP
Dim i As Long, HakU As Integer
Dim nilKosong As Boolean

Dim ColS As Integer
Dim ColDoNo As Integer, ColDODate As Integer
Dim ColItemCD As Integer, ColItemDesc As Integer, ColLotNo As Integer
Dim coldoqty As Integer, ColWH As Integer, ColReturnCls As Integer, ColUnit As Integer, ColReturn As Byte
Dim ColDORQty As Integer, ColSpoolQty As Integer
'Dim ColCurr As Integer, ColPrice As Integer, colService, colamount As Integer, colRemark As Byte
Dim ColHSeqNo As Integer, ColHDOSeqNo As Integer, ColCaseNo As Integer
Dim ColHCurr As Integer, ColHUnit As Integer, ColHItemCls As Integer
Dim ColHDtBef As Integer, ColHWHBef As Integer, ColHQtyBef As Integer, ColHQtyTemp As Integer, ColHSpoolQty As Integer
Dim ColHRowHeader As Integer, ColHTotChild As Integer, ColHUpdate As Integer, ColHChk As Integer
Dim ColHSpoolBef As Integer, ColHProdSeqNo As Integer, ColHProdResultDt As Integer

Dim KolDn As Byte, KolDate As Byte, KolItem As Byte, Koldesc As Byte
Dim KolLot As Byte, KolQty As Byte, KolRWh As Byte, KolRCls As Byte
Dim kolRQty As Byte, KolUnit As Byte, KolCurr As Byte, KolPrice As Byte
Dim kolService As Byte, kolAmount As Byte, KolRemarks As Byte
Dim KolPoNo As Byte: Dim koldoseqno As Byte: Dim KolSeqNo As Byte: Dim KolCust As Byte
Dim KolDnTmp As Byte: Dim KolReturnSeq
Dim KolWHCode As Byte 'Untuk menampung code WH
Dim kolReqSeqNO As Byte
Dim Str_User As String
Dim newCls As New clsMRP

Dim tempRow As Double, rowHeader As Double

'************************* Initial *************************
Sub Kosong(Optional stAwal As Byte)
nilKosong = True
    If stAwal = 1 Then
        cboCust = "": lblcust = ""
        DtStart = Format(Now, "dd MMM yyyy")
        DtEnd = Format(Now, "dd MMM yyyy")
        cboDO = ""

    End If
    tempRow = 0
    LblErrMsg = ""
nilKosong = False
End Sub

Sub SetCol()
grid.ColS = 24
grid.Rows = 2
KolDn = 1: KolDate = 2: KolItem = 3: Koldesc = 4
KolLot = 5: KolQty = 6: KolRWh = 8: KolRCls = 9
kolRQty = 7: KolUnit = 10: KolCurr = 11: KolPrice = 12
kolService = 13: kolAmount = 14: KolRemarks = 15
KolPoNo = 16: koldoseqno = 17: KolSeqNo = 18: KolCust = 19
KolDnTmp = 20: KolReturnSeq = 21: kolReqSeqNO = 22
KolWHCode = 23

With grid
.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True: .MergeRow(1) = True:
.TextMatrix(0, KolDn) = "DN No."
.TextMatrix(1, KolDn) = "Reff No."
 .MergeCol(0) = True
.Cell(flexcpText, 0, 0, 1, 0) = ".."
 .MergeCol(KolDate) = True
.Cell(flexcpText, 0, KolDate, 1, KolDate) = "DN Date"

.MergeCol(KolItem) = True
.Cell(flexcpText, 0, KolItem, 1, KolItem) = "Item Code"

.ColWidth(KolItem) = 2000
.ColWidth(Koldesc) = 2500
.ColWidth(KolRCls) = 1200
.ColWidth(KolWHCode) = 1000


.MergeCol(Koldesc) = True
.Cell(flexcpText, 0, Koldesc, 1, Koldesc) = "Description"

.MergeCol(KolLot) = True
.Cell(flexcpText, 0, KolLot, 1, KolLot) = "Lot No"

.MergeCol(KolQty) = True
.Cell(flexcpText, 0, KolQty, 1, KolQty) = "DN Qty"

.MergeCol(kolRQty) = True
.Cell(flexcpText, 0, kolRQty, 1, kolRQty) = "Return" & "Qty"


.MergeCol(KolRWh) = True
.Cell(flexcpText, 0, KolRWh, 1, KolRWh) = "Return WH"

.MergeCol(KolRCls) = True
.Cell(flexcpText, 0, KolRCls, 1, KolRCls) = "Return CLS"

.MergeCol(KolUnit) = True
.Cell(flexcpText, 0, KolUnit, 1, KolUnit) = "Unit"

.MergeCol(KolCurr) = True
.Cell(flexcpText, 0, KolCurr, 1, KolCurr) = "Curr"

.MergeCol(KolPrice) = True
.Cell(flexcpText, 0, KolPrice, 1, KolPrice) = "Price"

.MergeCol(kolService) = True
.Cell(flexcpText, 0, kolService, 1, kolService) = "Service"

.MergeCol(kolAmount) = True
.Cell(flexcpText, 0, kolAmount, 1, kolAmount) = "Amount"

.MergeCol(KolRemarks) = True
.Cell(flexcpText, 0, KolRemarks, 1, KolRemarks) = "Remark"

    





    

End With
grid.ColHidden(KolPoNo) = True
grid.ColHidden(koldoseqno) = True
grid.ColHidden(KolSeqNo) = True
grid.ColHidden(KolCust) = True
grid.ColHidden(KolDnTmp) = True
grid.ColHidden(KolReturnSeq) = True
grid.ColHidden(kolReqSeqNO) = True
grid.ColHidden(KolWHCode) = True
grid.FixedAlignment(KolRWh) = flexAlignCenterCenter
grid.FixedAlignment(KolRCls) = flexAlignCenterTop
grid.FixedAlignment(KolRCls) = flexAlignCenterCenter
grid.FixedAlignment(kolRQty) = flexAlignCenterCenter

'grid.ColAlignment(KolRWh) = flexAlignGeneralCenter

End Sub

Private Sub cbocust_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
nilKosong = True
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

    HakU = hakUpdate(Me.Name)
    Str_User = frmLogin.strUserID
    ComboLocAl
    Call SetCol
    Call Kosong(1)
    dtReturn = Now
    
nilKosong = False
End Sub

Sub ComboLocAl()

 cboCust.clear
 'CboCust.AddItem
 Call GetCombo(cboCust, "trade_Master", "Rtrim(Trade_Code)code,RTRIM(Trade_Name)name ", "WHERE Trade_CLS<>5 and Trade_CLS<>1", , "70;190", 260, True)
 
 cboDO.clear
 
 Call GetCombo(cboDO, "Do_Master", "Rtrim(DO_NO)", , , "10", 120)
End Sub
 Sub GetCombo(nmCombo As MSForms.ComboBox, ls_tablename As String, Optional ls_field As String, Optional ls_Condition As String, Optional lb_AllSelection As Boolean, Optional ls_ColWidth As String, Optional ld_ListWidth As Double, Optional lb_2rdField As Boolean, Optional lb_3rdField As Boolean, Optional lb_4rdField As Boolean)
    Dim ls_sql As String
    Dim i As Integer
    Dim lrs As New ADODB.Recordset
    
    With nmCombo
        
        If lb_4rdField = True Then
            .columnCount = 4
        ElseIf lb_3rdField = True Then
            .columnCount = 3
        ElseIf lb_2rdField = True Then
            .columnCount = 2
        Else
        .columnCount = 1
        End If
        
        If lrs.State <> adStateClosed Then lrs.Close
        
        If Trim(ls_field) = "" Then
            ls_sql = "Select Distinct * From " & ls_tablename
        Else
            ls_sql = "Select Distinct " & ls_field & " From " & ls_tablename & " " & ls_Condition
        End If
        
        lrs.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
        i = 0
        If lb_AllSelection = True Then
            .AddItem ""
            .List(i, 0) = strAll
            
            If lb_2rdField = True Then .List(i, 1) = strAll
            If lb_3rdField = True Then .List(i, 2) = strAll
            If lb_4rdField = True Then .List(i, 3) = strAll
            i = 1
        End If

        While lrs.EOF = False
            .AddItem ""
            .List(i, 0) = Trim(lrs(0) & "")
            
            If lb_2rdField = True Then .List(i, 1) = Trim(lrs(1))
            If lb_3rdField = True Then .List(i, 2) = Trim(lrs(2) & "")
            If lb_4rdField = True Then .List(i, 3) = Trim(lrs(3) & "")
            lrs.MoveNext
            i = i + 1
        Wend
        
        If lrs.State <> adStateClosed Then lrs.Close
         
        If ls_ColWidth = "" Or IsNull(ld_ListWidth) = True Then
            .ColumnWidths = "20;40"
            .ListWidth = 60
        Else
            .ColumnWidths = ls_ColWidth
            .ListWidth = ld_ListWidth
        End If
    End With

End Sub
'***********************************************************

'************************* View Dt *************************
Private Sub headerGrid()
End Sub
Sub GridView()
Dim rGrid As New ADODB.Recordset
Dim brs As Integer, Col As Integer
Dim sqlQ As String
Dim Sqldetail As String 'Query untuk simpen data string
Dim RDetail As New ADODB.Recordset 'Untuk Detail dari Return

sqlQ = "select DM.Cust_Code,DM.Do_Date,DM.Do_No,Dm.Remarks,DO.PO_NO,DO.DOSeq_No,DO.Seq_NO"
sqlQ = sqlQ & " ,DO.Item_Code,(SELECT Item_Name FROM ITem_Master WHERE Item_Code=DO.Item_Code) as DescITem,"
sqlQ = sqlQ & "  DO.MakerItem_Code ,DO.Qty,DO.LOt_No,DO.Currency_Code,(SELECT Description FROM Curr_Cls WHERE Curr_Cls=DO.Currency_Code) AS Currency_Name"
sqlQ = sqlQ & " ,DO.Unit_Cls,(SELECT Description FROM UNIT_CLS WHERE Unit_Cls=DO.Unit_Cls) aS Unit_Description"
sqlQ = sqlQ & " ,ISNULL(Do.Price,0)AS Price,isnull(do.service,0)as Service,Do.Amount"
sqlQ = sqlQ & " From DO_Master DM"
sqlQ = sqlQ & " INNER JOIN Delivery_Order DO ON dm.DO_No=do.DO_NO"
sqlQ = sqlQ & " WHERE DM.Do_Date BETWEEN '" & DtStart.Value & "' AND '" & DtEnd.Value & "'"
If cboCust.Text <> "" Then
sqlQ = sqlQ & " AND DM.Cust_Code='" & cboCust & "'"
End If
If cboDO.Text <> "" Then
sqlQ = sqlQ & " AND DM.Do_No='" & cboDO.Text & "'"
End If
sqlQ = sqlQ & " ORDER BY DM.DO_NO "

Set rGrid = Db.Execute(sqlQ)
Dim tmp As Integer
SetCol
tmp = grid.Rows
grid.Rows = 2
grid.Rows = tmp

If rGrid.EOF Then
    LblErrMsg = DisplayMsg(4006)
    Exit Sub
End If
 brs = 2
 MousePointer = vbHourglass
With grid

grid.Rows = 2
.ColComboList(KolRWh) = getwh
While Not rGrid.EOF
    .Rows = grid.Rows + 1
    .TextMatrix(brs, 0) = ""
    .Cell(flexcpBackColor, brs, 0) = vbWhite
    '.TextMatrix(brs, ColDODate) = Format(rGrid!DO_Date, "dd MMM yyyy")
    .TextMatrix(brs, 1) = rGrid.Fields("Do_NO")
    .TextMatrix(brs, 2) = Format(rGrid!do_date, "dd MMM yyyy")
    .TextMatrix(brs, 3) = rGrid.Fields("Item_Code")
    .TextMatrix(brs, 4) = rGrid!Descitem
    .TextMatrix(brs, 5) = rGrid!Lot_no
    .TextMatrix(brs, 6) = rGrid!Qty
    .TextMatrix(brs, 7) = ""
    .TextMatrix(brs, 8) = ""
    .TextMatrix(brs, 9) = ""
    .TextMatrix(brs, 10) = rGrid!unit_description
    .TextMatrix(brs, 11) = rGrid!Currency_Name
    
    
    .TextMatrix(brs, 12) = Format(IIf(IsNull(rGrid!Price), "", rGrid!Price), gs_formatQty)
    .TextMatrix(brs, 13) = Format(IIf(IsNull(rGrid!service), "", rGrid!service), gs_formatQty)
    .TextMatrix(brs, 14) = Format(IfNol(rGrid!Amount), gs_formatQty)
    .TextMatrix(brs, 15) = IIf(IsNull(rGrid!Remarks), "", rGrid!Remarks)
    .TextMatrix(brs, KolPoNo) = rGrid!po_no
    .TextMatrix(brs, koldoseqno) = rGrid!DOSeq_No
    .TextMatrix(brs, KolSeqNo) = rGrid!Seq_no
    .TextMatrix(brs, KolCust) = rGrid!Cust_CodE
       
       'Tampilkan Detail Dari Return
       Sqldetail = GetQueryDetail(brs) 'Ambil Query Detail
       Set RDetail = Db.Execute(Sqldetail)
       
       While Not RDetail.EOF
         brs = brs + 1
        .Rows = .Rows + 1
        '.Cell(flexcpBackColor, .Rows - 1, 0) = vbWhite 'kolom nol di warankan putih

        .TextMatrix(brs, KolDn) = (RDetail!Reference)
        .Cell(flexcpBackColor, .Rows - 1, KolDn) = vbWhite
        
        .TextMatrix(brs, KolDate) = Format(RDetail!Return_Date, "dd MMM yyyy")
        .Cell(flexcpBackColor, brs, KolDate) = vbWhite
        
        .TextMatrix(brs, KolItem) = (RDetail!Item_Code)
        .Cell(flexcpBackColor, brs, KolItem) = vbWhite
        
        .TextMatrix(brs, Koldesc) = (RDetail!itemname)
        .Cell(flexcpBackColor, brs, Koldesc) = vbWhite
       
         .TextMatrix(brs, kolRQty) = (RDetail!Return_Qty)
         .Cell(flexcpBackColor, brs, kolRQty) = vbWhite
         
         
         '.ColComboList(KolRWh) = getwh
         .TextMatrix(brs, KolRWh) = RTrim(RDetail!WH_Name) ' RTrim(Get_Record("SELECT WH_Name FROM WareHouse_master WHERE WH_Code='" & (RDetail!wh_code) & "'"))
         'menyimpan whcode
         .TextMatrix(brs, KolWHCode) = RTrim(RDetail!wh_code)
         .Cell(flexcpBackColor, brs, KolRWh) = vbWhite
         
         .ColComboList(KolRCls) = "#D1;DN Return|#D2;Sales Return"
         .TextMatrix(brs, KolRCls) = (RDetail!Return_Cls)
         .Cell(flexcpBackColor, brs, KolRCls) = vbWhite
         
         .Cell(flexcpBackColor, brs, KolRemarks) = vbWhite
         .TextMatrix(brs, KolRemarks) = (RDetail!Remarks)
         
         .TextMatrix(brs, KolReturnSeq) = (RDetail!ReturnSeq_no)
         
          
          'Turunkan data Header ke detail
          .TextMatrix(brs, KolPrice) = Format(IfNol(rGrid!Price), gs_formatQty)
          .TextMatrix(brs, KolUnit) = (rGrid!unit_description)
          .TextMatrix(brs, KolCurr) = (rGrid!Currency_Name)
          .TextMatrix(brs, kolService) = Format(IfNol(rGrid!service), gs_formatQty)
          .TextMatrix(brs, KolDnTmp) = (rGrid!do_no)
          .TextMatrix(brs, KolPoNo) = rGrid!po_no
          .TextMatrix(brs, KolSeqNo) = (rGrid!Seq_no)
          .TextMatrix(brs, koldoseqno) = (RDetail!DOSeq_No)
          .TextMatrix(brs, KolCust) = (RDetail!Cust_CodE)
          .TextMatrix(brs, kolAmount) = (RDetail!Amount)
          
         
                 IsiSumWH brs
       RDetail.MoveNext
       Wend
       
       
       
    brs = brs + 1
    rGrid.MoveNext
Wend
End With
LblErrMsg = ""
MousePointer = vbDefault
End Sub
Function IfNUllStr(Data)
Dim s As String
If Data = "" Or IsNull(Data) Then
IfNUllStr = ""
Else
IfNUllStr = Trim(Data)
End If


End Function
'Query mengambil detail dari return
Function GetQueryDetail(Baris)
     GetQueryDetail = "select *, (SELECT Item_Name FROM Item_Master WHERE Item_Code=DR.Item_Code)AS ItemName, " & _
     "(select WH_Name FROM warehouse_Master where WH_Code=DR.WH_Code) WH_Name FROm delivery_Return DR WHERE DR.Cust_Code='" & RTrim(grid.TextMatrix(Baris, KolCust)) & "'"
     GetQueryDetail = GetQueryDetail & " AND DR.DO_NO='" & RTrim(grid.TextMatrix(Baris, KolDn)) & "' AND PO_NO='" & RTrim(grid.TextMatrix(Baris, KolPoNo)) & "'"
     GetQueryDetail = GetQueryDetail & " AND DR.Seq_No='" & RTrim(grid.TextMatrix(Baris, KolSeqNo)) & "' AND doSeq_No='" & RTrim(grid.TextMatrix(Baris, koldoseqno)) & "'"
End Function
Sub IsiGrid(Optional stFilter As Byte)
End Sub

Function ConvDate(Tgl)
ConvDate = IIf(IsDate(Tgl), Format(Tgl, "yyyy-MM-dd"), "1900-01-31")

End Function

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

Dim pesandtAwal As String, pesandtAkhir As String
With grid

If Col = KolDate Then
End If


If (Col <> 0) And Col <> 1 And Col <> 7 And Col <> 9 And Col <> 8 And Col <> KolRemarks Then
Cancel = True
End If

'mengeck agar yang ada putih di awal tidak bisa di edit
    If .Cell(flexcpBackColor, Row, 1) <> vbWhite And Col <> 0 Then
        Cancel = True
    End If
    
 



    


End With
End Sub

Private Sub grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'mengecek  jika data telah Quantity telah di isi
 
End Sub

Private Sub grid_Click()

nilKosong = True
With grid
    LblErrMsg = ""
    If .Row > 0 Then
        If .Cell(flexcpBackColor, .Row, .Col) = vbWhite Then .FocusRect = flexFocusInset Else .FocusRect = flexFocusNone
        If .Cell(flexcpBackColor, .Row, 2) <> &HE0E0E0 Then
            If .Col = 2 Then
          If .Cell(flexcpBackColor, .Row, KolDn) <> vbWhite Then
                    dtReturn.Value = Format(Now, "dd MMM yyyy")
                dtReturn.Visible = False
                Else
                
                dtReturn.Visible = True
                If .TextMatrix(.Row, 2) <> "" Then
                dtReturn.Value = Format(.TextMatrix(.Row, 2), "dd MMM yyyy")
                End If
                dtReturn.Left = .Cell(flexcpLeft, .Row, 2)
                dtReturn.top = .Cell(flexcpTop, .Row, 2)
                dtReturn.Width = .CellWidth + 30
                dtReturn.SetFocus
                tempRow = .Row
                End If
                
                
            End If
        End If
    End If

 
 'Membuat Normal,Kembali mendjadi putih
 If .ColSel = KolDate And .Cell(flexcpBackColor, .Row, 0) <> vbWhite Then
    grid.Cell(flexcpBackColor, .RowSel, KolDate, .RowSel, KolDate) = vbWhite
 End If
 
 If .ColSel = kolRQty And .Cell(flexcpBackColor, .Row, 0) <> vbWhite Then
    grid.Cell(flexcpBackColor, .RowSel, kolRQty, .RowSel, kolRQty) = vbWhite
 End If
 
 If .ColSel = KolRWh And .Cell(flexcpBackColor, .Row, 0) <> vbWhite Then
    grid.Cell(flexcpBackColor, .RowSel, KolRWh, .RowSel, KolRWh) = vbWhite
 End If
 
 If .ColSel = KolDn And .Cell(flexcpBackColor, .Row, 0) <> vbWhite Then
    grid.Cell(flexcpBackColor, .RowSel, KolDn, .RowSel, KolDn) = vbWhite
 End If
 
 If .ColSel = KolRCls And .Cell(flexcpBackColor, .Row, 0) <> vbWhite Then
    grid.Cell(flexcpBackColor, .RowSel, KolRCls, .RowSel, KolRCls) = vbWhite
 End If
 
End With
nilKosong = False
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = kolRQty Then
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If
    
     If Col = 0 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii <> Asc("C") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
            KeyAscii = 0
        If KeyAscii = Asc(".") Then KeyAscii = 0
   End If
    



End Sub
Function getRow(Row)
Dim s As Boolean
Dim i As Integer
i = Row + 1
With grid
s = True
While s
    If i = .Rows Then
    '.Rows = .Rows + 1
    getRow = .Rows - Row
    Exit Function
    End If
    If .Cell(flexcpBackColor, i, 0) = vbWhite Then
        s = False
        getRow = i - Row
    End If
i = 1 + i
Wend
End With
End Function
Function getwh()
Dim s As String
Dim rd As New ADODB.Recordset
s = "select wh_code,Wh_name FROM warehouse_master"
getwh = ""
Set rd = Db.Execute(s)
'fg.ColComboList(1) = "#1;Full time|#23;Part time|#65;Contractor|#78;Intern|#0;Other"


While Not rd.EOF
If getwh = "" Then
  getwh = "#" & RTrim(rd!wh_code) & ";" & RTrim(rd!WH_Name) & ""
Else
getwh = getwh & "|#" & RTrim(rd!wh_code) & ";" & RTrim(rd!WH_Name) & ""
End If
rd.MoveNext
Wend

End Function
Function IfNol(Angka)
If Angka = "" Then
IfNol = 0
Else
IfNol = CDbl(Angka)
End If
'IfNol = (IIf(Angka = "", 0, CDbl(Angka)))
End Function
Sub IsiSumWH(Row)
Dim jml As Double
Dim barisatas As Integer
Dim Baris As Integer
jml = 0
'ambil jumlah atas dulu
Baris = Row
With grid
While .Cell(flexcpBackColor, Baris, 0) <> vbWhite
jml = jml + IfNol(.TextMatrix(Baris, kolRQty))
Baris = Baris - 1
Wend
barisatas = Baris
'ambil batas bawah
Baris = Row + 1
If CInt(Baris) < CInt(.Rows - 1) Then
    While .Cell(flexcpBackColor, Baris, 0) <> vbWhite And CInt(Baris) < CInt(.Rows - 1)
        jml = jml + IIf(.TextMatrix(Baris, kolRQty) = "", 0, .TextMatrix(Baris, kolRQty))
        Baris = Baris + 1
    Wend
End If
.TextMatrix(barisatas, kolRQty) = jml
'Mengecek apakah Jumalh yang di balikin sesuai dengan data yang ada
If CDbl(.TextMatrix(barisatas, kolRQty)) > CDbl(.TextMatrix(barisatas, KolQty)) Then
    .TextMatrix(barisatas, kolRQty) = .TextMatrix(barisatas, kolRQty) - .TextMatrix(Row, kolRQty)
    .TextMatrix(Row, kolRQty) = 0
    LblErrMsg = DisplayMsg(4043) & " " & .TextMatrix(barisatas, KolQty)
Else
LblErrMsg = ""
End If

End With
End Sub
Public Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim totRet As Double, totRemain As Double, setRow As Integer
Dim Baris As Byte
Dim WH As String

WH = ""

With grid
If UCase(.TextMatrix(Row, 0)) = "C" And .Cell(flexcpBackColor, Row, 0) = vbWhite And Col = 0 Then
    WH = getwh()
    Baris = getRow(Row) + Row
        .AddItem "", Baris 'getRow(Row) + Row  '+ Row 'grid.Rows - 2
        .Cell(flexcpBackColor, Baris, KolDn) = vbWhite
        .Cell(flexcpBackColor, Baris, KolDate) = vbWhite
        .Cell(flexcpBackColor, Baris, kolRQty) = vbWhite
        .Cell(flexcpBackColor, Baris, KolRWh) = vbWhite
        .Cell(flexcpBackColor, Baris, KolRCls) = vbWhite
        .Cell(flexcpBackColor, Baris, KolRemarks) = vbWhite
        .TextMatrix(Baris, KolPrice) = IfNUllStr(.TextMatrix(.RowSel, KolPrice))
        .TextMatrix(Baris, kolService) = IfNUllStr(.TextMatrix(.RowSel, kolService))
        .TextMatrix(Baris, KolUnit) = IfNUllStr(.TextMatrix(.RowSel, KolUnit))
        .TextMatrix(Baris, KolCurr) = IfNUllStr(.TextMatrix(.RowSel, KolCurr))
        .TextMatrix(Baris, KolPrice) = IfNUllStr(.TextMatrix(.RowSel, KolPrice))
        .TextMatrix(Baris, KolCust) = IfNUllStr(.TextMatrix(.RowSel, KolCust))
        .TextMatrix(Baris, KolPoNo) = IfNUllStr(.TextMatrix(.RowSel, KolPoNo))
        .TextMatrix(Baris, KolSeqNo) = IfNUllStr(.TextMatrix(.RowSel, KolSeqNo))
        .TextMatrix(Baris, koldoseqno) = IfNUllStr(.TextMatrix(.RowSel, koldoseqno))
        
        .TextMatrix(Baris, KolItem) = IfNUllStr(.TextMatrix(.RowSel, KolItem))
        'ambil wh code sesuai dengan setting item master,akhir januari 2009 By dudi
         sql = "select (SELECT WH_Name FROM Warehouse_master WHERE wh_Code=item_master.WH_CODE) WH_Name FROM item_Master WHERE Item_Code='" & IfNUllStr(.TextMatrix(.RowSel, KolItem)) & "'"
        .TextMatrix(Baris, KolRWh) = RTrim(Get_Record(sql))  ' IfNUllStr(.TextMatrix(.RowSel, KolItem))
        .TextMatrix(Baris, KolDnTmp) = IfNUllStr(.TextMatrix(.RowSel, KolDn))
        .TextMatrix(Baris, KolLot) = IfNUllStr(.TextMatrix(.RowSel, KolLot))
         
        .TextMatrix(Baris, kolService) = .TextMatrix(.RowSel, kolService)
        .TextMatrix(Baris, kolAmount) = grid.TextMatrix(.RowSel, kolAmount)
        .TextMatrix(Baris, Koldesc) = grid.TextMatrix(.RowSel, Koldesc)
        .TextMatrix(Baris, KolReturnSeq) = 0
        .ColComboList(KolRWh) = WH
    .ColComboList(KolRCls) = "#D1;DN Return|#D2;Sales Return"
ElseIf UCase(.TextMatrix(Row, 0)) = "D" And getRow(Row) <> 1 And .Cell(flexcpBackColor, Row, 0) = vbWhite And getRow(Row) <> 1 Then  'And getRow(Row) <> Row Then
    If .TextMatrix(getRow(Row) + Row - 1, KolReturnSeq) = "0" Then
        .RemoveItem getRow(Row) + Row - 1 ' 1 '+ Row
        IsiSumWH getRow(Row) + Row - 1
    End If
        
End If


If Col = kolRQty Then
    IsiSumWH Row
    grid.TextMatrix(Row, kolAmount) = Format((IfNol(grid.TextMatrix(Row, KolPrice)) + IfNol(grid.TextMatrix(Row, kolService))) * IfNol(.TextMatrix(Row, kolRQty)), "##,##0.#0")
    'grid.TextMatrix(Row, kolAmount) = IfNol(grid.TextMatrix(Row, kolAmount)) * IfNol(grid.TextMatrix(Row, kolRQty))
    Exit Sub
End If


  

End With
End Sub
Function TakAda(Baris)
TakAda = False
End Function
'***********************************************************

'************************* Process *************************
Function chkSave(Optional chkDetail As Byte) As Boolean
Dim rsCheck As New ADODB.Recordset

chkSave = True

End Function

Private Sub cmdSearch_Click()
    GridView
End Sub

Private Sub cmdprocess_Click(Index As Integer)
Dim tanya

Me.MousePointer = vbHourglass
Select Case Index
    Case 0: 'Save
        If chkSave(1) Then: LblErrMsg = "": Call SavingData '  ProcessSave
        
    Case 1:  'Cancel
       Call GridView: Call Kosong
        
    Case 2:  'Clear
        Call Kosong(1): cboCust.SetFocus
End Select
Me.MousePointer = vbDefault
End Sub
Function Get_Field(sql, Field)
Dim Rdata As New ADODB.Recordset
Set Rdata = Db.Execute(sql)
Get_Field = ""
If Not Rdata.EOF Then
 Get_Field = IIf(IsNull(Rdata.Fields(Field)), "", Rdata.Fields(Field))
End If
End Function

Function CekTanggal()
CekTanggal = False
Dim pesandtAwal As String, pesandtAkhir As String
With grid
For i = 2 To .Rows - 1
         If .Cell(flexcpBackColor, i, KolDn) = vbWhite Then
                pesandtAwal = up_ValidateDateRange(ConvDate(.TextMatrix(i, KolDate)), True)
                pesandtAkhir = up_ValidateDateRange(ConvDate(.TextMatrix(i, KolDate)), True)
                If pesandtAwal <> "" Or pesandtAkhir <> "" Then
                    CekTanggal = True
                    .Cell(flexcpBackColor, i, KolDate, i, KolDate) = vbRed
                End If
          End If
                      
Next
End With
End Function
Function NotComplete()
'mengecek kelengkapand data
            'cek tanggal
 NotComplete = False
    With grid
     For i = 2 To .Rows - 1
         If .Cell(flexcpBackColor, i, KolDn) = vbWhite Then
            'mengecek tanggal data terisi atau tidak
            If Not IsDate(.TextMatrix(i, KolDate)) Then
                .Cell(flexcpBackColor, i, KolDate, i, KolDate) = vbRed
                NotComplete = True
            End If
            
            
            If .TextMatrix(i, kolRQty) = "" Then
                .Cell(flexcpBackColor, i, kolRQty, i, kolRQty) = vbRed
                NotComplete = True
            End If
            If .TextMatrix(i, KolRWh) = "" Then
                .Cell(flexcpBackColor, i, KolRWh, i, KolRWh) = vbRed
                NotComplete = True
            End If
            If .TextMatrix(i, KolRCls) = "" Then
             .Cell(flexcpBackColor, i, KolRCls, i, KolRCls) = vbRed
             NotComplete = True
            End If
            If .TextMatrix(i, KolDn) = "" Then
                .Cell(flexcpBackColor, i, KolDn, i, KolDn) = vbRed
                NotComplete = True
            End If
            
            
          End If
     Next
    End With
End Function
Sub SavingData()
Dim rddat As New ADODB.Recordset
Dim i As Byte
Dim SqlInsert As String
Dim sqlcust, Cust_CodE As String
Dim sqlUpdate As String
Dim sDel As String
Dim Ket As String
Dim sField As String
Ket = "0"
Dim tampungBln As String
Dim blnFix As Integer
Dim jmlItems As Integer
Dim sq As String
Dim thnFix As Integer
Dim LblInput As String
tampungBln = newCls.blnAkhir()
                    blnFix = Split(tampungBln, ",")(0)
                    thnFix = Split(tampungBln, ",")(1)
With grid
For i = 2 To .Rows - 1
       
       'mengecek kelengkapan data terlebih dahulu
       
       
       
       If .Cell(flexcpBackColor, i, KolDn) = vbWhite Then 'Jika Reference nya berwana Putih
       
            'Hapus yang ada tanda D nye pada kolom pertama dan Yang Ada returnSeq_no nya di  kolom 22
            If .Cell(flexcpBackColor, i, KolDn) = vbWhite And UCase(.TextMatrix(i, 0)) = "D" And IfNol(.TextMatrix(i, KolReturnSeq)) <> 0 Then
                LblInput = MsgBox("Do you really to delete this Reference " & .TextMatrix(i, KolDn) & " ?", _
                vbYesNo + vbQuestion, "Confirmation")
                If LblInput = vbYes Then
                    sq = "SELECT Return_qty,WH_Code,Return_Cls,Item_Code FROM    Delivery_Return WHERE "
                    sq = sq & " Cust_Code='" & .TextMatrix(i, KolCust) & "' And PO_NO='" & .TextMatrix(i, KolPoNo) & "' AND "
                    sq = sq & " DO_NO='" & .TextMatrix(i, KolDnTmp) & "' And PO_NO='" & .TextMatrix(i, KolPoNo) & "' AND "
                    sq = sq & " DOSeq_No=" & .TextMatrix(i, koldoseqno) & " And Seq_No=" & .TextMatrix(i, KolSeqNo) & " AND Returnseq_No=" & IfNol(.TextMatrix(i, KolReturnSeq))
                
                    'Update
                    If Get_Field("select StockCOntrol_Cls FROM Item_master WHERE Item_code='" & .TextMatrix(i, KolItem) & "'", 0) = "01" _
                    And Get_Field("select StockCOntrol_Cls FROM Warehouse_Master WHERE Wh_Code='" & .TextMatrix(i, KolWHCode) & "'", 0) = "01" Then
                    'Delete Stokc
                    Call newCls.updateStock(Get_Field(sq, 1), Get_Field(sq, 3), IfNol(Get_Field(sq, 0)), IfNUllStr(.TextMatrix(i, KolLot)), Format(.TextMatrix(i, KolDate), "yyyy-mm-dd"), blnFix, thnFix, Db, "Supply", 0, 1)
                    End If
                
                    sDel = "DELETE FROM Delivery_Return WHERE "
                    sDel = sDel & " Cust_Code='" & .TextMatrix(i, KolCust) & "' And PO_NO='" & .TextMatrix(i, KolPoNo) & "' AND "
                    sDel = sDel & " DO_NO='" & .TextMatrix(i, KolDnTmp) & "' And PO_NO='" & .TextMatrix(i, KolPoNo) & "' AND "
                    sDel = sDel & " DOSeq_No=" & .TextMatrix(i, koldoseqno) & " And Seq_No=" & .TextMatrix(i, KolSeqNo) & " AND Returnseq_No=" & IfNol(.TextMatrix(i, KolReturnSeq))
                    Db.Execute (sDel)
                    Ket = Ket & ",1"
                    Delete_PartSupply i
                End If
                
                
                

            
            
            End If
                
            
            
            
        
            If .TextMatrix(i, 0) <> "D" Then
            
            
            If NotComplete Then
                LblErrMsg = DisplayMsg(5012)
                Exit Sub
            End If
            
            If CekTanggal Then
                LblErrMsg = DisplayMsg(1022)
                Exit Sub
            End If
            
            
                If .TextMatrix(i, KolReturnSeq) = 0 Then 'Insert Data Baru
                    SqlInsert = "INSERT INTO Delivery_Return (Cust_Code,DO_No,PO_NO,Seq_No,DOSEq_No,"
                    SqlInsert = SqlInsert & " Return_Date, Item_Code,"
                    SqlInsert = SqlInsert & " Reference,Return_Qty,WH_COde,"
                    SqlInsert = SqlInsert & " Return_Cls,Remarks,Last_Update,Last_User,Register_Date,"
                    SqlInsert = SqlInsert & "Unit_Cls,Lot_No,Curr_Code,Price,Service,Amount"
                    SqlInsert = SqlInsert & " ) Values ("
                    SqlInsert = SqlInsert & "'" & .TextMatrix(i, KolCust) & "','" & .TextMatrix(i, KolDnTmp) & "',"
                    SqlInsert = SqlInsert & "'" & .TextMatrix(i, KolPoNo) & "'," & IfNUllStr(.TextMatrix(i, KolSeqNo)) & ","
                    SqlInsert = SqlInsert & "" & .TextMatrix(i, koldoseqno) & ",'" & Format(.TextMatrix(i, KolDate), "yyyy-MM-dd") & "',"
                    SqlInsert = SqlInsert & "'" & .TextMatrix(i, KolItem) & "',"
                    SqlInsert = SqlInsert & "'" & .TextMatrix(i, KolDn) & "'," & .TextMatrix(i, kolRQty) & ","
                    SqlInsert = SqlInsert & "'" & .TextMatrix(i, KolWHCode) & "','" & .TextMatrix(i, KolRCls) & "',"
                    SqlInsert = SqlInsert & "'" & .TextMatrix(i, KolRemarks) & "',"
                    SqlInsert = SqlInsert & "'" & Format(Now(), "yyyy-MM-dd") & "','" & Str_User & "','" & Format(Now(), "yyyy-MM-dd") & "',"
                    sField = "select * FROM unit_Cls WHERE Description='" & .TextMatrix(i, KolUnit) & "'"
                    SqlInsert = SqlInsert & "'" & Get_Field(sField, 0) & "','" & .TextMatrix(i, KolLot) & "','"
                    sField = "select * FROM curr_Cls WHERE Description='" & .TextMatrix(i, KolCurr) & "'"
                    SqlInsert = SqlInsert & Get_Field(sField, 0) & "',"
                    SqlInsert = SqlInsert & Replace(.TextMatrix(i, KolPrice), ",", "") & "," & Replace(.TextMatrix(i, kolService), ",", "") & "," & Replace(IfNol(.TextMatrix(i, kolAmount)), ",", "") & ")"
                    Db.Execute (SqlInsert)
                    Ket = Ket & ",2"
                    Dim RetrunS As Integer
                    RetrunS = Get_Field("SELECT MAX(ReturnSeq_NO)AS MAXi  FROM Delivery_Return", 0)
                    Insert_PartSupplier i, RetrunS
                    
                    'mengecek status item dan status wh
                    If Get_Field("select StockCOntrol_Cls FROM Item_master WHERE Item_code='" & .TextMatrix(i, KolItem) & "'", 0) = "01" _
                    And Get_Field("select StockCOntrol_Cls FROM Warehouse_Master WHERE Wh_Code='" & .TextMatrix(i, KolWHCode) & "'", 0) = "01" Then
                    
                    'Update Master_STOCK --- Insert Baru
                    Call newCls.updateStock(.TextMatrix(i, KolWHCode), .TextMatrix(i, KolItem), -(.TextMatrix(i, kolRQty)), IfNUllStr(.TextMatrix(i, KolLot)), Format(.TextMatrix(i, KolDate), "yyyy-mm-dd"), blnFix, thnFix, Db, "Supply", 0, 1)
                    End If
                  
                  ElseIf IfNol(.TextMatrix(i, KolReturnSeq)) <> 0 Then 'Untuk UPDATE Data
                   
                   
                   'Status nya Hapus Terlebih dahulu data yang ada d master_stock
                   If Get_Field("select StockCOntrol_Cls FROM Item_master WHERE Item_code='" & .TextMatrix(i, KolItem) & "'", 0) = "01" _
                    And Get_Field("select StockCOntrol_Cls FROM Warehouse_Master WHERE Wh_Code='" & .TextMatrix(i, KolWHCode) & "'", 0) = "01" Then
                    'Query ambil data  Jumlah yang lama sebelum di ganti
                    sq = "SELECT  Return_Qty FROM dbo.Delivery_Return "
                    sq = sq & " WHERE "
                    sq = sq & " Cust_Code='" & .TextMatrix(i, KolCust) & "' And PO_NO='" & .TextMatrix(i, KolPoNo) & "' AND "
                    sq = sq & " DO_NO='" & .TextMatrix(i, KolDnTmp) & "' AND "
                    sq = sq & " DOSeq_No=" & .TextMatrix(i, koldoseqno) & " And Seq_No=" & .TextMatrix(i, KolSeqNo) & " AND Returnseq_No=" & IfNol(.TextMatrix(i, KolReturnSeq))
                    Call newCls.updateStock(.TextMatrix(i, KolWHCode), .TextMatrix(i, KolItem), IfNol(Get_Field(sq, 0)), IfNUllStr(.TextMatrix(i, KolLot)), Format(.TextMatrix(i, KolDate), "yyyy-mm-dd"), blnFix, thnFix, Db, "Supply", 0, 1)
                   
                   End If
                   
                   
                   
                   sqlUpdate = " update dbo.Delivery_Return "
                   sqlUpdate = sqlUpdate & " SET "
                   sqlUpdate = sqlUpdate & " Return_Date ='" & .TextMatrix(i, KolDate) & "'"
                   sqlUpdate = sqlUpdate & " ,Reference ='" & .TextMatrix(i, KolDn) & "'"
                   sqlUpdate = sqlUpdate & " ,Return_Qty = " & .TextMatrix(i, kolRQty)
                   sqlUpdate = sqlUpdate & " ,WH_Code ='" & .TextMatrix(i, KolWHCode) & "'"
                   sqlUpdate = sqlUpdate & " ,Return_Cls ='" & .TextMatrix(i, KolRCls) & "'"
                   sqlUpdate = sqlUpdate & " ,Remarks ='" & .TextMatrix(i, KolRemarks) & "'"
                   sqlUpdate = sqlUpdate & " ,Last_Update='" & Now() & "'"
                   sqlUpdate = sqlUpdate & " ,Last_User = '" & Str_User & "'"
                   sField = "select * FROM unit_Cls WHERE Description='" & .TextMatrix(i, KolUnit) & "'"
                   sqlUpdate = sqlUpdate & " ,Unit_Cls ='" & IfNUllStr(Get_Field(sField, 0)) & "'"
                   sqlUpdate = sqlUpdate & " ,Lot_No='" & .TextMatrix(i, KolLot) & "'"
                   
                   sField = "select * FROM curr_Cls WHERE Description='" & IfNUllStr(.TextMatrix(i, KolCurr)) & "'"
                   sqlUpdate = sqlUpdate & " ,Curr_Code ='" & Get_Field(sField, 0) & "'"
                   
                   sqlUpdate = sqlUpdate & " ,Price =" & IfNol(Replace(.TextMatrix(i, KolPrice), ",", ""))
                   sqlUpdate = sqlUpdate & " ,Service=" & IfNol(.TextMatrix(i, kolService))
                   sqlUpdate = sqlUpdate & " ,amount =" & IfNol(.TextMatrix(i, kolAmount))
                   sqlUpdate = sqlUpdate & " WHERE "
                   sqlUpdate = sqlUpdate & " Cust_Code='" & .TextMatrix(i, KolCust) & "' And PO_NO='" & .TextMatrix(i, KolPoNo) & "' AND "
                   sqlUpdate = sqlUpdate & " DO_NO='" & .TextMatrix(i, KolDnTmp) & "' AND "
                   sqlUpdate = sqlUpdate & " DOSeq_No=" & .TextMatrix(i, koldoseqno) & " And Seq_No=" & .TextMatrix(i, KolSeqNo) & " AND Returnseq_No=" & IfNol(.TextMatrix(i, KolReturnSeq))
                   Db.Execute (sqlUpdate)
                   Ket = Ket & ",3"
                   Update_PartSupplier i
                   
                   'Insert Stock setelah data nya di hapus
                     If Get_Field("select StockCOntrol_Cls FROM Item_master WHERE Item_code='" & .TextMatrix(i, KolItem) & "'", 0) = "01" _
                     And Get_Field("select StockCOntrol_Cls FROM Warehouse_Master WHERE Wh_Code='" & .TextMatrix(i, KolWHCode) & "'", 0) = "01" Then
                        Call newCls.updateStock(.TextMatrix(i, KolWHCode), .TextMatrix(i, KolItem), -.TextMatrix(i, kolRQty), IfNUllStr(.TextMatrix(i, KolLot)), Format(.TextMatrix(i, KolDate), "yyyy-mm-dd"), blnFix, thnFix, Db, "Supply", 0, 1)
                     End If
                   
                   
                    
                   
                  End If
                    
                End If
       
       
       End If

Next
If Ket = "0" Then
LblErrMsg = DisplayMsg(5012)
Else
Call GridView
    'If s Then
LblErrMsg = DisplayMsg(8005)
    'End If
End If


End With
End Sub

Sub Delete_PartSupply(brs)

Dim sqlDelete As String
With grid
    sqlDelete = "DELETE FROM part_Supply "
    sqlDelete = sqlDelete & " WHERE ChildItem_Code='" & .TextMatrix(brs, KolItem) & "'"
    sqlDelete = sqlDelete & " AND DO_NO='" & IfNUllStr(.TextMatrix(brs, KolDnTmp)) & "'"
    
    Db.Execute (sqlDelete)
End With
End Sub

Sub Update_PartSupplier(brs)
Dim d As String
Dim SSS As String
      With grid
      SSS = " update [Part_Supply]"
      SSS = SSS & "   SET "
      'SSS = SSS & ",[FromWarehouse_Code]='" & IfNUllStr(.TextMatrix(brs, KolCust)) & "'"
      SSS = SSS & " [ToWarehouse_Code] ='" & .TextMatrix(brs, KolWHCode) & "'"
      SSS = SSS & ",[ChildSupply_date]='" & Format(.TextMatrix(brs, KolDate), "MMM dd yyyy") & "'"
      SSS = SSS & ",[Supply_Cls] ='" & .TextMatrix(brs, KolRCls) & "'" & vbCrLf
      SSS = SSS & ",[ChildRequirement_Qty] =" & .TextMatrix(brs, kolRQty)
      SSS = SSS & ",[Consumption_Qty] =" & .TextMatrix(brs, kolRQty)
      d = "SELECT * FROM unit_Cls WHERE Description='" & .TextMatrix(brs, KolUnit) & "'"
      SSS = SSS & ",[ChildUnit_Cls]='" & Get_Field(d, 0) & "'" & vbCrLf
      d = "select * FROM curr_Cls where Description='" & .TextMatrix(brs, KolCurr) & "'"
      SSS = SSS & ",[Production_Date] ='" & Format(.TextMatrix(brs, KolDate), "MMM dd yyyy") & "'"
      SSS = SSS & ",[Remarks] ='" & .TextMatrix(brs, KolRemarks) & "'" & vbCrLf
      SSS = SSS & ",[SJNo] ='" & .TextMatrix(brs, KolDn) & "'"
      SSS = SSS & ",[Last_Update] ='" & Format(Now(), "MMM dd YYYY") & "'"
      SSS = SSS & ",[Last_User] ='" & Str_User & "'" & vbCrLf
      SSS = SSS & " WHERE ChildItem_Code='" & .TextMatrix(brs, KolItem) & "'"
      SSS = SSS & " AND DO_NO='" & .TextMatrix(brs, KolDnTmp) & "'"
      SSS = SSS & " AND seq_No='" & .TextMatrix(brs, KolSeqNo) & "'"
      
'      Db.Execute (SSS)

End With

End Sub

Sub Insert_PartSupplier(brs, ReturnS)
Dim sql As String
Dim s As String
Dim d As String
With grid
sql = " INSERT INTO [Part_Supply]( "
sql = sql & " [FromWarehouse_Code]" & vbCrLf
sql = sql & ",[From_Address],[ToWarehouse_Code],[ChildSupply_date]" & vbCrLf
sql = sql & ",[ChildItem_Code],[Supply_Cls],[ChildRequirement_Qty]" & vbCrLf
sql = sql & ",[Consumption_Qty],[ChildUnit_Cls],[Currency_Code]" & vbCrLf
sql = sql & ",[Price],[Service]" & vbCrLf
sql = sql & ",[Amount],[ParentItem_Code],[Lot_No]" & vbCrLf
sql = sql & ",[Production_Date],[DO_No],[Remarks]" & vbCrLf
sql = sql & ",[SJNo]" & vbCrLf
sql = sql & ",[recseq_NO]" & vbCrLf
sql = sql & ",[Last_Update],[Last_User],[Register_Date])" & vbCrLf
sql = sql & " Values ("
sql = sql & "'" & .TextMatrix(brs, KolCust) & "',"
sql = sql & "'Addres" & "','" & .TextMatrix(brs, KolWHCode) & "','" & Format(Now(), "MMM dd yyyy") & "'," & vbCrLf
sql = sql & "'" & .TextMatrix(brs, KolItem) & "','" & .TextMatrix(brs, KolRCls) & "'," & .TextMatrix(brs, kolRQty) & "," & vbCrLf
s = "SELECT * FROM unit_Cls WHERE Description='" & .TextMatrix(brs, KolUnit) & "'"
d = "select * FROM curr_Cls where Description='" & .TextMatrix(brs, KolCurr) & "'"
sql = sql & "" & .TextMatrix(brs, kolRQty) & ",'" & Get_Field(s, 0) & "','" & Get_Field(d, 0) & "'," & vbCrLf
sql = sql & "" & Replace(.TextMatrix(brs, KolPrice), ",", "") & "," & Replace(.TextMatrix(brs, kolService), ",", "") & "," & vbCrLf
sql = sql & "" & Replace(.TextMatrix(brs, kolAmount), ",", "") & ",'ParentItemOcde','" & .TextMatrix(brs, KolLot) & "'," & vbCrLf
sql = sql & "'" & Format(Now, "MMM dd yyyy") & "','" & .TextMatrix(brs, KolDnTmp) & "','" & .TextMatrix(brs, KolRemarks) & "'," & vbCrLf
sql = sql & "'" & .TextMatrix(brs, KolDn) & "',"
sql = sql & "" & ReturnS & ","
sql = sql & "'" & Format(Now, "MMM dd yyyy") & "','" & Str_User & "','" & (Now) & "')"
Db.Execute (sql)


End With
End Sub
Sub ProcessSave()
'Dim RsSave As New ADODB.Recordset
'
'    dbTrans.ConnectionTimeout = 0
'    dbTrans.CommandTimeout = 0
'    dbTrans.Open Db.ConnectionString
'    dbTrans.BeginTrans
'
'    With grid
'        For i = 1 To .Rows - 1
'            If .Cell(flexcpBackColor, i, ColDODate) = &HE0E0E0 Then
'                rowHeader = i
'            Else
'                If .TextMatrix(i, ColS) = "D" Then
'                    If .TextMatrix(i, ColHSeqNo) <> "" Then
'                        Sql = "Delete Part_Supply Where Seq_No = " & .TextMatrix(i, ColHSeqNo)
'                        Db.Execute Sql
'
'                        Call clsMRP.ProcessStock("DR", dbTrans, Format(.TextMatrix(i, ColHDtBef), "yyyy-MM-dd"), _
'                                .TextMatrix(i, ColHWHBef), .TextMatrix(rowHeader, ColItemCD), CDbl(.TextMatrix(i, ColHQtyBef)), _
'                                .TextMatrix(rowHeader, ColLotNo), .TextMatrix(rowHeader, ColCaseNo), _
'                                "Receipt", "-", "05", , 0, CboCust, CDbl(.TextMatrix(rowHeader, ColHProdSeqNo)), CDbl(.TextMatrix(i, ColHSpoolBef)))
'                    End If
'                Else
'                    If .TextMatrix(i, ColHUpdate) = "1" Then
'                        Call clsMRP.saveSupply(dbTrans, IIf(.TextMatrix(i, ColHSeqNo) = "", 0, .TextMatrix(i, ColHSeqNo)), _
'                            CboCust, "", CboWHCode, Format(.TextMatrix(i, ColDODate), "yyyy-MM-dd"), .TextMatrix(rowHeader, ColItemCD), _
'                            "D1", CDbl(.TextMatrix(i, ColDORQty)), CDbl(.TextMatrix(i, ColDORQty)), .TextMatrix(rowHeader, ColHUnit), _
'                            .TextMatrix(rowHeader, ColLotNo), .TextMatrix(rowHeader, ColCaseNo), CDbl(.TextMatrix(i, ColSpoolQty)), _
'                            .TextMatrix(rowHeader, ColHCurr), CDbl(.TextMatrix(i, ColPrice)), CDbl(.TextMatrix(i, ColAmount)), _
'                            , , .TextMatrix(rowHeader, ColHProdResultDt), .TextMatrix(rowHeader, ColDoNo), .TextMatrix(rowHeader, ColHDOSeqNo), , .TextMatrix(rowHeader, ColHProdSeqNo))
'
'                        If .TextMatrix(i, ColHSeqNo) <> "" Then
'                            Call clsMRP.ProcessStock("DR", dbTrans, Format(.TextMatrix(i, ColHDtBef), "yyyy-MM-dd"), _
'                                .TextMatrix(i, ColHWHBef), .TextMatrix(rowHeader, ColItemCD), CDbl(.TextMatrix(i, ColHQtyBef)), _
'                                .TextMatrix(rowHeader, ColLotNo), .TextMatrix(rowHeader, ColCaseNo), _
'                                "Receipt", "-", "05", , 0, CboCust, CDbl(.TextMatrix(rowHeader, ColHProdSeqNo)), CDbl(.TextMatrix(i, ColHSpoolBef)))
'                        End If
'
'                        Call clsMRP.ProcessStock("DR", dbTrans, Format(.TextMatrix(i, ColDODate), "yyyy-MM-dd"), _
'                            CboWHCode, .TextMatrix(rowHeader, ColItemCD), CDbl(.TextMatrix(i, ColDORQty)), _
'                            .TextMatrix(rowHeader, ColLotNo), .TextMatrix(rowHeader, ColCaseNo), _
'                            "Receipt", "+", "05", 0, CDbl(.TextMatrix(i, ColSpoolQty)), CboCust, CDbl(.TextMatrix(rowHeader, ColHProdSeqNo)), , .TextMatrix(rowHeader, ColHProdResultDt))
'                    End If
'                End If
'            End If
'        Next i
'    End With
'
'    dbTrans.CommitTrans
'    dbTrans.Close
'
'    Call kosong: Call isiGrid
'    lblErrMsg = DisplayMsg(1101) 'Update Record Success
End Sub

Private Sub cmdFilter_Click(Index As Integer)

End Sub

'***********************************************************


'************************* Validate ************************
Private Sub CboCust_Change()
 LblErrMsg.Caption = ""
    lblcust.Caption = ""
    If cboCust.MatchFound Then
        If cboCust.Text <> "" Then lblcust.Caption = cboCust.List(cboCust.ListIndex, 1)
        cboDO.clear
        Call GetCombo(cboDO, "Do_Master", "Rtrim(DO_NO)", "WHERE Cust_Code='" & cboCust & "' AND DO_date Between '" & DtStart & "' AND '" & DtEnd & "'", , "10", 120)
        
    End If
End Sub

Private Sub CboWHCode_Change()
   
End Sub

Private Sub dtStart_Change()
    Call CboCust_Change
End Sub

Private Sub dtEnd_Change()
    Call CboCust_Change
End Sub

Private Sub dtReturn_Change()
If nilKosong Then Exit Sub

With grid
    If tempRow > 0 Then
        .TextMatrix(tempRow, 2) = Format(dtReturn, "dd MMM yyyy")
        '.TextMatrix(tempRow, ColHUpdate) = 1
    End If
End With
End Sub

Private Sub dtReturn_LostFocus()
    Call dtReturn_Change
    dtReturn.Visible = False
End Sub

'***********************************************************

'************************* Out *****************************
Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then Unload Me Else LblErrMsg.Caption = ErrMsg
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub
'***********************************************************


