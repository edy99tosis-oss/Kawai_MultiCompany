VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmValuationPriceReportWH 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Valuation Price Report Per Warehouse"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmValuationPriceReportWH.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H0080FFFF&
      Caption         =   "Summar&y"
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
      Left            =   12780
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9840
      Width           =   1035
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13050
      TabIndex        =   20
      Top             =   360
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   714
   End
   Begin VB.CommandButton Cmd_Save 
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
      Index           =   1
      Left            =   11670
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9840
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton Cmd_Save 
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
      Index           =   9
      Left            =   3750
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2055
      Width           =   1035
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Sub &Menu"
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
      Index           =   8
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9840
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last Page"
      Enabled         =   0   'False
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
      Index           =   7
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Next Page"
      Enabled         =   0   'False
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
      Index           =   6
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Prev Page"
      Enabled         =   0   'False
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
      Index           =   5
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Page"
      Enabled         =   0   'False
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
      Left            =   2535
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   255
      TabIndex        =   7
      Top             =   9030
      Width           =   14685
      Begin VB.Label LblErrMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LblErrMsg"
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
         Left            =   75
         TabIndex        =   8
         Top             =   225
         Width           =   14430
      End
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
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
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9840
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Left            =   2310
      TabIndex        =   1
      Top             =   2055
      Width           =   1290
      _ExtentX        =   2275
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
      CustomFormat    =   "MMM yyyy"
      Format          =   141230083
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6345
      Left            =   255
      TabIndex        =   13
      Top             =   2610
      Width           =   14685
      _cx             =   2088789295
      _cy             =   2088774584
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
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
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
   Begin VB.Label LblPesan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rere"
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
      Left            =   360
      TabIndex        =   19
      Top             =   9345
      Width           =   11940
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Valuation Price Report Per Warehouse"
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
      Left            =   270
      TabIndex        =   18
      Top             =   315
      Width           =   14610
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   6675
      X2              =   9690
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Label LblLocationName 
      BackStyle       =   0  'Transparent
      Caption         =   "LblLocationName"
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
      Left            =   6675
      TabIndex        =   17
      Top             =   1665
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Warehouse Name"
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
      Left            =   4950
      TabIndex        =   16
      Top             =   1665
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date (Month)"
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
      Left            =   585
      TabIndex        =   15
      Top             =   2115
      Width           =   1335
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "WareHouse CD"
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
      Left            =   585
      TabIndex        =   14
      Top             =   1665
      Width           =   1335
   End
   Begin MSForms.ComboBox CboLocationCD 
      Height          =   315
      Left            =   2310
      TabIndex        =   0
      Top             =   1635
      Width           =   2370
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "4180;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FrmValuationPriceReportWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dateUp As Date

Dim bteColProdCod As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColGroupDesc As Byte
Dim bteColUnit As Byte

Dim bteColAddress As Byte
Dim bteColAddress2 As Byte
Dim bteColAddress3 As Byte

Dim bteColPreMonth As Byte
Dim bteColReceipt As Byte
Dim bteColSupply As Byte
Dim bteColLossReject As Byte
Dim bteColEnd As Byte
Dim bteColInventory As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColInTransit As Byte
Dim bteColInTransitAmount As Byte

Dim bytSort As Byte

Public setTglPrint As String

Private Sub Header()
    
    bteColProdCod = 0
    bteColPartNo = 1
    bteColDesc = 2
    bteColGroupDesc = 3
    bteColUnit = 4
    
    bteColAddress = 5
    bteColAddress2 = 6
    bteColAddress3 = 7
    
    bteColPreMonth = 8
    bteColReceipt = 9
    bteColSupply = 10
    bteColLossReject = 11
    bteColEnd = 12
    bteColInventory = 13
    bteColPrice = 14
    bteColAmount = 15
    bteColInTransit = 16
    bteColInTransitAmount = 17
    
    grid.Rows = 1
    grid.ColS = 18
    
    grid.TextMatrix(0, bteColProdCod) = "Product Code"
    grid.TextMatrix(0, bteColPartNo) = "Part Number"
    grid.TextMatrix(0, bteColDesc) = "Description"
    grid.TextMatrix(0, bteColGroupDesc) = "Group Desc"
    grid.TextMatrix(0, bteColUnit) = "Unit"
    
    grid.TextMatrix(0, bteColAddress) = "Address1"
    grid.TextMatrix(0, bteColAddress2) = "Address2"
    grid.TextMatrix(0, bteColAddress3) = "Address3"
    
    grid.TextMatrix(0, bteColPreMonth) = "Pre Month Stock"
    grid.TextMatrix(0, bteColReceipt) = "Receipt Total"
    grid.TextMatrix(0, bteColSupply) = "Supply Total"
    grid.TextMatrix(0, bteColLossReject) = "Loss/Reject"
    grid.TextMatrix(0, bteColEnd) = "End of Month Stock"
    grid.TextMatrix(0, bteColInventory) = "Inventory"
    grid.TextMatrix(0, bteColPrice) = "Price"
    grid.TextMatrix(0, bteColAmount) = "Amount"
    grid.TextMatrix(0, bteColInTransit) = "In Transit"
    grid.TextMatrix(0, bteColInTransitAmount) = "In Transit Amount"
    
    grid.ColWidth(bteColProdCod) = 1400
    grid.ColWidth(bteColPartNo) = 1400
    grid.ColWidth(bteColDesc) = 2000
    grid.ColWidth(bteColGroupDesc) = 1200
    grid.ColWidth(bteColUnit) = 600
    
    grid.ColWidth(bteColAddress) = 1000
    grid.ColWidth(bteColAddress2) = 1000
    grid.ColWidth(bteColAddress3) = 1000
    
    grid.ColWidth(bteColPreMonth) = 1500
    grid.ColWidth(bteColReceipt) = 1250
    grid.ColWidth(bteColSupply) = 1250
    grid.ColWidth(bteColLossReject) = 1200
    grid.ColWidth(bteColEnd) = 1800
    grid.ColWidth(bteColInventory) = 1500
    grid.ColWidth(bteColPrice) = 1500
    grid.ColWidth(bteColAmount) = 1500
    grid.ColWidth(bteColInTransit) = 1500
    grid.ColWidth(bteColInTransitAmount) = 1800
    
    grid.ColAlignment(bteColProdCod) = flexAlignLeftCenter
    grid.ColAlignment(bteColPartNo) = flexAlignLeftCenter
    grid.ColAlignment(bteColDesc) = flexAlignLeftCenter
    grid.ColAlignment(bteColGroupDesc) = flexAlignLeftCenter
    grid.ColAlignment(bteColUnit) = flexAlignCenterCenter
    
    grid.ColAlignment(bteColAddress) = flexAlignLeftCenter
    grid.ColAlignment(bteColAddress2) = flexAlignLeftCenter
    grid.ColAlignment(bteColAddress3) = flexAlignLeftCenter
    
    grid.ColAlignment(bteColPreMonth) = flexAlignRightCenter
    grid.ColAlignment(bteColReceipt) = flexAlignRightCenter
    grid.ColAlignment(bteColSupply) = flexAlignRightCenter
    grid.ColAlignment(bteColLossReject) = flexAlignRightCenter
    grid.ColAlignment(bteColEnd) = flexAlignRightCenter
    grid.ColAlignment(bteColInventory) = flexAlignRightCenter
    grid.ColAlignment(bteColPrice) = flexAlignRightCenter
    grid.ColAlignment(bteColAmount) = flexAlignRightCenter
    grid.ColAlignment(bteColInTransit) = flexAlignRightCenter
    grid.ColAlignment(bteColInTransitAmount) = flexAlignRightCenter
    
End Sub


Private Sub CboLocationCD_Change()
Call clearGrid

Cmd_save(0).Enabled = True
Cmd_save(9).Enabled = True
Cmd_save(2).Enabled = True

If CboLocationCD.MatchFound Then
   LblLocationName = CboLocationCD.List(CboLocationCD.ListIndex, 1)
   LblErrMsg = ""
   
   If CboLocationCD.ListIndex = 0 Then
    Cmd_save(0).Enabled = False
    Cmd_save(9).Enabled = False
    Cmd_save(2).Enabled = True
   Else
    Cmd_save(0).Enabled = True
    Cmd_save(9).Enabled = True
    Cmd_save(2).Enabled = False
   End If
Else
   LblLocationName = ""
   LblErrMsg = DisplayMsg(4014) '"Location CD is not found !"
End If

End Sub
Sub clearGrid()
grid.clear
grid.Rows = 1
Call Header
End Sub

Private Sub CboLocationCD_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim j As Integer
If KeyCode = 13 Then
Call clearGrid
  j = 0
For i = 0 To CboLocationCD.ListCount - 1
    If UCase(Trim(CboLocationCD)) = UCase(Trim(CboLocationCD.List(i, 0))) Then
        CboLocationCD = Trim(CboLocationCD.List(i, 0))
        LblLocationName = Trim(CboLocationCD.List(i, 1))
        j = 1: Exit For
    End If
Next

If j = 0 Then
    LblErrMsg = DisplayMsg(4018) '"Invalid warehouse code !"
    Exit Sub
Else
    LblErrMsg = ""
End If

End If
End Sub

Private Sub CboLocationCD_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Cmd_Save_Click(Index As Integer)
Dim strSQL As String
Select Case Index
       Case 1
            toExcel
       Case 8:
                
                frmMainMenu.Show
                
                Unload Me
       Case 9:
                
                Dim i As Long
                If CboLocationCD.Text = "" Then
                   LblErrMsg = DisplayMsg(1042) '"Please choose warehouse !"
                Else
                    Me.MousePointer = vbHourglass
'                    Strsql = "exec [sp_normalize_receipt_supply_BY_Warehouse] '" & Trim(CboLocationCD.Text) & "'"
'                    Db.Execute Strsql
                
                
                    LblErrMsg = ""
                    
                       Call Header
                       grid.Rows = 1
                       Call Browse
                       For i = 4 To 7
                           Cmd_save(i).Enabled = True
                       Next i
                    Me.MousePointer = vbDefault
                    
                End If
        Case 0: PrintReport
        Case 2: PrintSummaryReport
End Select
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub DMonth_Change()
Call clearGrid
If Format(DMonth.Value, "MM") < Format(dateUp, "MM") And Val(Format(DMonth.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then _
            DMonth.Year = DMonth.Year + 1: GoTo pass
    If Format(DMonth.Value, "MM") > Format(dateUp, "MM") And Val(Format(DMonth.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then _
            DMonth.Year = DMonth.Year - 1
pass:
    dateUp = Format(DMonth.Value, "dd MMM yyyy")
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)

DMonth = Format(Date, "MMM yyyy")
dateUp = DMonth.Value

CtrlMenu1.FormName = Me.Name
Me.Caption = "Inventory Report"
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"



Call StockLocation

DMonth = Format(Now, "mmmm yyyy")
Call Header

LblLocationName = ""

LblErrMsg = ""

End Sub

Private Sub StockLocation()
Dim sql As String, ls_sql As String, RsStock As New ADODB.Recordset
Dim i As Long

If RsStock.State <> adStateClosed Then RsStock.Close

'ls_sql = " select * from (select wh_code, wh_name, '' Company_Code  from warehouse_master where stockcontrol_cls='01' union  " & _
'      " select trade_code wh_code, trade_name wh_name, Company_Code from trade_master where trade_code in(select manufacture_code from manufacture_line) )tbWarehouse " & _
'      " where wh_code in (" & _
'      "         select code from (select username, trade_code code from user_factory union select username, wh_code code from user_warehouse) a where a.username = '" & userLogin & "' " & _
'      " ) " & _
'      "and Company_Code = '" & Trim(CboCompany.Text) & "' order by wh_code "

ls_sql = " Select * from (select wh_code, wh_name  from warehouse_master where stockcontrol_cls='01' union  " & vbCrLf & _
         " select trade_code wh_code, trade_name wh_name from trade_master where trade_code in(select manufacture_code from manufacture_line))tbWarehouse order by wh_code " & vbCrLf
RsStock.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
CboLocationCD.columnCount = 2
CboLocationCD.clear

CboLocationCD.AddItem
CboLocationCD.List(0, 0) = strAll
CboLocationCD.List(0, 1) = strAll
CboLocationCD.List(0, 2) = strAll
i = 1

Do While Not RsStock.EOF
   CboLocationCD.AddItem ""
   CboLocationCD.List(i, 0) = Trim(RsStock("wh_code"))
   CboLocationCD.List(i, 1) = Trim(RsStock("wh_name"))
   i = i + 1
   RsStock.MoveNext
Loop

CboLocationCD.ColumnWidths = "50 pt; 150 pt"
CboLocationCD.ListWidth = 200
CboLocationCD.ListRows = 15
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid
.TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gs_formatQty)
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Grid.Col <> bteColInventory Then
   Cancel = True
'End If
End Sub

Private Sub Browse()

Dim RsStock As New ADODB.Recordset
Dim sqlControl As String
Dim RsInvControl As New ADODB.Recordset

If RsStock.State <> adStateClosed Then RsStock.Close

'sql = "select rtrim(gc.description) groupDesc, rtrim(uc.description) Unit_Desc, rtrim(item_name) item_name, rtrim(address) address, rtrim(wh_name) wh_name, rtrim(makeritem_code) makeritem_code, sm.*, " & _
'      vbLf & "  isnull(ip.inventory_price, 0) price, isnull(it.in_transit, 0) in_transit " & _
'      vbLf & "from stock_master sm " & _
'      vbLf & "  left join warehouse_master wm on sm.warehouse_code=wm.wh_code " & _
'      vbLf & "  left join item_master im on im.item_code=sm.item_code " & _
'      vbLf & "  left join group_cls gc on im.group_cls=gc.group_cls " & _
'      vbLf & "  left join unit_cls uc on im.unit_cls=uc.unit_cls " & _
'      vbLf & "  left join ( " & _
'      vbLf & "      select item_code, inventory_price " & _
'      vbLf & "          From (Select * From Inventory_Price " & _
'      vbLf & "                  Union All " & _
'      vbLf & "               Select * From InventoryPrice_History ) Inventory_Price " & _
'      vbLf & "          where inventory_year = '" & DMonth.year & "' and inventory_month = '" & DMonth.month & "' " & _
'      vbLf & "              and duty_status in ('0', '2', '3') " & _
'      vbLf & "      ) ip on sm.item_code = ip.item_code " & _
'      vbLf & "  left join ( " & _
'      vbLf & "      select whcode, item_code, sum(qty) in_transit from packing_master pm " & _
'      vbLf & "          inner join packing_detail pd on pm.packing_no = pd.packing_no " & _
'      vbLf & "              Where DateDiff(month, pm.stuffing_date, etd) > 0 " & _
'      vbLf & "                  and year(pm.stuffing_date) = '" & DMonth.year & "' and month(pm.stuffing_date) = '" & DMonth.month & "' " & _
'      vbLf & "          group by whcode, item_code " & _
'      vbLf & "      ) it on sm.item_code = it.item_code and sm.warehouse_code = it.whcode " & _
'      vbLf & "where warehouse_code = '" & CboLocationCD & "' " & _
'      vbLf & "order by  warehouse_code, sm.item_Code"
    
    sql = "  " & vbCrLf & _
                " DECLARE @LastClosing AS DATETIME " & vbCrLf & _
                " DECLARE @GetDiff AS INTEGER " & vbCrLf & _
                " DECLARE @SelectPeriod AS DATETIME " & vbCrLf & _
                " DECLARE @Location AS VARCHAR(30) " & vbCrLf & _
                " DECLARE @Company_Code AS CHAR(5) " & vbCrLf & _
                "  " & vbCrLf & _
                " " & vbCrLf & _
                " SET @Location='" & CboLocationCD & "' " & vbCrLf & _
                " SET @SelectPeriod='" & DMonth.Year & "-" & DMonth.Month & "-01' " & vbCrLf & _
                " SET @LastClosing= " & vbCrLf & _
                "                           (SELECT CONVERT(DATETIME,CAST(Inventory_Year AS VARCHAR(4)) +'-' +CAST(Inventory_Month AS varchar(2)) +'-01') " & vbCrLf & _
                "                               FROM Inventory_Control " & vbCrLf
    
    sql = sql + "                                   WHERE Inventory_Year=(SELECT MAX(Inventory_Year) FROM Inventory_Control) " & vbCrLf & _
                "                                       AND Inventory_Month=(SELECT MAX(Inventory_Month) FROM Inventory_Control " & vbCrLf & _
                "                                                                               WHERE Inventory_Year=(SELECT MAX(Inventory_Year) FROM Inventory_Control)) " & vbCrLf & _
                "                           )                                                    " & vbCrLf & _
                "  " & vbCrLf & _
                " SET @GetDiff=DATEDIFF(M,@LastClosing,@SelectPeriod) " & vbCrLf & _
                "  " & vbCrLf & _
                " SELECT  RTRIM(gc.description) groupDesc , " & vbCrLf & _
                "         RTRIM(uc.description) Unit_Desc , " & vbCrLf & _
                "         RTRIM(item_name) item_name , " & vbCrLf & _
                "         RTRIM(address) address , " & vbCrLf
    
    sql = sql + "         RTRIM(wh_name) wh_name , " & vbCrLf & _
                "         RTRIM(makeritem_code) makeritem_code , " & vbCrLf & _
                "         sm.* , " & vbCrLf & _
                "         ISNULL(ip.inventory_price, 0) price , " & vbCrLf & _
                "         ISNULL(it.in_transit, 0) in_transit " & vbCrLf & _
                "         --TAMBAH FIELD DiffClosing UTK HANDLE AMOUNT JIKA KONDISI SETELAH CLOSING MENGAMBIL INVENTORY (LIHAT LOOPING INTO GRID)  " & vbLf & _
                "         ,@GetDiff DiffClosing, " & vbLf & _
                "         -- Tambah Lokasi 2 dan 3 -- " & vbCrLf & _
                "         '' Location1, " & vbCrLf & _
                "         '' Location2, " & vbCrLf & _
                "         '' Location3 " & vbCrLf & _
                " FROM    " & vbCrLf & _
                "   (SELECT * FROM  " & vbCrLf & _
                "       (   SELECT warehouse_code, item_code, " & vbCrLf & _
                "                       SUM(Pre_MonthS) Pre_MonthS, SUM(ReceiptS) ReceiptS, " & vbCrLf & _
                "                       SUM(SupplyS) SupplyS, SUM(LossRejectS) LossRejectS, " & vbCrLf & _
                "                       SUM(CurrentS) CurrentS, SUM(InventoryS) InventoryS " & vbCrLf
    
    sql = sql + "           FROM             " & vbCrLf & _
                "           ( " & vbCrLf & _
                "               SELECT warehouse_code, item_code, " & vbCrLf & _
                "                   COALESCE(CASE WHEN @GetDiff=0 THEN LM_PreMonth " & vbCrLf & _
                "                            WHEN @GetDiff=1 THEN TM_PreMonth " & vbCrLf & _
                "                            WHEN @GetDiff=2 THEN NM_PreMonth END,0) Pre_MonthS, " & vbCrLf & _
                "                   COALESCE(CASE WHEN @GetDiff=0 THEN LM_Receipt " & vbCrLf & _
                "                            WHEN @GetDiff=1 THEN TM_Receipt " & vbCrLf & _
                "                            WHEN @GetDiff=2 THEN NM_Receipt END,0) ReceiptS, " & vbCrLf & _
                "                   COALESCE(CASE WHEN @GetDiff=0 THEN LM_Supply " & vbCrLf & _
                "                            WHEN @GetDiff=1 THEN TM_Supply " & vbCrLf
    
    sql = sql + "                            WHEN @GetDiff=2 THEN NM_Supply END,0) SupplyS, " & vbCrLf & _
                "                   COALESCE(CASE WHEN @GetDiff=0 THEN LM_LossReject " & vbCrLf & _
                "                            WHEN @GetDiff=1 THEN TM_LossReject " & vbCrLf & _
                "                            WHEN @GetDiff=2 THEN NM_LossReject END,0) LossRejectS, " & vbCrLf & _
                "                   COALESCE(CASE WHEN @GetDiff=0 THEN LM_Current " & vbCrLf & _
                "                            WHEN @GetDiff=1 THEN TM_Current " & vbCrLf & _
                "                            WHEN @GetDiff=2 THEN NM_Current END,0) CurrentS, " & vbCrLf & _
                "                   COALESCE(CASE WHEN @GetDiff=0 THEN LM_Inventory " & vbCrLf & _
                "                            WHEN @GetDiff=1 THEN TM_Inventory " & vbCrLf & _
                "                            WHEN @GetDiff=2 THEN NM_Inventory END,0) InventoryS " & vbCrLf & _
                "               FROM Stock_Master " & vbCrLf
    
    sql = sql + "                   WHERE Warehouse_Code=@Location            " & vbCrLf & _
                "  " & vbCrLf & _
                "               UNION ALL " & vbCrLf & _
                "  " & vbCrLf & _
                "               SELECT Warehouse_Code, Item_code, PreMonth, Receipt,Supply, LossReject, [Current], COALESCE(Inventory,[Current]) " & vbCrLf & _
                "                   FROM Stock_history " & vbCrLf & _
                "                       WHERE Warehouse_Code=@Location " & vbCrLf & _
                "                                   AND Stock_Year=YEAR(@SelectPeriod) " & vbCrLf & _
                "                                   AND Stock_Month=MONTH(@SelectPeriod) " & vbCrLf & _
                "           ) Stock GROUP BY warehouse_code, item_code " & vbCrLf & _
                "       )   SM " & vbCrLf
    
    sql = sql + "   )sm " & vbCrLf & _
                "         LEFT JOIN warehouse_master wm ON sm.warehouse_code = wm.wh_code " & vbCrLf & _
                "         LEFT JOIN item_master im ON im.item_code = sm.item_code " & vbCrLf & _
                "         LEFT JOIN group_cls gc ON im.group_cls = gc.group_cls " & vbCrLf & _
                "         LEFT JOIN unit_cls uc ON im.unit_cls = uc.unit_cls " & vbCrLf & _
                "         LEFT JOIN ( SELECT  item_code , " & vbCrLf & _
                "                             inventory_price " & vbCrLf & _
                "                     FROM     " & vbCrLf & _
                "                       (SELECT * FROM Inventory_Price   " & vbCrLf & _
                "                           UNION " & vbCrLf & _
                "                           SELECT * FROM InventoryPrice_History  " & vbCrLf
    
    sql = sql + "                       )Inventory_Price " & vbCrLf & _
                "                     WHERE   inventory_year = YEAR(@SelectPeriod) " & vbCrLf & _
                "                             AND inventory_month = MONTH(@SelectPeriod) " & vbCrLf & _
                "                             AND duty_status IN ( '0', '2', '3' ) " & vbCrLf & _
                "                   ) ip ON sm.item_code = ip.item_code " & vbCrLf & _
                "         LEFT JOIN ( SELECT  whcode , " & vbCrLf & _
                "                             item_code , " & vbCrLf & _
                "                             SUM(qty) in_transit " & vbCrLf & _
                "                     FROM    packing_master pm " & vbCrLf & _
                "                             INNER JOIN packing_detail pd ON pm.packing_no = pd.packing_no " & vbCrLf & _
                "                     WHERE   DATEDIFF(month, pm.stuffing_date, etd) > 0 " & vbCrLf
    
    sql = sql + "                             AND YEAR(pm.stuffing_date) = YEAR(@SelectPeriod) " & vbCrLf & _
                "                             AND MONTH(pm.stuffing_date) = MONTH(@SelectPeriod) " & vbCrLf & _
                "                     GROUP BY whcode , " & vbCrLf & _
                "                             item_code " & vbCrLf & _
                "                   ) it ON sm.item_code = it.item_code " & vbCrLf & _
                "                           AND sm.warehouse_code = it.whcode " & vbCrLf & _
                " WHERE   warehouse_code = @Location " & vbCrLf & _
                " ORDER BY warehouse_code , " & vbCrLf & _
                "         sm.item_Code " & vbCrLf & _
                "  " & vbCrLf
    
    
RsStock.Open sql, Db, adOpenDynamic, adLockOptimistic

sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year,inventory_month"

If RsInvControl.State <> adStateClosed Then RsInvControl.Close
RsInvControl.Open sqlControl, Db, adOpenKeyset, adLockOptimistic

If RsInvControl.EOF = True And RsInvControl.BOF = True Then
    LblErrMsg = DisplayMsg(4022) '"Inventory Stock hasn't been closed !"
    Exit Sub
End If

RsInvControl.MoveLast

'lblErrMsg = up_ValidateDateRange(DMonth.Value, False)
'If Trim(lblErrMsg) <> "" Then Exit Sub
'
'Select Case up_GetDateRange(DMonth)
'    Case 0:
'
'            i = 0
'            Do While Not RsStock.EOF
'                i = i + 1
'                grid.AddItem i
'                grid.TextMatrix(i, bteColProdCod) = Trim(RsStock!Item_Code)
'                grid.TextMatrix(i, bteColPartNo) = Trim(RsStock!makeritem_code)
'                grid.TextMatrix(i, bteColDesc) = uf_GetItemDescription(Trim(RsStock!Item_Code))
'                grid.TextMatrix(i, bteColGroupDesc) = Trim(RsStock!groupdesc & "")
'                grid.TextMatrix(i, bteColUnit) = Trim(RsStock!unit_desc)
'                grid.TextMatrix(i, bteColAddress) = Trim(RsStock!Address)
'                grid.TextMatrix(i, bteColPreMonth) = Format(RsStock!lm_premonth, gs_formatQty)
'                grid.TextMatrix(i, bteColReceipt) = Format(RsStock!lm_receipt, gs_formatQty)
'                grid.TextMatrix(i, bteColSupply) = Format(RsStock!lm_supply, gs_formatQty)
'                grid.TextMatrix(i, bteColLossReject) = Format(RsStock!lm_lossreject, gs_formatQty)
'                grid.TextMatrix(i, bteColEnd) = Format(RsStock!lm_current, gs_formatQty)
'                grid.TextMatrix(i, bteColInventory) = IIf(RsStock!lm_inventory = Null, "", Format(RsStock!lm_inventory, gs_formatQty))
'                grid.TextMatrix(i, bteColPrice) = Format(RsStock!Price, gs_formatPriceIDR)
'                grid.TextMatrix(i, bteColAmount) = Format(RsStock!Price * Val(RsStock!lm_inventory & ""), gs_formatPriceIDR)
'                grid.TextMatrix(i, bteColInTransit) = Format(RsStock!in_transit, gs_formatQty)
'                grid.TextMatrix(i, bteColInTransitAmount) = Format(RsStock!Price * Val(RsStock!in_transit & ""), gs_formatPriceIDR)
'                RsStock.MoveNext
'            Loop
'
'    Case 1:
'
'            i = 0
'            Do While Not RsStock.EOF
'                i = i + 1
'                grid.AddItem i
'                grid.TextMatrix(i, bteColProdCod) = Trim(RsStock!Item_Code)
'                grid.TextMatrix(i, bteColPartNo) = Trim(RsStock!makeritem_code)
'                grid.TextMatrix(i, bteColDesc) = uf_GetItemDescription(Trim(RsStock!Item_Code))
'                grid.TextMatrix(i, bteColAddress) = Trim(RsStock!Address)
'                grid.TextMatrix(i, bteColGroupDesc) = Trim(RsStock!groupdesc & "")
'                grid.TextMatrix(i, bteColUnit) = Trim(RsStock!unit_desc)
'                grid.TextMatrix(i, bteColPreMonth) = Format(RsStock!tm_premonth, gs_formatQty)
'                grid.TextMatrix(i, bteColReceipt) = Format(RsStock!tm_receipt, gs_formatQty)
'                grid.TextMatrix(i, bteColSupply) = Format(RsStock!tm_supply, gs_formatQty)
'                grid.TextMatrix(i, bteColLossReject) = Format(RsStock!tm_lossreject, gs_formatQty)
'                grid.TextMatrix(i, bteColEnd) = Format(RsStock!tm_current, gs_formatQty)
'                grid.TextMatrix(i, bteColInventory) = IIf(RsStock!tm_inventory = Null, "", Format(RsStock!tm_inventory, gs_formatQty))
'                grid.TextMatrix(i, bteColPrice) = Format(RsStock!Price, gs_formatPriceIDR)
'                grid.TextMatrix(i, bteColAmount) = Format(RsStock!Price * Val(RsStock!tm_current & ""), gs_formatPriceIDR)
'                grid.TextMatrix(i, bteColInTransit) = Format(RsStock!in_transit, gs_formatQty)
'                grid.TextMatrix(i, bteColInTransitAmount) = Format(RsStock!Price * Val(RsStock!in_transit & ""), gs_formatPriceIDR)
'                RsStock.MoveNext
'            Loop
'
'
'    Case 2:
'
'            i = 0
'            Do While Not RsStock.EOF
'                i = i + 1
'                grid.AddItem i
'                grid.TextMatrix(i, bteColProdCod) = Trim(RsStock!Item_Code)
'                grid.TextMatrix(i, bteColPartNo) = Trim(RsStock!makeritem_code)
'                grid.TextMatrix(i, bteColDesc) = uf_GetItemDescription(Trim(RsStock!Item_Code))
'                grid.TextMatrix(i, bteColAddress) = Trim(RsStock!Address)
'                grid.TextMatrix(i, bteColGroupDesc) = Trim(RsStock!groupdesc & "")
'                grid.TextMatrix(i, bteColUnit) = Trim(RsStock!unit_desc)
'                grid.TextMatrix(i, bteColPreMonth) = Format(RsStock!nm_premonth, gs_formatQty)
'                grid.TextMatrix(i, bteColReceipt) = Format(RsStock!nm_receipt, gs_formatQty)
'                grid.TextMatrix(i, bteColSupply) = Format(RsStock!nm_supply, gs_formatQty)
'                grid.TextMatrix(i, bteColLossReject) = Format(RsStock!nm_lossreject, gs_formatQty)
'                grid.TextMatrix(i, bteColEnd) = Format(RsStock!nm_current, gs_formatQty)
'                grid.TextMatrix(i, bteColInventory) = IIf(RsStock!nm_inventory = Null, "", Format(RsStock!nm_inventory, gs_formatQty))
'                grid.TextMatrix(i, bteColPrice) = Format(RsStock!Price, gs_formatPriceIDR)
'                grid.TextMatrix(i, bteColAmount) = Format(RsStock!Price * Val(RsStock!nm_current & ""), gs_formatPriceIDR)
'                grid.TextMatrix(i, bteColInTransit) = Format(RsStock!in_transit, gs_formatQty)
'                grid.TextMatrix(i, bteColInTransitAmount) = Format(RsStock!Price * Val(RsStock!in_transit & ""), gs_formatPriceIDR)
'                RsStock.MoveNext
'            Loop
'
' End Select
'
    i = 0
    Do While Not RsStock.EOF
        i = i + 1
        grid.AddItem i
        grid.TextMatrix(i, bteColProdCod) = Trim(RsStock!Item_Code)
        grid.TextMatrix(i, bteColPartNo) = Trim(RsStock!MakerItem_Code)
        grid.TextMatrix(i, bteColDesc) = uf_GetItemDescription(Trim(RsStock!Item_Code))

        grid.TextMatrix(i, bteColAddress) = Trim(RsStock!Location1)
        grid.TextMatrix(i, bteColAddress2) = Trim(RsStock!Location2)
        grid.TextMatrix(i, bteColAddress3) = Trim(RsStock!Location3)

        grid.TextMatrix(i, bteColGroupDesc) = Trim(RsStock!groupdesc & "")
        grid.TextMatrix(i, bteColUnit) = Trim(RsStock!Unit_Desc)
        grid.TextMatrix(i, bteColPreMonth) = Format(RsStock!pre_months, gs_formatQty)
        grid.TextMatrix(i, bteColReceipt) = Format(RsStock!receipts, gs_formatQty)
        grid.TextMatrix(i, bteColSupply) = Format(RsStock!supplys, gs_formatQty)
        grid.TextMatrix(i, bteColLossReject) = Format(RsStock!lossrejects, gs_formatQty)
        grid.TextMatrix(i, bteColEnd) = Format(RsStock!Currents, gs_formatQty)
        grid.TextMatrix(i, bteColInventory) = IIf(RsStock!inventoryS = Null, "", Format(RsStock!inventoryS, gs_formatQty))
        grid.TextMatrix(i, bteColPrice) = Format(RsStock!Price, gs_formatPriceIDR)
        grid.TextMatrix(i, bteColAmount) = Format(RsStock!Price * IIf(RsStock!DiffClosing <= 0, Val(RsStock!inventoryS) & "", Val(RsStock!Currents & "")), gs_formatPriceIDR)
        grid.TextMatrix(i, bteColInTransit) = Format(RsStock!in_transit, gs_formatQty)
        grid.TextMatrix(i, bteColInTransitAmount) = Format(RsStock!Price * Val(RsStock!in_transit & ""), gs_formatPriceIDR)
        RsStock.MoveNext
    Loop

sql = "select"
End Sub


Private Sub Grid_DblClick()
    If grid.Row = 1 Then
        If bytSort = 0 Then
            grid.Sort = flexSortGenericDescending
            bytSort = 1
        Else
            grid.Sort = flexSortGenericAscending
            bytSort = 0
        End If
    End If
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Sub toExcel()
Dim xlapp As New Excel.application
Dim rsCek As New Recordset, Idx As Long, tempi As String, tempcust As String
Dim bolcust As Boolean, bolinv As Boolean
Dim rsCompany As New Recordset, sql_plus As String, sqlP As String
Dim sqlControl As String, RsInvControl As New ADODB.Recordset

CboLocationCD = Trim(CboLocationCD)
If CboLocationCD.MatchFound Then
    LblLocationName = Trim(CboLocationCD.Column(1))
Else
    LblErrMsg = DisplayMsg(4018)
    Exit Sub
End If

        

sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year,inventory_month"

If RsInvControl.State <> adStateClosed Then RsInvControl.Close
RsInvControl.Open sqlControl, Db, adOpenKeyset, adLockOptimistic

If RsInvControl.EOF = True And RsInvControl.BOF = True Then LblErrMsg = "Inventory Stock hasn't been closed !": Exit Sub
RsInvControl.MoveLast
     
If DateDiff("M", RsInvControl!Inventory_Year & "-" & Format(RsInvControl!Inventory_Month, "00") & "-01", Year(DMonth) & "-" & Format(Month(DMonth), "00") & "-01") = 0 Then
    sqlP = "lm_premonth PreMonth, LM_Receipt Receipt, LM_supply Supply, LM_lossReject LossReject, Lm_Current Currents, LM_Inventory Inventory "
ElseIf DateDiff("M", RsInvControl!Inventory_Year & "-" & Format(RsInvControl!Inventory_Month, "00") & "-01", Year(DMonth) & "-" & Format(Month(DMonth), "00") & "-01") = 1 Then
    sqlP = "TM_premonth PreMonth, TM_Receipt Receipt, TM_supply Supply, TM_lossReject LossReject, TM_Current Currents, TM_Inventory Inventory "
Else
    sqlP = "NM_premonth PreMonth, NM_Receipt Receipt, NM_supply Supply, NM_lossReject LossReject, NM_Current Currents, NM_Inventory Inventory "
End If

sql = "select gc.description groupDesc, uc.description Unit_Desc,sm.item_code, " & _
        " descriptions=case isnull(im.sheetcoil_cls,0) when 0 then  im.item_name else rtrim(im.item_name) + ' (' + rtrim(sh.description) + ', T' + cast(im.thickness as varchar(15)) + ' x W' + cast(im.width as varchar(15)) + ' x L' + cast(im.length as varchar(15)) + ')'  end " & _
          ",address, wh_name, makeritem_code, " & sqlP & _
          " from stock_master sm left join warehouse_master wm on " & _
          " sm.warehouse_code = wm.wh_code " & _
          " left join item_master im on sm.item_code=im.item_code " & _
          " left join sheetcoil_cls sh on im.sheetcoil_cls = sh.sheetcoil_cls " & _
            " left join group_cls gc on im.group_cls=gc.group_cls " & _
            " left join unit_cls uc on im.unit_cls=uc.unit_cls " & _
          " where warehouse_code='" & Trim(CboLocationCD) & "' order by  warehouse_code,sm.item_Code"


If rsCek.State <> adStateClosed Then rsCek.Close
rsCek.CursorLocation = adUseClient
rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic


If Not rsCek.EOF Then
Screen.MousePointer = vbHourglass
With xlapp

    sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax  From company_profile "
    If rsCompany.State <> adStateClosed Then rsCompany.Close
    rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
    If rsCompany.EOF Then Screen.MousePointer = vbDefault: Exit Sub
    .Workbooks.Add
    
    .Range("a2", "l2").Merge
    .Range("a2") = rsCompany!company_name
    .Range("a3", "l3").Merge
    .Range("a3") = rsCompany!address1 & " " & rsCompany!address2 & " " & rsCompany!City & " " & rsCompany!Province & " " & rsCompany!postal_code
    .Range("a4", "l4").Merge
    .Range("a4") = "Phone: " & rsCompany!phone1 & " " & rsCompany!phone2 & " Fax: " & rsCompany!fax
    
    .Range("a6") = "Inventory Report"
    .Range("b6") = ""
    .Range("a6", "b6").Merge
    .Range("a6").horizontalAlignment = xlLeft
    .Range("a7") = "Warehouse Code"
    .Range("b7", "l7").Merge
    .Range("b7") = ": " & Trim(CboLocationCD.Text) & " / " & Trim(LblLocationName)
    .Range("B7").horizontalAlignment = xlLeft
    .Range("a8") = "Date(Month)"
    .Range("b8") = ": " & Format(DMonth, "MMMM YYYY")
    
    
    Idx = 10
       
    Do While Not rsCek.EOF
        If Idx = 10 Then
            .Range("a" & Idx) = "Product Code"
            .Range("b" & Idx) = "Part Number"
            .Range("c" & Idx) = "Description"
            .Range("d" & Idx) = "Group Desc"
            .Range("e" & Idx) = "Unit"
            .Range("f" & Idx) = "Pre Month Stock"
            .Range("g" & Idx) = "Receipt Total"
            .Range("h" & Idx) = "Supply Total"
            .Range("i" & Idx) = "Loss/Reject"
            .Range("j" & Idx) = "Difference"
            .Range("k" & Idx) = "End of Month Stock"
            .Range("l" & Idx) = "Inventory"
            .Range("a" & Idx, "l" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range("a" & Idx, "l" & Idx).Borders(xlEdgeBottom).LineStyle = xlDouble
            Idx = Idx + 1
        End If
        
        Idx = Idx
        'Content
        .Range("a" & Idx) = Trim(rsCek!Item_Code)
        .Range("b" & Idx) = Trim(rsCek!MakerItem_Code)
        .Range("c" & Idx) = Trim(rsCek!descriptions)
        .Range("d" & Idx) = Trim(rsCek!groupdesc & "")
        .Range("e" & Idx) = Trim(rsCek!Unit_Desc)
        .Range("f" & Idx) = Format(rsCek!Premonth, gs_formatQty)
        .Range("g" & Idx) = Format(rsCek!Receipt, gs_formatQty)
        .Range("h" & Idx) = Format(rsCek!supply, gs_formatQty)
        .Range("i" & Idx) = Format(rsCek!LossReject, gs_formatQty)
        .Range("j" & Idx) = Format(rsCek!Currents - rsCek!inventory, gs_formatQty)
        .Range("k" & Idx) = Format(rsCek!Currents, gs_formatQty)
        .Range("l" & Idx) = Format(rsCek!inventory, gs_formatQty)
        Idx = Idx + 1
        rsCek.MoveNext
    Loop
        
    .Range("a1", "l" & Idx + 3).Columns.Font.Name = "Arial"
    .Range("a1", "l" & Idx + 3).Columns.Font.Size = 8
    
    .Range("a2", "l2").Columns.Font.Name = "Arial"
    .Range("a2", "l2").Columns.Font.Size = "10"
    .Range("a2", "l2").Columns.Font.Bold = True
    .Range("a2", "l4").horizontalAlignment = xlCenter
    .Range("a6", "b6").Columns.Font.Bold = True
    
    .Visible = True
    .ActiveSheet.PageSetup.PaperSize = xlPaperA4
    .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
    .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.4)
    .ActiveSheet.PageSetup.Orientation = xlLandscape
    .Range("A:l").Columns.AutoFit
    .Range("C11:I" & Idx).Select
    .Selection.NumberFormat = gs_formatQty
    .Range("A1").Select
    .WindowState = xlMaximized
End With
Else
    LblPesan = DisplayMsg(4006)
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub PrintReport()

    Dim j As Integer
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsRpt As New ADODB.Recordset
    Dim rsrpt2 As New ADODB.Recordset
    Dim Rpt As New FrmRpt3
    Dim sqlControl As String
    Dim RsInvControl As New ADODB.Recordset
    Dim intDiffClosing As Integer
    
    sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year,inventory_month"
    
    If RsInvControl.State <> adStateClosed Then RsInvControl.Close
    RsInvControl.Open sqlControl, Db, adOpenKeyset, adLockOptimistic
    
    If RsInvControl.EOF = True And RsInvControl.BOF = True Then LblErrMsg = "Inventory Stock hasn't been closed !": Exit Sub
    RsInvControl.MoveLast
    
    j = 0
    For i = 0 To CboLocationCD.ListCount - 1
        If UCase(Trim(CboLocationCD)) = UCase(Trim(CboLocationCD.List(i, 0))) Then
            CboLocationCD = Trim(CboLocationCD.List(i, 0))
            LblLocationName = Trim(CboLocationCD.List(i, 1))
            j = 1: Exit For
        End If
    Next
    If j = 0 Then LblErrMsg = DisplayMsg(4018): Exit Sub  '"Invalid warehouse code !"
    
'    LblErrMsg = up_ValidateDateRange(DMonth.Value, False)
'    If Trim(LblErrMsg) <> "" Then Exit Sub
    
    LblErrMsg = ""
    Me.MousePointer = vbHourglass
    
    dtMPList = DMonth.Value
    datePiList = Format(DMonth.Value, "MMM yyyy")
    
'    sql = "select rtrim(sm.item_code) item_code, rtrim(makeritem_code) makeritem_code, rtrim(item_name) item_name, " & _
'          vbLf & "rtrim(sm.warehouse_code) warehouse_code, rtrim(wh_name) wh_name, rtrim(address) address, " & _
'          vbLf & "rtrim(gc.description) group_desc, rtrim(uc.description) unit_desc, " & _
'          vbLf & "sm.lm_premonth, sm.lm_receipt, sm.lm_supply, sm.lm_lossreject, sm.lm_current, isnull(sm.lm_inventory,0) lm_inventory, " & _
'          vbLf & "sm.tm_premonth, sm.tm_receipt, sm.tm_supply, sm.tm_lossreject, sm.tm_current, sm.tm_inventory, " & _
'          vbLf & "sm.nm_premonth, sm.nm_receipt, sm.nm_supply, sm.nm_lossreject, sm.nm_current, sm.nm_inventory, " & _
'          vbLf & "isnull(ip.inventory_price, 0) price, isnull(it.in_transit, 0) in_transit, " & _
'          vbLf & "descriptions = case isnull(im.sheetcoil_cls,0) when 0 then  im.item_name else rtrim(im.item_name) + ' (' + rtrim(sh.description) + ', T' + cast(im.thickness as varchar(15)) + ' x W' + cast(im.width as varchar(15)) + ' x L' + cast(im.length as varchar(15)) + ')'  end, " & _
'          vbLf & "case when warehouse_code = 'FG' then " & _
'          vbLf & "case when left(sm.item_code, 1) = 'E' then 'Export' else 'Local' end else " & _
'          vbLf & "case when isnull(tm.country_cls, 0) = 0 then 'Local' else 'Import' end " & _
'          vbLf & "end local, im.material_cls "
'
'    sql = sql & _
'          vbLf & "from stock_master sm " & _
'          vbLf & "left join warehouse_master wm on sm.warehouse_code=wm.wh_code " & _
'          vbLf & "left join item_master im on im.item_code=sm.item_code " & _
'          vbLf & "left join trade_master tm on im.supplier_code = tm.trade_code " & _
'          vbLf & "left join group_cls gc on im.group_cls=gc.group_cls " & _
'          vbLf & "left join unit_cls uc on im.unit_cls=uc.unit_cls " & _
'          vbLf & "left join sheetcoil_cls sh on im.sheetcoil_cls = sh.sheetcoil_cls " & _
'          vbLf & "left join ( " & _
'          vbLf & "select item_code, inventory_price From inventory_price " & _
'          vbLf & "where inventory_year = '" & DMonth.year & "' and inventory_month = '" & DMonth.month & "' " & _
'          vbLf & "and duty_status in ('0','2','3') " & _
'          vbLf & ") ip on sm.item_code = ip.item_code " & _
'          vbLf & "left join ( " & _
'          vbLf & "select whcode, item_code, sum(qty) in_transit " & _
'          vbLf & "from packing_master pm " & _
'          vbLf & "inner join packing_detail pd on pm.packing_no = pd.packing_no " & _
'          vbLf & "Where DateDiff(month, pm.stuffing_date, etd) > 0 " & _
'          vbLf & "and year(pm.stuffing_date) = '" & DMonth.year & "' and month(pm.stuffing_date) = '" & DMonth.month & "' " & _
'          vbLf & "group by whcode, item_code " & _
'          vbLf & ") it on sm.item_code = it.item_code and sm.warehouse_code = it.whcode " & _
'          vbLf & "where warehouse_code = '" & CboLocationCD & "' " & _
'          vbLf & "order by sm.warehouse_code, local desc, im.material_cls, sm.item_code "
    
    
    ' --------------------------------------------------
    ' Update Query Include Valuation History
    ' --------------------------------------------------
    
    sql = " DECLARE @LastClosing AS DATETIME " & vbCrLf & _
                " DECLARE @GetDiff AS INTEGER " & vbCrLf & _
                " DECLARE @SelectPeriod AS DATETIME " & vbCrLf & _
                " DECLARE @Location AS VARCHAR(30) " & vbCrLf & _
                "  " & vbCrLf & _
                "  " & vbCrLf & _
                "  " & vbCrLf & _
                " SET @SelectPeriod='" & Format(DMonth, "yyyy-MM-01") & "' " & vbCrLf & _
                " SET @Location='" & Trim(CboLocationCD) & "' " & vbCrLf & _
                "  " & vbCrLf & _
                " SET @LastClosing= " & vbCrLf & _
                "                           (SELECT CONVERT(DATETIME,CAST(Inventory_Year AS VARCHAR(4)) +'-' +CAST(Inventory_Month AS varchar(2)) +'-01') " & vbCrLf & _
                "                               FROM Inventory_Control " & vbCrLf
    
    sql = sql + "                                   WHERE Inventory_Year=(SELECT MAX(Inventory_Year) FROM Inventory_Control) " & vbCrLf & _
                "                                       AND Inventory_Month=(SELECT MAX(Inventory_Month) FROM Inventory_Control " & vbCrLf & _
                "                                                                               WHERE Inventory_Year=(SELECT MAX(Inventory_Year) FROM Inventory_Control)) " & vbCrLf & _
                "                           )                                                    " & vbCrLf & _
                "  " & vbCrLf & _
                " SET @GetDiff=DATEDIFF(M,@LastClosing,@SelectPeriod) " & vbCrLf & _
                "  " & vbCrLf & _
                " SELECT  RTRIM(sm.item_code) item_code , " & vbCrLf & _
                "         RTRIM(makeritem_code) makeritem_code , " & vbCrLf & _
                "         RTRIM(item_name) item_name , " & vbCrLf & _
                "         RTRIM(sm.warehouse_code) warehouse_code , " & vbCrLf
    
    sql = sql + "         RTRIM(wh_name) wh_name , " & vbCrLf & _
                "         RTRIM(address) address , " & vbCrLf & _
                "         '' Location1 , " & vbCrLf & _
                "         '' Location2 , " & vbCrLf & _
                "         '' Location3 , " & vbCrLf & _
                "         RTRIM(gc.description) group_desc , " & vbCrLf & _
                "         RTRIM(uc.description) unit_desc , " & vbCrLf & _
                "         sm.Pre_MonthS lm_premonth , " & vbCrLf & _
                "         sm.ReceiptS lm_receipt , " & vbCrLf & _
                "         sm.SupplyS lm_supply , " & vbCrLf & _
                "         sm.LossRejectS lm_lossreject , " & vbCrLf & _
                "         sm.CurrentS lm_current , " & vbCrLf & _
                "         ISNULL(sm.InventoryS, 0) lm_inventory , " & vbCrLf & _
                "         sm.Pre_MonthS tm_premonth , " & vbCrLf
    
    sql = sql + "         sm.ReceiptS tm_receipt , " & vbCrLf & _
                "         sm.SupplyS tm_supply , " & vbCrLf & _
                "         sm.LossRejectS tm_lossreject , " & vbCrLf & _
                "         sm.CurrentS tm_current , " & vbCrLf & _
                "         sm.InventoryS tm_inventory , " & vbCrLf & _
                "         sm.Pre_MonthS nm_premonth , " & vbCrLf & _
                "         sm.ReceiptS nm_receipt , " & vbCrLf & _
                "         sm.SupplyS nm_supply , " & vbCrLf & _
                "         sm.LossRejectS nm_lossreject , " & vbCrLf & _
                "         sm.CurrentS nm_current , " & vbCrLf & _
                "         sm.InventoryS nm_inventory , " & vbCrLf
    
    sql = sql + "         ISNULL(ip.inventory_price, 0) price , " & vbCrLf & _
                "         ISNULL(it.in_transit, 0) in_transit , " & vbCrLf & _
                "         descriptions = CASE ISNULL(im.sheetcoil_cls, 0) " & vbCrLf & _
                "                          WHEN 0 THEN im.item_name " & vbCrLf & _
                "                          ELSE RTRIM(im.item_name) + ' (' " & vbCrLf & _
                "                               + RTRIM(sh.description) + ', T' " & vbCrLf & _
                "                               + CAST(im.thickness AS VARCHAR(15)) + ' x W' " & vbCrLf & _
                "                               + CAST(im.width AS VARCHAR(15)) + ' x L' " & vbCrLf & _
                "                               + CAST(im.length AS VARCHAR(15)) + ')' " & vbCrLf & _
                "                        END , " & vbCrLf & _
                "         " & vbCrLf
    
    sql = sql + "              " & vbCrLf & _
                "               " & vbCrLf & _
                "              " & vbCrLf & _
                "              " & vbCrLf & _
                "              " & vbCrLf & _
                "                    " & vbCrLf & _
                "         '' local , " & vbCrLf & _
                "         im.material_cls " & vbCrLf & _
                "         ,@GetDiff DiffClosing " & vbLf & _
                " FROM     " & vbCrLf & _
                "          (  " & vbCrLf & _
                "           SELECT warehouse_code, item_code, " & vbCrLf
    
    sql = sql + "                        SUM(Pre_MonthS) Pre_MonthS, SUM(ReceiptS) ReceiptS,  " & vbCrLf & _
                "                        SUM(SupplyS) SupplyS, SUM(LossRejectS) LossRejectS,  " & vbCrLf & _
                "                        SUM(CurrentS) CurrentS, SUM(InventoryS) InventoryS  " & vbCrLf & _
                "           FROM             " & vbCrLf & _
                "           ( " & vbCrLf & _
                "               SELECT warehouse_code, item_code, " & vbCrLf & _
                "                    COALESCE(CASE WHEN @GetDiff=0 THEN LM_PreMonth  " & vbCrLf & _
                "                             WHEN @GetDiff=1 THEN TM_PreMonth  " & vbCrLf & _
                "                             WHEN @GetDiff=2 THEN NM_PreMonth END,0) Pre_MonthS,  " & vbCrLf & _
                "                    COALESCE(CASE WHEN @GetDiff=0 THEN LM_Receipt  " & vbCrLf & _
                "                             WHEN @GetDiff=1 THEN TM_Receipt  " & vbCrLf
    
    sql = sql + "                             WHEN @GetDiff=2 THEN NM_Receipt END,0) ReceiptS,  " & vbCrLf & _
                "                    COALESCE(CASE WHEN @GetDiff=0 THEN LM_Supply  " & vbCrLf & _
                "                             WHEN @GetDiff=1 THEN TM_Supply  " & vbCrLf & _
                "                             WHEN @GetDiff=2 THEN NM_Supply END,0) SupplyS,  " & vbCrLf & _
                "                    COALESCE(CASE WHEN @GetDiff=0 THEN LM_LossReject  " & vbCrLf & _
                "                             WHEN @GetDiff=1 THEN TM_LossReject  " & vbCrLf & _
                "                             WHEN @GetDiff=2 THEN NM_LossReject END,0) LossRejectS,  " & vbCrLf & _
                "                    COALESCE(CASE WHEN @GetDiff=0 THEN LM_Current  " & vbCrLf & _
                "                             WHEN @GetDiff=1 THEN TM_Current  " & vbCrLf & _
                "                             WHEN @GetDiff=2 THEN NM_Current END,0) CurrentS,  " & vbCrLf & _
                "                    COALESCE(CASE WHEN @GetDiff=0 THEN LM_Inventory  " & vbCrLf
    
    sql = sql + "                             WHEN @GetDiff=1 THEN TM_Inventory  " & vbCrLf & _
                "                             WHEN @GetDiff=2 THEN NM_Inventory END,0) InventoryS  " & vbCrLf & _
                "               FROM Stock_Master " & vbCrLf & _
                "                    " & vbCrLf & _
                "                UNION ALL  " & vbCrLf & _
                "    " & vbCrLf & _
                "                SELECT Warehouse_Code, Item_code, PreMonth, Receipt, Supply, LossReject, [Current], COALESCE(Inventory,[Current])  " & vbCrLf & _
                "                    FROM Stock_history  " & vbCrLf & _
                "                        WHERE Warehouse_Code=@Location  " & vbCrLf & _
                "                                    AND Stock_Year=YEAR(@SelectPeriod)  " & vbCrLf & _
                "                                    AND Stock_Month=MONTH(@SelectPeriod)  " & vbCrLf
    
    sql = sql + "  " & vbCrLf & _
                "           ) Stock GROUP BY warehouse_code, item_code " & vbCrLf & _
                "          ) sm " & vbCrLf & _
                "  " & vbCrLf & _
                "         LEFT JOIN warehouse_master wm ON sm.warehouse_code = wm.wh_code " & vbCrLf & _
                "         LEFT JOIN item_master im ON im.item_code = sm.item_code " & vbCrLf & _
                "         LEFT JOIN trade_master tm ON im.supplier_code = tm.trade_code " & vbCrLf & _
                "         LEFT JOIN group_cls gc ON im.group_cls = gc.group_cls " & vbCrLf & _
                "         LEFT JOIN unit_cls uc ON im.unit_cls = uc.unit_cls " & vbCrLf & _
                "         LEFT JOIN sheetcoil_cls sh ON im.sheetcoil_cls = sh.sheetcoil_cls " & vbCrLf & _
                "         LEFT JOIN ( SELECT  item_code , " & vbCrLf
    
    sql = sql + "                             inventory_price " & vbCrLf & _
                "                     FROM   " & vbCrLf & _
                "                       (SELECT * From Inventory_Price " & vbCrLf & _
                "                           UNION   ALL " & vbCrLf & _
                "                         SELECT * From InventoryPrice_History " & vbCrLf & _
                "                       ) Inventory_Price                            " & vbCrLf & _
                "                     WHERE   inventory_year = YEAR(@SelectPeriod) " & vbCrLf & _
                "                             AND inventory_month = MONTH(@SelectPeriod) " & vbCrLf & _
                "                             AND duty_status IN ( '0', '2', '3' )  " & vbCrLf & _
                "                   ) ip ON sm.item_code = ip.item_code " & vbCrLf & _
                "         LEFT JOIN ( SELECT  whcode , " & vbCrLf
    
    sql = sql + "                             item_code , " & vbCrLf & _
                "                             SUM(qty) in_transit " & vbCrLf & _
                "                     FROM    packing_master pm " & vbCrLf & _
                "                             INNER JOIN packing_detail pd ON pm.packing_no = pd.packing_no " & vbCrLf & _
                "                     WHERE   DATEDIFF(month, pm.stuffing_date, etd) > 0 " & vbCrLf & _
                "                             AND YEAR(pm.stuffing_date) = YEAR(@SelectPeriod) " & vbCrLf & _
                "                             AND MONTH(pm.stuffing_date) = MONTH(@SelectPeriod) " & vbCrLf & _
                "                     GROUP BY whcode , " & vbCrLf & _
                "                             item_code " & vbCrLf & _
                "                   ) it ON sm.item_code = it.item_code " & vbCrLf & _
                "                           AND sm.warehouse_code = it.whcode " & vbCrLf
    
    sql = sql + " WHERE   warehouse_code = @Location " & vbCrLf & _
                " ORDER BY sm.warehouse_code , " & vbCrLf & _
                "         local DESC , " & vbCrLf & _
                "         im.material_cls , " & vbCrLf & _
                "         sm.item_code  " & vbCrLf & _
                "  "
    ' --------------------------------------------------
    
    
    If rsRpt.State <> adStateClosed Then rsRpt.Close
    rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
    If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
        
    Set report = application.OpenReport(App.path & "\Reports\rpt_pi_reportwh.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    
    'intDiffClosing = up_GetDateRange(DMonth.Value)
    intDiffClosing = 1
    
    report.ReportTitle = "Valuation Price Report Per Warehouse"
    report.FormulaFields(1).Text = "'" & datePiList & "'"
    report.FormulaFields(5).Text = "" & gi_decimalDigitQty & ""
    report.FormulaFields(6).Text = "" & gi_decimalDigitAmountIDR & ""
    report.FormulaFields(11).Text = "" & intDiffClosing & ""
    
'    sqlprint2 = "select a.local, a.material_cls, sum(a.lm_inventory) lm_inventory, sum(a.tm_current) tm_current, sum(a.nm_current) nm_current, sum(a.in_transit) in_transit " & _
'        vbLf & "from( " & _
'            vbLf & "select case when warehouse_code = 'FG' " & _
'                vbLf & "then case when left(sm.item_code, 1) = 'E' then 'Export' else 'Local' end " & _
'                vbLf & "else case when isnull(tm.country_cls, 0) = 0 then 'Local' else 'Import' end " & _
'            vbLf & "end local, im.material_cls, " & _
'            vbLf & "isnull(sm.lm_inventory, 0) * isnull(ip.inventory_price, 0) lm_inventory, " & _
'            vbLf & "isnull(sm.tm_current, 0) * isnull(ip.inventory_price, 0) tm_current, " & _
'            vbLf & "isnull(sm.nm_current, 0) * isnull(ip.inventory_price, 0) nm_current, " & _
'            vbLf & "isnull(it.in_transit, 0) * isnull(ip.inventory_price, 0) in_transit "
'
'    sqlprint2 = sqlprint2 & _
'            vbLf & "from stock_master sm " & _
'            vbLf & "left join item_master im on im.item_code=sm.item_code " & _
'            vbLf & "left join trade_master tm on im.supplier_code = tm.trade_code " & _
'            vbLf & "left join ( " & _
'                vbLf & "select item_code, inventory_price From inventory_price " & _
'                vbLf & "where inventory_year = '" & DMonth.year & "' and inventory_month = '" & DMonth.month & "' " & _
'                vbLf & "and duty_status in ('0', '2', '3') " & _
'            vbLf & ") ip on sm.item_code = ip.item_code " & _
'            vbLf & "left join ( " & _
'                vbLf & "select whcode, item_code, sum(qty) in_transit " & _
'                vbLf & "from packing_master pm " & _
'                vbLf & "inner join packing_detail pd on pm.packing_no = pd.packing_no " & _
'                vbLf & "Where DateDiff(month, pm.stuffing_date, etd) > 0 " & _
'                vbLf & "and year(pm.stuffing_date) = '" & DMonth.year & "' and month(pm.stuffing_date) = '" & DMonth.month & "' " & _
'                vbLf & "group by whcode, item_code " & _
'            vbLf & ") it on sm.item_code = it.item_code and sm.warehouse_code = it.whcode " & _
'            vbLf & "where warehouse_code = '" & CboLocationCD & "' " & _
'        vbLf & ") a group by a.local, a.material_cls " & _
'        vbLf & "order by a.local desc, a.material_cls "
    
    ' --------------------------------------------------
    ' Update Query Include Valuation History
    ' --------------------------------------------------
        sqlprint2 = " DECLARE @LastClosing AS DATETIME " & vbCrLf & _
                    " DECLARE @GetDiff AS INTEGER " & vbCrLf & _
                    " DECLARE @SelectPeriod AS DATETIME " & vbCrLf & _
                    " DECLARE @Location AS VARCHAR(30) " & vbCrLf & _
                    "  " & vbCrLf & _
                    " SET @SelectPeriod='" & Format(DMonth, "yyyy-MM-01") & "' " & vbCrLf & _
                    " SET @LastClosing= " & vbCrLf & _
                    "                           (SELECT CONVERT(DATETIME,CAST(Inventory_Year AS VARCHAR(4)) +'-' +CAST(Inventory_Month AS varchar(2)) +'-01') " & vbCrLf & _
                    "                               FROM Inventory_Control " & vbCrLf & _
                    "                                   WHERE Inventory_Year=(SELECT MAX(Inventory_Year) FROM Inventory_Control) " & vbCrLf & _
                    "                                       AND Inventory_Month=(SELECT MAX(Inventory_Month) FROM Inventory_Control " & vbCrLf
        
        sqlprint2 = sqlprint2 + "                                                                               WHERE Inventory_Year=(SELECT MAX(Inventory_Year) FROM Inventory_Control)) " & vbCrLf & _
                    "                           )                                                    " & vbCrLf & _
                    "  " & vbCrLf & _
                    " SET @GetDiff=DATEDIFF(M,@LastClosing,@SelectPeriod) " & vbCrLf & _
                    "  " & vbCrLf & _
                    " SELECT  " & vbCrLf & _
                    "         a.local , " & vbCrLf & _
                    "         a.material_cls , " & vbCrLf & _
                    "         SUM(a.lm_inventory) lm_inventory , " & vbCrLf & _
                    "         SUM(a.tm_current) tm_current , " & vbCrLf & _
                    "         SUM(a.nm_current) nm_current , " & vbCrLf
        
        sqlprint2 = sqlprint2 + "         SUM(a.in_transit) in_transit " & vbCrLf & _
                    "          ,@GetDiff DiffClosing " & vbLf & _
                    " FROM    ( SELECT    warehouse_code , " & vbCrLf & _
                    "                     CASE WHEN warehouse_code = 'FG' " & vbCrLf & _
                    "                          THEN CASE WHEN LEFT(sm.item_code, 1) = 'E' " & vbCrLf & _
                    "                                    THEN 'Export' " & vbCrLf & _
                    "                                    ELSE 'Local' " & vbCrLf & _
                    "                               END " & vbCrLf & _
                    "                          ELSE CASE WHEN ISNULL(tm.country_cls, 0) = 0 " & vbCrLf & _
                    "                                    THEN 'Local' " & vbCrLf & _
                    "                                    ELSE 'Import' " & vbCrLf & _
                    "                               END " & vbCrLf
        
        sqlprint2 = sqlprint2 + "                     END local , " & vbCrLf & _
                    "                     im.material_cls , " & vbCrLf & _
                    "                     ISNULL(sm.lm_inventory, 0) * ISNULL(ip.inventory_price, 0) lm_inventory , " & vbCrLf & _
                    "                     ISNULL(sm.tm_current, 0) * ISNULL(ip.inventory_price, 0) tm_current , " & vbCrLf & _
                    "                     ISNULL(sm.nm_current, 0) * ISNULL(ip.inventory_price, 0) nm_current , " & vbCrLf & _
                    "                     ISNULL(it.in_transit, 0) * ISNULL(ip.inventory_price, 0) in_transit " & vbCrLf & _
                    "           FROM       " & vbCrLf & _
                    "          (  " & vbCrLf & _
                    "           SELECT warehouse_code, item_code, " & vbCrLf & _
                    "                       SUM(LM_inventory) LM_inventory, " & vbCrLf & _
                    "                       SUM(TM_Current) TM_Current, " & vbCrLf
        
        sqlprint2 = sqlprint2 + "                       SUM(NM_Current) NM_Current " & vbCrLf & _
                    "           FROM             " & vbCrLf & _
                    "           ( " & vbCrLf & _
                    "               SELECT warehouse_code, item_code, " & vbCrLf & _
                    "                   COALESCE(CASE WHEN @GetDiff=0 THEN 0 " & vbCrLf & _
                    "                            WHEN @GetDiff=1 THEN LM_Inventory " & vbCrLf & _
                    "                            WHEN @GetDiff=2 THEN TM_Current END,0) LM_inventory, " & vbCrLf & _
                    "                   COALESCE(CASE WHEN @GetDiff=0 THEN LM_inventory " & vbCrLf & _
                    "                            WHEN @GetDiff=1 THEN TM_Current " & vbCrLf & _
                    "                            WHEN @GetDiff=2 THEN NM_Current END,0) TM_Current, " & vbCrLf & _
                    "                   COALESCE(CASE WHEN @GetDiff=0 THEN TM_Current " & vbCrLf
        
        sqlprint2 = sqlprint2 + "                            WHEN @GetDiff=1 THEN NM_Current " & vbCrLf & _
                    "                            WHEN @GetDiff=2 THEN NM_Current END,0) NM_Current " & vbCrLf & _
                    "               FROM Stock_Master " & vbCrLf & _
                    "                    " & vbCrLf & _
                    "               UNION ALL " & vbCrLf & _
                    "               -- Last Month " & vbCrLf & _
                    "               SELECT Warehouse_Code, Item_code,  COALESCE(Inventory,[Current]) LM_inventory, 0 TM_Inventory, 0 NM_Inventory " & vbCrLf & _
                    "                   FROM Stock_history " & vbCrLf & _
                    "                       WHERE Stock_Year=YEAR(@SelectPeriod) " & vbCrLf & _
                    "                                   AND Stock_Month=MONTH(@SelectPeriod)-1 " & vbCrLf & _
                    "               UNION ALL " & vbCrLf
        
        sqlprint2 = sqlprint2 + "               -- This Month " & vbCrLf & _
                    "               SELECT Warehouse_Code, Item_code,  0 LM_inventory, COALESCE(Inventory,[Current]) TM_Inventory, 0 NM_Inventory " & vbCrLf & _
                    "                   FROM Stock_history " & vbCrLf & _
                    "                       WHERE Stock_Year=YEAR(@SelectPeriod) " & vbCrLf & _
                    "                                   AND Stock_Month=MONTH(@SelectPeriod) " & vbCrLf & _
                    "               UNION ALL                                    " & vbCrLf & _
                    "               -- Next Month " & vbCrLf & _
                    "               SELECT Warehouse_Code, Item_code,  0 LM_inventory, 0 TM_Inventory, COALESCE(Inventory,[Current]) NM_Inventory " & vbCrLf & _
                    "                   FROM Stock_history " & vbCrLf & _
                    "                       WHERE Stock_Year=YEAR(@SelectPeriod) " & vbCrLf & _
                    "                                   AND Stock_Month=MONTH(@SelectPeriod)+1 " & vbCrLf
        
        sqlprint2 = sqlprint2 + "           ) Stock GROUP BY warehouse_code, item_code " & vbCrLf & _
                    "          ) sm " & vbCrLf & _
                    "                     LEFT JOIN item_master im ON im.item_code = sm.item_code " & vbCrLf & _
                    "                     LEFT JOIN trade_master tm ON im.supplier_code = tm.trade_code " & vbCrLf & _
                    "                     LEFT JOIN ( SELECT  item_code , " & vbCrLf & _
                    "                                         inventory_price " & vbCrLf & _
                    "                                 FROM   " & vbCrLf & _
                    "                                   (SELECT * From Inventory_Price  " & vbCrLf & _
                    "                                       UNION   ALL " & vbCrLf & _
                    "                                     SELECT * From InventoryPrice_History " & vbCrLf & _
                    "                                   ) Inventory_Price                            " & vbCrLf
         
        sqlprint2 = sqlprint2 + "                                 WHERE   inventory_year = YEAR(@SelectPeriod) " & vbCrLf & _
                    "                                         AND inventory_month = MONTH(@SelectPeriod) " & vbCrLf & _
                    "                                         AND duty_status IN ( '0', '2', '3' )  " & vbCrLf & _
                    "                               ) ip ON sm.item_code = ip.item_code " & vbCrLf & _
                    "                     LEFT JOIN ( SELECT  whcode , " & vbCrLf & _
                    "                                         item_code , " & vbCrLf & _
                    "                                         SUM(qty) in_transit " & vbCrLf & _
                    "                                 FROM    packing_master pm " & vbCrLf & _
                    "                                         INNER JOIN packing_detail pd ON pm.packing_no = pd.packing_no " & vbCrLf & _
                    "                                 WHERE   DATEDIFF(month, pm.stuffing_date, etd) > 0 " & vbCrLf & _
                    "                                         AND YEAR(pm.stuffing_date) = YEAR(@SelectPeriod) " & vbCrLf
        
        sqlprint2 = sqlprint2 + "                                         AND MONTH(pm.stuffing_date) = MONTH(@SelectPeriod) " & vbCrLf & _
                    "                                 GROUP BY whcode , " & vbCrLf & _
                    "                                         item_code " & vbCrLf & _
                    "                               ) it ON sm.item_code = it.item_code " & vbCrLf & _
                    "                                       AND sm.warehouse_code = it.whcode " & vbCrLf
                    
                    If UCase(CboLocationCD.Text) <> "ALL" Then
       sqlprint2 = sqlprint2 + "   WHERE sm.warehouse_code = '" & Trim(CboLocationCD.Text) & "' " & vbLf
                    End If
                    
       sqlprint2 = sqlprint2 & _
                    "         ) a " & vbCrLf & _
                    " GROUP BY  " & vbCrLf & _
                    "         a.local , " & vbCrLf & _
                    "         a.material_cls " & vbCrLf & _
                    " ORDER BY  " & vbCrLf & _
                    "         a.local DESC , " & vbCrLf
        
        sqlprint2 = sqlprint2 + "         a.material_cls ASC  " & vbCrLf & _
                    "  " & vbCrLf
    
    '---------------------------------------------------
    
    If rsrpt2.State <> adStateClosed Then rsrpt2.Close
    rsrpt2.Open sqlprint2, Db, adOpenDynamic, adLockOptimistic
    
    report.OpenSubreport("summary").Database.Tables(1).SetDataSource rsrpt2
    report.OpenSubreport("summary").FormulaFields(1).Text = "" & intDiffClosing & ""
    
    reportcode = "pireportwh"
    printorient = "2"
    sqlprint = sql
            
    Rpt.CRViewer1.ReportSource = report
    Rpt.CRViewer1.ViewReport
    Rpt.CRViewer1.Zoom 1
    
    Rpt.WindowState = 2
    Rpt.Show 1
    
    Me.MousePointer = vbDefault

End Sub

Private Sub PrintSummaryReport()

    Dim j As Integer
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsRpt As New ADODB.Recordset
    Dim rsrpt2 As New ADODB.Recordset
    Dim Rpt As New FrmRpt3
    Dim sqlControl As String
    Dim RsInvControl As New ADODB.Recordset
    Dim intDiffClosing As Integer
    
    sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year,inventory_month"
    
    If RsInvControl.State <> adStateClosed Then RsInvControl.Close
    RsInvControl.Open sqlControl, Db, adOpenKeyset, adLockOptimistic
    
    If RsInvControl.EOF = True And RsInvControl.BOF = True Then LblErrMsg = "Inventory Stock hasn't been closed !": Exit Sub
    RsInvControl.MoveLast
    
    j = 0
    For i = 0 To CboLocationCD.ListCount - 1
        If UCase(Trim(CboLocationCD)) = UCase(Trim(CboLocationCD.List(i, 0))) Then
            CboLocationCD = Trim(CboLocationCD.List(i, 0))
            LblLocationName = Trim(CboLocationCD.List(i, 1))
            j = 1: Exit For
        End If
    Next
    If j = 0 Then LblErrMsg = DisplayMsg(4018): Exit Sub  '"Invalid warehouse code !"
    
    
'    LblErrMsg = up_ValidateDateRange(DMonth.Value, False)
'    If Trim(LblErrMsg) <> "" Then Exit Sub
    
    LblErrMsg = ""
    Me.MousePointer = vbHourglass
    
    dtMPList = DMonth.Value
    datePiList = Format(DMonth.Value, "MMM yyyy")
    
'    sql = "select a.warehouse_code, a.local, a.material_cls, sum(a.lm_inventory) lm_inventory, sum(a.tm_current) tm_current, sum(a.nm_current) nm_current, sum(a.in_transit) in_transit " & _
'          vbLf & "from( " & _
'          vbLf & "select warehouse_code, case when warehouse_code = 'FG' " & _
'          vbLf & "then case when left(sm.item_code, 1) = 'E' then 'Export' else 'Local' end " & _
'          vbLf & "else case when isnull(tm.country_cls, 0) = 0 then 'Local' else 'Import' end " & _
'          vbLf & "end local, im.material_cls, " & _
'          vbLf & "isnull(sm.lm_inventory, 0) * isnull(ip.inventory_price, 0) lm_inventory, " & _
'          vbLf & "isnull(sm.tm_current, 0) * isnull(ip.inventory_price, 0) tm_current, " & _
'          vbLf & "isnull(sm.nm_current, 0) * isnull(ip.inventory_price, 0) nm_current, " & _
'          vbLf & "isnull(it.in_transit, 0) * isnull(ip.inventory_price, 0) in_transit "
'
'    sql = sql & _
'          vbLf & "from stock_master sm " & _
'          vbLf & "left join item_master im on im.item_code=sm.item_code " & _
'          vbLf & "left join trade_master tm on im.supplier_code = tm.trade_code " & _
'          vbLf & "left join ( " & _
'          vbLf & "select item_code, inventory_price From inventory_price " & _
'          vbLf & "where inventory_year = '" & DMonth.year & "' and inventory_month = '" & DMonth.month & "' " & _
'          vbLf & "and duty_status in ('0', '2', '3') " & _
'          vbLf & ") ip on sm.item_code = ip.item_code " & _
'          vbLf & "left join ( " & _
'          vbLf & "select whcode, item_code, sum(qty) in_transit " & _
'          vbLf & "from packing_master pm " & _
'          vbLf & "inner join packing_detail pd on pm.packing_no = pd.packing_no " & _
'          vbLf & "Where DateDiff(month, pm.stuffing_date, etd) > 0 " & _
'          vbLf & "and year(pm.stuffing_date) = '" & DMonth.year & "' and month(pm.stuffing_date) = '" & DMonth.month & "' " & _
'          vbLf & "group by whcode, item_code " & _
'          vbLf & ") it on sm.item_code = it.item_code and sm.warehouse_code = it.whcode " & _
'          vbLf & ") a group by a.warehouse_code, a.local, a.material_cls " & _
'          vbLf & "order by a.warehouse_code asc, a.local desc, a.material_cls asc "

' --------------------------------------------------
' Update Query Include Valuation History
' --------------------------------------------------
'    sql = " DECLARE @LastClosing AS DATETIME " & vbCrLf & _
'                " DECLARE @GetDiff AS INTEGER " & vbCrLf & _
'                " DECLARE @SelectPeriod AS DATETIME " & vbCrLf & _
'                " DECLARE @Location AS VARCHAR(30) " & vbCrLf & _
'                "  " & vbCrLf & _
'                " SET @SelectPeriod='" & Format(DMonth, "yyyy-MM-01") & "' " & vbCrLf & _
'                " SET @LastClosing= " & vbCrLf & _
'                "                           (SELECT CONVERT(DATETIME,CAST(Inventory_Year AS VARCHAR(4)) +'-' +CAST(Inventory_Month AS varchar(2)) +'-01') " & vbCrLf & _
'                "                               FROM Inventory_Control " & vbCrLf & _
'                "                                   WHERE Inventory_Year=(SELECT MAX(Inventory_Year) FROM Inventory_Control) " & vbCrLf & _
'                "                                       AND Inventory_Month=(SELECT MAX(Inventory_Month) FROM Inventory_Control " & vbCrLf
'
'    sql = sql + "                                                                               WHERE Inventory_Year=(SELECT MAX(Inventory_Year) FROM Inventory_Control)) " & vbCrLf & _
'                "                           )                                                    " & vbCrLf & _
'                "  " & vbCrLf & _
'                " SET @GetDiff=DATEDIFF(M,@LastClosing,@SelectPeriod) " & vbCrLf & _
'                "  " & vbCrLf & _
'                " SELECT  a.warehouse_code , " & vbCrLf & _
'                "         a.local , " & vbCrLf & _
'                "         a.material_cls , " & vbCrLf & _
'                "         SUM(a.lm_inventory) lm_inventory , " & vbCrLf & _
'                "         SUM(a.tm_current) tm_current , " & vbCrLf & _
'                "         SUM(a.nm_current) nm_current , " & vbCrLf
'
'    sql = sql + "         SUM(a.in_transit) in_transit " & vbCrLf & _
'                " FROM    ( SELECT    warehouse_code , " & vbCrLf & _
'                "                     CASE WHEN warehouse_code = 'FG' " & vbCrLf & _
'                "                          THEN CASE WHEN LEFT(sm.item_code, 1) = 'E' " & vbCrLf & _
'                "                                    THEN 'Export' " & vbCrLf & _
'                "                                    ELSE 'Local' " & vbCrLf & _
'                "                               END " & vbCrLf & _
'                "                          ELSE CASE WHEN ISNULL(tm.country_cls, 0) = 0 " & vbCrLf & _
'                "                                    THEN 'Local' " & vbCrLf & _
'                "                                    ELSE 'Import' " & vbCrLf & _
'                "                               END " & vbCrLf
'
'    sql = sql + "                     END local , " & vbCrLf & _
'                "                     im.material_cls , " & vbCrLf & _
'                "                     ISNULL(sm.lm_inventory, 0) * ISNULL(ip.inventory_price, 0) lm_inventory , " & vbCrLf & _
'                "                     ISNULL(sm.tm_current, 0) * ISNULL(ip.inventory_price, 0) tm_current , " & vbCrLf & _
'                "                     ISNULL(sm.nm_current, 0) * ISNULL(ip.inventory_price, 0) nm_current , " & vbCrLf & _
'                "                     ISNULL(it.in_transit, 0) * ISNULL(ip.inventory_price, 0) in_transit " & vbCrLf & _
'                "           FROM       " & vbCrLf & _
'                "          (  " & vbCrLf & _
'                "           SELECT warehouse_code, item_code, " & vbCrLf & _
'                "                       SUM(LM_inventory) LM_inventory, " & vbCrLf & _
'                "                       SUM(TM_Current) TM_Current, " & vbCrLf
'
'    sql = sql + "                       SUM(NM_Current) NM_Current " & vbCrLf & _
'                "           FROM             " & vbCrLf & _
'                "           ( " & vbCrLf & _
'                "               SELECT warehouse_code, item_code, " & vbCrLf & _
'                "                   COALESCE(CASE WHEN @GetDiff=0 THEN 0 " & vbCrLf & _
'                "                            WHEN @GetDiff=1 THEN LM_Inventory " & vbCrLf & _
'                "                            WHEN @GetDiff=2 THEN TM_Current END,0) LM_inventory, " & vbCrLf & _
'                "                   COALESCE(CASE WHEN @GetDiff=0 THEN LM_inventory " & vbCrLf & _
'                "                            WHEN @GetDiff=1 THEN TM_Current " & vbCrLf & _
'                "                            WHEN @GetDiff=2 THEN NM_Current END,0) TM_Current, " & vbCrLf & _
'                "                   COALESCE(CASE WHEN @GetDiff=0 THEN TM_Current " & vbCrLf
'
'    sql = sql + "                            WHEN @GetDiff=1 THEN NM_Current " & vbCrLf & _
'                "                            WHEN @GetDiff=2 THEN NM_Current END,0) NM_Current " & vbCrLf & _
'                "               FROM Stock_Master " & vbCrLf & _
'                "                    " & vbCrLf & _
'                "               UNION ALL " & vbCrLf & _
'                "               -- Last Month " & vbCrLf & _
'                "               SELECT Warehouse_Code, Item_code,  COALESCE(Inventory,[Current]) LM_inventory, 0 TM_Inventory, 0 NM_Inventory " & vbCrLf & _
'                "                   FROM Stock_history " & vbCrLf & _
'                "                       WHERE Stock_Year=YEAR(@SelectPeriod) " & vbCrLf & _
'                "                                   AND Stock_Month=MONTH(@SelectPeriod)-1 " & vbCrLf & _
'                "               UNION ALL " & vbCrLf
'
'    sql = sql + "               -- This Month " & vbCrLf & _
'                "               SELECT Warehouse_Code, Item_code,  0 LM_inventory, COALESCE(Inventory,[Current]) TM_Inventory, 0 NM_Inventory " & vbCrLf & _
'                "                   FROM Stock_history " & vbCrLf & _
'                "                       WHERE Stock_Year=YEAR(@SelectPeriod) " & vbCrLf & _
'                "                                   AND Stock_Month=MONTH(@SelectPeriod) " & vbCrLf & _
'                "               UNION ALL                                    " & vbCrLf & _
'                "               -- Next Month " & vbCrLf & _
'                "               SELECT Warehouse_Code, Item_code,  0 LM_inventory, 0 TM_Inventory, COALESCE(Inventory,[Current]) NM_Inventory " & vbCrLf & _
'                "                   FROM Stock_history " & vbCrLf & _
'                "                       WHERE Stock_Year=YEAR(@SelectPeriod) " & vbCrLf & _
'                "                                   AND Stock_Month=MONTH(@SelectPeriod)+1 " & vbCrLf
'
'    sql = sql + "           ) Stock GROUP BY warehouse_code, item_code " & vbCrLf & _
'                "          ) sm " & vbCrLf & _
'                "                     LEFT JOIN item_master im ON im.item_code = sm.item_code " & vbCrLf & _
'                "                     LEFT JOIN trade_master tm ON im.supplier_code = tm.trade_code " & vbCrLf & _
'                "                     LEFT JOIN ( SELECT  item_code , " & vbCrLf & _
'                "                                         inventory_price " & vbCrLf & _
'                "                                 FROM   " & vbCrLf & _
'                "                                   (SELECT * From Inventory_Price " & vbCrLf & _
'                "                                       UNION   ALL " & vbCrLf & _
'                "                                     SELECT * From InventoryPrice_History " & vbCrLf & _
'                "                                   ) Inventory_Price                            " & vbCrLf
'
'    sql = sql + "                                 WHERE   inventory_year = YEAR(@SelectPeriod) " & vbCrLf & _
'                "                                         AND inventory_month = MONTH(@SelectPeriod) " & vbCrLf & _
'                "                                         AND duty_status IN ( '0', '2', '3' ) " & vbCrLf & _
'                "                               ) ip ON sm.item_code = ip.item_code " & vbCrLf & _
'                "                     LEFT JOIN ( SELECT  whcode , " & vbCrLf & _
'                "                                         item_code , " & vbCrLf & _
'                "                                         SUM(qty) in_transit " & vbCrLf & _
'                "                                 FROM    packing_master pm " & vbCrLf & _
'                "                                         INNER JOIN packing_detail pd ON pm.packing_no = pd.packing_no " & vbCrLf & _
'                "                                 WHERE   DATEDIFF(month, pm.stuffing_date, etd) > 0 " & vbCrLf & _
'                "                                         AND YEAR(pm.stuffing_date) = YEAR(@SelectPeriod) " & vbCrLf
'
'    sql = sql + "                                         AND MONTH(pm.stuffing_date) = MONTH(@SelectPeriod) " & vbCrLf & _
'                "                                 GROUP BY whcode , " & vbCrLf & _
'                "                                         item_code " & vbCrLf & _
'                "                               ) it ON sm.item_code = it.item_code " & vbCrLf & _
'                "                                       AND sm.warehouse_code = it.whcode " & vbCrLf & _
'                "         ) a " & vbCrLf & _
'                " GROUP BY a.warehouse_code , " & vbCrLf & _
'                "         a.local , " & vbCrLf & _
'                "         a.material_cls " & vbCrLf & _
'                " ORDER BY a.warehouse_code ASC , " & vbCrLf & _
'                "         a.local DESC , " & vbCrLf
'
'    sql = sql + "         a.material_cls ASC  " & vbCrLf & _
'                "  " & vbCrLf

    sql = "  DECLARE @LastClosing AS DATETIME  " & vbCrLf & _
            "  DECLARE @GetDiff AS INTEGER  " & vbCrLf & _
            "  DECLARE @SelectPeriod AS DATETIME  " & vbCrLf & _
            "  DECLARE @Location AS VARCHAR(30)  " & vbCrLf & _
            "   " & vbCrLf & _
            "  " & vbCrLf & _
            "  " & vbCrLf & _
            "  SET @SelectPeriod = '" & Format(DMonth, "yyyy-MM-01") & "'  " & vbCrLf & _
            "  SET @LastClosing = ( SELECT    CONVERT(DATETIME, CAST(Inventory_Year AS VARCHAR(4)) + '-' + CAST(Inventory_Month AS VARCHAR(2)) + '-01') " & vbCrLf & _
            "                       FROM      Inventory_Control " & vbCrLf & _
            "                       WHERE     Inventory_Year = ( SELECT   MAX(Inventory_Year) FROM     Inventory_Control) " & vbCrLf

    sql = sql + "                                 AND Inventory_Month = ( SELECT MAX(Inventory_Month) " & vbCrLf & _
                "                                                         FROM  Inventory_Control " & vbCrLf & _
                "                                                         WHERE Inventory_Year = ( SELECT " & vbCrLf & _
                "                                                               MAX(Inventory_Year) " & vbCrLf & _
                "                                                               FROM " & vbCrLf & _
                "                                                               Inventory_Control " & vbCrLf & _
                "                                                               ) " & vbCrLf & _
                "                                                       ) " & vbCrLf & _
                "                     )                                                     " & vbCrLf & _
                "    " & vbCrLf & _
                "  SET @GetDiff = DATEDIFF(M, @LastClosing, @SelectPeriod)  " & vbCrLf
    
    sql = sql + "    " & vbCrLf & _
                "  SELECT a.warehouse_code ,a.local ,a.material_cls ,SUM(a.lm_inventory) lm_inventory , " & vbCrLf & _
                "         SUM(a.tm_current) tm_current ,SUM(a.nm_current) nm_current ,SUM(a.in_transit) in_transit " & vbCrLf & _
                "  FROM   ( SELECT    warehouse_code , " & vbCrLf & _
                "                     CASE WHEN warehouse_code = 'FG' " & vbCrLf & _
                "                          THEN CASE WHEN LEFT(sm.item_code, 1) = 'E' " & vbCrLf & _
                "                                    THEN 'Export' " & vbCrLf & _
                "                                    ELSE 'Local' " & vbCrLf & _
                "                               END " & vbCrLf & _
                "                          ELSE CASE WHEN ISNULL(tm.country_cls, 0) = 0 " & vbCrLf & _
                "                                    THEN 'Local' " & vbCrLf
    
    sql = sql + "                                    ELSE 'Import' " & vbCrLf & _
                "                               END " & vbCrLf & _
                "                     END local , " & vbCrLf & _
                "                     im.material_cls ,ISNULL(sm.lm_inventory, 0) * ISNULL(ip.inventory_price, 0) lm_inventory , " & vbCrLf & _
                "                     ISNULL(sm.tm_current, 0) * ISNULL(ip.inventory_price, 0) tm_current ,ISNULL(sm.nm_current, 0) * ISNULL(ip.inventory_price, 0) nm_current , " & vbCrLf & _
                "                     ISNULL(it.in_transit, 0) * ISNULL(ip.inventory_price, 0) in_transit " & vbCrLf & _
                "           FROM      ( SELECT    warehouse_code ,item_code ,SUM(LM_inventory) LM_inventory ,SUM(TM_Current) TM_Current ,SUM(NM_Current) NM_Current " & vbCrLf & _
                "                       FROM      ( SELECT    warehouse_code ,item_code , " & vbCrLf & _
                "                                             COALESCE(CASE WHEN @GetDiff = 0 " & vbCrLf & _
                "                                                           THEN 0 " & vbCrLf & _
                "                                                           WHEN @GetDiff = 1 " & vbCrLf
    
    sql = sql + "                                                           THEN LM_Inventory " & vbCrLf & _
                "                                                           WHEN @GetDiff = 2 " & vbCrLf & _
                "                                                           THEN TM_Current " & vbCrLf & _
                "                                                      END, 0) LM_inventory , " & vbCrLf & _
                "                                             COALESCE(CASE WHEN @GetDiff = 0 " & vbCrLf & _
                "                                                           THEN LM_inventory " & vbCrLf & _
                "                                                           WHEN @GetDiff = 1 " & vbCrLf & _
                "                                                           THEN TM_Current " & vbCrLf & _
                "                                                           WHEN @GetDiff = 2 " & vbCrLf & _
                "                                                           THEN NM_Current " & vbCrLf & _
                "                                                      END, 0) TM_Current , " & vbCrLf
    
    sql = sql + "                                             COALESCE(CASE WHEN @GetDiff = 0 " & vbCrLf & _
                "                                                           THEN TM_Current " & vbCrLf & _
                "                                                           WHEN @GetDiff = 1 " & vbCrLf & _
                "                                                           THEN NM_Current " & vbCrLf & _
                "                                                           WHEN @GetDiff = 2 " & vbCrLf & _
                "                                                           THEN NM_Current " & vbCrLf & _
                "                                                      END, 0) NM_Current " & vbCrLf & _
                "                                   FROM      Stock_Master " & vbCrLf
                                                   
    
    sql = sql + "                                              " & vbCrLf & _
                "                                   UNION ALL  " & vbCrLf & _
                "                -- Last Month  " & vbCrLf & _
                "                                   SELECT    Warehouse_Code ,Item_code ,COALESCE(Inventory, [Current]) LM_inventory ,0 TM_Inventory ,0 NM_Inventory " & vbCrLf & _
                "                                   FROM      Stock_history " & vbCrLf & _
                "                                   WHERE     Stock_Year = YEAR(@SelectPeriod) " & vbCrLf
    
    sql = sql + "                                             AND Stock_Month = MONTH(@SelectPeriod)- 1 " & vbCrLf & _
                "                                   UNION ALL  " & vbCrLf
    
    sql = sql + "                -- This Month  " & vbCrLf & _
                "                                   SELECT    Warehouse_Code ,Item_code ,0 LM_inventory ,COALESCE(Inventory, [Current]) TM_Inventory ,0 NM_Inventory " & vbCrLf & _
                "                                   FROM      Stock_history " & vbCrLf & _
                "                                   WHERE     Stock_Year = YEAR(@SelectPeriod) AND Stock_Month = MONTH(@SelectPeriod) " & vbCrLf
    
    sql = sql + "                                             " & vbCrLf & _
                "                                   UNION ALL  " & vbCrLf & _
                "                -- Next Month  " & vbCrLf & _
                "                                   SELECT    Warehouse_Code ,Item_code , 0 LM_inventory , 0 TM_Inventory ,COALESCE(Inventory, [Current]) NM_Inventory " & vbCrLf & _
                "                                   FROM      Stock_history " & vbCrLf & _
                "                                   WHERE     Stock_Year = YEAR(@SelectPeriod) AND Stock_Month = MONTH(@SelectPeriod) + 1 " & vbCrLf

    
    sql = sql + "                                             " & vbCrLf & _
                "                                 ) Stock " & vbCrLf & _
                "                       GROUP BY  warehouse_code ,item_code " & vbCrLf & _
                "                     ) sm " & vbCrLf & _
                "                     LEFT JOIN item_master im ON im.item_code = sm.item_code " & vbCrLf & _
                "                     LEFT JOIN trade_master tm ON im.supplier_code = tm.trade_code " & vbCrLf & _
                "                     LEFT JOIN ( SELECT  item_code ,inventory_price " & vbCrLf
    
    sql = sql + "                                 FROM    ( SELECT *FROM Inventory_Price  where company_Code=@Company_Code" & vbCrLf & _
                "                                           UNION ALL " & vbCrLf & _
                "                                           SELECT * FROM InventoryPrice_History  where company_Code=@Company_Code" & vbCrLf & _
                "                                         ) Inventory_Price " & vbCrLf & _
                "                                 WHERE   inventory_year = YEAR(@SelectPeriod) AND inventory_month = MONTH(@SelectPeriod) AND duty_status IN ( '0', '2', '3' ) " & vbCrLf & _
                "                               ) ip ON sm.item_code = ip.item_code " & vbCrLf & _
                "                     LEFT JOIN ( SELECT  whcode , item_code ,SUM(qty) in_transit " & vbCrLf & _
                "                                 FROM    packing_master pm " & vbCrLf & _
                "                                         INNER JOIN packing_detail pd ON pm.packing_no = pd.packing_no " & vbCrLf & _
                "                                 WHERE   DATEDIFF(month, pm.stuffing_date, etd) > 0 AND YEAR(pm.stuffing_date) = YEAR(@SelectPeriod) AND MONTH(pm.stuffing_date) = MONTH(@SelectPeriod) " & vbCrLf & _
                "                                 GROUP BY whcode ,item_code " & vbCrLf
    
    sql = sql + "                               ) it ON sm.item_code = it.item_code AND sm.warehouse_code = it.whcode " & vbCrLf & _
                "         ) a " & vbCrLf & _
                "  GROUP BY a.warehouse_code , " & vbCrLf & _
                "         a.local , " & vbCrLf & _
                "         a.material_cls " & vbCrLf & _
                "  ORDER BY a.warehouse_code ASC , " & vbCrLf & _
                "         a.local DESC , " & vbCrLf & _
                "         a.material_cls ASC "
'---------------------------------------------------
    
    
    If rsRpt.State <> adStateClosed Then rsRpt.Close
    rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
    If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
        
    sqlprint = sql
        
        
    Set report = application.OpenReport(App.path & "\Reports\rptValuationPriceWarehouse.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    
    'intDiffClosing = up_GetDateRange(DMonth.Value)
    
    intDiffClosing = 1
    
    report.ReportTitle = "Valuation Price Report Per Warehouse (Summary)"
    report.FormulaFields.GetItemByName("Period").Text = "'" & datePiList & "'"
    report.FormulaFields.GetItemByName("01_Diff_Closing").Text = "" & intDiffClosing & ""
    report.FormulaFields.GetItemByName("DigitDesimalAmountIDR").Text = "" & gi_decimalDigitAmountIDR & ""
    report.FormulaFields.GetItemByName("CurrentDate").Text = "'" & Format(Now, "dddd, dd mmmm, yyyy  hh:mm:ss") & "'"
    setTglPrint = Format(Now, "dddd, dd mmmm, yyyy  hh:mm:ss")
    
'    sql = "select a.local, a.material_cls, sum(a.lm_inventory) lm_inventory, sum(a.tm_current) tm_current, sum(a.nm_current) nm_current, sum(a.in_transit) in_transit " & _
'          vbLf & "from( " & _
'          vbLf & "select warehouse_code, case when warehouse_code = 'FG' " & _
'          vbLf & "then case when left(sm.item_code, 1) = 'E' then 'Export' else 'Local' end " & _
'          vbLf & "else case when isnull(tm.country_cls, 0) = 0 then 'Local' else 'Import' end " & _
'          vbLf & "end local, im.material_cls, " & _
'          vbLf & "isnull(sm.lm_inventory, 0) * isnull(ip.inventory_price, 0) lm_inventory, " & _
'          vbLf & "isnull(sm.tm_current, 0) * isnull(ip.inventory_price, 0) tm_current, " & _
'          vbLf & "isnull(sm.nm_current, 0) * isnull(ip.inventory_price, 0) nm_current, " & _
'          vbLf & "isnull(it.in_transit, 0) * isnull(ip.inventory_price, 0) in_transit "
'
'    sql = sql & _
'          vbLf & "from stock_master sm " & _
'          vbLf & "left join item_master im on im.item_code=sm.item_code " & _
'          vbLf & "left join trade_master tm on im.supplier_code = tm.trade_code " & _
'          vbLf & "left join ( " & _
'          vbLf & "select item_code, inventory_price From inventory_price " & _
'          vbLf & "where inventory_year = '" & DMonth.year & "' and inventory_month = '" & DMonth.month & "' " & _
'          vbLf & "and duty_status in ('0', '2', '3') " & _
'          vbLf & ") ip on sm.item_code = ip.item_code " & _
'          vbLf & "left join ( " & _
'          vbLf & "select whcode, item_code, sum(qty) in_transit " & _
'          vbLf & "from packing_master pm " & _
'          vbLf & "inner join packing_detail pd on pm.packing_no = pd.packing_no " & _
'          vbLf & "Where DateDiff(month, pm.stuffing_date, etd) > 0 " & _
'          vbLf & "and year(pm.stuffing_date) = '" & DMonth.year & "' and month(pm.stuffing_date) = '" & DMonth.month & "' " & _
'          vbLf & "group by whcode, item_code " & _
'          vbLf & ") it on sm.item_code = it.item_code and sm.warehouse_code = it.whcode " & _
'          vbLf & ") a group by a.local, a.material_cls " & _
'          vbLf & "order by a.local desc, a.material_cls asc "
   
' --------------------------------------------------
' Update Query Include Valuation History
' --------------------------------------------------
'sql = "  " & vbCrLf & _
'            "  DECLARE @LastClosing AS DATETIME  " & vbCrLf & _
'            "  DECLARE @GetDiff AS INTEGER  " & vbCrLf & _
'            "  DECLARE @SelectPeriod AS DATETIME  " & vbCrLf & _
'            "  DECLARE @Location AS VARCHAR(30)  " & vbCrLf & _
'            "  DECLARE @Company_Code AS CHAR(5) " & vbCrLf & _
'            "    " & vbCrLf & _
'            "  SET @SelectPeriod = '" & Format(DMonth, "yyyy-MM-01") & "'  " & vbCrLf & _
'            "  SET @Company_Code = '" & Trim(CboCompany.Text) & "' " & vbCrLf & _
'            "  SET @LastClosing = ( SELECT    CONVERT(DATETIME, CAST(Inventory_Year AS VARCHAR(4)) " & vbCrLf & _
'            "                                 + '-' + CAST(Inventory_Month AS VARCHAR(2)) "
'
'sql = sql + "                                 + '-01') " & vbCrLf & _
'            "                       FROM      Inventory_Control " & vbCrLf & _
'            "                       WHERE     Inventory_Year = ( SELECT   MAX(Inventory_Year) " & vbCrLf & _
'            "                                                    FROM     Inventory_Control " & vbCrLf & _
'            "                                                  ) " & vbCrLf & _
'            "                                 AND Inventory_Month = ( SELECT " & vbCrLf & _
'            "                                                               MAX(Inventory_Month) " & vbCrLf & _
'            "                                                         FROM  Inventory_Control " & vbCrLf & _
'            "                                                         WHERE Inventory_Year = ( SELECT " & vbCrLf & _
'            "                                                               MAX(Inventory_Year) " & vbCrLf & _
'            "                                                               FROM "
'
'sql = sql + "                                                               Inventory_Control " & vbCrLf & _
'            "                                                               ) " & vbCrLf & _
'            "                                                       ) " & vbCrLf & _
'            "                     )                                                     " & vbCrLf & _
'            "    " & vbCrLf & _
'            "  SET @GetDiff = DATEDIFF(M, @LastClosing, @SelectPeriod)  " & vbCrLf & _
'            "    " & vbCrLf & _
'            "  SELECT a.local , " & vbCrLf & _
'            "         a.material_cls , " & vbCrLf & _
'            "         SUM(a.lm_inventory) lm_inventory , " & vbCrLf & _
'            "         SUM(a.tm_current) tm_current , "
'
'sql = sql + "         SUM(a.nm_current) nm_current , " & vbCrLf & _
'            "         SUM(a.in_transit) in_transit " & vbCrLf & _
'            "  FROM   ( SELECT    warehouse_code , " & vbCrLf & _
'            "                     CASE WHEN warehouse_code = 'FG' " & vbCrLf & _
'            "                          THEN CASE WHEN LEFT(sm.item_code, 1) = 'E' " & vbCrLf & _
'            "                                    THEN 'Export' " & vbCrLf & _
'            "                                    ELSE 'Local' " & vbCrLf & _
'            "                               END " & vbCrLf & _
'            "                          ELSE CASE WHEN ISNULL(tm.country_cls, 0) = 0 " & vbCrLf & _
'            "                                    THEN 'Local' " & vbCrLf & _
'            "                                    ELSE 'Import' "
'
'sql = sql + "                               END " & vbCrLf & _
'            "                     END local , " & vbCrLf & _
'            "                     im.material_cls , " & vbCrLf & _
'            "                     ISNULL(sm.lm_inventory, 0) * ISNULL(ip.inventory_price, 0) lm_inventory , " & vbCrLf & _
'            "                     ISNULL(sm.tm_current, 0) * ISNULL(ip.inventory_price, 0) tm_current , " & vbCrLf & _
'            "                     ISNULL(sm.nm_current, 0) * ISNULL(ip.inventory_price, 0) nm_current , " & vbCrLf & _
'            "                     ISNULL(it.in_transit, 0) * ISNULL(ip.inventory_price, 0) in_transit " & vbCrLf & _
'            "           FROM      ( SELECT    warehouse_code , " & vbCrLf & _
'            "                                 item_code , " & vbCrLf & _
'            "                                 SUM(LM_inventory) LM_inventory , " & vbCrLf & _
'            "                                 SUM(TM_Current) TM_Current , "
'
'sql = sql + "                                 SUM(NM_Current) NM_Current " & vbCrLf & _
'            "                       FROM      ( SELECT    warehouse_code , " & vbCrLf & _
'            "                                             item_code , " & vbCrLf & _
'            "                                             COALESCE(CASE WHEN @GetDiff = 0 " & vbCrLf & _
'            "                                                           THEN 0 " & vbCrLf & _
'            "                                                           WHEN @GetDiff = 1 " & vbCrLf & _
'            "                                                           THEN LM_Inventory " & vbCrLf & _
'            "                                                           WHEN @GetDiff = 2 " & vbCrLf & _
'            "                                                           THEN TM_Current " & vbCrLf & _
'            "                                                      END, 0) LM_inventory , " & vbCrLf & _
'            "                                             COALESCE(CASE WHEN @GetDiff = 0 "
'
'sql = sql + "                                                           THEN LM_inventory " & vbCrLf & _
'            "                                                           WHEN @GetDiff = 1 " & vbCrLf & _
'            "                                                           THEN TM_Current " & vbCrLf & _
'            "                                                           WHEN @GetDiff = 2 " & vbCrLf & _
'            "                                                           THEN NM_Current " & vbCrLf & _
'            "                                                      END, 0) TM_Current , " & vbCrLf & _
'            "                                             COALESCE(CASE WHEN @GetDiff = 0 " & vbCrLf & _
'            "                                                           THEN TM_Current " & vbCrLf & _
'            "                                                           WHEN @GetDiff = 1 " & vbCrLf & _
'            "                                                           THEN NM_Current " & vbCrLf & _
'            "                                                           WHEN @GetDiff = 2 "
'
'sql = sql + "                                                           THEN NM_Current " & vbCrLf & _
'            "                                                      END, 0) NM_Current " & vbCrLf & _
'            "                                   FROM      Stock_Master " & vbCrLf & _
'            "                                   WHERE     Warehouse_Code IN ( " & vbCrLf & _
'            "                                             SELECT  WH_Code " & vbCrLf & _
'            "                                             FROM    dbo.WareHouse_Master A " & vbCrLf & _
'            "                                                     INNER JOIN dbo.Trade_Master B ON A.Adm_Group = B.Trade_Code " & vbCrLf & _
'            "                                             WHERE   Company_Code = @Company_Code ) " & vbCrLf & _
'            "                                   UNION ALL  " & vbCrLf & _
'            "                -- Last Month  " & vbCrLf & _
'            "                                   SELECT    Warehouse_Code , "
'
'sql = sql + "                                             Item_code , " & vbCrLf & _
'            "                                             COALESCE(Inventory, [Current]) LM_inventory , " & vbCrLf & _
'            "                                             0 TM_Inventory , " & vbCrLf & _
'            "                                             0 NM_Inventory " & vbCrLf & _
'            "                                   FROM      Stock_history " & vbCrLf & _
'            "                                   WHERE     Stock_Year = YEAR(@SelectPeriod) " & vbCrLf & _
'            "                                             AND Stock_Month = MONTH(@SelectPeriod) " & vbCrLf & _
'            "                                             - 1 " & vbCrLf & _
'            "                                             AND Warehouse_Code IN ( " & vbCrLf & _
'            "                                             SELECT  WH_Code " & vbCrLf & _
'            "                                             FROM    dbo.WareHouse_Master A "
'
'sql = sql + "                                                     INNER JOIN dbo.Trade_Master B ON A.Adm_Group = B.Trade_Code " & vbCrLf & _
'            "                                             WHERE   Company_Code = @Company_Code ) " & vbCrLf & _
'            "                                   UNION ALL  " & vbCrLf & _
'            "                -- This Month  " & vbCrLf & _
'            "                                   SELECT    Warehouse_Code , " & vbCrLf & _
'            "                                             Item_code , " & vbCrLf & _
'            "                                             0 LM_inventory , " & vbCrLf & _
'            "                                             COALESCE(Inventory, [Current]) TM_Inventory , " & vbCrLf & _
'            "                                             0 NM_Inventory " & vbCrLf & _
'            "                                   FROM      Stock_history " & vbCrLf & _
'            "                                   WHERE     Stock_Year = YEAR(@SelectPeriod) "
'
'sql = sql + "                                             AND Stock_Month = MONTH(@SelectPeriod) " & vbCrLf & _
'            "                                             AND Warehouse_Code IN ( " & vbCrLf & _
'            "                                             SELECT  WH_Code " & vbCrLf & _
'            "                                             FROM    dbo.WareHouse_Master A " & vbCrLf & _
'            "                                                     INNER JOIN dbo.Trade_Master B ON A.Adm_Group = B.Trade_Code " & vbCrLf & _
'            "                                             WHERE   Company_Code = @Company_Code ) " & vbCrLf & _
'            "                                   UNION ALL                                     " & vbCrLf & _
'            "                -- Next Month  " & vbCrLf & _
'            "                                   SELECT    Warehouse_Code , " & vbCrLf & _
'            "                                             Item_code , " & vbCrLf & _
'            "                                             0 LM_inventory , "
'
'sql = sql + "                                             0 TM_Inventory , " & vbCrLf & _
'            "                                             COALESCE(Inventory, [Current]) NM_Inventory " & vbCrLf & _
'            "                                   FROM      Stock_history " & vbCrLf & _
'            "                                   WHERE     Stock_Year = YEAR(@SelectPeriod) " & vbCrLf & _
'            "                                             AND Stock_Month = MONTH(@SelectPeriod) " & vbCrLf & _
'            "                                             + 1 " & vbCrLf & _
'            "                                             AND Warehouse_Code IN ( " & vbCrLf & _
'            "                                             SELECT  WH_Code " & vbCrLf & _
'            "                                             FROM    dbo.WareHouse_Master A " & vbCrLf & _
'            "                                                     INNER JOIN dbo.Trade_Master B ON A.Adm_Group = B.Trade_Code " & vbCrLf & _
'            "                                             WHERE   Company_Code = @Company_Code ) "
'
'sql = sql + "                                 ) Stock " & vbCrLf & _
'            "                       GROUP BY  warehouse_code , " & vbCrLf & _
'            "                                 item_code " & vbCrLf & _
'            "                     ) sm " & vbCrLf & _
'            "                     LEFT JOIN item_master im ON im.item_code = sm.item_code " & vbCrLf & _
'            "                     LEFT JOIN trade_master tm ON im.supplier_code = tm.trade_code " & vbCrLf & _
'            "                     LEFT JOIN ( SELECT  item_code , " & vbCrLf & _
'            "                                         inventory_price " & vbCrLf & _
'            "                                 FROM    ( SELECT    * " & vbCrLf & _
'            "                                           FROM      Inventory_Price " & vbCrLf & _
'            "                                           UNION   ALL "
'
'sql = sql + "                                           SELECT    * " & vbCrLf & _
'            "                                           FROM      InventoryPrice_History " & vbCrLf & _
'            "                                         ) Inventory_Price " & vbCrLf & _
'            "                                 WHERE   inventory_year = YEAR(@SelectPeriod) " & vbCrLf & _
'            "                                         AND inventory_month = MONTH(@SelectPeriod) " & vbCrLf & _
'            "                                         AND duty_status IN ( '0', '2', '3' ) " & vbCrLf & _
'            "                               ) ip ON sm.item_code = ip.item_code " & vbCrLf & _
'            "                     LEFT JOIN ( SELECT  whcode , " & vbCrLf & _
'            "                                         item_code , " & vbCrLf & _
'            "                                         SUM(qty) in_transit " & vbCrLf & _
'            "                                 FROM    packing_master pm "
'
'sql = sql + "                                         INNER JOIN packing_detail pd ON pm.packing_no = pd.packing_no " & vbCrLf & _
'            "                                 WHERE   DATEDIFF(month, pm.stuffing_date, etd) > 0 " & vbCrLf & _
'            "                                         AND YEAR(pm.stuffing_date) = YEAR(@SelectPeriod) " & vbCrLf & _
'            "                                         AND MONTH(pm.stuffing_date) = MONTH(@SelectPeriod) " & vbCrLf & _
'            "                                 GROUP BY whcode , " & vbCrLf & _
'            "                                         item_code " & vbCrLf & _
'            "                               ) it ON sm.item_code = it.item_code " & vbCrLf & _
'            "                                       AND sm.warehouse_code = it.whcode " & vbCrLf & _
'            "         ) a " & vbCrLf & _
'            "  GROUP BY a.local , " & vbCrLf & _
'            "         a.material_cls "
'
'sql = sql + "  ORDER BY a.local DESC , " & vbCrLf & _
'            "         a.material_cls ASC   " & vbCrLf & _
'            "  "
sql = "    " & vbCrLf & _
            "   DECLARE @LastClosing AS DATETIME   " & vbCrLf & _
            "   DECLARE @GetDiff AS INTEGER   " & vbCrLf & _
            "   DECLARE @SelectPeriod AS DATETIME   " & vbCrLf & _
            "   DECLARE @Location AS VARCHAR(30)   " & vbCrLf & _
            "   " & vbCrLf & _
            "      " & vbCrLf & _
            "   SET @SelectPeriod = '" & Format(DMonth, "yyyy-MM-01") & "'   " & vbCrLf & _
            "   " & vbCrLf & _
            "   SET @LastClosing = ( SELECT   CONVERT(DATETIME, CAST(Inventory_Year AS VARCHAR(4)) " & vbCrLf & _
            "                                 + '-' + CAST(Inventory_Month AS VARCHAR(2)) "

sql = sql + "                                 + '-01') " & vbCrLf & _
            "                        FROM     Inventory_Control " & vbCrLf & _
            "                        WHERE    Inventory_Year = ( SELECT   MAX(Inventory_Year) " & vbCrLf & _
            "                                                    FROM     Inventory_Control " & vbCrLf & _
            "                                                  ) " & vbCrLf & _
            "                                 AND Inventory_Month = ( SELECT " & vbCrLf & _
            "                                                               MAX(Inventory_Month) " & vbCrLf & _
            "                                                         FROM  Inventory_Control " & vbCrLf & _
            "                                                         WHERE Inventory_Year = ( SELECT " & vbCrLf & _
            "                                                               MAX(Inventory_Year) " & vbCrLf & _
            "                                                               FROM "

sql = sql + "                                                               Inventory_Control " & vbCrLf & _
            "                                                               ) " & vbCrLf & _
            "                                                       ) " & vbCrLf & _
            "                      )                                                      " & vbCrLf & _
            "      " & vbCrLf & _
            "   SET @GetDiff = DATEDIFF(M, @LastClosing, @SelectPeriod)   " & vbCrLf & _
            "      " & vbCrLf & _
            "   SELECT    a.local , " & vbCrLf & _
            "             a.material_cls , " & vbCrLf & _
            "             SUM(a.lm_inventory) lm_inventory , " & vbCrLf & _
            "             SUM(a.tm_current) tm_current , "

sql = sql + "             SUM(a.nm_current) nm_current , " & vbCrLf & _
            "             SUM(a.in_transit) in_transit " & vbCrLf & _
            "   FROM      ( SELECT    warehouse_code , " & vbCrLf & _
            "                         CASE WHEN warehouse_code = 'FG' " & vbCrLf & _
            "                              THEN CASE WHEN LEFT(sm.item_code, 1) = 'E' " & vbCrLf & _
            "                                        THEN 'Export' " & vbCrLf & _
            "                                        ELSE 'Local' " & vbCrLf & _
            "                                   END " & vbCrLf & _
            "                              ELSE CASE WHEN ISNULL(tm.country_cls, 0) = 0 " & vbCrLf & _
            "                                        THEN 'Local' " & vbCrLf & _
            "                                        ELSE 'Import' "

sql = sql + "                                   END " & vbCrLf & _
            "                         END local , " & vbCrLf & _
            "                         im.material_cls , " & vbCrLf & _
            "                         ISNULL(sm.lm_inventory, 0) * ISNULL(ip.inventory_price, " & vbCrLf & _
            "                                                             0) lm_inventory , " & vbCrLf & _
            "                         ISNULL(sm.tm_current, 0) * ISNULL(ip.inventory_price, " & vbCrLf & _
            "                                                           0) tm_current , " & vbCrLf & _
            "                         ISNULL(sm.nm_current, 0) * ISNULL(ip.inventory_price, " & vbCrLf & _
            "                                                           0) nm_current , " & vbCrLf & _
            "                         ISNULL(it.in_transit, 0) * ISNULL(ip.inventory_price, " & vbCrLf & _
            "                                                           0) in_transit "

sql = sql + "               FROM      ( SELECT    warehouse_code , " & vbCrLf & _
            "                                     item_code , " & vbCrLf & _
            "                                     SUM(LM_inventory) LM_inventory , " & vbCrLf & _
            "                                     SUM(TM_Current) TM_Current , " & vbCrLf & _
            "                                     SUM(NM_Current) NM_Current " & vbCrLf & _
            "                           FROM      ( SELECT    warehouse_code , " & vbCrLf & _
            "                                                 item_code , " & vbCrLf & _
            "                                                 COALESCE(CASE WHEN @GetDiff = 0 " & vbCrLf & _
            "                                                               THEN 0 " & vbCrLf & _
            "                                                               WHEN @GetDiff = 1 " & vbCrLf & _
            "                                                               THEN LM_Inventory "

sql = sql + "                                                               WHEN @GetDiff = 2 " & vbCrLf & _
            "                                                               THEN TM_Current " & vbCrLf & _
            "                                                          END, 0) LM_inventory , " & vbCrLf & _
            "                                                 COALESCE(CASE WHEN @GetDiff = 0 " & vbCrLf & _
            "                                                               THEN LM_inventory " & vbCrLf & _
            "                                                               WHEN @GetDiff = 1 " & vbCrLf & _
            "                                                               THEN TM_Current " & vbCrLf & _
            "                                                               WHEN @GetDiff = 2 " & vbCrLf & _
            "                                                               THEN NM_Current " & vbCrLf & _
            "                                                          END, 0) TM_Current , " & vbCrLf & _
            "                                                 COALESCE(CASE WHEN @GetDiff = 0 "

sql = sql + "                                                               THEN TM_Current " & vbCrLf & _
            "                                                               WHEN @GetDiff = 1 " & vbCrLf & _
            "                                                               THEN NM_Current " & vbCrLf & _
            "                                                               WHEN @GetDiff = 2 " & vbCrLf & _
            "                                                               THEN NM_Current " & vbCrLf & _
            "                                                          END, 0) NM_Current " & vbCrLf & _
            "                                       FROM      Stock_Master " & vbCrLf & _
            "                                       " & vbCrLf & _
            "                                                 " & vbCrLf & _
            "                                                 " & vbCrLf & _
            "                                                         "

sql = sql + "                                                  " & vbCrLf & _
            "                                                  " & vbCrLf & _
            "                                               " & vbCrLf & _
            "                                           " & vbCrLf & _
            "                                               " & vbCrLf & _
            "                                                " & vbCrLf & _
            "                                               " & vbCrLf & _
            "                                       UNION ALL   " & vbCrLf & _
            "                 -- Last Month   " & vbCrLf & _
            "                                       SELECT    Warehouse_Code , " & vbCrLf & _
            "                                                 Item_code , "

sql = sql + "                                                 COALESCE(Inventory, [Current]) LM_inventory , " & vbCrLf & _
            "                                                 0 TM_Inventory , " & vbCrLf & _
            "                                                 0 NM_Inventory " & vbCrLf & _
            "                                       FROM      Stock_history " & vbCrLf & _
            "                                       WHERE     Stock_Year = YEAR(@SelectPeriod) " & vbCrLf & _
            "                                                 AND Stock_Month = MONTH(@SelectPeriod) " & vbCrLf & _
            "                                                 - 1 " & vbCrLf & _
            "                                                 " & vbCrLf & _
            "                                                " & vbCrLf & _
            "                                                  " & vbCrLf & _
                                                                     ""

sql = sql + "                                                 " & vbCrLf & _
            "                                                " & vbCrLf & _
            "  " & vbCrLf & _
            "                                               " & vbCrLf & _
            "                                               " & vbCrLf & _
            "                                                   " & vbCrLf & _
            "                                       UNION ALL   " & vbCrLf & _
            "                 -- This Month   " & vbCrLf & _
            "                                       SELECT    Warehouse_Code , " & vbCrLf & _
            "                                                 Item_code , " & vbCrLf & _
            "                                                 0 LM_inventory , "

sql = sql + "                                                 COALESCE(Inventory, [Current]) TM_Inventory , " & vbCrLf & _
            "                                                 0 NM_Inventory " & vbCrLf & _
            "                                       FROM      Stock_history " & vbCrLf & _
            "                                       WHERE     Stock_Year = YEAR(@SelectPeriod) " & vbCrLf & _
            "                                                 AND Stock_Month = MONTH(@SelectPeriod) " & vbCrLf & _
            "                                               "

sql = sql + "  " & vbCrLf & _
            "                                       UNION ALL                                      " & vbCrLf & _
            "                 -- Next Month   " & vbCrLf & _
            "                                       SELECT    Warehouse_Code , " & vbCrLf & _
            "                                                 Item_code , " & vbCrLf & _
            "                                                 0 LM_inventory , " & vbCrLf & _
            "                                                 0 TM_Inventory , " & vbCrLf & _
            "                                                 COALESCE(Inventory, [Current]) NM_Inventory "

sql = sql + "                                       FROM      Stock_history " & vbCrLf & _
            "                                       WHERE     Stock_Year = YEAR(@SelectPeriod) " & vbCrLf & _
            "                                                 AND Stock_Month = MONTH(@SelectPeriod) " & vbCrLf & _
            "                                                 + 1 " & vbCrLf & _
            "                                                  "

sql = sql + "                                              " & vbCrLf & _
            "                                                " & vbCrLf & _
            "                                     ) Stock " & vbCrLf & _
            "                           GROUP BY  warehouse_code , " & vbCrLf & _
            "                                     item_code " & vbCrLf & _
            "                         ) sm " & vbCrLf & _
            "                         LEFT JOIN item_master im ON im.item_code = sm.item_code " & vbCrLf & _
            "                         LEFT JOIN trade_master tm ON im.supplier_code = tm.trade_code " & vbCrLf & _
            "                         LEFT JOIN ( SELECT  item_code , " & vbCrLf & _
            "                                             inventory_price "

sql = sql + "                                     FROM    ( SELECT    * " & vbCrLf & _
            "                                               FROM      Inventory_Price WHERE Company_Code=@Company_Code " & vbCrLf & _
            "                                               UNION   ALL " & vbCrLf & _
            "                                               SELECT    * " & vbCrLf & _
            "                                               FROM      InventoryPrice_History WHERE Company_Code=@Company_Code " & vbCrLf & _
            "                                             ) Inventory_Price " & vbCrLf & _
            "                                     WHERE   inventory_year = YEAR(@SelectPeriod) " & vbCrLf & _
            "                                             AND inventory_month = MONTH(@SelectPeriod) " & vbCrLf & _
            "                                             AND duty_status IN ( '0', '2', '3' ) " & vbCrLf & _
            "                                   ) ip ON sm.item_code = ip.item_code " & vbCrLf & _
            "                         LEFT JOIN ( SELECT  whcode , "

sql = sql + "                                             item_code , " & vbCrLf & _
            "                                             SUM(qty) in_transit " & vbCrLf & _
            "                                     FROM    packing_master pm " & vbCrLf & _
            "                                             INNER JOIN packing_detail pd ON pm.packing_no = pd.packing_no " & vbCrLf & _
            "                                             INNER JOIN dbo.OrderEntry_Master OM ON PD.Order_No=OM.PO_No " & vbCrLf & _
            "                                     WHERE   OM.Company_Code=@Company_Code AND DATEDIFF(month, pm.stuffing_date, " & vbCrLf & _
            "                                                      etd) > 0 " & vbCrLf & _
            "                                             AND YEAR(pm.stuffing_date) = YEAR(@SelectPeriod) " & vbCrLf & _
            "                                             AND MONTH(pm.stuffing_date) = MONTH(@SelectPeriod) " & vbCrLf & _
            "                                     GROUP BY whcode , " & vbCrLf & _
            "                                             item_code "

sql = sql + "                                   ) it ON sm.item_code = it.item_code " & vbCrLf & _
            "                                           AND sm.warehouse_code = it.whcode " & vbCrLf & _
            "             ) a " & vbCrLf & _
            "   GROUP BY  a.local , " & vbCrLf & _
            "             a.material_cls " & vbCrLf & _
            "   ORDER BY  a.local DESC , " & vbCrLf & _
            "             a.material_cls ASC    " & vbCrLf & _
            "    " & vbCrLf & _
            "  "


'---------------------------------------------------
    
    If rsrpt2.State <> adStateClosed Then rsrpt2.Close
    rsrpt2.Open sql, Db, adOpenDynamic, adLockOptimistic
    
    sqlprint2 = sql
        
    report.OpenSubreport("summary").Database.Tables(1).SetDataSource rsrpt2
    report.OpenSubreport("summary").FormulaFields(1).Text = "" & intDiffClosing & ""
    
    
    reportcode = "pireportwhsummary"
    printorient = "1"
                
    Rpt.CRViewer1.ReportSource = report
    Rpt.CRViewer1.ViewReport
    Rpt.CRViewer1.Zoom 1
    
    Rpt.WindowState = 2
    Rpt.Show 1
    
    Me.MousePointer = vbDefault

End Sub


