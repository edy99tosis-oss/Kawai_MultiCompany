VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form F_StockLocation 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Stock Inquiry (Location)"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "F_StockLocation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   600
      TabIndex        =   16
      Top             =   9030
      Width           =   13965
      Begin VB.Label LblPesan 
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
         Height          =   240
         Left            =   135
         TabIndex        =   17
         Top             =   210
         Width           =   13575
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   12690
      TabIndex        =   15
      Top             =   405
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sea&rch"
      Height          =   375
      Index           =   9
      Left            =   4170
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1950
      Width           =   1035
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Sub &Menu"
      Height          =   375
      Index           =   8
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9705
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Last Page"
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   13455
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9705
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Next Page"
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   12135
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9705
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Prev Page"
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   10845
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9705
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&First Page"
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   9540
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9705
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Left            =   2693
      TabIndex        =   1
      Top             =   1965
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
      Format          =   146538499
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6300
      Left            =   600
      TabIndex        =   13
      Top             =   2610
      Width           =   13950
      _cx             =   24606
      _cy             =   11112
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
      HighLight       =   1
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
   Begin MSComCtl2.DTPicker DTgl 
      Height          =   315
      Left            =   2693
      TabIndex        =   14
      Top             =   1965
      Visible         =   0   'False
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
      Format          =   146538499
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Inquiry (Location)"
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
      Left            =   645
      TabIndex        =   8
      Top             =   420
      Width           =   13890
   End
   Begin VB.Line Line1 
      X1              =   6968
      X2              =   9983
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label LblLocationName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   6968
      TabIndex        =   7
      Top             =   1575
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Left            =   5378
      TabIndex        =   6
      Top             =   1575
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date (Month)"
      Height          =   255
      Left            =   968
      TabIndex        =   5
      Top             =   2025
      Width           =   1335
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "WareHouse"
      Height          =   255
      Left            =   968
      TabIndex        =   4
      Top             =   1575
      Width           =   1335
   End
   Begin MSForms.ComboBox CboLocationCD 
      Height          =   315
      Left            =   2693
      TabIndex        =   0
      Top             =   1545
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
Attribute VB_Name = "F_StockLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BulanTahun, HitungBulan As String, RsLast As New ADODB.Recordset
Dim dateUp As Date

Dim bteColProduct As Byte
Dim bteColDesc As Byte
Dim bteColAddress As Byte
Dim bteColPreMonth As Byte
Dim bteColReceipt As Byte
Dim bteColSupply As Byte
Dim bteColLossReject As Byte
Dim bteColCurrent As Byte

Private Sub Header()
    bteColProduct = 0
    bteColDesc = 1
    bteColAddress = 2
    bteColPreMonth = 3
    bteColReceipt = 4
    bteColSupply = 5
    bteColLossReject = 6
    bteColCurrent = 7
    
    Grid.Rows = 1
    Grid.ColS = 8
    
    Grid.TextMatrix(0, bteColProduct) = "Product Code"
    Grid.TextMatrix(0, bteColDesc) = "Description"
    Grid.TextMatrix(0, bteColAddress) = "Address"
    Grid.TextMatrix(0, bteColPreMonth) = "Pre Month"
    Grid.TextMatrix(0, bteColReceipt) = "Receipt Total"
    Grid.TextMatrix(0, bteColSupply) = "Supply Total"
    Grid.TextMatrix(0, bteColLossReject) = "Loss/Reject"
    Grid.TextMatrix(0, bteColCurrent) = "Current Stock"
    
    Grid.ColWidth(bteColProduct) = 2000
    Grid.ColWidth(bteColDesc) = 2500
    Grid.ColWidth(bteColAddress) = 800
    Grid.ColWidth(bteColPreMonth) = 1500
    Grid.ColWidth(bteColReceipt) = 1500
    Grid.ColWidth(bteColSupply) = 1500
    Grid.ColWidth(bteColLossReject) = 1500
    Grid.ColWidth(bteColCurrent) = 1300
    
    Grid.ColAlignment(bteColProduct) = flexAlignLeftCenter
    Grid.ColAlignment(bteColDesc) = flexAlignLeftCenter
    Grid.ColAlignment(bteColAddress) = flexAlignLeftCenter
    Grid.ColAlignment(bteColPreMonth) = flexAlignRightCenter
    Grid.ColAlignment(bteColReceipt) = flexAlignRightCenter
    Grid.ColAlignment(bteColSupply) = flexAlignRightCenter
    Grid.ColAlignment(bteColLossReject) = flexAlignRightCenter
    Grid.ColAlignment(bteColCurrent) = flexAlignRightCenter
End Sub

Private Sub CboLocationCD_Change()
    If CboLocationCD.MatchFound Then
       LblLocationName = CboLocationCD.List(CboLocationCD.ListIndex, 1)
       LblPesan = ""
    Else
       LblLocationName = ""
       LblPesan = DisplayMsg(4014) '"Location CD is not found !"
    End If
    Call Header
End Sub

Private Sub Cmd_Save_Click(Index As Integer)
    Select Case Index
    Case 8:
        frmMainMenu.Show
        Unload Me
    Case 9:
        Dim i As Integer
        If CboLocationCD.Text = "" Then
            LblPesan = DisplayMsg(1009) '"Please choose Product Code !"
        Else
            Call SettingGrid
        End If
    End Select
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblPesan.Caption = ErrMsg
    End If
End Sub

Private Sub DMonth_Change()
    Dim sql As String, RsMonth As New ADODB.Recordset, BMonth, BYear, BTgl As String
    If Format(DMonth.Value, "MM") < Format(dateUp, "MM") And Val(Format(DMonth.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then
        DMonth.Year = DMonth.Year + 1
        GoTo pass
    End If
    If Format(DMonth.Value, "MM") > Format(dateUp, "MM") And Val(Format(DMonth.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then DMonth.Year = DMonth.Year - 1
pass:
    dateUp = Format(DMonth.Value, "dd MMM yyyy")
    LblPesan = up_ValidateDateRange(DMonth, False)
    Call Header
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = "Stock Inquiry (Location)"
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    dateUp = Date
    DMonth = Format(Now, "mmmm yyyy")
    Call setting
    Call Header
End Sub

Private Sub setting()
    Dim sql As String, RsStock As New ADODB.Recordset
    Dim i As Integer
    
    If RsStock.State <> adStateClosed Then RsStock.Close
    ls_sql = " select * from (select wh_code, wh_name  from warehouse_master where stockcontrol_cls='01' union  " & _
    " select trade_code wh_code, trade_name wh_name from trade_master where trade_code in(select manufacture_code from manufacture_line))tbWarehouse order by wh_code "
    RsStock.Open ls_sql, Db, adOpenDynamic, adLockOptimistic
    
    CboLocationCD.columnCount = 2
    CboLocationCD.clear
    i = 0
    Do While Not RsStock.EOF
        CboLocationCD.AddItem ""
        CboLocationCD.List(i, 0) = Trim(RsStock("wh_code"))
        CboLocationCD.List(i, 1) = Trim(RsStock("wh_name"))
        i = i + 1
        RsStock.MoveNext
    Loop
    CboLocationCD.ColumnWidths = "50 pt; 300 pt"
    CboLocationCD.ListWidth = 350
    CboLocationCD.ListRows = 15
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Grid.Col >= bteColProduct Then Cancel = True
End Sub

Private Sub SettingGrid()
    Dim Simbol As String, RsSimbol As New ADODB.Recordset
    Dim sqlControl As String, RsInvControl As New ADODB.Recordset
    
    sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year,inventory_month"
    If RsInvControl.State <> adStateClosed Then RsInvControl.Close
    RsInvControl.Open sqlControl, Db, adOpenKeyset, adLockOptimistic
    If RsInvControl.EOF = True And RsInvControl.BOF = True Then
        LblErrMsg = DisplayMsg(4022) '"Inventory Stock hasn't been closed !"
        Exit Sub
    End If
    RsInvControl.MoveLast
    
    LblErrMsg = up_ValidateDateRange(DMonth.Value, False)
    If Trim(LblErrMsg) <> "" Then Exit Sub
    LblErrMsg = ""
    
    With Grid
        If RsLast.State <> adStateClosed Then RsLast.Close
        sql = " select address,item_name,stock_master.* From stock_master left join item_master " & _
            " on stock_master.item_code=item_master.item_code " & _
            " where stock_master.warehouse_code='" & Trim(CboLocationCD) & "'"
            RsLast.Open sql, Db, adOpenDynamic, adLockOptimistic
    
        Call Header
        Select Case up_GetDateRange(DMonth.Value) 'Val(Format(DMonth.Value, "MM"))
        Case 0:
            With RsLast
                i = 0
                Do While Not .EOF
                    i = i + 1
                    Grid.AddItem i
                    Grid.TextMatrix(i, bteColProduct) = Trim(!Item_Code)
                    Grid.TextMatrix(i, bteColDesc) = Trim(!item_name)
                    Grid.TextMatrix(i, bteColAddress) = IIf(IsNull(!Address), "", Trim(!Address))
                    Grid.TextMatrix(i, bteColPreMonth) = IIf(IsNull(!lm_premonth), "0.00", Format(!lm_premonth, gs_formatQty))
                    Grid.TextMatrix(i, bteColReceipt) = IIf(IsNull(!lm_receipt), "0.00", Format(!lm_receipt, gs_formatQty))
                    Grid.TextMatrix(i, bteColSupply) = IIf(IsNull(!lm_supply), "0.00", Format(!lm_supply, gs_formatQty))
                    Grid.TextMatrix(i, bteColLossReject) = IIf(IsNull(!lm_lossreject), "0.00", Format(!lm_lossreject, gs_formatQty))
                    Grid.TextMatrix(i, bteColCurrent) = IIf(IsNull(!lm_inventory), "0.00", Format(!lm_inventory, gs_formatQty))
                    .MoveNext
                Loop
            End With
        Case 1:

            With RsLast
                i = 0
                Do While Not .EOF
                    i = i + 1
                    Grid.AddItem i
                    Grid.TextMatrix(i, bteColProduct) = Trim(!Item_Code)
                    Grid.TextMatrix(i, bteColDesc) = Trim(!item_name)
                    Grid.TextMatrix(i, bteColAddress) = IIf(IsNull(!Address), "", Trim(!Address))
                    Grid.TextMatrix(i, bteColPreMonth) = IIf(IsNull(!tm_premonth), "0.00", Format(!tm_premonth, gs_formatQty))
                    Grid.TextMatrix(i, bteColReceipt) = IIf(IsNull(!tm_receipt), "0.00", Format(!tm_receipt, gs_formatQty))
                    Grid.TextMatrix(i, bteColSupply) = IIf(IsNull(!tm_supply), "0.00", Format(!tm_supply, gs_formatQty))
                    Grid.TextMatrix(i, bteColLossReject) = IIf(IsNull(!tm_lossreject), "0.00", Format(!tm_lossreject, gs_formatQty))
                    Grid.TextMatrix(i, bteColCurrent) = IIf(IsNull(!tm_current), "0.00", Format(!tm_current, gs_formatQty))
                    .MoveNext
                Loop
            End With
        Case 2:

            With RsLast
                i = 0
                Do While Not .EOF
                    i = i + 1
                    Grid.AddItem i
                    Grid.TextMatrix(i, bteColProduct) = Trim(!Item_Code)
                    Grid.TextMatrix(i, bteColDesc) = Trim(!item_name)
                    Grid.TextMatrix(i, bteColAddress) = IIf(IsNull(!Address), "", Trim(!Address))
                    Grid.TextMatrix(i, bteColPreMonth) = IIf(IsNull(!nm_premonth), "0.00", Format(!nm_premonth, gs_formatQty))
                    Grid.TextMatrix(i, bteColReceipt) = IIf(IsNull(!nm_receipt), "0.00", Format(!nm_receipt, gs_formatQty))
                    Grid.TextMatrix(i, bteColSupply) = IIf(IsNull(!nm_supply), "0.00", Format(!nm_supply, gs_formatQty))
                    Grid.TextMatrix(i, bteColLossReject) = IIf(IsNull(!nm_lossreject), "0.00", Format(!nm_lossreject, gs_formatQty))
                    Grid.TextMatrix(i, bteColCurrent) = IIf(IsNull(!nm_current), "0.00", Format(!nm_current, gs_formatQty))
                    .MoveNext
                Loop
            End With

        End Select
    End With
End Sub

