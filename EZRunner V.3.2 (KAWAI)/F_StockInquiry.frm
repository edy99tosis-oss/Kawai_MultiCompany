VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form F_StockInquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Stock Inquiry (Item Code)"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14955
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "F_StockInquiry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10380
   ScaleWidth      =   14955
   StartUpPosition =   1  'CenterOwner
   Tag             =   " "
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5100
      TabIndex        =   18
      Top             =   1560
      Width           =   300
   End
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
      Top             =   360
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sea&rch"
      Height          =   375
      Index           =   9
      Left            =   4380
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
      Left            =   577
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9765
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Last Page"
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9765
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Next Page"
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9765
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Prev Page"
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   10830
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9765
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&First Page"
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   9525
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9765
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Left            =   2670
      TabIndex        =   1
      Top             =   1950
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
      CustomFormat    =   "MMM yyyy"
      Format          =   141230083
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6270
      Left            =   600
      TabIndex        =   13
      Top             =   2610
      Width           =   13950
      _cx             =   24606
      _cy             =   11060
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
      Left            =   2670
      TabIndex        =   14
      Top             =   1950
      Visible         =   0   'False
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
      CustomFormat    =   "MMM yyyy"
      Format          =   141230083
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Inquiry (Item Code)"
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
      Left            =   600
      TabIndex        =   8
      Top             =   360
      Width           =   13830
   End
   Begin VB.Line Line1 
      X1              =   6930
      X2              =   11880
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label LblDesc 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   6930
      TabIndex        =   7
      Top             =   1575
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   1575
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      Height          =   255
      Left            =   930
      TabIndex        =   5
      Top             =   2025
      Width           =   1335
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Left            =   930
      TabIndex        =   4
      Top             =   1575
      Width           =   1335
   End
   Begin MSForms.ComboBox CboItemCD 
      Height          =   315
      Left            =   2655
      TabIndex        =   0
      Top             =   1545
      Width           =   2370
      VariousPropertyBits=   612386843
      MaxLength       =   25
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
Attribute VB_Name = "F_StockInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BulanTahun, HitungBulan As String, RsLast As New ADODB.Recordset
Dim dateUp As Date

Dim bteColWarehouse As Byte
Dim bteColDesc As Byte
Dim bteColAddress As Byte
Dim bteColPreMonth As Byte
Dim bteColReceipt As Byte
Dim bteColSupply As Byte
Dim bteColLossReject As Byte
Dim bteColCurrent As Byte

Private Sub Header()
    bteColWarehouse = 0
    bteColDesc = 1
    bteColAddress = 2
    bteColPreMonth = 3
    bteColReceipt = 4
    bteColSupply = 5
    bteColLossReject = 6
    bteColCurrent = 7
    
    grid.Rows = 1
    grid.ColS = 8
    
    grid.TextMatrix(0, bteColWarehouse) = "WareHouse"
    grid.TextMatrix(0, bteColDesc) = "Description"
    grid.TextMatrix(0, bteColAddress) = "Address"
    grid.TextMatrix(0, bteColPreMonth) = "Pre Month"
    grid.TextMatrix(0, bteColReceipt) = "Receipt Total"
    grid.TextMatrix(0, bteColSupply) = "Supply Total"
    grid.TextMatrix(0, bteColLossReject) = "Loss/Reject"
    grid.TextMatrix(0, bteColCurrent) = "Current Stock"
    
    grid.ColWidth(bteColWarehouse) = 1100
    grid.ColWidth(bteColDesc) = 3200
    grid.ColWidth(bteColAddress) = 800
    grid.ColWidth(bteColPreMonth) = 1500
    grid.ColWidth(bteColReceipt) = 1500
    grid.ColWidth(bteColSupply) = 1500
    grid.ColWidth(bteColLossReject) = 1500
    grid.ColWidth(bteColCurrent) = 1300
    
    grid.ColAlignment(bteColWarehouse) = flexAlignLeftCenter
    grid.ColAlignment(bteColDesc) = flexAlignLeftCenter
    grid.ColAlignment(bteColAddress) = flexAlignLeftCenter
    grid.ColAlignment(bteColPreMonth) = flexAlignRightCenter
    grid.ColAlignment(bteColReceipt) = flexAlignRightCenter
    grid.ColAlignment(bteColSupply) = flexAlignRightCenter
    grid.ColAlignment(bteColLossReject) = flexAlignRightCenter
    grid.ColAlignment(bteColCurrent) = flexAlignRightCenter
End Sub

Private Sub CboItemCD_Change()
    If CboItemCD.MatchFound Then
        lbldesc = CboItemCD.List(CboItemCD.ListIndex, 1)
        LblPesan = ""
    Else
        lbldesc = ""
        LblPesan = DisplayMsg(4003) '"Item Code is not found !"
    End If
    Call Header
End Sub

Private Sub Cmd_Save_Click(Index As Integer)
Dim strSQL As String
    Select Case Index
    Case 8:
        frmMainMenu.Show
        Unload Me
    Case 9:
        If CboItemCD.Text = "" Then
            LblPesan = DisplayMsg(1009) '"Please choose Product Code !"
        Else
            Me.MousePointer = vbHourglass
            
'            strSQL = "exec [sp_normalize_receipt_supply_BY_Item] '" & Trim(CboItemCD.Text) & "'"
'            Db.Execute strSQL
            Call SettingGrid
            Me.MousePointer = vbDefault
        End If
    End Select
End Sub

Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = CboItemCD.Text
 frm_BrowseItem.Show 1
 CboItemCD.Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblPesan.Caption = ErrMsg
    End If
End Sub

Private Sub DMonth_Change()
    If Format(DMonth.Value, "MM") < Format(dateUp, "MM") And Val(Format(DMonth.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then
        DMonth.Year = DMonth.Year + 1
        GoTo pass
    End If
    If Format(DMonth.Value, "MM") > Format(dateUp, "MM") And Val(Format(DMonth.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then DMonth.Year = DMonth.Year - 1
    
pass:
    dateUp = Format(DMonth.Value, "dd MMM yyyy")
    LblPesan = up_ValidateDateRange(DMonth, False)
    If Trim(LblPesan) <> "" Then Call Header: Exit Sub
    LblPesan = ""
    Call Header
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = "Stock Inquiry (Item Code)"
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    dateUp = Date
    DMonth = Format(Now, "mmm yyyy")
    Call setting
    Call Header
End Sub

Private Sub setting()
    Dim sql As String, RsItem As New ADODB.Recordset
    Dim i As Long
    
    If RsItem.State <> adStateClosed Then RsItem.Close
    sql = "select IM.*,wh_name from item_master IM,warehouse_master WM where IM.wh_code=WM.wh_code and IM.use_endday > convert(char(8), getdate(), 112)"
    RsItem.Open sql, Db, adOpenDynamic, adLockOptimistic
    
    CboItemCD.columnCount = 2
    CboItemCD.clear
    i = 0
    Do While Not RsItem.EOF
        CboItemCD.AddItem ""
        CboItemCD.List(i, 0) = Trim(RsItem!Item_Code)
        CboItemCD.List(i, 1) = Trim(RsItem!item_name) & " " & Trim(RsItem!WH_Name)
        i = i + 1
        RsItem.MoveNext
    Loop
    CboItemCD.ColumnWidths = "120 pt; 300 pt"
    CboItemCD.ListWidth = 430
    CboItemCD.ListRows = 15
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Col >= bteColWarehouse Then Cancel = True
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
    
    With grid
        If RsLast.State <> adStateClosed Then RsLast.Close
        sql = " select isnull(wh_name,'')wh_name,stock_master.* From stock_master " & _
        " left join  " & _
        " (select wh_code,wh_name from warehouse_master union all  " & _
        " select trade_code wh_code,trade_name wh_name from trade_master )warehouse_master " & _
        " on stock_master.warehouse_code=warehouse_master.wh_code " & _
        " where stock_master.item_code='" & Trim(CboItemCD) & "'"
        RsLast.Open sql, Db, adOpenDynamic, adLockOptimistic
        Call Header
    
        Select Case up_GetDateRange(DMonth.Value) 'Val(Format(DMonth.Value, "MM"))
        Case 0:
            With RsLast
                i = 0
                Do While Not .EOF
                    If RsSimbol.State <> adStateClosed Then RsSimbol.Close
                    RsSimbol.Open "item_master where item_code='" & Trim(CboItemCD) & "' and wh_code='" & Trim(!Warehouse_Code) & "'", Db, adOpenDynamic, adLockOptimistic, adCmdTable
                    If Not (RsSimbol.BOF And RsSimbol.EOF) Then
                        Simbol = IIf(IsNull(RsSimbol!Address), "", Trim(RsSimbol!Address))
                    Else
                        Simbol = ""
                    End If
                    
                    i = i + 1
                    grid.AddItem i
                    grid.TextMatrix(i, bteColWarehouse) = Trim(!Warehouse_Code)
                    grid.TextMatrix(i, bteColDesc) = Trim(!WH_Name)
                    grid.TextMatrix(i, bteColAddress) = Simbol
                    grid.TextMatrix(i, bteColPreMonth) = IIf(IsNull(!lm_premonth), "0.00", Format(!lm_premonth, gs_formatQty))
                    grid.TextMatrix(i, bteColReceipt) = IIf(IsNull(!lm_receipt), "0.00", Format(!lm_receipt, gs_formatQty))
                    grid.TextMatrix(i, bteColSupply) = IIf(IsNull(!lm_supply), "0.00", Format(!lm_supply, gs_formatQty))
                    grid.TextMatrix(i, bteColLossReject) = IIf(IsNull(!lm_lossreject), "0.00", Format(!lm_lossreject, gs_formatQty))
                    grid.TextMatrix(i, bteColCurrent) = IIf(IsNull(!lm_inventory), "0.00", Format(!lm_inventory, gs_formatQty))
                    .MoveNext
                Loop
            End With
        Case 1:

            With RsLast
                i = 0
                Do While Not .EOF
                    If RsSimbol.State <> adStateClosed Then RsSimbol.Close
                    RsSimbol.Open "item_master where item_code='" & Trim(CboItemCD) & "' and wh_code='" & Trim(!Warehouse_Code) & "'", Db, adOpenDynamic, adLockOptimistic, adCmdTable
                    If Not (RsSimbol.BOF And RsSimbol.EOF) Then
                        Simbol = IIf(IsNull(Trim(RsSimbol!Address)), "", Trim(RsSimbol!Address))
                    Else
                        Simbol = ""
                    End If
                    
                    i = i + 1
                    grid.AddItem i
                    grid.TextMatrix(i, bteColWarehouse) = Trim(!Warehouse_Code)
                    grid.TextMatrix(i, bteColDesc) = Trim(!WH_Name)
                    grid.TextMatrix(i, bteColAddress) = Simbol
                    grid.TextMatrix(i, bteColPreMonth) = IIf(IsNull(!tm_premonth), "0.00", Format(!tm_premonth, gs_formatQty))
                    grid.TextMatrix(i, bteColReceipt) = IIf(IsNull(!tm_receipt), "0.00", Format(!tm_receipt, gs_formatQty))
                    grid.TextMatrix(i, bteColSupply) = IIf(IsNull(!tm_supply), "0.00", Format(!tm_supply, gs_formatQty))
                    grid.TextMatrix(i, bteColLossReject) = IIf(IsNull(!tm_lossreject), "0.00", Format(!tm_lossreject, gs_formatQty))
                    grid.TextMatrix(i, bteColCurrent) = IIf(IsNull(!tm_current), "0.00", Format(!tm_current, gs_formatQty))
                    .MoveNext
                Loop
            End With
        Case 2:

            With RsLast
                i = 0
                Do While Not .EOF
                    If RsSimbol.State <> adStateClosed Then RsSimbol.Close
                    RsSimbol.Open "item_master where item_code='" & Trim(CboItemCD) & "' and wh_code='" & Trim(!Warehouse_Code) & "'", Db, adOpenDynamic, adLockOptimistic, adCmdTable
                    If Not (RsSimbol.BOF And RsSimbol.EOF) Then
                        Simbol = IIf(IsNull(RsSimbol!Address), "", Trim(RsSimbol!Address))
                    Else
                        Simbol = ""
                    End If
                    
                    i = i + 1
                    grid.AddItem i
                    grid.TextMatrix(i, bteColWarehouse) = Trim(!Warehouse_Code)
                    grid.TextMatrix(i, bteColDesc) = Trim(!WH_Name)
                    grid.TextMatrix(i, bteColAddress) = Simbol
                    grid.TextMatrix(i, bteColPreMonth) = IIf(IsNull(!nm_premonth), "0.00", Format(!nm_premonth, gs_formatQty))
                    grid.TextMatrix(i, bteColReceipt) = IIf(IsNull(!nm_receipt), "0.00", Format(!nm_receipt, gs_formatQty))
                    grid.TextMatrix(i, bteColSupply) = IIf(IsNull(!nm_supply), "0.00", Format(!nm_supply, gs_formatQty))
                    grid.TextMatrix(i, bteColLossReject) = IIf(IsNull(!nm_lossreject), "0.00", Format(!nm_lossreject, gs_formatQty))
                    grid.TextMatrix(i, bteColCurrent) = IIf(IsNull(!nm_current), "0.00", Format(!nm_current, gs_formatQty))
                    .MoveNext
                Loop
            End With
      
        End Select
        RsLast.Close
    End With
End Sub
