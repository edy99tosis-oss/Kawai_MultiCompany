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
   ScaleHeight     =   10380
   ScaleWidth      =   14955
   StartUpPosition =   1  'CenterOwner
   Tag             =   " "
   WindowState     =   2  'Maximized
   Begin VB.Timer tmSuggest 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   225
      Top             =   1890
   End
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
      TabIndex        =   17
      Top             =   1560
      Width           =   300
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   600
      TabIndex        =   15
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
         TabIndex        =   16
         Top             =   210
         Width           =   13575
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   12690
      TabIndex        =   14
      Top             =   360
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sea&rch"
      Height          =   375
      Index           =   9
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   1
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
      TabIndex        =   2
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   9765
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Left            =   2670
      TabIndex        =   0
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
      Format          =   129302531
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6270
      Left            =   600
      TabIndex        =   12
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
      TabIndex        =   13
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
      Format          =   129302531
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin MSForms.ComboBox CboItemCD 
      Height          =   315
      Left            =   2640
      TabIndex        =   18
      Tag             =   "TTFF*/"
      Top             =   1560
      Width           =   2370
      VariousPropertyBits=   612386843
      MaxLength       =   30
      DisplayStyle    =   3
      Size            =   "4180;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   1575
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   1575
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      Height          =   255
      Left            =   930
      TabIndex        =   4
      Top             =   2025
      Width           =   1335
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Left            =   930
      TabIndex        =   3
      Top             =   1575
      Width           =   1335
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

'===================================================================
Private IsInternalChange As Boolean
Private IsSelecting As Boolean
Private IsListOpen As Boolean
'===================================================================

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
    'Call setting
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

Private Sub SettingGridOld()
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
        " INNER join  " & _
        " (select wh_code,wh_name from warehouse_master WHERE Company_Code = (SELECT TOP 1 Factory_Code FROM dbo.App_FactoryPrivilege WHERE UserID = '" & userLogin & "' ORDER BY UpdateDate DESC ) union all  " & _
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

Private Sub SettingGrid()
    Dim Simbol As String
    Dim RsSimbol As New ADODB.Recordset
    Dim sqlControl As String
    Dim RsInvControl As New ADODB.Recordset
    Dim sql As String

    sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year,inventory_month"
    If RsInvControl.State <> adStateClosed Then RsInvControl.Close
    RsInvControl.Open sqlControl, Db, adOpenKeyset, adLockOptimistic
    If RsInvControl.EOF And RsInvControl.BOF Then
        LblErrMsg = DisplayMsg(4022)
        Exit Sub
    End If
    RsInvControl.MoveLast

    LblErrMsg = up_ValidateDateRange(DMonth.Value, False)
    If Trim(LblErrMsg) <> "" Then Exit Sub
    LblErrMsg = ""

    With grid
        If RsLast.State <> adStateClosed Then RsLast.Close

        sql = " SELECT ISNULL(wh_name,'') wh_name, stock_master.* "
        sql = sql & " FROM stock_master "
        sql = sql & " INNER JOIN ( "
        sql = sql & "   SELECT wh_code, wh_name, Company_Code "
        sql = sql & "   FROM warehouse_master "
        sql = sql & "   WHERE Company_Code = ( "
        sql = sql & "     SELECT TOP 1 Factory_Code "
        sql = sql & "     FROM dbo.App_FactoryPrivilege "
        sql = sql & "     WHERE UserID = '" & userLogin & "' "
        sql = sql & "     ORDER BY UpdateDate DESC "
        sql = sql & "   ) "
        sql = sql & "   UNION ALL "
        sql = sql & "   SELECT trade_code, trade_name, '00000' "
        sql = sql & "   FROM trade_master "
        sql = sql & "   UNION ALL "
        sql = sql & "   SELECT trade_code, trade_name, '11111' "
        sql = sql & "   FROM trade_master "
        sql = sql & "   WHERE trade_code LIKE '%FC%' "
        sql = sql & " ) warehouse_master "
        sql = sql & " ON stock_master.warehouse_code = warehouse_master.wh_code "
        sql = sql & " WHERE stock_master.item_code = '" & Trim(CboItemCD) & "' "
        sql = sql & " AND warehouse_master.Company_Code = ( "
        sql = sql & "   SELECT TOP 1 Factory_Code "
        sql = sql & "   FROM dbo.App_FactoryPrivilege "
        sql = sql & "   WHERE UserID = '" & userLogin & "' "
        sql = sql & "   ORDER BY UpdateDate DESC "
        sql = sql & " ) "

        RsLast.Open sql, Db, adOpenDynamic, adLockOptimistic
        Call Header

        Select Case up_GetDateRange(DMonth.Value)

        Case 0
            With RsLast
                i = 0
                Do While Not .EOF
                    If RsSimbol.State <> adStateClosed Then RsSimbol.Close
                    RsSimbol.Open "item_master where item_code='" & Trim(CboItemCD) & _
                                  "' and wh_code='" & Trim(!Warehouse_Code) & "'", _
                                  Db, adOpenDynamic, adLockOptimistic, adCmdTable

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

        Case 1
            With RsLast
                i = 0
                Do While Not .EOF
                    If RsSimbol.State <> adStateClosed Then RsSimbol.Close
                    RsSimbol.Open "item_master where item_code='" & Trim(CboItemCD) & _
                                  "' and wh_code='" & Trim(!Warehouse_Code) & "'", _
                                  Db, adOpenDynamic, adLockOptimistic, adCmdTable

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
                    grid.TextMatrix(i, bteColPreMonth) = IIf(IsNull(!tm_premonth), "0.00", Format(!tm_premonth, gs_formatQty))
                    grid.TextMatrix(i, bteColReceipt) = IIf(IsNull(!tm_receipt), "0.00", Format(!tm_receipt, gs_formatQty))
                    grid.TextMatrix(i, bteColSupply) = IIf(IsNull(!tm_supply), "0.00", Format(!tm_supply, gs_formatQty))
                    grid.TextMatrix(i, bteColLossReject) = IIf(IsNull(!tm_lossreject), "0.00", Format(!tm_lossreject, gs_formatQty))
                    grid.TextMatrix(i, bteColCurrent) = IIf(IsNull(!tm_current), "0.00", Format(!tm_current, gs_formatQty))
                    .MoveNext
                Loop
            End With

        Case 2
            With RsLast
                i = 0
                Do While Not .EOF
                    If RsSimbol.State <> adStateClosed Then RsSimbol.Close
                    RsSimbol.Open "item_master where item_code='" & Trim(CboItemCD) & _
                                  "' and wh_code='" & Trim(!Warehouse_Code) & "'", _
                                  Db, adOpenDynamic, adLockOptimistic, adCmdTable

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

Private Sub CboItemCD_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     Dim SqlI As String
    Dim RsI As ADODB.Recordset
    Dim filterText As String
    Dim ir As Long
    Dim oldSelStart As Long

    ' Simpan posisi cursor
    oldSelStart = CboItemCD.SelStart

    ' Ambil teks user
    filterText = Trim$(CboItemCD.Text)

    ' Minimal 2 karakter sebelum search
    If Len(filterText) < 2 Then Exit Sub

    SqlI = "EXEC dbo.sp_PartReceiptGetItemListByUser @UserID = '" & userLogin & "', @FilterText = '" & filterText & "'"

    Set RsI = New ADODB.Recordset
    RsI.Open SqlI, Db, adOpenKeyset, adLockReadOnly

    ' Build list sementara tanpa Clear
    CboItemCD.clear
    CboItemCD.columnCount = 2
    CboItemCD.TextColumn = 1
    ir = 0
       While Not RsI.EOF
        CboItemCD.AddItem ""
        CboItemCD.List(ir, 0) = RsI!ICd
        CboItemCD.List(ir, 1) = Trim$(RsI!inm)
        ir = ir + 1
        RsI.MoveNext
    Wend
    CboItemCD.ColumnWidths = "130 pt; 250 pt"
    CboItemCD.ListWidth = 380
    CboItemCD.ListRows = 15
    RsI.Close
    Set RsI = Nothing

    ' Kembalikan teks dan posisi cursor
    CboItemCD.Text = filterText
    CboItemCD.SelStart = oldSelStart
    CboItemCD.SelLength = 0

    ' Tampilkan dropdown
    If ir > 0 Then CboItemCD.DropDown
End Sub

Private Sub CboItemCD_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub CboItemCD_Change()
    CboItemCD_Click
End Sub

Private Sub CboItemCD_Click()
On Error GoTo ErrHandler

    Dim RsItem As ADODB.Recordset
    Dim rsPrice As ADODB.Recordset

    ' Pastikan ada match
    If Not CboItemCD.matchFound Then
        LblDesc.Caption = ""
        Exit Sub
    End If
    ' Ambil deskripsi item
    LblDesc = Trim$(uf_GetItemDescription(Trim$(CboItemCD)))

Cleanup:
    ' Tutup recordset jika masih terbuka
    If Not RsItem Is Nothing Then
        If RsItem.State = adStateOpen Then RsItem.Close
        Set RsItem = Nothing
    End If

    If Not rsPrice Is Nothing Then
        If rsPrice.State = adStateOpen Then rsPrice.Close
        Set rsPrice = Nothing
    End If
    Exit Sub

ErrHandler:
    ' Jika error, reset alamat
'    Lbladdress = ""
    Resume Cleanup
End Sub

''===================================================================
'Private Sub LoadSuggestion(ByVal keyword As String)
'    Dim RS As New ADODB.Recordset
'    Dim sql As String
'    Dim cleanKeyword As String
'
'    On Error GoTo ErrHandler
'
'    ' Cek Koneksi
'    If Db Is Nothing Then Exit Sub
'    If Db.State = 0 Then Exit Sub
'
'    cleanKeyword = Replace(keyword, "'", "''")
'
'    sql = "SELECT DISTINCT TOP 15 IM.Item_Code, IM.item_name, WM.wh_name " & _
'          "FROM item_master IM " & _
'          "JOIN warehouse_master WM ON IM.wh_code = WM.wh_code " & _
'          "WHERE IM.use_endday > CONVERT(char(8), GETDATE(), 112) " & _
'          "AND (IM.Item_Code LIKE '%" & cleanKeyword & "%' " & _
'          "OR IM.item_name LIKE '%" & cleanKeyword & "%' " & _
'          "OR WM.wh_name LIKE '%" & cleanKeyword & "%')"
'
'    RS.CursorLocation = adUseClient
'    RS.Open sql, Db, adOpenStatic, adLockReadOnly
'
'    With CboItemCD
'        .clear
'        .columnCount = 3
'        .ColumnWidths = "70 pt;175 pt;0 pt"
'        .ListWidth = 240
'    End With
'
'    If Not RS.EOF Then
'        Do While Not RS.EOF
'            CboItemCD.AddItem Trim(RS!Item_Code & "")
'            CboItemCD.List(CboItemCD.ListCount - 1, 1) = Trim(RS!item_name & "")
'            CboItemCD.List(CboItemCD.ListCount - 1, 2) = Trim(RS!WH_Name & "")
'            RS.MoveNext
'        Loop
'    End If
'
'CleanExit:
'    If RS.State = 1 Then RS.Close
'    Set RS = Nothing
'    Exit Sub
'
'ErrHandler:
'    Debug.Print "Error LoadSuggestion: " & err.Description
'    Resume CleanExit
'End Sub

'Private Sub CboItemCD_Change()
'    If IsInternalChange Or IsSelecting Then Exit Sub
'    If Trim(CboItemCD.Text) = "" Then
'        lbldesc.Caption = ""
'        Exit Sub
'    End If
'
'    If CboItemCD.ListIndex > -1 Then Exit Sub
'
'    tmSuggest.Enabled = False
'
'    If Len(CboItemCD.Text) >= 3 Then
'        tmSuggest.Interval = 500
'        tmSuggest.Enabled = True
'    End If
'End Sub
'
'Private Sub CboItemCD_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'    ' 13 adalah kode ASCII untuk tombol ENTER
'    If KeyCode = 13 Then
'        KeyCode = 0
'        tmSuggest.Enabled = False
'        If Trim(CboItemCD.Text) <> "" Then
'            CheckExactItem Trim(CboItemCD.Text)
'        End If
'    End If
'End Sub
'
'Private Sub CheckExactItem(ByVal ItemCode As String)
'    Dim RS As New ADODB.Recordset
'    Dim sql As String
'    Dim cleanCode As String
'
'    On Error GoTo ErrHandler
'
'    CboItemCD.DropDown
'    cleanCode = Replace(ItemCode, "'", "''")
'
'    sql = "SELECT IM.Item_Code, IM.item_name, WM.wh_name " & _
'          "FROM item_master IM " & _
'          "JOIN warehouse_master WM ON IM.wh_code = WM.wh_code " & _
'          "WHERE IM.Item_Code = '" & cleanCode & "' " & _
'          "AND IM.use_endday > CONVERT(char(8), GETDATE(), 112)"
'
'    RS.CursorLocation = adUseClient
'    RS.Open sql, Db, adOpenStatic, adLockReadOnly
'
'    If Not RS.EOF Then
'        IsInternalChange = True
'
'        CboItemCD.Text = Trim(RS!Item_Code & "")
'        lbldesc.Caption = Trim(RS!item_name & "")
'        DMonth.SetFocus
'
'        IsInternalChange = False
'
'        On Error Resume Next
'        Call Header
'        On Error GoTo 0
'
'    Else
'        lbldesc.Caption = ""
'        CboItemCD.SelStart = 0
'        CboItemCD.SelLength = Len(CboItemCD.Text)
'    End If
'
'CleanExit:
'    If RS.State = 1 Then RS.Close
'    Set RS = Nothing
'    Exit Sub
'
'ErrHandler:
'    Resume CleanExit
'End Sub

'Private Sub tmSuggest_Timer()
'    Dim kw As String
'    Dim curPos As Long
'
'    ' 1. Matikan Timer & Kunci
'    tmSuggest.Enabled = False
'    IsInternalChange = True
'
'    On Error Resume Next
'
'    kw = CboItemCD.Text
'    curPos = CboItemCD.SelStart
'    If err.number <> 0 Then curPos = Len(kw)
'    err.clear
'
'    ' 2. Load Data (Ingat: LoadSuggestion melakukan .Clear di dalamnya)
'    LoadSuggestion kw
'
'    ' 3. Restore Text & Cursor
'    CboItemCD.Text = kw
'    If curPos > Len(kw) Then curPos = Len(kw)
'    CboItemCD.SelStart = curPos
'
'    If CboItemCD.ListCount > 0 Then
'        ' Hanya buka jika statusnya "Belum Terbuka"
'        If IsListOpen = False Then
'            CboItemCD.DropDown
'            IsListOpen = True ' Tandai sudah terbuka
'        End If
'    Else
'        IsListOpen = False
'    End If
'
'    IsInternalChange = False
'    On Error GoTo 0
'End Sub

'Private Sub CboItemCD_Click()
'    IsSelecting = True
'    tmSuggest.Enabled = False
'
'    IsListOpen = False
'
'    If CboItemCD.ListIndex >= 0 Then
'        UpdateDescriptionFromList
'        On Error Resume Next
'        Call Header
'        On Error GoTo 0
'    End If
'
'    IsSelecting = False
'End Sub
'
'Private Sub UpdateDescriptionFromList()
'    If CboItemCD.ListIndex > -1 Then
'        lbldesc.Caption = CboItemCD.List(CboItemCD.ListIndex, 1) & ""
'        LblPesan.Caption = ""
'    End If
'End Sub
