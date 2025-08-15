VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRM_SheetDODD 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Delivery Note Print Out (Customer && DN Date)"
   ClientHeight    =   10740
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
   Icon            =   "FRM_SheetDODD.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   10740
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&D Instruction"
      Height          =   375
      Index           =   3
      Left            =   13290
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   10320
      Visible         =   0   'False
      Width           =   1260
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
      Height          =   1275
      Left            =   615
      TabIndex        =   9
      Top             =   1335
      Width           =   13935
      Begin VB.TextBox Text1 
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   300
         Width           =   5130
      End
      Begin MSComCtl2.DTPicker edate 
         Height          =   315
         Left            =   3195
         TabIndex        =   2
         Top             =   750
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
         Format          =   287637507
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker sdate 
         Height          =   315
         Left            =   1380
         TabIndex        =   1
         Top             =   750
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
         Format          =   287637507
         CurrentDate     =   37798
      End
      Begin VB.Line Line1 
         X1              =   2985
         X2              =   8100
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust CD"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN Date"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   14
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   195
         Index           =   3
         Left            =   2955
         TabIndex        =   13
         Top             =   810
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Cls"
         Height          =   195
         Index           =   4
         Left            =   4845
         TabIndex        =   12
         Top             =   810
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.New"
         Height          =   195
         Index           =   5
         Left            =   6675
         TabIndex        =   11
         Top             =   810
         Width           =   525
      End
      Begin MSForms.ComboBox cbodealer 
         Height          =   315
         Left            =   1380
         TabIndex        =   0
         Top             =   300
         Width           =   1500
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboisu 
         Height          =   315
         Left            =   5730
         TabIndex        =   3
         Top             =   750
         Width           =   855
         VariousPropertyBits=   746604571
         MaxLength       =   1
         DisplayStyle    =   3
         Size            =   "1508;556"
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
         Caption         =   " 2.Reissue"
         Height          =   195
         Index           =   7
         Left            =   7230
         TabIndex        =   10
         Top             =   810
         Width           =   885
      End
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
      Height          =   615
      Left            =   615
      TabIndex        =   7
      Top             =   8985
      Width           =   13935
      Begin VB.Label lblerror 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
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
         Left            =   90
         TabIndex        =   8
         Top             =   240
         Width           =   13695
      End
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Index           =   0
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9780
      Width           =   1260
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   375
      Index           =   2
      Left            =   13290
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9780
      Width           =   1260
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Index           =   1
      Left            =   11910
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9780
      Width           =   1260
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12750
      TabIndex        =   17
      Top             =   510
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6255
      Left            =   615
      TabIndex        =   18
      Top             =   2685
      Width           =   13935
      _cx             =   24580
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
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FDDFE3&
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Note Print Out (Customer && DN Date)"
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
      Index           =   6
      Left            =   615
      TabIndex        =   16
      Top             =   510
      Width           =   13935
   End
End
Attribute VB_Name = "FRM_SheetDODD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstcust As Recordset
Dim rst As Recordset
Dim i As Integer, Y As Integer

Dim bteColSelect As Byte
Dim bteColSJNo As Byte
Dim bteColSJDate As Byte
Dim bteColSJAmount As Byte
Dim bteColFix As Byte

Dim bteHakPrice As Byte

Sub Header()
With grid
    bteColSelect = 0
    bteColSJNo = 1
    bteColSJDate = 2
    bteColSJAmount = 3
    bteColFix = 4
    
    .ColS = 5
    .Rows = 1
    
    .TextMatrix(0, bteColSelect) = ""
    .TextMatrix(0, bteColSJNo) = "DN No (Ref No.)"
    .TextMatrix(0, bteColSJDate) = "DN Date"
    .TextMatrix(0, bteColSJAmount) = "Total DN Amount"
    .TextMatrix(0, bteColFix) = "Fix"
    
    .ColWidth(bteColSelect) = 300
    .ColWidth(bteColSJNo) = 2600
    .ColWidth(bteColSJDate) = 1500
    .ColWidth(bteColSJAmount) = 2800
    .ColWidth(bteColFix) = 800
    
    .ColHidden(bteColSJAmount) = (bteHakPrice = 0)
        
    .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
    .EditMaxLength = 1
End With
End Sub

Private Sub cbodealer_Change()
    cbodealer_Click
End Sub

Private Sub cbodealer_Click()
cbodealer = Trim(cbodealer)
If cbodealer.MatchFound Then
    Text1 = cbodealer.Column(1)
    If cboisu.ListIndex <> -1 Then
        MousePointer = vbHourglass
        display
        MousePointer = vbDefault
    End If
End If
End Sub

Private Sub cbodealer_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
        If Trim(cbodealer) <> "" Then
            cbodealer = Trim(cbodealer)
            If cbodealer.MatchFound Then
                cbodealer_Click
            Else
                lblerror = DisplayMsg(4072)
            End If
        Else
            lblerror.Caption = DisplayMsg(1033)
            Header
        End If
End If
End Sub

Private Sub cbodealer_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboisu_Change()
If cboisu.ListIndex = 0 Then
    Label1(5).FontBold = True
    Label1(7).FontBold = False
ElseIf cboisu.ListIndex = 1 Then
    Label1(5).FontBold = False
    Label1(7).FontBold = True
Else
    Exit Sub
End If
MousePointer = vbHourglass
display
MousePointer = vbDefault
End Sub

Private Sub cmdAction_Click(Index As Integer)
    Dim bolrpt As Boolean
    Me.MousePointer = vbHourglass
    Select Case Index
    Case 0
        Unload Me
        frmMainMenu.Show
    Case 2, 3
        do_no = ""
        bolrpt = False
        For i = 1 To grid.Rows - 1
            If grid.Cell(flexcpChecked, i, bteColSelect) = 1 Then
                If do_no = "" Then
                    do_no = "'" & Trim(grid.TextMatrix(i, bteColSJNo)) & "'"
                Else
                    do_no = do_no + ",'" & Trim(grid.TextMatrix(i, bteColSJNo)) & "'"
                End If
                bolrpt = True
            End If
        Next
        If bolrpt Then
            If Index = 2 Then Call DOReport(do_no) Else Call DIReport(do_no)
        Else
            lblerror = DisplayMsg("8011")
        End If
        If Trim(lblerror) = "" And cboisu.Text = "1" Then
            Header
            display
            do_no = ""
        End If
    Case 1
        blank
        Me.CtrlMenu1.MenuText = ""
        Header
        lblerror = ""
    End Select
    Me.MousePointer = vbDefault
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    lblerror.Caption = ErrMsg
End If
End Sub

Private Sub edate_Change()
rstcust.Requery
rstcust.Find "trade_code = '" & cbodealer.Text & "'"
If Not rstcust.EOF And cboisu.ListIndex <> -1 Then
    MousePointer = vbHourglass
    display
    MousePointer = vbDefault
End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
bteHakPrice = hakPrice(Me.Name)
Header
adtocombo
SDate = Format(Now, "dd mmm yyyy")
EDate = Format(Now, "dd mmm yyyy")
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
End Sub

Sub adtocombo()
'Dealer
sql = "select * from trade_master where trade_cls='2' "
Set rstcust = New Recordset
rstcust.Open sql, Db, adOpenKeyset, adLockOptimistic
With cbodealer
    .clear
    .columnCount = 2
    .ColumnWidths = "50 pt;280 pt"
    .ListWidth = 350
    .ListRows = 15
i = 0
Do Until rstcust.EOF
    .AddItem ""
    .List(i, 0) = Trim(rstcust!Trade_Code)
    .List(i, 1) = Trim(rstcust!trade_name)
    i = i + 1
    rstcust.MoveNext
Loop
End With

With cboisu
    .clear
    .columnCount = 2
    .ColumnWidths = "50 pt;75 pt"
    .ListWidth = 150
    .AddItem ""
    .List(0, 0) = "1"
    .List(0, 1) = "New"
    .AddItem ""
    .List(1, 0) = "2"
    .List(1, 1) = "ReIssue"
End With
End Sub

Sub display()
'cek date
Dim sqlP As String
Me.MousePointer = vbHourglass
If CDate(SDate) > CDate(EDate) Then
    lblerror.Caption = DisplayMsg("4068")
    Me.MousePointer = vbDefault
    Exit Sub
ElseIf CDate(EDate) < CDate(SDate) Then
    lblerror.Caption = DisplayMsg("4066")
    Me.MousePointer = vbDefault
    Exit Sub
End If
If cboisu.Text = "1" Then
    sqlP = "AND (DO_Master.Reissue_Cls IS NULL OR DO_Master.Reissue_Cls = '0')"
Else
    sqlP = "AND (DO_Master.Reissue_Cls = '1')"
End If
sql = "select Do_no,Do_date,amount,fix_cls from do_master " & _
    "where cust_code ='" & cbodealer.Text & "' and do_date >= '" & SDate & "' and do_date <='" & EDate & "' " & _
    sqlP & "order by do_no"
    
Set rst = New Recordset
rst.CursorLocation = adUseClient
rst.Open sql, Db, adOpenDynamic, adLockOptimistic
If Not rst.EOF Then
    With grid
        .Refresh
        .Rows = rst.RecordCount + 1
        .Row = .Rows - 1
    For i = 1 To rst.RecordCount
        .Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
        .TextMatrix(i, bteColSJNo) = Trim(rst!do_no)
        .TextMatrix(i, bteColSJDate) = Format(rst!do_date, "dd MMM yyyy")
        .TextMatrix(i, bteColSJAmount) = Format(rst!Amount, gs_formatAmountIDR)
       If IsNull(rst!fix_cls) Or rst!fix_cls = 0 Then
        .Cell(flexcpChecked, i, bteColFix) = flexUnchecked
       Else
        .Cell(flexcpChecked, i, bteColFix) = 1
       End If
       .Cell(flexcpAlignment, i, bteColSJNo, i, bteColSJDate) = flexAlignLeftCenter
       .Cell(flexcpAlignment, i, bteColSJAmount, i, bteColSJAmount) = flexAlignRightCenter
       'whitecols
       .Cell(flexcpBackColor, i, bteColSelect) = &HFFFFFF
    rst.MoveNext
    Next i
    End With
    lblerror.Caption = ""
Else
    Header
    lblerror.Caption = DisplayMsg(4006)
End If
Me.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
rst.Close
Set rst = Nothing
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = bteColSelect Then
    If grid.Cell(flexcpChecked, Row, bteColFix) = flexUnchecked Then
        grid.Cell(flexcpChecked, Row, bteColSelect) = flexUnchecked
        lblerror = DisplayMsg("0042")
    Else
        lblerror = ""
    End If
Else
    lblerror = ""
End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col <> bteColSelect Then Cancel = True
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If grid.Col = bteColSelect Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
      KeyAscii = 0
   End If
If KeyAscii = Asc(".") Then KeyAscii = 0
End If
End Sub

Private Sub sdate_Change()
rstcust.Requery
rstcust.Find "trade_code = '" & cbodealer.Text & "'"
If Not rstcust.EOF And cboisu.ListIndex <> -1 Then
    MousePointer = vbHourglass
    display
    MousePointer = vbDefault
End If
End Sub

Sub blank()
cbodealer.ListIndex = -1
cboisu.ListIndex = -1
SDate = Format(Now, "dd MMM YYYY")
EDate = Format(Now, "dd MMM YYYY")
End Sub



