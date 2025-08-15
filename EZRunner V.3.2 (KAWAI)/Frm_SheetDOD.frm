VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_SheetDOD 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Delivery Note Print Out (DN Date)"
   ClientHeight    =   10605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14835
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_SheetDOD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10605
   ScaleWidth      =   14835
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&D Instruction"
      Height          =   375
      Index           =   3
      Left            =   12900
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10230
      Visible         =   0   'False
      Width           =   1260
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12300
      TabIndex        =   15
      Top             =   360
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   375
      Index           =   2
      Left            =   12915
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9825
      Width           =   1260
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Index           =   0
      Left            =   690
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9825
      Width           =   1260
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
      Left            =   690
      TabIndex        =   12
      Top             =   9075
      Width           =   13485
      Begin VB.Label lblerror 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   60
         TabIndex        =   13
         Top             =   180
         Width           =   13185
      End
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
      Height          =   1155
      Left            =   690
      TabIndex        =   6
      Top             =   1395
      Width           =   13485
      Begin MSComCtl2.DTPicker edate 
         Height          =   315
         Left            =   3390
         TabIndex        =   1
         Top             =   270
         Width           =   1485
         _ExtentX        =   2619
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
         Format          =   287703043
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker sdate 
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   270
         Width           =   1485
         _ExtentX        =   2619
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
         Format          =   287703043
         CurrentDate     =   37798
      End
      Begin MSForms.ComboBox cboisu 
         Height          =   315
         Left            =   1530
         TabIndex        =   2
         Top             =   675
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
         Caption         =   "1.New"
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   11
         Top             =   735
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Cls"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   735
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   195
         Index           =   3
         Left            =   3090
         TabIndex        =   9
         Top             =   330
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN Date"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " 2.Reissue"
         Height          =   195
         Index           =   7
         Left            =   3120
         TabIndex        =   7
         Top             =   735
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Index           =   1
      Left            =   11580
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9825
      Width           =   1260
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6270
      Left            =   690
      TabIndex        =   16
      Top             =   2715
      Width           =   13485
      _cx             =   23786
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
      Caption         =   "Delivery Note Print Out (DN Date)"
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
      Height          =   435
      Index           =   6
      Left            =   690
      TabIndex        =   14
      Top             =   330
      Width           =   13485
   End
End
Attribute VB_Name = "Frm_SheetDOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstcust As Recordset
Dim rst As Recordset
Dim i As Integer, Y As Integer

Dim bteColSelect As Byte
Dim bteColCustCode As Byte
Dim bteColCustName As Byte
Dim bteColSJNo As Byte
Dim bteColSJDate As Byte
Dim bteColSJAmount As Byte
Dim bteColFix As Byte

Dim bteHakPrice As Byte

Sub Header()
    bteColSelect = 0
    bteColCustCode = 1
    bteColCustName = 2
    bteColSJNo = 3
    bteColSJDate = 4
    bteColSJAmount = 5
    bteColFix = 6
    
    With grid
        .ColS = 7
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColCustCode) = "Customer Code"
        .TextMatrix(0, bteColCustName) = "Customer Abbr"
        .TextMatrix(0, bteColSJNo) = "DN No. (Ref No.)"
        .TextMatrix(0, bteColSJDate) = "DN Date"
        .TextMatrix(0, bteColSJAmount) = "Total DN Amount"
        .TextMatrix(0, bteColFix) = "Fix"
        
        .ColWidth(bteColSelect) = 250
        .ColWidth(bteColCustCode) = 1500
        .ColWidth(bteColCustName) = 1500
        .ColWidth(bteColSJNo) = 2100
        .ColWidth(bteColSJDate) = 1500
        .ColWidth(bteColSJAmount) = 2100
        .ColWidth(bteColFix) = 800
        
        .ColHidden(bteColSJAmount) = (bteHakPrice = 0)
        
        .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
        .EditMaxLength = 1
    End With
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
display
End Sub

Private Sub cboisu_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub cmdAction_Click(Index As Integer)
Dim bolrpt As Boolean
Select Case Index
        Case 0
            Unload Me
            frmMainMenu.Show
        Case 2, 3
            Me.MousePointer = vbHourglass
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
            Me.MousePointer = vbDefault
        Case 1
            blank
            Me.CtrlMenu1.MenuText = ""
            Header
            lblerror = ""
End Select
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    lblerror.Caption = ErrMsg
End If
End Sub

Private Sub edate_Change()
If cboisu.ListIndex <> -1 Then
    display
End If
End Sub

Private Sub edate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
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
Me.MousePointer = vbHourglass
Dim sqlP As String
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

sql = " select trade_code,trade_abbr,Do_no,Do_date,amount, Fix_Cls from do_master,trade_master " & _
        "Where do_master.cust_code = trade_master.trade_code " & _
        "and do_date >= '" & SDate & "' " & _
        "and do_date <='" & EDate & "' " & _
        sqlP & " order by do_no"
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
        .TextMatrix(i, bteColCustCode) = Trim(rst!Trade_Code)
        .TextMatrix(i, bteColCustName) = Trim(rst!trade_abbr)
        .TextMatrix(i, bteColSJNo) = rst!do_no
        .TextMatrix(i, bteColSJDate) = Format(rst!do_date, "dd mmm yyyy")
        .TextMatrix(i, bteColSJAmount) = Format(rst!Amount, gs_formatAmountIDR)
        If IsNull(rst!fix_cls) Or rst!fix_cls = 0 Then
            .Cell(flexcpChecked, i, bteColFix) = flexUnchecked
        Else
            .Cell(flexcpChecked, i, bteColFix) = 1
        End If
       .Cell(flexcpAlignment, i, bteColCustCode, i, bteColSJDate) = flexAlignLeftCenter
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
If Col = bteColSelect Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
      KeyAscii = 0
   End If
If KeyAscii = Asc(".") Then KeyAscii = 0
End If
End Sub

Private Sub sdate_Change()
If cboisu.ListIndex <> -1 Then
    display
End If
End Sub

Sub blank()
cboisu.ListIndex = -1
SDate = Format(Now, "dd MMM YYYY")
EDate = Format(Now, "dd MMM YYYY")
End Sub

Private Sub sdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub






