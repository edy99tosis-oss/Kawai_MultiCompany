VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBC27List 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BC 2.7"
   ClientHeight    =   10830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBC27List.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10830
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1755
      Left            =   360
      TabIndex        =   6
      Tag             =   "TTTF*/"
      Top             =   1200
      Width           =   14565
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
         Height          =   375
         Index           =   3
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "TTFF*/"
         Top             =   1200
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   360
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
         Format          =   293208067
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   3480
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   360
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
         Format          =   293208067
         CurrentDate     =   37798
      End
      Begin MSForms.ComboBox cbotrade 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Tag             =   "TTFF*/"
         Top             =   750
         Width           =   1500
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2646;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboInterface 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Tag             =   "TTFF*/"
         Top             =   1200
         Width           =   1500
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2646;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3120
         X2              =   6840
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interface"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Tag             =   "TTFF*/"
         Top             =   1230
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Code"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Tag             =   "TTFF*/"
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   195
         Index           =   4
         Left            =   3120
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   165
      End
      Begin VB.Label lblTradeName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3120
         TabIndex        =   10
         Top             =   840
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   360
      TabIndex        =   4
      Tag             =   "TFTT*/"
      Top             =   9120
      Width           =   14490
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
         Left            =   90
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   180
         Width           =   14325
      End
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "TFFT*/"
      Top             =   10080
      Width           =   1140
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Create"
      Height          =   375
      Index           =   1
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "FFTT*/"
      Top             =   10080
      Width           =   1140
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Detail"
      Height          =   375
      Index           =   2
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "FFTT*/"
      Top             =   10080
      Width           =   1140
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   495
      Left            =   12720
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      _extentx        =   3625
      _extenty        =   873
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5775
      Left            =   360
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   3120
      Width           =   14565
      _cx             =   25691
      _cy             =   10186
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FDDFE3&
      Caption         =   "BC 2.7 List"
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
      Left            =   6000
      TabIndex        =   18
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmBC27List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const ColCheck As Integer = 0
Const ColTradeCode As Integer = 1
Const ColTradeName As Integer = 2
Const colSuratJalanNo As Integer = 3
Const colReceiptDate As Integer = 4
Const colNoPengajuan As Integer = 5
Const ColBCNO As Integer = 6
Const colBCDate As Integer = 7
Const colcount As Integer = 8

Public SuratJalan As String, NoPengajuan As String

Private Sub cboInterface_Change()
    up_GridHeader
End Sub

Private Sub cbotrade_Change()
LblErrMsg = ""

    If cbotrade.ListIndex <> -1 Then
        lblTradeName.Caption = cbotrade.Column(1)
        up_GridHeader
'        If combo1.ListIndex = 1 Then
'            adtocboapno
'        End If
'        Header
'        cboapno.Text = ""
    Else
'        kosong
'        LblErrMsg.Caption = DisplayMsg(4011) '"Record with this Customer Code not Exist"
        lblTradeName.Caption = ""
        cbotrade.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(Me)

up_FillCombo
up_GridHeader

HakU = hakUpdate(Me.Name)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

'With Anchor1
'    .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
'    .DoInit
'End With

End Sub

Private Sub up_FillCombo()
    dtpStartDate.Value = Format(Now, "yyyy-MM-01")
    dtpEndDate.Value = Now
    
    up_FillComboTrade
    
    With cboInterface
        .clear
        .AddItem "ALL"
        .AddItem "Yes"
        .AddItem "No"
        
        .ListWidth = 60
        .ListRows = 15
        
        .ListIndex = 0
    End With
End Sub

Private Sub up_FillComboTrade()
Dim sql As String
Dim RS As New Recordset

    sql = "Select Trade_Code, Trade_Name From Trade_Master Where Country_Cls = 1"
    Set RS = Db.Execute(sql)

    With cbotrade
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;300pt"
        .ListWidth = 350
        .ListRows = 15
        .AddItem
        .List(0, 0) = "ALL"
        .List(0, 1) = "ALL"
        
        i = 1
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS("Trade_code"))
            .List(i, 1) = IIf(IsNull(RS("trade_name")), " ", Trim(RS("Trade_Name")))
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = 0
        
    End With
End Sub

Private Sub up_GridHeader()
    LblErrMsg.Caption = ""

    With grid
        .ColS = colcount
        .Rows = 1
        
        .TextMatrix(0, ColCheck) = ""
        .TextMatrix(0, ColTradeCode) = "Trade Code"
        .TextMatrix(0, ColTradeName) = "Trade Name"
        .TextMatrix(0, colSuratJalanNo) = "Surat Jalan No"
        .TextMatrix(0, colReceiptDate) = "Receipt Date"
        .TextMatrix(0, colNoPengajuan) = "No Pengajuan"
        .TextMatrix(0, ColBCNO) = "BC No"
        .TextMatrix(0, colBCDate) = "BC Date"
        
        .ColWidth(ColCheck) = 300
        .ColWidth(ColTradeCode) = 1200
        .ColWidth(ColTradeName) = 3000
        .ColWidth(colSuratJalanNo) = 1800
        .ColWidth(colReceiptDate) = 1300
        .ColWidth(colNoPengajuan) = 3200
        .ColWidth(ColBCNO) = 1200
        .ColWidth(colBCDate) = 1300
        
        .FrozenCols = 3
        
        .ColFormat(colReceiptDate) = "dd MMM yyyy"
        .ColFormat(colBCDate) = "dd MMM yyyy"
        
        .Cell(flexcpAlignment, 0, 0, 0, colcount - 1) = flexAlignCenterCenter
        .ColAlignment(ColCheck) = flexAlignCenterCenter
        .ColAlignment(ColBCNO) = flexAlignLeftCenter
    End With
End Sub

Private Sub up_GridLoad()
    Dim ls_status As String
    
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer

    up_GridHeader
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27List_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("StartDate", adDBTime, adParamInput, , dtpStartDate.Value)
    cmd.Parameters.append cmd.CreateParameter("EndDate", adDBTime, adParamInput, , dtpEndDate.Value)
    cmd.Parameters.append cmd.CreateParameter("TradeCode", adVarChar, adParamInput, 6, cbotrade.Text)
    
    If cboInterface.Text = "ALL" Then
        ls_status = "-"
    ElseIf cboInterface.Text = "Yes" Then
        ls_status = "1"
    ElseIf cboInterface.Text = "No" Then
        ls_status = "0"
    End If
    
    cmd.Parameters.append cmd.CreateParameter("Status", adVarChar, adParamInput, 1, ls_status)
    
    Set RS = cmd.Execute
    
    If RS.EOF Then
         LblErrMsg.Caption = "[8012] Data is not found !"
         Exit Sub
    End If
    
    With grid
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
        
            .Cell(flexcpChecked, li_Row, ColCheck) = flexUnchecked

            .TextMatrix(li_Row, ColTradeCode) = Trim(RS!Supplier_Code)
            .TextMatrix(li_Row, ColTradeName) = Trim(RS!trade_name)
            .TextMatrix(li_Row, colSuratJalanNo) = Trim(RS!SuratJalan_No)
            .TextMatrix(li_Row, colReceiptDate) = RS!Receipt_Date
            .TextMatrix(li_Row, colNoPengajuan) = RS!No_Pengajuan
            .TextMatrix(li_Row, ColBCNO) = Trim(RS!BC40_No)
            .TextMatrix(li_Row, colBCDate) = RS!BC40_Date

'            .Cell(flexcpFontBold, li_row, colNoPengajuan) = True

            li_Row = .Rows - 1
            
        RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End With
    
End Sub

Private Sub up_CreateNoPengajuan()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
      
    For i = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, ColCheck) = flexChecked Then
            Set cmd = New ADODB.Command
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 0
            cmd.ActiveConnection = Db
            cmd.CommandText = "sp_BC27ListGenerateNoAju_Sel"
        
            cmd.Parameters.append cmd.CreateParameter("SuratJalanNo", adVarChar, adParamInput, 30, grid.TextMatrix(i, colSuratJalanNo))
            cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, grid.TextMatrix(i, colNoPengajuan))
        
            cmd.Execute
        End If

    Next
        
    up_GridLoad
    
    
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim cek As Integer

    If Col = ColCheck Then
        If grid.Cell(flexcpChecked, Row, Col) = flexChecked Then
            cek = 1
        Else
            cek = 2
        End If
        
        For i = 1 To grid.Rows - 1
            grid.Cell(flexcpChecked, i, 0) = flexUnchecked
        Next i
        
        grid.Cell(flexcpChecked, Row, Col) = cek
    End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> ColCheck Then
        Cancel = True
    Else
        For i = 1 To grid.Rows - 1
            If grid.Cell(flexcpChecked, i, Col) = flexChecked Then
                grid.Cell(flexcpChecked, i, 0) = flexChecked
            Else
                grid.Cell(flexcpChecked, i, 0) = flexUnchecked
            End If
        Next i
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub cmdAction_Click(Index As Integer)
If Index = 0 Then
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
ElseIf Index = 1 Then
    Dim icek As Integer
    LblErrMsg.Caption = ""
    For i = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, ColCheck) = flexChecked Then
            icek = icek + 1
        End If
    Next i
    
    If icek = 0 Then
        LblErrMsg.Caption = "[5012] There is no data to created !"
        Exit Sub
    End If
    
    Dim Aa As String
    Aa = MsgBox("Are you sure want to created?", vbYesNo + vbQuestion + vbDefaultButton2, "Question")
    If Aa = vbYes Then
        up_CreateNoPengajuan
    End If
    
ElseIf Index = 2 Then
    If grid.TextMatrix(grid.RowSel, colNoPengajuan) = "" Then
        Exit Sub
    End If
    SuratJalan = grid.TextMatrix(grid.RowSel, colSuratJalanNo)
    NoPengajuan = grid.TextMatrix(grid.RowSel, colNoPengajuan)
    
    frmBC27Detail.txtNoPengajuan.Text = grid.TextMatrix(grid.RowSel, colNoPengajuan)
    frmBC27Detail.up_LoadDataBC27 (grid.TextMatrix(grid.RowSel, colNoPengajuan))
    
    frmBC27Detail.Show
    Me.Hide
ElseIf Index = 3 Then
    up_GridLoad
End If
End Sub

