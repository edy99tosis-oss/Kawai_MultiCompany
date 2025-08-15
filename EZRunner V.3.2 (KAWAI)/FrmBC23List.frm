VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBC23List 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BC 23 List"
   ClientHeight    =   10395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmBC23List.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10395
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   1875
      Left            =   120
      TabIndex        =   10
      Tag             =   "TTTF*/"
      Top             =   960
      Width           =   14655
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H0080FFFF&
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
         Index           =   3
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   1320
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   294125571
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   13
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   294125571
         CurrentDate     =   37798
      End
      Begin MSForms.ComboBox cbointeface 
         Height          =   345
         Left            =   1680
         TabIndex        =   22
         Tag             =   "TTFF*/"
         Top             =   1335
         Width           =   1530
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2699;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interface Cls"
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
         Left            =   240
         TabIndex        =   21
         Tag             =   "TTFF*/"
         Top             =   1410
         Width           =   1110
      End
      Begin VB.Line Line1 
         X1              =   3360
         X2              =   7620
         Y1              =   1150
         Y2              =   1150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   20
         Tag             =   "TTFF*/"
         Top             =   450
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Code"
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
         Left            =   240
         TabIndex        =   19
         Tag             =   "TTFF*/"
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label LblCustomer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   3420
         TabIndex        =   18
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   240
         TabIndex        =   17
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   405
      End
      Begin MSForms.ComboBox cboTradeCode 
         Height          =   345
         Left            =   1680
         TabIndex        =   16
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1530
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2699;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblTradeName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   3360
         TabIndex        =   15
         Tag             =   "TTFF*/"
         Top             =   960
         Width           =   60
      End
      Begin VB.Label lblTampung 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   9000
         TabIndex        =   14
         Tag             =   "TTFF*/"
         Top             =   360
         Visible         =   0   'False
         Width           =   60
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   120
      TabIndex        =   8
      Tag             =   "TFTT*/"
      Top             =   9180
      Width           =   14640
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Tag             =   "TTTF*/"
         Top             =   180
         Width           =   14325
      End
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Create"
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
      Left            =   12420
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "FFTT*/"
      Top             =   9900
      Width           =   1125
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "TFFT*/"
      Top             =   9900
      Width           =   1125
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Detail"
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "FFTT*/"
      Top             =   9900
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   120
      TabIndex        =   3
      Tag             =   "TFTT*/"
      Top             =   9180
      Width           =   14640
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
         TabIndex        =   4
         Tag             =   "TTTF*/"
         Top             =   180
         Width           =   14325
      End
   End
   Begin VB.CommandButton CmdCreate 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Create"
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
      Left            =   12420
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "FFTT*/"
      Top             =   9900
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_SubMenu 
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "TFFT*/"
      Top             =   9900
      Width           =   1125
   End
   Begin VB.CommandButton CmdDetail 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Detail"
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "FFTT*/"
      Top             =   9900
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5910
      Left            =   135
      TabIndex        =   23
      Tag             =   "TTTT*/"
      Top             =   2895
      Width           =   14640
      _cx             =   25823
      _cy             =   10425
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12930
      TabIndex        =   24
      Tag             =   "FTTF*/"
      Top             =   240
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin MSForms.Label LblRecord 
      Height          =   255
      Left            =   11430
      TabIndex        =   26
      Tag             =   "FFTT*/"
      Top             =   8880
      Visible         =   0   'False
      Width           =   3345
      BackColor       =   16637923
      VariousPropertyBits=   8388627
      Caption         =   "0 Record(s)"
      Size            =   "5900;450"
      FontName        =   "Verdana"
      FontEffects     =   1073741827
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BC 23 List"
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
      Index           =   1
      Left            =   150
      TabIndex        =   25
      Tag             =   "TTTF*/"
      Top             =   270
      Width           =   14610
   End
End
Attribute VB_Name = "FrmBC23List"
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

Private Sub cboInterface_Change()
    up_GridHeader
End Sub

Private Sub cbotrade_Change()
'LblErrMsg = ""
'
'    If cbotrade.ListIndex <> -1 Then
'        lblTradeName.Caption = cbotrade.Column(1)
'        up_GridHeader
'    Else
'        lblTradeName.Caption = ""
'        cbotrade.SetFocus
'        Exit Sub
'    End If
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
    DtpFrom = Format(Now, "yyyy-MM-01")
    DtpTo.Value = Now
    
    up_FillComboTrade
    
    With cbointeface
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

    With cboTradeCode
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
    cmd.CommandText = "sp_BC23List_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("StartDate", adDBTime, adParamInput, , DtpFrom.Value)
    cmd.Parameters.append cmd.CreateParameter("EndDate", adDBTime, adParamInput, , DtpTo.Value)
    cmd.Parameters.append cmd.CreateParameter("TradeCode", adVarChar, adParamInput, 6, cboTradeCode.Text)
    
    If cbointeface.Text = "ALL" Then
        ls_status = "-"
    ElseIf cbointeface.Text = "Yes" Then
        ls_status = "1"
    ElseIf cbointeface.Text = "No" Then
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
            cmd.CommandText = "sp_BC23ListGenerateNoAju_Sel"
        
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
    FrmBC23Detail.txtNoPengajuan.Text = grid.TextMatrix(grid.RowSel, colNoPengajuan)
    FrmBC23Detail.up_LoadDataBC23 (grid.TextMatrix(grid.RowSel, colNoPengajuan))

    FrmBC23Detail.Show
    Me.Hide
ElseIf Index = 3 Then
    up_GridLoad
End If
End Sub

