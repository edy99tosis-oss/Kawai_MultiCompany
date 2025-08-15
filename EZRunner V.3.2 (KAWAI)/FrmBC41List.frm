VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBC41List 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BC 41 List"
   ClientHeight    =   10680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14970
   Icon            =   "FrmBC41List.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   37644
   ScaleMode       =   0  'User
   ScaleWidth      =   14970
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTampung 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9000
      TabIndex        =   22
      Tag             =   "TTFF*/"
      Top             =   10080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1875
      Left            =   150
      TabIndex        =   5
      Tag             =   "TTTF*/"
      Top             =   1080
      Width           =   14655
      Begin VB.CommandButton CmdSearch 
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "TTFF*/"
         Top             =   1320
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
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
         Format          =   293732355
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   8
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
         Format          =   293732355
         CurrentDate     =   37798
      End
      Begin MSForms.ComboBox cbointeface 
         Height          =   345
         Left            =   1680
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   405
      End
      Begin MSForms.ComboBox cboTradeCode 
         Height          =   345
         Left            =   1680
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   360
         Visible         =   0   'False
         Width           =   60
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   120
      TabIndex        =   3
      Tag             =   "TFTT*/"
      Top             =   9300
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
      Top             =   10020
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
      Top             =   10020
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
      Top             =   10020
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5910
      Left            =   135
      TabIndex        =   18
      Tag             =   "TTTT*/"
      Top             =   3015
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12930
      TabIndex        =   19
      Tag             =   "FTTF*/"
      Top             =   360
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin MSForms.Label LblRecord 
      Height          =   255
      Left            =   11430
      TabIndex        =   21
      Tag             =   "FFTT*/"
      Top             =   9000
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
      Caption         =   "BC 41 List"
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
      Index           =   0
      Left            =   150
      TabIndex        =   20
      Tag             =   "TTTF*/"
      Top             =   390
      Width           =   14610
   End
End
Attribute VB_Name = "FrmBC41List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim nilKosong As Boolean

Const ColCheck As Integer = 0
Const ColTradeCode As Integer = 1
Const ColTradeName As Integer = 2
Const colSuratJalanNo As Integer = 3
Const colReceiptDate As Integer = 4
Const colNoPengajuan As Integer = 5
Const ColBCNO As Integer = 6
Const colBCDate As Integer = 7
Const colcount As Integer = 8
Dim tampung As String

Private Sub cboTradeCode_Change()
LblErrMsg = ""

    If cboTradeCode.ListIndex <> -1 Then
        lblTradeName.Caption = cboTradeCode.Column(1)
        up_GridHeader
    Else
        lblTradeName.Caption = ""
        cboTradeCode.SetFocus
        Exit Sub
    End If
End Sub

Sub Kosong()

    DtpFrom = Format(Year(Now) & "-" & Format(Month(Now), "#0") & "-01", "dd MMM yyyy")
    DtpTo = Format(Now, "dd MMM yyyy")
    
    With cbointeface
        .clear
        .AddItem "ALL"
        .AddItem "Yes"
        .AddItem "No"
        
        .ListIndex = 0
    End With

End Sub

Private Sub CmdCreate_Click()
    Dim ls_sql As String
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer
    Dim JmlTran As Integer
    Dim SuratJalanNo As String
    Dim NoPengajuan As String
        
    With grid
        li_Row = 0
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, ColCheck) = flexChecked Then
                li_Row = i
                Exit For
            End If
        Next i
        
        If li_Row = 0 Then
            '#If from supply request then close form Prod
            'If Index = 1 Then GoTo frmMaterialRequest Else
            LblErrMsg = DisplayMsg(8011)
            Exit Sub
        Else
            DoEvents
                SuratJalanNo = Trim(.TextMatrix(li_Row, colSuratJalanNo))
                NoPengajuan = Trim(.TextMatrix(li_Row, colNoPengajuan))
            DoEvents
        End If
    End With
    
    Dim Aa As String
    Aa = MsgBox("Are you sure want to created?", vbYesNo + vbQuestion + vbDefaultButton2, "Question")
    If Aa = vbYes Then
    
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC41ListGenerateNoAju_Sel"
        
        cmd.Parameters.append cmd.CreateParameter("SuratJalanNo", adVarChar, adParamInput, 50, SuratJalanNo)
        cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, NoPengajuan)
        
        Set RS = cmd.Execute
        
        up_GridLoad
        
    End If
    
End Sub

Private Sub cmdDetail_Click()
Dim cek As Integer
Dim rsCek As New ADODB.Recordset

With grid
        cek = 0
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, ColCheck) = flexChecked Then
                cek = i
                Exit For
            End If
        Next i
        
        If cek = 0 Then
            
            LblErrMsg = DisplayMsg(8011)
            Exit Sub
        Else
            DoEvents
               txtTampung = Trim(.TextMatrix(cek, colNoPengajuan))
               FrmBC41Detail.txtNoPengajuan = Trim(.TextMatrix(cek, colNoPengajuan))
               If FrmBC41Detail.txtNoPengajuan = "" Then
                    LblErrMsg = DisplayMsg(8122)
                    Exit Sub
                Else
                DoEvents
               End If
        End If
    End With
    
    Unload Me
    FrmBC41Detail.Show
End Sub

Private Sub cmdSearch_Click()
    up_GridLoad
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(Me)

up_FillComboTrade
up_GridHeader
Kosong

HakU = hakUpdate(Me.Name)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

End Sub


Private Sub up_FillComboTrade()
Dim sql As String
Dim RS As New Recordset

    sql = "Select Trade_code, trade_name From Trade_Master Where Epte_Cls = 0 and Country_Cls=0 "
    Set RS = Db.Execute(sql)

    With cboTradeCode
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;300pt"
        .ListWidth = 350
        .ListRows = 15
        
    cboTradeCode.AddItem ""
    cboTradeCode.List(0, 0) = "ALL"
    cboTradeCode.List(0, 1) = "ALL"
    
        i = 1
        Do While Not RS.EOF
            .AddItem ""
            .List(i, 0) = Trim(RS("Trade_code"))
            .List(i, 1) = IIf(IsNull(RS("trade_name")), " ", Trim(RS("Trade_Name")))
            RS.MoveNext
            i = i + 1
        Loop
        cboTradeCode.ListIndex = 0
    End With
End Sub

Private Sub up_GridHeader()
    With grid
        .ColS = colcount
        .Rows = 1
'        .EditMaxLength = 1
        
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
        .ColWidth(ColTradeName) = 3500
        .ColWidth(colSuratJalanNo) = 2200
        .ColWidth(colReceiptDate) = 1500
        .ColWidth(colNoPengajuan) = 3100
        .ColWidth(ColBCNO) = 1000
        .ColWidth(colBCDate) = 1500
                
        .Cell(flexcpAlignment, 0, 0, 0, 7) = flexAlignCenterCenter
        .ColAlignment(colReceiptDate) = flexAlignCenterCenter
        .ColAlignment(ColBCNO) = flexAlignCenterCenter
        .ColAlignment(colBCDate) = flexAlignCenterCenter
    End With
End Sub

Private Sub up_GridLoad()
    Dim ls_sql As String
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer
    Dim JmlTran As Integer
    
    LblErrMsg.Caption = ""
    
    up_GridHeader
    
    Me.MousePointer = vbHourglass
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41List_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("StartDate", adDBTime, adParamInput, , DtpFrom.Value)
    cmd.Parameters.append cmd.CreateParameter("EndDate", adDBTime, adParamInput, , DtpTo.Value)
    cmd.Parameters.append cmd.CreateParameter("TradeCode", adVarChar, adParamInput, 6, cboTradeCode.Text)
    cmd.Parameters.append cmd.CreateParameter("Status", adVarChar, adParamInput, 3, cbointeface.Text)
    
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
    
        i = 1
        With grid
            While Not RS.EOF
                .Rows = .Rows + 1
                
                .Cell(flexcpChecked, i, ColCheck) = flexUnchecked
                .Cell(flexcpBackColor, i, ColCheck) = vbWhite
                .TextMatrix(i, ColTradeCode) = Trim(RS("Supplier_Code"))
                .TextMatrix(i, ColTradeName) = Trim(RS("trade_name"))
                .TextMatrix(i, colSuratJalanNo) = Trim(RS("SuratJalan_No"))
                .TextMatrix(i, colNoPengajuan) = Format(RS("No_Pengajuan"), gs_formatNoAju)
                .TextMatrix(i, colReceiptDate) = Format(Trim(RS("Receipt_Date")), "dd MMM yyyy")
                .TextMatrix(i, ColBCNO) = Trim(RS("BC40_No"))
                .TextMatrix(i, colBCDate) = Format(Trim(RS("BC40_Date")), "dd MMM yyyy")
                i = i + 1
            RS.MoveNext
            Wend
        End With
        
        LblRecord = Format(i - 1, "#,##0") & " Record(s)"
        
        Me.MousePointer = vbDefault
    
    Else
    
        LblErrMsg.Caption = DisplayMsg(13)
        
        Me.MousePointer = vbDefault
    
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

Private Sub Cmd_SubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

