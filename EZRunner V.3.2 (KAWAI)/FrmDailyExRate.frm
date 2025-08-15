VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDailyExRate 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Exchange Rate"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   15270
   Icon            =   "FrmDailyExRate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
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
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9750
      Width           =   1065
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   645
      TabIndex        =   2
      Top             =   8595
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
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
      CurrentDate     =   38483
   End
   Begin VB.CommandButton CmdMenu 
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
      Left            =   518
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9780
      Width           =   1185
   End
   Begin VB.CommandButton CmdSubmit 
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
      Left            =   13598
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9750
      Width           =   1155
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   518
      TabIndex        =   15
      Top             =   9075
      Width           =   14235
      Begin VB.Label LblErr 
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
         Left            =   105
         TabIndex        =   16
         Top             =   195
         Width           =   14010
      End
   End
   Begin VB.TextBox TxtRate 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4028
      MaxLength       =   12
      TabIndex        =   4
      Top             =   8595
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1155
      Left            =   518
      TabIndex        =   13
      Top             =   1320
      Width           =   14235
      Begin MSComCtl2.DTPicker Mperiod 
         Height          =   345
         Left            =   1290
         TabIndex        =   1
         Top             =   660
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
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
         Format          =   294125571
         UpDown          =   -1  'True
         CurrentDate     =   37868
      End
      Begin VB.Label LblTax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
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
         Left            =   180
         TabIndex        =   19
         Top             =   315
         Width           =   795
      End
      Begin VB.Line Line2 
         X1              =   2655
         X2              =   3435
         Y1              =   555
         Y2              =   555
      End
      Begin MSForms.ComboBox ComboBox1 
         Height          =   315
         Left            =   1290
         TabIndex        =   0
         Top             =   255
         Width           =   1260
         VariousPropertyBits=   746604571
         MaxLength       =   2
         DisplayStyle    =   3
         Size            =   "2222;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
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
         Left            =   2655
         TabIndex        =   18
         Top             =   285
         Width           =   795
      End
      Begin VB.Label LblTax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
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
         Left            =   180
         TabIndex        =   14
         Top             =   735
         Width           =   540
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5445
      Left            =   525
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2580
      Width           =   14235
      _cx             =   25109
      _cy             =   9604
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
      AllowUserResizing=   3
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
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
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
      Left            =   12915
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   420
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label LblTax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ex. Rate Date"
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
      Left            =   690
      TabIndex        =   20
      Top             =   8220
      Width           =   1185
   End
   Begin VB.Line Line1 
      X1              =   3135
      X2              =   3915
      Y1              =   8910
      Y2              =   8910
   End
   Begin VB.Label LblCurr 
      BackStyle       =   0  'Transparent
      Caption         =   "Currency"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3135
      TabIndex        =   17
      Top             =   8655
      Width           =   795
   End
   Begin MSForms.ComboBox CboCurr 
      Height          =   345
      Left            =   2235
      TabIndex        =   3
      Top             =   8595
      Width           =   795
      VariousPropertyBits=   746604571
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "1402;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      Height          =   555
      Index           =   0
      Left            =   525
      Top             =   8490
      Width           =   5925
   End
   Begin VB.Label LblTax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exchange Rate"
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
      Index           =   5
      Left            =   4560
      TabIndex        =   12
      Top             =   8220
      Width           =   1275
   End
   Begin VB.Label LblTax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Currency"
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
      Left            =   2235
      TabIndex        =   11
      Top             =   8220
      Width           =   795
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Exchange Rate"
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
      Left            =   525
      TabIndex        =   10
      Top             =   435
      Width           =   14235
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   525
      Top             =   8130
      Width           =   5925
   End
End
Attribute VB_Name = "FrmDailyExRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RSTR As ADODB.Recordset
Dim baru As Boolean
Dim Pos As Integer, jml As Integer
Dim StaErr As Boolean, nErr As Integer, StrWDel As String, nOK As Integer
Dim HakU As Integer
Dim Updet, Delet As Integer, Start1 As String, Finis As String, NO As Integer
Dim Test As String, dtDate As Boolean
Dim LDay As String, ADay As String
Dim tgl_sb As String
Dim Tgl_sb2 As String
Dim StLDay As Boolean
Dim PrevWeekEmpty As Boolean
Dim TglMP As String
Dim gs_status As String

Dim bteColSelect As Byte
Dim bteColDate As Byte
Dim bteColCurr As Byte
Dim bteColRate As Byte
Dim bteColCurrCode As Byte

Sub Header()
    Dim C As Byte
    
    bteColSelect = 0
    bteColDate = 1
    bteColCurr = 2
    bteColRate = 3
    bteColCurrCode = 4
    
    With grid
        .ColS = 5
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColDate) = "Ex. Rate Date"
        .TextMatrix(0, bteColCurr) = "Currency"
        .TextMatrix(0, bteColRate) = "Exchange Rate"
        .TextMatrix(0, bteColCurrCode) = "CurrCode"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColDate) = 1500
        .ColWidth(bteColCurr) = 1800
        .ColWidth(bteColRate) = 1800
        
        .ColHidden(bteColCurrCode) = True
        
        .ColAlignment(bteColDate) = flexAlignCenterCenter
        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        
        .EditMaxLength = 1
    End With
End Sub

Private Sub cbocurr_Change()
LblCurr = ""
cbocurr = cbocurr
If cbocurr.MatchFound Then
    LblCurr = cbocurr.List(cbocurr.ListIndex, 1)
Else
    LblErr = ""
End If
TxtRate = ""
TxtRate.SetFocus
End Sub

Private Sub cbocurr_Click()
LblCurr = ""
cbocurr = cbocurr
If cbocurr.MatchFound Then
    LblCurr = cbocurr.List(cbocurr.ListIndex, 1)
    LblErr = ""
Else
    LblErr = DisplayMsg(1028)
End If
TxtRate = Format(0, gs_formatExchangeRate)
TxtRate.SetFocus
End Sub

Private Sub cmdClear_Click()
ComboBox1.ListIndex = -1
Kosong
End Sub

Private Sub ComboBox1_Change()
If ComboBox1.MatchFound Then
    Label1 = ComboBox1.List(ComboBox1.ListIndex, 1)
    Browse
Else
    Label1 = ""
    Header
End If
End Sub

Private Sub CmdMenu_Click()
frmMainMenu.Show
Unload Me
DoEvents
End Sub

Private Sub Combobox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then Cancel = 1

End Sub
Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErr.Caption = ErrMsg
End If
End Sub

Private Sub CmdSubmit_Click()
Dim code As String, ls_sql As String

If hakUpdate(Me.Name) = 0 Then LblErr = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
If gs_status = "insert" Then If cek = False Then Exit Sub
If gs_status = "insert" Then
    Db.BeginTrans
    ls_sql = "INSERT INTO [Daily_ExchangeRate]([ExchangeRate_Date], [Currency_Code], [Daily_ExchangeRate], Last_Update, Last_User) " & _
                "VALUES('" & Format(DTPicker1, "yyyy-MM-dd") & "', '" & Trim(cbocurr) & "', '" & CDbl(TxtRate) & "', getdate(), '" & userLogin & "')"
    
    Db.Execute ls_sql
    Db.CommitTrans
    code = 1000
ElseIf gs_status = "update" Then
    Db.BeginTrans
    ls_sql = " update [Daily_ExchangeRate] " & _
            "set [Daily_ExchangeRate]='" & CDbl(TxtRate) & "', " & _
            "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
            "where [Currency_Code]='" & Trim(cbocurr) & "' " & _
            "and [ExchangeRate_Date] ='" & Format(DTPicker1, "yyyy-MM-dd") & "'"
    
    Db.Execute ls_sql
    Db.CommitTrans
    code = 1101
ElseIf gs_status = "delete" Then
    Db.BeginTrans
    
    For i = 0 To grid.Rows - 1
        If grid.TextMatrix(i, bteColSelect) = "D" Then
            ls_sql = " delete [Daily_ExchangeRate] where " & _
                              " [Currency_Code]='" & Trim(grid.TextMatrix(i, bteColCurrCode)) & "' " & _
                              " and [ExchangeRate_date] ='" & Format(grid.TextMatrix(i, bteColDate), "yyyy-MM-dd") & "'"
            Db.Execute ls_sql
        End If
    Next
    Db.CommitTrans

    code = 1201
End If

Call Kosong
LblErr = DisplayMsg(code)
Call Browse
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
Dim RS As Recordset, ir As Integer

CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
HakU = hakUpdate(Me.Name)

Header
'Browse

Call Kosong

Mperiod = Format(Now(), "MMM YYYY")
DTPicker1 = Format(Date, "dd MMM YYYY")

ComboBox1.AddItem ""
ComboBox1.List(0, 0) = "==ALL=="
ComboBox1.List(0, 1) = ""

Call up_FillCombo(ComboBox1, "curr_cls")
Call up_FillCombo(cbocurr, "curr_cls")

cbocurr = ""
LblCurr = ""

ComboBox1 = ""
Label1 = ""
End Sub

Sub Browse()
Dim RSTR As Recordset

sql = "select * from daily_exchangerate where month(exchangerate_date)='" & Format(Mperiod, "MM") & "'" & _
      " and year(exchangerate_date)='" & Format(Mperiod, "yyyy") & "' " & IIf(ComboBox1 = "==ALL==", "", " and currency_code='" & Trim(ComboBox1) & "'")

Set RSTR = New ADODB.Recordset
RSTR.Open sql, Db, adOpenKeyset, adLockOptimistic
i = 0
Header
While Not RSTR.EOF
        i = i + 1
        grid.AddItem ""
        grid.TextMatrix(i, bteColDate) = Format(RSTR!ExchangeRate_Date, "dd MMM yyyy")
            grid.Cell(flexcpAlignment, i, bteColDate) = flexAlignLeftCenter
        grid.TextMatrix(i, bteColCurr) = uf_GetCurrencyDescription(RSTR!currency_code)
        grid.TextMatrix(i, bteColCurrCode) = RSTR!currency_code
            grid.Cell(flexcpAlignment, i, bteColCurr) = flexAlignCenterCenter
        grid.TextMatrix(i, bteColRate) = Format(RSTR!daily_ExchangeRate, gs_formatExchangeRate)
        
        grid.Cell(flexcpBackColor, i, bteColSelect) = vbWhite
        
   RSTR.MoveNext
Wend

Kosong
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrGrid As String
Dim AdaS As Boolean, Pakai As Boolean, brs As Integer, id As Integer
StrGrid = grid.Text
AdaS = False
Pakai = False
brs = 0

If StrGrid = "S" Then
    For id = 1 To grid.Rows - 1
        If id <> Row Then grid.TextMatrix(id, bteColSelect) = ""
    Next id
    
    cbocurr = grid.TextMatrix(grid.Row, bteColCurrCode)
    TxtRate = grid.TextMatrix(grid.Row, bteColRate)
    DTPicker1 = Format(grid.TextMatrix(grid.Row, 1), "dd MMM yyyy")
    cbocurr.Enabled = False
    DTPicker1.Enabled = False
    baru = False
    LblErr = ""
    
    gs_status = "update"
    
ElseIf StrGrid = "D" Then
    Pakai = True
    For id = 1 To grid.Rows - 1
        'Jika ada S maka , hapus yg S
        If grid.TextMatrix(id, bteColSelect) = "S" Then grid.TextMatrix(id, bteColSelect) = "": Exit For
    Next id
    cbocurr.Enabled = True
    DTPicker1.Enabled = True
    baru = False
    LblErr = ""
    
    gs_status = "delete"
Else
    Pakai = True
    Kosong
End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid.Col > bteColSelect Then Cancel = True
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If grid.Col = bteColSelect Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
   End If
End If
End Sub

Function cek() As Boolean

Dim RS As New ADODB.Recordset

cek = True
cbocurr = cbocurr
If cbocurr.MatchFound = False Then cbocurr.SetFocus: LblErr = DisplayMsg(1028): cek = False: Exit Function
If CDbl(TxtRate) = 0 Then TxtRate.SetFocus: LblErr = DisplayMsg("0073"): cek = False: Exit Function
If RS.State <> adStateClosed Then RS.Close
sql = "select * from daily_exchangerate where currency_code='" & Trim(cbocurr) & "' and exchangeRate_date='" & Format(DTPicker1, "yyyy-MM-dd") & "'"
RS.Open sql, Db, adOpenKeyset, adLockOptimistic
If RS.EOF = False Then TxtRate.SetFocus: LblErr = DisplayMsg(1023): cek = False: Exit Function
If RS.State <> adStateClosed Then RS.Close

End Function

Sub Kosong()

TxtRate = Format(0, gs_formatExchangeRate)
cbocurr = ""
LblCurr = ""
DTPicker1.Value = Format(Date, "dd MMM yyyy")
cbocurr.Enabled = True
DTPicker1.Enabled = True
gs_status = "insert"
LblErr = ""

End Sub

Private Sub MPeriod_Change()
Call MPeriod_Click
Browse
TglMP = Mperiod.Month
End Sub

Private Sub MPeriod_Click()
If Mperiod.Month = 1 And Val(TglMP) = 12 Then Mperiod.Year = Mperiod.Year + 1
If Mperiod.Month = 12 And Val(TglMP) = 1 Then Mperiod.Year = Mperiod.Year - 1
LDay = Format(DateSerial(Mperiod.Year, Mperiod.Month + 1, 0), "YYYY-MM-DD")
End Sub

Private Sub TxtRate_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
End If
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtRate_LostFocus()
On Error GoTo errhndl
If IsNumeric(TxtRate) = False Then TxtRate = Format(0, gs_formatExchangeRate)
If TxtRate > gd_MaxExchangeRate Then TxtRate = gd_MaxExchangeRate
If Round(CDbl(TxtRate)) / CDbl(TxtRate) = 1 Then
    TxtRate = Format(CDbl(TxtRate), gs_formatExchangeRate)
Else
    TxtRate = Format(CDbl(TxtRate), gs_formatExchangeRate)
End If
Exit Sub
errhndl:
DoEvents
End Sub
