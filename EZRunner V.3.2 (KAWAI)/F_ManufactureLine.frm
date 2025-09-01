VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.dll"
Begin VB.Form F_ManufactureLine 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manufacture Master"
   ClientHeight    =   10350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15120
   Icon            =   "F_ManufactureLine.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCompanyName 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
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
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   840
      Width           =   6615
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   510
      TabIndex        =   17
      Top             =   9030
      Width           =   14220
      Begin VB.Label Lblpesan 
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
         Left            =   210
         TabIndex        =   18
         Top             =   210
         Width           =   13785
      End
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
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
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   6615
   End
   Begin VB.TextBox TxtCode 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   645
      MaxLength       =   10
      TabIndex        =   4
      Top             =   8400
      Width           =   1560
   End
   Begin VB.TextBox TxtDesc 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   5
      Top             =   8400
      Width           =   6720
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12885
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   405
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.CommandButton CmdManufacture 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sea&rching"
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
      Left            =   10804
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9735
      Width           =   1200
   End
   Begin VB.CommandButton CmdManufacture 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Cancel"
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
      Left            =   12154
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9735
      Width           =   1200
   End
   Begin VB.CommandButton CmdManufacture 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Refresh"
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
      Left            =   9454
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9735
      Width           =   1200
   End
   Begin VB.CommandButton CmdManufacture 
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
      Left            =   514
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9735
      Width           =   1200
   End
   Begin VB.CommandButton CmdManufacture 
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
      Index           =   4
      Left            =   13519
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9735
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5985
      Left            =   525
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1785
      Width           =   14205
      _cx             =   25056
      _cy             =   10557
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
   Begin VB.Line Line2 
      X1              =   3960
      X2              =   10560
      Y1              =   1200
      Y2              =   1200
   End
   Begin MSForms.ComboBox TxtCc 
      Height          =   345
      Left            =   2325
      TabIndex        =   0
      Top             =   840
      Width           =   1560
      VariousPropertyBits=   746604571
      MaxLength       =   10
      DisplayStyle    =   3
      Size            =   "2752;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Code"
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
      Left            =   600
      TabIndex        =   19
      Top             =   900
      Width           =   1635
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   10560
      Y1              =   1605
      Y2              =   1605
   End
   Begin MSForms.ComboBox TxtMc 
      Height          =   345
      Left            =   2325
      TabIndex        =   1
      Top             =   1305
      Width           =   1560
      VariousPropertyBits=   746604571
      MaxLength       =   10
      DisplayStyle    =   3
      Size            =   "2752;609"
      ListRows        =   5
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacture Code"
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
      Left            =   600
      TabIndex        =   16
      Top             =   1350
      Width           =   1635
   End
   Begin VB.Label Isian 
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
      Height          =   255
      Left            =   2115
      TabIndex        =   15
      Top             =   9795
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacture Master"
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
      TabIndex        =   14
      Top             =   240
      Width           =   14205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   2280
      TabIndex        =   13
      Top             =   7995
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   645
      TabIndex        =   12
      Top             =   7995
      Width           =   450
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   615
      Index           =   2
      Left            =   525
      Top             =   8265
      Width           =   8610
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   525
      Top             =   7905
      Width           =   8610
   End
End
Attribute VB_Name = "F_ManufactureLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DtFile, ChangeDelete As Boolean, DtRec As Integer
Dim Dami, DamiUpd As String, Test As String

Dim bteColId As Byte
Dim bteColSelect As Byte
Dim bteColLineCode As Byte
Dim bteColLineName As Byte

Private Sub CmdManufacture_Click(Index As Integer)
    Select Case Index
    Case 4:
        If hakUpdate(Me.Name) = 0 Then LblPesan = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
        If DtRec = 1 Then
            Call DtUpdate
            'DtRec = 3
            txtCode.Text = ""
            txtDesc.Text = ""
            txtCode.Enabled = True
            DtRec = 3
        ElseIf DtRec = 2 Then
            Call DtDelete
            DtRec = 3
        ElseIf DtRec = 3 Then
            Call DtSave
            txtCode.Text = ""
            txtDesc.Text = ""
        End If
    Case 2:
        Call Hidden
        Call ClearS
        LblPesan = ""
    Case 0:
        DoEvents
        frmMainMenu.Show
        DoEvents
        Unload Me
    Case 1: Call Browse
    Case 3: Call Find
    End Select
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblPesan.Caption = ErrMsg
    End If
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = "Manufacture Master"
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    Call CompanyMaster
    Call Header
    Dami = 0
    DtRec = 3
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim TextGrid As String, Data1 As String, rtrec As Integer, i As Integer
    
    TextGrid = grid.Text
    If TextGrid = "S" Then
        Data1 = "S"
        txtCode.Enabled = False
        txtCode = Trim(grid.TextMatrix(Row, bteColLineCode))
        txtDesc = Trim(grid.TextMatrix(Row, bteColLineName))
        Isian = Trim(txtCode)
        DtRec = 1
        Call ClearS
    ElseIf TextGrid = "D" Then
        Data1 = "D"
        DtRec = 2
        Call ClearS("S")
        'Call Hidden
    Else
        For i = 1 To grid.Rows - 1
            If grid.TextMatrix(i, bteColSelect) = "S" Then
                rtrec = 3
                Call ClearS
                Call Hidden
                GoTo Olah
            End If
        Next i
        For i = 1 To grid.Rows - 1
            If grid.TextMatrix(i, bteColSelect) = "D" Then
                rtrec = 2
                GoTo Olah
            End If
        Next i
        DtRec = 3
        Call Hidden
    End If
    
Olah:
    grid.TextMatrix(Row, Col) = TextGrid
    grid.Col = Col
    grid.Row = Row
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

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    If KeyAscii = 39 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtCode_LostFocus()
    Dim sql As String, RsCode As New ADODB.Recordset
    If RsCode.State <> adStateClosed Then RsCode.Close
    RsCode.Open "select * from manufacture_line where Company_Code='" & Trim(TxtCC) & "' and manufacture_code='" & Trim(TxtMc) & "' and line_code='" & Trim(txtCode) & "'", Db, adOpenDynamic, adLockOptimistic, adCmdText
    If Not (RsCode.BOF And RsCode.EOF) Then
        txtDesc = Trim(RsCode("line_name"))
        LblPesan = DisplayMsg(3004)
        Test = Trim(txtCode)
    End If
    Test = Trim(txtCode)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub TxtCc_Change()
    If TxtCC.matchFound Then
        TxtCompanyName = TxtCC.List(TxtCC.ListIndex, 1)
        DamiUpd = Trim(TxtCC)
    Else
        TxtCompanyName = ""
        LblPesan = DisplayMsg(4069)  '"Record is not found"
        DamiUpd = Trim(TxtCC)
    End If
    Call TradeMaster
End Sub

Private Sub TxtMc_Change()
    If TxtMc.matchFound Then
        txtName = TxtMc.List(TxtMc.ListIndex, 1)
        DamiUpd = Trim(TxtMc)
    Else
        txtName = ""
        LblPesan = DisplayMsg(4069)  '"Record is not found"
        DamiUpd = Trim(TxtMc)
    End If
    Call Browse
End Sub

Private Sub TxtMc_Click()
    Call Browse
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    If KeyAscii = 39 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Header()
    
    bteColId = 0
    bteColSelect = 1
    bteColLineCode = 2
    bteColLineName = 3
    
    grid.ColS = 4
    grid.Rows = 1
    
    grid.TextMatrix(0, bteColSelect) = " "
    grid.TextMatrix(0, bteColLineCode) = "Line CD"
    grid.TextMatrix(0, bteColLineName) = "Line Name"
    
    grid.ColWidth(bteColId) = 0
    grid.ColWidth(bteColSelect) = 300
    grid.ColWidth(bteColLineCode) = 1000
    grid.ColWidth(bteColLineName) = 6600
    
    grid.ColAlignment(bteColSelect) = flexAlignCenterCenter
    grid.ColAlignment(bteColLineCode) = flexAlignLeftCenter
    grid.ColAlignment(bteColLineName) = flexAlignLeftCenter
    
    grid.EditMaxLength = 1

End Sub

Private Sub DtSave()
    Dim sql As String, RsSave As New ADODB.Recordset, LblInput As String
    Dim rsCek As New ADODB.Recordset
    
    Call OlahDt
    If txtCode <> "" And txtDesc <> "" And TxtMc <> "" Then
        sql = "select * from trade_master where trade_code='" & Trim(TxtMc) & "'"
        If rsCek.State <> adStateClosed Then rsCek.Close
        rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic
        If (rsCek.BOF And rsCek.EOF) Then
            LblPesan = DisplayMsg(1052)
            Exit Sub
        End If
        sql = "select * from manufacture_line where Company_Code ='" & Trim(TxtCC) & "' and manufacture_code='" & Trim(TxtMc) & "' and line_code='" & Trim(txtCode) & "'"
        If RsSave.State <> adStateClosed Then RsSave.Close
        RsSave.Open sql, Db, adOpenDynamic, adLockOptimistic
        If Not (RsSave.BOF And RsSave.EOF) Then
            LblInput = MsgBox("Do you really to update line code ?", vbYesNo + vbQuestion, "Confirmation")
            If LblInput = vbYes Then
                sql = "update manufacture_line set line_name='" & Trim(txtDesc) & "', Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                    "where Company_Code ='" & Trim(TxtCC) & "' and manufacture_code='" & Trim(TxtMc) & "' and line_code='" & Trim(txtCode) & "'"
                Db.Execute (sql)
                LblPesan = DisplayMsg(1101)
            Else
                LblPesan = ""
            End If
        Else
            sql = "insert manufacture_line(Company_Code, Manufacture_Code, Line_Code, Line_Name, Last_Update, Last_User) values('" & Trim(TxtCC) & "', '" & Trim(TxtMc) & "','" & Trim(txtCode) & "', '" & Trim(txtDesc) & "', getdate(), '" & userLogin & "')"
            Db.Execute (sql)
            LblPesan = DisplayMsg(1000)
        End If
    End If
    Call Browse
End Sub

Private Sub Browse()
    Dim sql As String, RsBros As New ADODB.Recordset
    Call Header
    If RsBros.State <> adStateClosed Then RsBros.Close
    RsBros.Open "select * from manufacture_line where Company_Code ='" & Trim(TxtCC) & "' and manufacture_code='" & Trim(TxtMc) & "'", Db, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not RsBros.EOF
        grid.AddItem ""
        grid.TextMatrix(grid.Rows - 1, bteColSelect) = ""
        grid.TextMatrix(grid.Rows - 1, bteColLineCode) = Trim(RsBros("line_code"))
        grid.TextMatrix(grid.Rows - 1, bteColLineName) = Trim(RsBros("line_name"))
        grid.Cell(flexcpBackColor, grid.Rows - 1, bteColSelect) = &HFFFFFF
        RsBros.MoveNext
    Loop
    Call Clearin
    Call HeaderText
End Sub

Private Sub HeaderText()
    Dim i As Integer
    For i = 1 To grid.Rows - 1
        If Trim(grid.TextMatrix(i, bteColLineCode)) = Trim(Test) Then
            grid.Row = i
            grid.SetFocus
            grid.TopRow = i - 1
        End If
    Next i
End Sub

Private Sub Clearin()
    txtCode = ""
    txtDesc = ""
    txtCode.Enabled = True
End Sub

Private Sub DtUpdate()
    Dim sql As String, RsUd As New ADODB.Recordset
    sql = "update manufacture_line set line_name='" & Trim(txtDesc) & "', Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
        "where Company_Code ='" & Trim(TxtCC) & "' and manufacture_code='" & Trim(TxtMc) & "' and line_code='" & Trim(txtCode) & "'"
    Db.Execute (sql)
    LblPesan = DisplayMsg(1101)
    Test = Trim(txtCode)
    Call Browse
End Sub

Private Sub ClearS(Optional C As String)
    Dim i As Integer
    grid.Col = bteColSelect
    If C <> "" Then
        For i = 1 To grid.Rows - 1
            grid.Row = i
            If grid.Text = C Then grid.Text = ""
        Next i
    Else
        For i = 1 To grid.Rows - 1
            grid.Row = i
            grid.Text = ""
        Next i
    End If
End Sub

Private Sub DtDelete()
    Dim sql As String, RsDel As New ADODB.Recordset
    Dim i As Integer, IMCodeAda, IMLineAda As String, LblInput As String
    
    IMCodeAda = ""
    IMLineAda = ""
    
    LblInput = MsgBox("Do you really to delete line code ?", vbYesNo + vbQuestion, "Confirmation")
    If LblInput = vbYes Then
        For i = 1 To grid.Rows - 1
            If grid.TextMatrix(i, bteColSelect) = "D" Then
                sql = "select * from item_master where manufacture_code='" & Trim(TxtMc) & "' and line_code='" & Trim(grid.TextMatrix(i, bteColLineCode)) & "'"
                If RsDel.State <> adStateClosed Then RsDel.Close
                RsDel.Open sql, Db, adOpenDynamic, adLockOptimistic
                If Not (RsDel.BOF And RsDel.EOF) Then
                    IMCodeAda = IMCodeAda & " " & RsDel("manufacture_code") & ","
                    IMLineAda = IMLineAda & " " & RsDel("line_code") & ","
                Else
                    sql = "delete manufacture_line where Company_Code ='" & Trim(TxtCC) & "' and manufacture_code='" & Trim(TxtMc) & "' and line_code='" & Trim(grid.TextMatrix(i, bteColLineCode)) & "'"
                    Db.Execute (sql)
                End If
            End If
        Next i
        If IMCodeAda <> "" Then
            LblPesan = "Update Failed!. This record is used in table 'Item Master'"
        Else
            LblPesan = DisplayMsg(1201)
        End If
        Call Browse
        Call Isi(IMLineAda)
    Else
        Call ClearS
        Call Hidden
    End If
End Sub

Private Sub Isi(DtLine$)
    Dim i, X As Integer, LmNo, DLine, Panjang, TLine As String
    
    LmNo = ""
    Panjang = 0
    For i = 1 To grid.Rows - 1
        TLine = ""
        For X = 1 To Len(Trim(DtLine)) - Panjang
            DLine = Mid(Trim(DtLine), X, 1)
            If DLine = "," Then
                Panjang = Len(TLine + DLine)
                GoTo Masuk
            Else
                TLine = TLine + DLine
            End If
        Next X
    
Masuk:
        If grid.TextMatrix(i, bteColLineCode) = TLine Then
            grid.TextMatrix(i, bteColSelect) = "D"
        End If
    Next i
End Sub

Private Sub OlahDt()
    If TxtCC = "" Then
        LblPesan = DisplayMsg(1052)  '"Please input Company code!"
        TxtCC.SetFocus
        Dami = 1
        Exit Sub
    ElseIf TxtMc = "" Then
        LblPesan = DisplayMsg(1052)  '"Please input manufacture code!"
        TxtMc.SetFocus
        Dami = 1
        Exit Sub
    ElseIf txtCode = "" Then
        LblPesan = DisplayMsg(1041)  '"Please input line code!"
        txtCode.SetFocus
        Dami = 1
        Exit Sub
    ElseIf txtDesc = "" Then
        LblPesan = DisplayMsg(1053)  '"Please input line name!"
        txtDesc.SetFocus
        Dami = 1
        Exit Sub
    End If
End Sub

Private Sub Find()
    Dim SSql As String, RsFind As New ADODB.Recordset, RsFind1 As New ADODB.Recordset
    Dim LblInput As String
    
    LblInput = InputBox("Input manufacture code or line code", "Search")
    If InStr(1, LblInput, "'") > 0 Then
        LblPesan = DisplayMsg(4073)   '"Record is not found"
        Exit Sub
    End If
    If vbOK Then
        Call Header
        sql = "select * from manufacture_line where Company_Code ='" & Trim(TxtCC) & "' and manufacture_code='" & Trim(TxtMc) & "' and line_code like '" & Trim(LblInput) & "%'"
        If RsFind.State <> adStateClosed Then RsFind.Close
        RsFind.Open sql, Db, adOpenDynamic, adLockOptimistic, adCmdText
        If Not (RsFind.BOF And RsFind.EOF) Then
            Do While Not RsFind.EOF
                grid.AddItem ""
                grid.TextMatrix(grid.Rows - 1, bteColSelect) = ""
                grid.TextMatrix(grid.Rows - 1, bteColLineCode) = Trim(RsFind("line_code"))
                grid.TextMatrix(grid.Rows - 1, bteColLineName) = Trim(RsFind("line_name"))
                grid.Cell(flexcpBackColor, grid.Rows - 1, bteColSelect) = &HFFFFFF
                RsFind.MoveNext
            Loop
            LblPesan = ""
        Else
            LblPesan = DisplayMsg(4073) '"Record is not found!"
        End If
    Else
        Call Browse
        LblPesan = ""
    End If
End Sub

Private Sub CompanyMaster()
    Dim sql As String, rsCompany As New ADODB.Recordset
    Dim i As Integer
    
    If rsCompany.State <> adStateClosed Then rsCompany.Close
    rsCompany.CursorLocation = adUseClient
    rsCompany.Open "Company_Profile order by Company_Code asc", Db, adOpenDynamic, adLockOptimistic, adCmdTable
    TxtCC.columnCount = 2
    TxtCC.TextColumn = 1
    i = 0
    Do While Not rsCompany.EOF
        TxtCC.AddItem ""
        TxtCC.List(i, 0) = Trim(rsCompany("Company_Code"))
        TxtCC.List(i, 1) = Trim(rsCompany("Company_Name"))
        i = i + 1
        rsCompany.MoveNext
    Loop
    TxtCC.ColumnWidths = "50 pt; 300 pt"
    TxtCC.ListWidth = 350
    TxtCC.ListRows = 15
End Sub

Private Sub TradeMaster()
    TxtMc.clear
    Dim sql As String, rstrade As New ADODB.Recordset
    Dim i As Integer
    
    If rstrade.State <> adStateClosed Then rstrade.Close
    rstrade.CursorLocation = adUseClient
    rstrade.Open "trade_master  where trade_cls='1' order by trade_code asc", Db, adOpenDynamic, adLockOptimistic, adCmdTable
    TxtMc.columnCount = 2
    TxtMc.TextColumn = 1
    i = 0
    Do While Not rstrade.EOF
        TxtMc.AddItem ""
        TxtMc.List(i, 0) = Trim(rstrade("trade_code"))
        TxtMc.List(i, 1) = Trim(rstrade("trade_name"))
        i = i + 1
        rstrade.MoveNext
    Loop
    TxtMc.ColumnWidths = "50 pt; 300 pt"
    TxtMc.ListWidth = 350
    TxtMc.ListRows = 15
End Sub

Private Sub SiPutih()
    Dim i As Integer
    For i = 1 To grid.Rows - 1
        grid.Cell(flexcpBackColor, i, bteColSelect) = &HFFFFFF
    Next
End Sub

Private Sub Hidden()
    txtCode = ""
    txtDesc = ""
    txtCode.Enabled = True
    txtCode.SetFocus
End Sub
