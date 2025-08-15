VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Begin VB.Form F_Manufacture 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manufacture Master"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin NIC.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   6930
      TabIndex        =   16
      Top             =   1020
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
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6570
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
      Left            =   6210
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6570
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
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6570
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
      Left            =   570
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6570
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
      Left            =   7575
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6570
      Width           =   1200
   End
   Begin VB.TextBox TxtName 
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
      Left            =   5918
      MaxLength       =   20
      TabIndex        =   3
      Top             =   5190
      Width           =   2640
   End
   Begin VB.TextBox TxtLine 
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
      Left            =   4973
      MaxLength       =   3
      TabIndex        =   2
      Top             =   5190
      Width           =   885
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
      Left            =   1778
      MaxLength       =   30
      TabIndex        =   1
      Top             =   5190
      Width           =   3135
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
      Left            =   833
      MaxLength       =   6
      TabIndex        =   0
      Top             =   5190
      Width           =   885
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grid 
      Height          =   3030
      Left            =   585
      TabIndex        =   17
      Top             =   1530
      Width           =   8190
      _cx             =   14446
      _cy             =   5345
      _ConvInfo       =   1
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
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
      Left            =   2160
      TabIndex        =   18
      Top             =   6615
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label LblPesan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   330
      Left            =   660
      TabIndex        =   11
      Top             =   5955
      Width           =   8025
   End
   Begin MSForms.Label Label5 
      Height          =   555
      Left            =   570
      TabIndex        =   10
      Top             =   5850
      Width           =   8205
      BackColor       =   16637923
      Size            =   "14473;979"
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacture Master"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   540
      TabIndex        =   9
      Top             =   435
      Width           =   8250
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Line Name"
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
      Left            =   5925
      TabIndex        =   8
      Top             =   4785
      Width           =   2640
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Line Code"
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
      Left            =   4950
      TabIndex        =   7
      Top             =   4785
      Width           =   885
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   1785
      TabIndex        =   6
      Top             =   4785
      Width           =   3135
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   4785
      Width           =   885
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   570
      Top             =   5085
      Width           =   8205
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   570
      Top             =   4725
      Width           =   8205
   End
End
Attribute VB_Name = "F_Manufacture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'<<<<<<<<<<<<<<<<<<<<<<< >>>>>>>>>>>>>>>>>>>>>>>
'                 Ny.Risdiana
'     Penyerahan Pekerjaan Tgl 16 Juli 2003
'          Selesai Tgl 16 Juli 2003
' Bener atau tidak pekerjaan yg penting sombong
'<<<<<<<<<<<<<<<<<<<<<<< >>>>>>>>>>>>>>>>>>>>>>>

Dim Data1 As String, DSave As Boolean
Dim Isi As String

Private Sub CmdManufacture_Click(Index As Integer)
Select Case Index
    Case 0: 'Sub menu
            DoEvents
            frmMainMenu.Show
            frmMainMenu.VerticalMenu1.MenuCur = 1
            DoEvents
            Unload Me
    Case 1: 'Refresh
            Call Browse
            LblPesan = ""
    Case 2: 'Cancel
            Call Clean
            Call ClearS
    Case 3: 'Searching
            Call Find
    Case 4: 'Submit
            If Data1 = "" Then Call DataSave
            If Data1 = "S" Then Call DataUpdate
            If Data1 = "D" Then Call DataDelete
End Select
End Sub

Private Sub Form_Load()

CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

Call Browse
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim TextGrid As String

TextGrid = grid.Text

If TextGrid = "S" Then
   Data1 = "S"
   TxtCode.Enabled = False
   TxtLine.Enabled = False
   TxtCode = grid.TextMatrix(Row, 2)
   TxtDesc = grid.TextMatrix(Row, 3)
   TxtLine = grid.TextMatrix(Row, 4)
   TxtName = grid.TextMatrix(Row, 5)
   Isian = Trim(TxtCode)
   Call ClearS
Else
   Data1 = "D"
   Call ClearS("S")
End If

grid.TextMatrix(Row, Col) = TextGrid
grid.Col = Col
grid.Row = Row
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid.Col > 1 Then Cancel = True
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If grid.Col = 1 Then
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

Private Sub Header()
With grid
    .Cols = 6
    .Rows = 1
    
    .TextMatrix(0, 2) = "Code"
    .TextMatrix(0, 3) = "Description"
    .TextMatrix(0, 4) = "Line Cd"
    .TextMatrix(0, 5) = "Line Name"
    
    .ColWidth(0) = 0
    .ColWidth(1) = 250
    .ColWidth(2) = 900
    .ColWidth(3) = 3200
    .ColWidth(4) = 900
    .ColWidth(5) = 2600
    
    .ColAlignment(2) = 1
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(5) = 1
    .EditMaxLength = 1
End With
End Sub

Private Sub cek()
Dim Sql As String, rsCek As New ADODB.Recordset

Sql = "select * from manufacture_master " & _
      "where manufacture_code='" & Trim(TxtCode) & "' " & _
      "and line_code='" & Trim(TxtLine) & "'"
If rsCek.State <> adStateClosed Then rsCek.Close
rsCek.Open Sql, Db, adOpenDynamic, adLockOptimistic, adCmdText
If Not (rsCek.BOF And rsCek.EOF) Then
   LblPesan = "Data already exist"
   Call Clean
   DSave = True
Else
   DSave = False
End If
End Sub

Private Sub Clean()
TxtCode = ""
TxtLine = ""
TxtDesc = ""
TxtName = ""
TxtCode.Enabled = True
TxtLine.Enabled = True
TxtCode.SetFocus
End Sub

Private Sub TxtCode_LostFocus()
If TxtCode <> "" Then
   Isian = TxtCode.Text
End If
Isi = TxtCode.Text
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub TxtLine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call cek
'Grid.Row = Grid.FindRow(TxtWHCode, FixedRows, 1, False)
If DSave = False Then SendKeys vbTab
End If
If KeyAscii = 39 Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub ClearS(Optional C As String)
Dim i As Integer

grid.Col = 1
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

Private Sub Find()
Dim SSql As String, RsFind As New ADODB.Recordset, RsFind1 As New ADODB.Recordset
Dim Lblinput As String

Lblinput = InputBox("Input manufacture code or line code", "Search")
If vbOK Then
   Call Header
   Sql = "select * from manufacture_master where manufacture_code like '" & Trim(Lblinput) & "%'"
   If RsFind.State <> adStateClosed Then RsFind.Close
   RsFind.Open Sql, Db, adOpenDynamic, adLockOptimistic, adCmdText
   If Not (RsFind.BOF And RsFind.EOF) Then
        i = 0
        Do While Not RsFind.EOF
           grid.AddItem i & _
           Chr(9) & "" & _
           Chr(9) & RsFind("manufacture_code") & _
           Chr(9) & RsFind("manufacture_name") & _
           Chr(9) & Trim(RsFind("line_code")) & _
           Chr(9) & RsFind("Line_name")
           i = i + 1
           RsFind.MoveNext
        Loop
        LblPesan = ""
   Else
        Call Header
        Sql = "select * from manufacture_master where line_code like '" & Trim(Lblinput) & "%'"
        If RsFind1.State <> adStateClosed Then RsFind1.Close
        RsFind1.Open Sql, Db, adOpenDynamic, adLockOptimistic, adCmdText
        If Not (RsFind1.BOF And RsFind1.EOF) Then
             i = 0
             Do While Not RsFind1.EOF
                grid.AddItem i & _
                Chr(9) & "" & _
                Chr(9) & RsFind1("manufacture_code") & _
                Chr(9) & RsFind1("manufacture_name") & _
                Chr(9) & RsFind1("line_code") & _
                Chr(9) & RsFind1("Line_name")
                i = i + 1
                RsFind1.MoveNext
             Loop
             LblPesan = ""
        Else
            LblPesan = "Record is not found !"
        End If
   End If
Else
End If
End Sub

Private Sub Browse()
Dim Sql As String, RsB As New ADODB.Recordset
Dim i As Integer

Call Header
If RsB.State <> adStateClosed Then RsB.Close
RsB.Open "manufacture_master order by manufacture_code", Db, adOpenDynamic, adLockOptimistic, adCmdTable
i = 0
Do While Not RsB.EOF
   grid.AddItem i & _
   Chr(9) & "" & _
   Chr(9) & RsB("manufacture_code") & _
   Chr(9) & RsB("manufacture_name") & _
   Chr(9) & Trim(RsB("line_code")) & _
   Chr(9) & RsB("line_name")
   i = i + 1
   RsB.MoveNext
   DoEvents
Loop

For i = 1 To grid.Rows - 1
    grid.Cell(flexcpBackColor, i, 1) = vbWhite
Next i
End Sub

Private Sub DataSave()
Dim Sql As String, RsSave As New ADODB.Recordset

If TxtCode.Text = "" Then
   LblPesan = "Please input code !"
   TxtCode.SetFocus
   Exit Sub
ElseIf TxtDesc.Text = "" Then
   LblPesan = "Please input Description !"
   TxtDesc.SetFocus
   Exit Sub
ElseIf TxtLine.Text = "" Then
   LblPesan = "Please input line code !"
   TxtLine.SetFocus
   Exit Sub
ElseIf TxtName.Text = "" Then
   LblPesan = "Please input line name !"
   TxtName.SetFocus
   Exit Sub
End If

Sql = "insert into manufacture_master(manufacture_code,manufacture_name," & _
      "line_code,line_name) values('" & Trim(TxtCode) & "'," & _
      "'" & Trim(TxtDesc) & "','" & Trim(TxtLine) & "'," & _
      "'" & Trim(TxtName) & "')"
If RsSave.State <> adStateClosed Then RsSave.Close
RsSave.Open Sql, Db, adOpenDynamic, adLockOptimistic
LblPesan = "Data Saved Success !"
Call Browse
Call Clean
Isi = ""
'Grid.Row = Grid.FindRow(TxtCode, FixedRows, 2, False)
'Grid.SetFocus
Call CariText
Isi = ""
'Call Clean
End Sub

Private Sub DataUpdate()
Dim SSql As String, RsU As New ADODB.Recordset

Data1 = ""
If RsU.State <> adStateClosed Then RsU.Close
SSql = "update manufacture_master set manufacture_name='" & Trim(TxtDesc) & "'," & _
      "line_name='" & Trim(TxtName) & "' where manufacture_code='" & TxtCode & "' " & _
      "and line_code='" & TxtLine & "'"
RsU.Open SSql, Db, adOpenDynamic, adLockOptimistic
LblPesan = "Update Record Success !"
Call Browse
Call Clean
Call CariText
End Sub

Private Sub DataDelete()
Dim Sql As String, RsD As New ADODB.Recordset, RsM As New ADODB.Recordset
Dim Master, LMaster As String, Manu, LManu As String, i As Integer
Dim Lblinput As String

Master = ""
Manu = ""
Data1 = ""
Lblinput = MsgBox("Do You really want to delete this record ?", vbYesNo, "Delete")
If Lblinput = vbYes Then
For i = 1 To grid.Rows - 1
    If grid.TextMatrix(i, 1) = "D" Then
       Sql = "select * from item_master where manufacture_code='" & grid.TextMatrix(i, 2) & "' " & _
             "and line_code='" & grid.TextMatrix(i, 4) & "'"
       If RsM.State <> adStateClosed Then RsM.Close
       RsM.Open Sql, Db, adOpenDynamic, adLockOptimistic
       If Not (RsM.BOF And RsM.EOF) Then
          Master = Master & "'" & grid.TextMatrix(i, 2) & "',"
          LMaster = LMaster & "'" & grid.TextMatrix(i, 4) & "',"
       Else
          Manu = Manu & "'" & grid.TextMatrix(i, 2) & "',"
          LManu = LManu & "'" & grid.TextMatrix(i, 4) & "',"
       End If
    End If
Next i
Else
Call Browse
Call Clean
Exit Sub
End If
If Manu <> "" And Master = "" Then 'Yg tidak terdapat di tabel item master
   If RsM.State <> adStateClosed Then RsM.Close
   Sql = "delete manufacture_master where manufacture_code in(" & Mid(Manu, 1, Len(Manu) - 1) & ") " & _
         "and line_code in(" & Mid(LManu, 1, Len(LManu) - 1) & ")"
   RsM.Open Sql, Db, adOpenDynamic, adLockOptimistic
   LblPesan = "Delete Record Success !"
   Call Browse
   Call Clean
ElseIf Manu <> "" And Master <> "" Then 'Record yg terdapat oleh tabel item master dan manufacture master
   If RsM.State <> adStateClosed Then RsM.Close
   Sql = "delete manufacture_master where manufacture_code in(" & Mid(Manu, 1, Len(Manu) - 1) & ") " & _
         "and line_code in(" & Mid(LManu, 1, Len(LManu) - 1) & ")"
   RsM.Open Sql, Db, adOpenDynamic, adLockOptimistic
   LblPesan = "Delete Failed!. This record is used in table 'Item Master'"
   Call Browse
   Call Clean
   GoTo TampilD
ElseIf Manu = "" And Master <> "" Then 'Record yg terdapat hanya di item master
   LblPesan = "Delete Failed!. This record is used in table 'Item Master'"
   Call Clean
TampilD:
   Call DisplayD(Mid(Master, 1, Len(Master) - 1), Mid(LMaster, 1, Len(LMaster) - 1))
End If
End Sub

Private Sub DisplayD(DMaster As String, DLine As String)
Dim Sql As String, RsD As New ADODB.Recordset
Dim Master, Linee As String

If RsD.State <> adStateClosed Then RsD.Close
Sql = "select * from item_master where manufacture_code in(" & DMaster & ") and " & _
      "line_code in(" & DLine & ")"
RsD.Open Sql, Db, adOpenDynamic, adLockOptimistic

Do While Not RsD.EOF
   i = 0
Ulang:
   i = i + 1
   Master = RsD("manufacture_code")
   Linee = RsD("line_code")
   If grid.TextMatrix(i, 2) = Trim(Master) And grid.TextMatrix(i, 4) = Trim(Linee) Then
      grid.TextMatrix(i, 1) = "D"
   Else
      GoTo Ulang
   End If
   RsD.MoveNext
Loop
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub CariText()
Dim i As Integer

For i = 1 To grid.Rows - 1
    If Trim(grid.TextMatrix(i, 2)) = Trim(Isian) Then
       grid.Row = i
       grid.SetFocus
       grid.TopRow = i - 1
    End If
Next i
End Sub
