VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmPOSetMaster 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Purchase Set Maste"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPOSetMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7620
      TabIndex        =   3
      Top             =   8340
      Width           =   1455
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
      Height          =   375
      Index           =   4
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9960
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Index           =   2
      Left            =   11190
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9960
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Index           =   0
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9960
      Width           =   1140
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   1
      Left            =   13725
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9960
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   615
      Left            =   450
      TabIndex        =   12
      Top             =   9000
      Width           =   14430
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
         Height          =   330
         Left            =   90
         TabIndex        =   13
         Top             =   180
         Width           =   14220
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1365
      Left            =   450
      TabIndex        =   8
      Top             =   930
      Width           =   14400
      Begin VB.TextBox txtdesc 
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   810
         Width           =   5085
      End
      Begin MSForms.ComboBox CboSetCode 
         Height          =   315
         Left            =   2280
         TabIndex        =   0
         Top             =   270
         Width           =   1485
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2619;556"
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
         Caption         =   "Purchase Description"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   10
         Top             =   870
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Set Code"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   9
         Top             =   345
         Width           =   1635
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5115
      Left            =   450
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2430
      Width           =   14415
      _cx             =   25426
      _cy             =   9022
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
      Left            =   12975
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   150
      Width           =   1875
      _extentx        =   3307
      _extenty        =   714
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   7350
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   8400
      Width           =   5070
   End
   Begin MSForms.ComboBox cboitem_code 
      Height          =   315
      Left            =   750
      TabIndex        =   2
      Top             =   8340
      Width           =   1485
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2619;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Index           =   3
      Left            =   8130
      TabIndex        =   16
      Top             =   7920
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
      Height          =   195
      Index           =   0
      Left            =   750
      TabIndex        =   14
      Top             =   7920
      Width           =   915
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      Height          =   1035
      Index           =   0
      Left            =   450
      Top             =   7860
      Width           =   14415
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Set Master"
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
      Height          =   330
      Left            =   6375
      TabIndex        =   11
      Top             =   150
      Width           =   2280
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   330
      Index           =   1
      Left            =   450
      Top             =   7860
      Width           =   14415
   End
End
Attribute VB_Name = "FrmPOSetMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Idx As Long, HakU As Byte
Dim stopmsg As Boolean, X As Long, temp(10000) As String

Dim bteColSelect As Byte
Dim bteColItem_Code As Byte
Dim bteColItem_Name As Byte
Dim bteColQty As Byte

Sub Header()
    
    bteColSelect = 0
    bteColItem_Code = 1
    bteColItem_Name = 2
    bteColQty = 3
    
    With grid
        .ColS = 4
        .Rows = 1
        .EditMaxLength = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColItem_Code) = "Item Code"
        .TextMatrix(0, bteColItem_Name) = "Item Name"
        .TextMatrix(0, bteColQty) = "Qty"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColItem_Code) = 1700
        .ColWidth(bteColItem_Name) = 5000
        .ColWidth(bteColQty) = 2500
        
        
        .Cell(flexcpAlignment, 0, 0, 0, 3) = flexAlignCenterCenter
        
    End With
End Sub

Private Sub cbotrade_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboitem_code_Change()
If cboitem_code.ListIndex <> -1 Then
   LblDesc.Caption = cboitem_code.Column(1)
Else
   LblDesc.Caption = ""
   txtQty.Text = Format(0, gs_formatQty)
End If
End Sub

Private Sub cboitem_code_KeyPress(KeyAscii As MSForms.ReturnInteger)
    LblErrMsg = ""
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CboSetCode_Change()
If CboSetCode.ListIndex <> -1 Then
    Call search
Else
    txtDesc.Text = ""
    cboitem_code.ListIndex = -1
    LblDesc.Caption = ""
    txtQty.Text = ""
    LblErrMsg = ""
    Header
End If
End Sub

Private Sub CboSetCode_KeyPress(KeyAscii As MSForms.ReturnInteger)
    LblErrMsg = ""
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cmdAction_Click(Index As Integer)
Select Case Index
    Case 0
        Unload Me
        frmMainMenu.Show
    Case 1
        If HakU = 0 Then _
            LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
        Call Submit
        Call search
    Case 2
        Header
        clear
    Case 3
    Case 4
        clearmark ("S")
        clearmark ("D")

End Select
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Sub adtocombo()
Dim rstcust As Recordset
sql = "SELECT  rtrim(Item_Code) Item_Code, rtrim(Item_Name) Item_Name " & _
    "From Item_Master where MakeBuy_Cls='02' "
Set rstcust = New Recordset
rstcust.Open sql, Db, adOpenKeyset, adLockOptimistic
With cboitem_code
    .clear
    .columnCount = 2
    .ColumnWidths = "55 pt;120 pt"
    .ListWidth = 300
    .ListRows = 15
    
Idx = 0
Do While Not rstcust.EOF
    .AddItem ""
    .List(Idx, 0) = Trim(rstcust!Item_Code)
    .List(Idx, 1) = Trim(rstcust!item_name)
    
    Idx = Idx + 1
    rstcust.MoveNext
Loop
End With

'Add Parent Code

sql = "SELECT  Distinct Parent_SetCode Item_Code " & _
    "From PO_Set_Master"
Set rstcust = New Recordset
rstcust.Open sql, Db, adOpenKeyset, adLockOptimistic
With CboSetCode
    .clear
    .columnCount = 1
    .ColumnWidths = "50 pt;15 pt"
    .ListWidth = 200
    .ListRows = 15
Idx = 0
Do While Not rstcust.EOF
    .AddItem ""
    .List(Idx, 0) = Trim(rstcust!Item_Code)
    Idx = Idx + 1
    rstcust.MoveNext
Loop
End With

End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
adtocombo
txtQty.Text = Format(0, gs_formatQty)
'clear
Header
HakU = hakUpdate(Me.Name)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
End Sub

Sub clear()
CboSetCode.ListIndex = -1
txtDesc.Text = ""
cboitem_code.ListIndex = -1
LblDesc.Caption = ""
txtQty.Text = ""
LblErrMsg = ""
Header
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = True
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim temp As String

temp = grid.Text
If temp = "S" Then
clearmark (temp)
cboitem_code.Text = Trim(grid.TextMatrix(Row, bteColItem_Code))
LblDesc.Caption = Trim(grid.TextMatrix(Row, bteColItem_Name))
txtQty.Text = Trim(grid.TextMatrix(Row, bteColQty))
cboitem_code.Enabled = False
ElseIf temp = "D" Then
    clearmark (temp)
    cboitem_code.ListIndex = -1
End If
grid.TextMatrix(Row, Col) = temp
grid.Col = Col
grid.Row = Row
LblErrMsg.Caption = ""
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
If KeyAscii = Asc(".") Then KeyAscii = 0
End If
End Sub

Private Sub clearmark(nilai$)
Dim k As Long
If nilai = "D" Then
For k = 1 To grid.Rows - 1
    If grid.TextMatrix(k, bteColSelect) = "S" Then
        grid.TextMatrix(k, bteColSelect) = ""
    End If
Next
Else
For k = 1 To grid.Rows - 1
        grid.TextMatrix(k, bteColSelect) = ""
Next
End If
cboitem_code.Enabled = True
End Sub



Private Sub Submit()
Dim Q As String, X As Integer
Dim LS_Raff As Integer, LS_Raff1 As Integer, LS_Raff2 As Integer, LS_Raff3 As Integer

If MsgInsert = False Then
    LblErrMsg.Caption = MdlEZRunner.DisplayMsg("1023")  'Data Already Exists
    Exit Sub
End If

LblErrMsg.Caption = ""
Me.MousePointer = vbHourglass

If grid.FindRow("S", 0, 0, , True) > 0 Then
    
            Q = " UPDATE [PO_Set_Master] " & vbCrLf & _
                        "    SET [Description] ='" & Trim(txtDesc.Text) & "' " & vbCrLf & _
                        "       ,[Qty] = '" & CDbl(IIf(txtQty.Text = "", 0, txtQty.Text)) & "' " & vbCrLf & _
                        "       ,[Last_Update] = Getdate() " & vbCrLf & _
                        "       ,[Last_User] = '" & userLogin & "' " & vbCrLf & _
                        "  WHERE [Parent_SetCode]='" & Trim(CboSetCode) & "' And Item_Code='" & cboitem_code & "' "
            Db.Execute Q, LS_Raff
Else
    
    If LS_Raff = 0 And cboitem_code.ListIndex <> -1 Then
            Q = " INSERT INTO [PO_Set_Master] " & vbCrLf & _
                    "            ([Parent_SetCode] " & vbCrLf & _
                    "            ,[Description] " & vbCrLf & _
                    "            ,[Item_Code] " & vbCrLf & _
                    "            ,[Item_Name] " & vbCrLf & _
                    "            ,[Qty] " & vbCrLf & _
                    "            ,[Register_Date] " & vbCrLf & _
                    "            ,[Register_User] " & vbCrLf & _
                    "           ) " & vbCrLf & _
                    "      VALUES " & vbCrLf & _
                    "            ('" & Trim(CboSetCode.Text) & "' "
            
            Q = Q + "            ,'" & Trim(txtDesc.Text) & "' " & vbCrLf & _
                    "            ,'" & Trim(cboitem_code.Text) & "' " & vbCrLf & _
                    "            ,'" & Trim(LblDesc.Caption) & "' " & vbCrLf & _
                    "            ,'" & CDbl(IIf(txtQty.Text = "", 0, txtQty.Text)) & "'" & vbCrLf & _
                    "            ,Getdate()" & vbCrLf & _
                    "            ,'" & userLogin & "' " & vbCrLf & _
                    "          ) "
            Db.Execute Q, LS_Raff1
    End If

End If

    For X = 1 To grid.Rows - 1
    
        If grid.TextMatrix(X, bteColSelect) = "D" Then
            Db.Execute "Delete PO_Set_Master Where Parent_SetCode='" & Trim(CboSetCode.Text) & "' And Item_Code='" & Trim(grid.TextMatrix(X, bteColItem_Code)) & "'"
        End If
        
    Next X

    If LS_Raff <> 0 Then
        LblErrMsg.Caption = DisplayMsg("1101")
        cboitem_code.ListIndex = -1
        cboitem_code.Enabled = True
        LblDesc.Caption = ""
        txtQty.Text = Format(0, gs_formatQty)
    ElseIf LS_Raff1 <> 0 Then
        LblErrMsg.Caption = DisplayMsg("1000")
        cboitem_code.ListIndex = -1
        cboitem_code.Enabled = True
        LblDesc.Caption = ""
        txtQty.Text = Format(0, gs_formatQty)
        
    ElseIf LS_Raff2 <> 0 Then
        LblErrMsg.Caption = DisplayMsg("1201")
    End If
    
    
 Me.MousePointer = vbDefault

    
End Sub
 
Function MsgInsert() As Boolean
Dim sql As String
Dim RS As New Recordset

sql = " SELECT * FROM [PO_Set_Master]" & vbCrLf & _
      " WHERE [Parent_SetCode]='" & Trim(CboSetCode) & "' And Item_Code='" & cboitem_code & "' "
If RS.State = 1 Then RS.Close
RS.Open sql, Db, adOpenKeyset, adLockOptimistic

If RS.RecordCount > 0 Then
    MsgInsert = False
Else
    MsgInsert = True
End If


End Function

Private Sub search()
Dim Q As String, i As Integer
Dim RsQ As New ADODB.Recordset


Q = "Select * From PO_SET_Master Where Parent_SetCode='" & Trim(CboSetCode.Text) & "'"

If RsQ.State <> adStateClosed Then RsQ.Close
RsQ.Open Q, Db, adOpenDynamic, adLockOptimistic

Header
With grid

i = 1
Do While Not RsQ.EOF
   txtDesc = IIf(IsNull(RsQ("Description")), "", Trim(RsQ("Description")))
  .Rows = .Rows + 1
  
  .TextMatrix(i, bteColItem_Code) = Trim(RsQ("Item_COde"))
  .TextMatrix(i, bteColItem_Name) = Trim(RsQ("Item_name"))
  .TextMatrix(i, bteColQty) = Format(RsQ("QTY"), "#,##0.00")
  
  grid.Cell(flexcpAlignment, i, bteColItem_Code, i, bteColItem_Name) = flexAlignLeftCenter
  .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
  i = i + 1
  RsQ.MoveNext
Loop

End With


End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    LblErrMsg = ""
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub


Private Sub txtQty_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
End If
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtQty_LostFocus()
txtQty.Text = Format(txtQty.Text, gs_formatQty)
End Sub
