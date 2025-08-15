VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmCostMaster 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Cost Classification Master"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmCostMaster.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   270
      TabIndex        =   16
      Top             =   9120
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
         Left            =   90
         TabIndex        =   17
         Top             =   180
         Width           =   14325
      End
   End
   Begin VB.TextBox txtcost 
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
      Left            =   525
      MaxLength       =   2
      TabIndex        =   0
      Top             =   8640
      Width           =   945
   End
   Begin VB.TextBox Txttitle 
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
      Left            =   1590
      MaxLength       =   15
      TabIndex        =   1
      Top             =   8640
      Width           =   2160
   End
   Begin VB.CommandButton Cmd_Clear 
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
      Left            =   12570
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9840
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
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9840
      Width           =   1125
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
      Left            =   3840
      MaxLength       =   25
      TabIndex        =   2
      Top             =   8640
      Width           =   3180
   End
   Begin VB.CommandButton Cmd_save 
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9840
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6300
      Left            =   285
      TabIndex        =   7
      Top             =   1620
      Width           =   14640
      _cx             =   25823
      _cy             =   11112
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
      Left            =   13080
      TabIndex        =   15
      Top             =   240
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Title"
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
      Index           =   4
      Left            =   1590
      TabIndex        =   14
      Top             =   8280
      Width           =   810
   End
   Begin MSForms.ComboBox cboAdd 
      Height          =   345
      Left            =   7110
      TabIndex        =   3
      Top             =   8640
      Width           =   1290
      VariousPropertyBits=   746604571
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "2275;609"
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
      Caption         =   "Additional Cls"
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
      Left            =   7125
      TabIndex        =   13
      Top             =   8280
      Width           =   1170
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
      Index           =   2
      Left            =   3855
      TabIndex        =   12
      Top             =   8280
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Cls"
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
      Left            =   540
      TabIndex        =   11
      Top             =   8280
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   270
      Top             =   8205
      Width           =   8265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Classification Master"
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
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   14610
   End
   Begin VB.Label LblCode 
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
      Left            =   480
      TabIndex        =   9
      Top             =   8265
      Width           =   975
   End
   Begin VB.Label LblName 
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
      Left            =   1560
      TabIndex        =   8
      Top             =   8265
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   270
      Top             =   8565
      Width           =   8265
   End
End
Attribute VB_Name = "FrmCostMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim RS As New Recordset
Dim ubah As Boolean, hapus As Boolean, gavalid As Boolean, ubahedate As Boolean
Dim SDate, EDate, sdateawal, edateakhir
Dim i As Integer

Dim bteColSelect As Byte
Dim bteColCostCls As Byte
Dim bteColCostTittle As Byte
Dim bteColDesc As Byte
Dim bteColAddCls As Byte
Dim bteColAddDesc As Byte

Sub Header()

    bteColSelect = 0
    bteColCostCls = 1
    bteColCostTittle = 2
    bteColDesc = 3
    bteColAddCls = 4
    bteColAddDesc = 5
  
    With grid
        .clear
        
        .Rows = 1
        .ColS = 6
        
        .TextMatrix(0, bteColSelect) = "S"
        .TextMatrix(0, bteColCostCls) = "Cost Cls"
        .TextMatrix(0, bteColCostTittle) = "Cost Title"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColAddCls) = "Additional Cls"
        .TextMatrix(0, bteColAddDesc) = "Additional Cls"
        
        .ColWidth(bteColSelect) = 250
        .ColWidth(bteColCostCls) = 850
        .ColWidth(bteColCostTittle) = 1500
        .ColWidth(bteColDesc) = 2750
        .ColWidth(bteColAddCls) = 1250
        .ColWidth(bteColAddDesc) = 1250
        
        .ColAlignment(bteColSelect) = flexAlignLeftCenter
        .ColAlignment(bteColCostCls) = flexAlignLeftCenter
        .ColAlignment(bteColCostTittle) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColAddCls) = flexAlignCenterCenter
        .ColAlignment(bteColAddDesc) = flexAlignCenterCenter
        
        .ColHidden(bteColAddCls) = True
        .EditMaxLength = 1
    End With

End Sub





Private Sub cboAdd_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    Kosong
    Header
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    sql = "select * from inventorycost_master order by cost_cls, cost_title"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    Browse
End Sub

Sub Kosong()
    txtCost = ""
    txtCost.Enabled = True
    TxtTitle.Text = ""
    txtDesc.Text = ""
    cboAdd.clear
    cboAdd.AddItem ""
    cboAdd.List(0, 0) = "Yes"
    cboAdd.List(0, 1) = "1"
    cboAdd.AddItem ""
    cboAdd.List(1, 0) = "No"
    cboAdd.List(1, 1) = "0"
End Sub

Private Sub Cmd_SubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Sub Browse()
    RS.filter = ""
    RS.Requery
    i = 1
    If Not (RS.BOF And RS.EOF) Then
        With grid
        Do While Not RS.EOF
            .Rows = .Rows + 1
            .TextMatrix(i, bteColCostCls) = Trim(RS("cost_cls"))
            .TextMatrix(i, bteColCostTittle) = IIf(IsNull(RS("cost_title")), "", Trim(RS("cost_title")))
            .TextMatrix(i, bteColDesc) = IIf(IsNull(RS("description")), "", Trim(RS("description")))
            .TextMatrix(i, bteColAddCls) = IIf(IsNull(RS("additional_cls")), "", Trim(RS("additional_cls")))
            .TextMatrix(i, bteColAddDesc) = IIf(IsNull(RS("additional_cls")) Or Trim(RS("additional_cls")) = "0", "No", "Yes")
            .Cell(flexcpBackColor, i, 0) = &HFFFFFF
            RS.MoveNext
            i = i + 1
        Loop
        End With
    Else
        Header
    End If
End Sub

Private Sub Cmd_Save_Click()
Dim tanya
Dim sql1 As String
On Error Resume Next
hapus = False
If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub

      With grid
          For i = 1 To .Rows - 1
            If .TextMatrix(i, bteColSelect) = "D" Then
              If IsEmpty(tanya) Then tanya = MsgBox("Do You Really Want To Delete This Data ?", vbQuestion & vbYesNo, "Confirmation")
              If tanya = vbYes Then
                    sql1 = "delete from inventorycost_master where cost_cls='" & .TextMatrix(i, bteColCostCls) & "'  "
                           '"additional_cls='" & .TextMatrix(i, bteColAddCls) & "'"
                    Db.Execute sql1
               If err.number = "-2147217873" Then
                    LblErrMsg.Caption = DisplayMsg(1204) '"Can't delete this record!This record used by another application "
                    kosongColGrid
                    Exit Sub
               End If
                    hapus = True
              Else
                    txtCost.Enabled = True
                  Exit For
              End If
            End If
          Next i
        
         If (hapus) Then Kosong: Header: Browse: LblErrMsg = DisplayMsg(1201): Exit Sub

        If txtCost.Text = "" Then
          txtCost.SetFocus
          LblErrMsg = DisplayMsg(8053) '"Please Input Cost"
          Exit Sub
        ElseIf cboAdd = "" Then
          cboAdd.SetFocus
          LblErrMsg = DisplayMsg(8052) '"Please Input Additional Cls"
          Exit Sub
        Else
        End If

          If ubah = False Then
              RS.filter = "cost_cls='" & txtCost & "' and additional_cls='" & cboAdd.List(cboAdd.ListIndex, 1) & "'"
              If Not (RS.EOF And RS.BOF) Then
                  LblErrMsg = DisplayMsg(1023):  Exit Sub
              Else
                  RS.AddNew
                  RS("cost_cls") = txtCost
              End If
          Else
              RS.filter = "cost_cls='" & txtCost & "' "
          End If
          
          RS("cost_title") = TxtTitle.Text
          RS("description") = txtDesc.Text
          RS("additional_cls") = cboAdd.List(cboAdd.ListIndex, 1)
          RS("Last_Update") = Now
          RS("Last_User") = userLogin
          RS.update

      RS.Requery
      RS.filter = ""

      Kosong
      Header
      Browse

      LblErrMsg = DisplayMsg(IIf((ubah = False), 1000, 1101))

      ubah = False
        txtCost.Enabled = True
      End With
End Sub

Private Sub cmd_clear_Click()
    Kosong
    Header
    Browse
    txtCost.Enabled = True
    txtCost.SetFocus
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim TextGrid As String
Dim k As Boolean
Dim j As Integer

k = False
With grid
    TextGrid = grid.Text

    If TextGrid = "S" Then
        txtCost.Text = .TextMatrix(Row, bteColCostCls)
        TxtTitle.Text = .TextMatrix(Row, bteColCostTittle)
        
        txtDesc.Text = .TextMatrix(Row, bteColDesc)
        If .TextMatrix(Row, bteColAddCls) = "1" Then
        cboAdd.Text = "Yes"
        Else
        cboAdd.Text = "No"
        End If
        ubah = True
       Call kosongColGrid
    ElseIf TextGrid = "D" Then
       Call kosongColGrid("S")
    End If

    .TextMatrix(Row, Col) = TextGrid
    
    For j = 1 To .Rows - 1
        If .TextMatrix(j, bteColSelect) <> "" Then
            k = True
        End If
    Next j
    
    If k = False Then Kosong
End With
txtCost.Enabled = False
End Sub

Private Sub kosongColGrid(Optional Kolom As String)
    Dim i As Integer
    
    With grid
        .Col = bteColSelect
    
        If Kolom <> "" Then
           For i = 1 To .Rows - 1
              If .Text = Kolom Then .Text = ""
              If .TextMatrix(i, bteColSelect) <> "D" Then .TextMatrix(i, bteColSelect) = ""
           Next i
           Kosong
        Else
           For i = 1 To .Rows - 1
              If .TextMatrix(i, bteColSelect) <> "" Then .TextMatrix(i, bteColSelect) = ""
           Next i
        End If
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 If grid.Col <> bteColSelect Then Cancel = True
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
  If grid.Col = bteColSelect Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> Asc("S") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
  End If
End Sub

Private Sub txtcost_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

