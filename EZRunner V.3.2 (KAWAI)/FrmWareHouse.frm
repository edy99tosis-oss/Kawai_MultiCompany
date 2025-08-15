VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmWarehouse 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warehouse Master"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   14010
   Icon            =   "FrmWareHouse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   14010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LblAdm 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   5685
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7425
      Width           =   3600
   End
   Begin MSMask.MaskEdBox Tgl1 
      Height          =   315
      Left            =   12075
      TabIndex        =   4
      Top             =   7410
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
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
      Left            =   11183
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8490
      Width           =   1155
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   495
      Left            =   413
      TabIndex        =   15
      Top             =   7860
      Width           =   13155
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
         Height          =   180
         Left            =   105
         TabIndex        =   16
         Top             =   195
         Width           =   12930
      End
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
      Left            =   12443
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8490
      Width           =   1155
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
      Left            =   413
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8490
      Width           =   1185
   End
   Begin VB.CommandButton CmdData 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last Page"
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
      Left            =   7455
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8490
      Width           =   1275
   End
   Begin VB.CommandButton CmdData 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Next Page"
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
      Left            =   6105
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8490
      Width           =   1275
   End
   Begin VB.CommandButton CmdData 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Prev Page"
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
      Left            =   4755
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8490
      Width           =   1275
   End
   Begin VB.CommandButton CmdData 
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Page"
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
      Left            =   3405
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8490
      Width           =   1275
   End
   Begin VB.TextBox TxtWh 
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
      Index           =   1
      Left            =   1545
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "Warehouse Name"
      Top             =   7410
      Width           =   2925
   End
   Begin VB.TextBox TxtWh 
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
      Index           =   0
      Left            =   510
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "WH Code"
      Top             =   7410
      Width           =   990
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   11730
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   525
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5445
      Left            =   420
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1395
      Width           =   13155
      _cx             =   23204
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   12090
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "Use End"
      Top             =   7410
      Width           =   1395
      _ExtentX        =   2461
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
      Format          =   295305219
      CurrentDate     =   37818
   End
   Begin VB.Label LblWh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NG Cls"
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
      Left            =   10770
      TabIndex        =   27
      Top             =   7050
      Width           =   585
   End
   Begin VB.Label lblNG 
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
      Height          =   285
      Left            =   11520
      TabIndex        =   26
      Top             =   7410
      Width           =   465
   End
   Begin VB.Line Line3 
      X1              =   11490
      X2              =   11960
      Y1              =   7710
      Y2              =   7710
   End
   Begin MSForms.ComboBox cboNG 
      Height          =   315
      Left            =   10710
      TabIndex        =   25
      Tag             =   "Stock"
      Top             =   7410
      Width           =   735
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "1296;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CboTrade 
      Height          =   315
      Left            =   4545
      TabIndex        =   2
      Tag             =   "Adm group"
      Top             =   7410
      Width           =   1095
      VariousPropertyBits=   746604571
      MaxLength       =   10
      DisplayStyle    =   3
      Size            =   "1931;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      X1              =   5685
      X2              =   9300
      Y1              =   7710
      Y2              =   7710
   End
   Begin MSForms.ComboBox CboStock 
      Height          =   315
      Left            =   9375
      TabIndex        =   3
      Tag             =   "Stock"
      Top             =   7410
      Width           =   735
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "1296;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label LblWh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Use End Date"
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
      Left            =   11925
      TabIndex        =   23
      Top             =   7050
      Width           =   1155
   End
   Begin VB.Label LblWh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Cls"
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
      Left            =   9300
      TabIndex        =   22
      Top             =   7080
      Width           =   810
   End
   Begin VB.Label LblWh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adm Group"
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
      Left            =   4680
      TabIndex        =   21
      Top             =   7050
      Width           =   975
   End
   Begin VB.Label LblWh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warehouse Name"
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
      Left            =   2055
      TabIndex        =   20
      Top             =   7050
      Width           =   1515
   End
   Begin VB.Label LblWh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WH Code"
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
      Left            =   615
      TabIndex        =   19
      Top             =   7050
      Width           =   795
   End
   Begin VB.Label LblStock 
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
      Height          =   285
      Left            =   10170
      TabIndex        =   18
      Top             =   7425
      Width           =   465
   End
   Begin VB.Line Line2 
      X1              =   10140
      X2              =   10610
      Y1              =   7710
      Y2              =   7710
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warehouse Master"
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
      Left            =   420
      TabIndex        =   12
      Top             =   480
      Width           =   13155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   420
      Top             =   6960
      Width           =   13155
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   495
      Index           =   2
      Left            =   420
      Top             =   7320
      Width           =   13155
   End
End
Attribute VB_Name = "FrmWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rswh As ADODB.Recordset
Dim baru As Boolean
Dim Pos As Integer, jml As Integer
Dim StaErr As Boolean, nErr As Integer, StrWDel As String, nOK As Integer
Dim HakU As Integer

Dim bteColSelect As Byte
Dim bteColWHCode As Byte
Dim bteColWHName As Byte
Dim bteColAdmGroup As Byte
Dim bteColAdmName As Byte
Dim bteColStockCls As Byte
Dim bteColUseEndDate As Byte
Dim bteColLastUpdate As Byte
Dim bteColNGCls As Byte

Sub Header()
    Dim C As Byte
    
    bteColSelect = 0
    bteColWHCode = 1
    bteColWHName = 2
    bteColAdmGroup = 3
    bteColAdmName = 4
    bteColStockCls = 5
    bteColUseEndDate = 7
    bteColLastUpdate = 8
    bteColNGCls = 6
    
    With grid
        .ColS = 9
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColWHCode) = "WH Code"
        .TextMatrix(0, bteColWHName) = "Warehouse Name"
        .TextMatrix(0, bteColAdmGroup) = "Adm Group"
        .TextMatrix(0, bteColAdmName) = "Adm Name"
        .TextMatrix(0, bteColStockCls) = "Stock Cls"
        .TextMatrix(0, bteColUseEndDate) = "Use End Date"
        .TextMatrix(0, bteColLastUpdate) = "Last Update"
        .TextMatrix(0, bteColNGCls) = "NG Cls"
        
        For C = 0 To bteColLastUpdate
            .ColAlignment(C) = flexAlignLeftCenter
        Next C
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColWHCode) = 1200
        .ColWidth(bteColWHName) = 2800
        .ColWidth(bteColAdmGroup) = 1200
        .ColWidth(bteColAdmName) = 4000
        .ColWidth(bteColStockCls) = 1000
        .ColWidth(bteColUseEndDate) = 1500
        .ColWidth(bteColStockCls) = 1000
        
        .ColHidden(bteColLastUpdate) = True
        
        .EditMaxLength = 1
    End With
End Sub

Private Sub cboNG_Change()
If cboNG.MatchFound Then
    lblNG.Caption = cboNG.List(cboNG.ListIndex, 1)
Else
    lblNG.Caption = ""
End If
End Sub

Private Sub CboStock_Change()
    On Error Resume Next
    LblStock = CboStock.List(CboStock.ListIndex, 1)
End Sub

Private Sub CboStock_Click()
    LblStock = CboStock.List(CboStock.ListIndex, 1)
End Sub

Private Sub cbotrade_Change()
    On Error Resume Next
    LblAdm = cbotrade.List(cbotrade.ListIndex, 1)
    If cbotrade.MatchFound = False Then cbotrade.SetFocus: LblAdm = ""
End Sub

Private Sub CboTrade_Click()
    LblAdm = cbotrade.List(cbotrade.ListIndex, 1)
    If cbotrade.MatchFound = False Then cbotrade.SetFocus: LblAdm = ""
End Sub

Private Sub cmdClear_Click()
    Kosong
    Browse
    Pakai True
    LblErr = ""
    baru = True
    Dim IK As Long
    For IK = 1 To grid.Rows - 1
        grid.TextMatrix(IK, bteColSelect) = ""
    Next
    TxtWh(0).SetFocus
    End Sub

Private Sub CmdMenu_Click()
    frmMainMenu.Show
    Unload Me
    DoEvents
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    If Tgl1 <> "99/99/9999" Then
        DTPicker1 = Tgl1
    End If
End Sub

Private Sub DTPicker1_Change()
    Tgl1 = Format(DTPicker1.Month, "00") & "/" & Format(DTPicker1.Day, "00") & "/" & DTPicker1.Year
End Sub

Private Sub DTPicker1_Click()
    Tgl1 = Format(DTPicker1.Month, "00") & "/" & Format(DTPicker1.Day, "00") & "/" & DTPicker1.Year
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

Private Sub CmdData_Click(Index As Integer)
Dim IC As Integer
For IC = 0 To grid.Rows - 1
    grid.TextMatrix(IC, bteColSelect) = ""
Next
Select Case Index
    Case 0
        grid.TopRow = 1
        grid.Row = 1
    Case 1
        If grid.Row > 1 Then
            grid.Row = grid.Row - 1
            grid.TopRow = grid.Row
        Else
            grid.TopRow = 1
        End If
    Case 2
        If grid.Row < grid.Rows - 1 Then
            grid.Row = grid.Row + 1
            grid.TopRow = grid.Row
        Else
            grid.TopRow = grid.Rows - 1
        End If
    Case 3
        grid.Row = grid.Rows - 1
        grid.TopRow = grid.Rows - 1
        
End Select
Pos = grid.Row
jml = grid.Rows - 1
If Pos = jml Then
    LblErr = DisplayMsg("4021")
ElseIf Pos = 1 Then
    LblErr = DisplayMsg("4020")
Else
    LblErr = ""
End If
grid.SetFocus
End Sub

Private Sub CmdSubmit_Click()
    Dim strS As Integer, Jawab As Integer, CekS As Boolean, CekD As Boolean
    Dim strD As Integer, ie As Integer
    
    CekS = False
    CekD = False
    StaErr = False
    
    If HakU = 0 Then LblErr = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
    strS = 0
    strD = 0
    
    If baru = False Then
        strS = grid.FindRow("S", 0, bteColSelect, False)
        strD = grid.FindRow("D", 0, bteColSelect, False, False)
        If strD > 0 Then CekD = True: Jawab = MsgBox("Do you really want to Delete this Record", vbInformation + vbYesNo + vbDefaultButton2, "Confirmation")
        If Jawab = vbYes Then DataGrid
        If strS > 0 And cek Then
            DataGrid
            If StaErr = False Then
                LblErr = DisplayMsg(1101)
                baru = True
                Pakai True
                Browse
                Kosong
                Dim IK As Long
                For IK = 1 To grid.Rows - 1
                    grid.TextMatrix(IK, bteColSelect) = ""
                Next
                TxtWh(0).SetFocus
            Else
                LblErr = DisplayMsg(1102)
            End If
        End If
    
        Dim PRec As Integer
        If CekD Then
            If Jawab = vbYes Then
                If StaErr = False Then
                    LblErr = DisplayMsg(1201)
                    baru = True
                    TxtWh(0).SetFocus
                Else
                    LblErr = DisplayMsg(1000)
                    If Trim$(StrWDel) <> "" Then
                        For ie = 0 To nErr - 1
                            PRec = grid.FindRow(Trim$(Split(StrWDel, ",")(ie)), 0, 1, False)
                            grid.TextMatrix(PRec, bteColSelect) = "D"
                        Next ie
                    End If
                End If
            ElseIf Jawab = vbNo Then
                LblErr = ""
                baru = True
                Kosong
                Dim Ikd As Long
                For Ikd = 1 To grid.Rows - 1
                    grid.TextMatrix(Ikd, bteColSelect) = ""
                Next
            End If
        End If
        strS = 0
    Else
        Dim SqlU As String, PosRec As Integer
        If cek Then
        
            '#Check Code in Trade_Master
            Dim rs2 As New ADODB.Recordset
            If rs2.State = 1 Then rs2.Close
            rs2.CursorLocation = adUseClient
            rs2.Open "select * from trade_master where trade_code='" & Trim(TxtWh(0)) & "'", Db, adOpenKeyset, adLockOptimistic
            If rs2.EOF = False Then
                LblErr = DisplayMsg(3004)
                If rs2.State = 1 Then rs2.Close
                Me.MousePointer = vbDefault: Exit Sub
            End If
            If rs2.State = 1 Then rs2.Close
        
            If Tgl1 <> "99/99/9999" Then
                DTPicker1.Value = Format(Tgl1, "dd-mmm-yyyy")
            End If
            
            Dim UseEnd As String
            If Tgl1 = "99/99/9999" Then
                UseEnd = "99999999"
            Else
                UseEnd = Format(Tgl1, "YYYYMMDD")
            End If
            SqlU = " insert into WareHouse_master " & vbLf & _
                   "       ( " & vbLf & _
                   "          WH_Code, WH_Name, Adm_Group, StockControl_Cls, Use_EndDay, " & vbLf & _
                   "          Last_Update, Last_User, NG_Cls " & vbLf & _
                   "        ) " & vbLf & _
                   " values ( " & vbLf & _
                   "          '" & TxtWh(0) & "','" & TxtWh(1) & "','" & cbotrade.List(cbotrade.ListIndex, 0) & "', " & vbLf & _
                   "          '" & CboStock.List(CboStock.ListIndex, 0) & "','" & UseEnd & "', getdate(), " & vbLf & _
                   "          '" & userLogin & "', '" & cboNG.List(cboNG.ListIndex, 0) & "' " & vbLf & _
                   "        ) "
            PosRec = grid.FindRow(Trim$(TxtWh(0)), 0, bteColWHCode, False)
            If PosRec < 0 Then
                Db.Execute SqlU
                LblErr = DisplayMsg(1000)
                Kosong
                TxtWh(0).SetFocus
            Else
                LblErr = "Warehouse " & DisplayMsg("3004")
                TxtWh(0).SetFocus
            End If
        End If
        baru = True
        Browse
        SqlU = ""
    End If
End Sub

Private Sub Form_Load()
 If gb_Simulation = True Then Call up_InitSimulation(Me)
Dim RS As Recordset, ir As Integer, StG As String
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
HakU = hakUpdate(Me.Name)
StrWDel = ""
Header
Browse
baru = True
StG = ""
Tgl1 = "99/99/9999"
Set RS = Db.Execute("Select Trade_code as TC, Trade_name as TN from trade_master order by trade_code")

cbotrade.clear
cbotrade.columnCount = 2
cbotrade.TextColumn = 1
ir = 0
While Not RS.EOF
    cbotrade.AddItem ""
    cbotrade.List(ir, 0) = Trim(RS!TC)
    cbotrade.List(ir, 1) = Trim$(RS!TN)
    ir = ir + 1
    RS.MoveNext
Wend
cbotrade.ColumnWidths = "60 pt; 300 pt"
cbotrade.ListWidth = 360
cbotrade.ListRows = 15

CboStock.clear
CboStock.columnCount = 2
CboStock.TextColumn = 1
CboStock.AddItem ""
CboStock.List(0, 0) = "01"
CboStock.List(0, 1) = "Yes"
CboStock.AddItem ""
CboStock.List(1, 0) = "02"
CboStock.List(1, 1) = "No"
CboStock.ColumnWidths = "20 pt; 40 pt"
CboStock.ListWidth = 60
CboStock.ListRows = 3

cboNG.clear
cboNG.columnCount = 2
cboNG.TextColumn = 1
cboNG.AddItem ""
cboNG.List(0, 0) = "01"
cboNG.List(0, 1) = "Yes"
cboNG.AddItem ""
cboNG.List(1, 0) = "02"
cboNG.List(1, 1) = "No"
cboNG.ColumnWidths = "20 pt; 40 pt"
cboNG.ListWidth = 60
cboNG.ListRows = 3

DTPicker1.Value = Now()

End Sub

Sub Browse()
sql = "select WH_Code,WH_Name,Adm_group,trade_name,Use_EndDay,warehouse_master.Last_Update, " & vbLf & _
    " SC= " & vbLf & _
    " Case StockControl_Cls " & vbLf & _
    "   when '01' then StockControl_Cls + ' - Yes' " & vbLf & _
    "   when '02' then StockControl_Cls + ' - No' " & vbLf & _
    " end, " & vbLf & _
    " NGCls = " & vbLf & _
    " Case warehouse_master.NG_Cls " & vbLf & _
    "   when '01' then warehouse_master.NG_Cls + ' - Yes' " & vbLf & _
    "   when '02' then warehouse_master.NG_Cls + ' - No' " & vbLf & _
    " end " & vbLf & _
    " from warehouse_master INNER JOIN trade_master  " & vbLf & _
    " On trade_master.trade_code=warehouse_master.adm_group order by wh_code"
Set rswh = New ADODB.Recordset
rswh.Open sql, Db, adOpenKeyset, adLockOptimistic
Dim RSA As Recordset
i = 0
Header
While Not rswh.EOF
        i = i + 1
        grid.AddItem ""
        grid.TextMatrix(i, bteColWHCode) = Trim$(rswh!wh_code)
        grid.TextMatrix(i, bteColWHName) = Trim$(rswh!WH_Name)
        grid.TextMatrix(i, bteColAdmGroup) = Trim$(rswh!adm_group)
        If IsNull(rswh!trade_name) Then
            grid.TextMatrix(i, bteColAdmName) = ""
        Else
            grid.TextMatrix(i, bteColAdmName) = Trim$(rswh!trade_name)
        End If
        grid.TextMatrix(i, bteColStockCls) = Trim$(rswh!SC)
        grid.TextMatrix(i, bteColNGCls) = Trim$(rswh!NGCls)
        If rswh!Use_EndDay = "99999999" Then
            grid.TextMatrix(i, bteColUseEndDate) = "99/99/9999"
        Else
            grid.TextMatrix(i, bteColUseEndDate) = Format(Mid(Trim$(rswh!Use_EndDay), 1, 4) + "/" + Mid(Trim$(rswh!Use_EndDay), 5, 2) + "/" + Mid(Trim$(rswh!Use_EndDay), 7, 2), "dd mmm yyyy")
        End If
        grid.TextMatrix(i, bteColLastUpdate) = Format(Trim$(rswh!Last_Update), "dd mmm yyyy hh:mm:ss AM/PM")
        grid.Cell(flexcpBackColor, i, bteColWHCode, i, bteColUseEndDate) = &HDFFFFF
        grid.Cell(flexcpBackColor, i, bteColSelect) = vbWhite
        
        
        
   rswh.MoveNext
Wend


End Sub
Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrGrid As String
Dim AdaS As Boolean, brs As Integer, id As Integer
StrGrid = grid.Text
AdaS = False
Pakai False
brs = 0
If StrGrid = "S" Then
    For id = 1 To grid.Rows - 1
        If id <> Row Then grid.TextMatrix(id, bteColSelect) = ""
    Next id
    TxtWh(0) = grid.TextMatrix(grid.Row, bteColWHCode)
    TxtWh(1) = grid.TextMatrix(grid.Row, bteColWHName)
    cbotrade = grid.TextMatrix(grid.Row, bteColAdmGroup)
    CboStock = Left(grid.TextMatrix(grid.Row, bteColStockCls), 2)
    cboNG.Text = Left(grid.TextMatrix(grid.Row, bteColNGCls), 2)
    LblAdm = grid.TextMatrix(grid.Row, bteColAdmName)
    If grid.TextMatrix(grid.Row, bteColUseEndDate) = "99/99/9999" Then
        Tgl1.Text = "99/99/9999"
    Else
        Tgl1.Text = Format(Month(Format(grid.TextMatrix(grid.Row, bteColUseEndDate), "mmm-dd-yyyy")), "00") & "/" & Format(Day(Format(grid.TextMatrix(grid.Row, bteColUseEndDate), "mmm-dd-yyyy")), "00") & "/" & Year(Format(grid.TextMatrix(grid.Row, bteColUseEndDate), "mmm-dd-yyyy"))
    End If
    TxtWh(0).Enabled = False
    TxtWh(1).Enabled = True
    cbotrade.Enabled = True
    CboStock.Enabled = True
    DTPicker1.Enabled = True
    TxtWh(1).SetFocus
    baru = False
    LblErr = ""
ElseIf StrGrid = "D" Then
    Pakai True
    For id = 1 To grid.Rows - 1
        'Jika ada S maka , hapus yg S
        If grid.TextMatrix(id, bteColSelect) = "S" Then grid.TextMatrix(id, bteColSelect) = "": Exit For
    Next id
    baru = False
    LblErr = ""
Else
    Pakai True
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
    If KeyAscii = vbKeyEscape Then KeyAscii = 0
    If KeyAscii = Asc(".") Then KeyAscii = 0
End If
End Sub

Private Sub TxtWh_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = 0
If Index <> 1 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Function cek() As Boolean
Dim tmpi As Byte, rsc As Recordset
cek = False
    For tmpi = 0 To 1
        If Trim$(TxtWh(tmpi)) <> "" Then
            cek = True
        Else
            cek = False
            TxtWh(tmpi).SetFocus
            LblErr = DisplayMsg("0001") & " " & TxtWh(tmpi).Tag & " ! "
            Exit Function
        End If
    Next
If Trim$(cbotrade) = "" Then cbotrade.SetFocus: LblErr = DisplayMsg("0002"): cek = False: Exit Function
cbotrade = Trim(cbotrade)
If cbotrade.MatchFound = False Then cbotrade.SetFocus: LblErr = DisplayMsg("0003"): cek = False: Exit Function

If Trim$(CboStock) = "" Then CboStock.SetFocus: LblErr = DisplayMsg("0004"): cek = False: Exit Function
CboStock = CboStock
If CboStock.MatchFound = False Then CboStock.SetFocus: LblErr = DisplayMsg("0005"): cek = False: Exit Function
If Tgl1 <> "99/99/9999" Then
    On Error GoTo tgx
    DTPicker1.Month = Mid(Tgl1, 1, 2)
    DTPicker1.Day = Mid(Tgl1, 4, 2)
    DTPicker1.Year = Right(Tgl1, 4)
End If

Exit Function
tgx:
Tgl1.SetFocus: LblErr = DisplayMsg(1022): cek = False: Exit Function
End Function
Sub Pakai(stat As Boolean)
    TxtWh(0).Enabled = stat
    TxtWh(1).Enabled = stat
    cbotrade.Enabled = stat
    CboStock.Enabled = stat
    DTPicker1.Enabled = stat
End Sub
Sub Kosong()
Dim tmpi As Integer
For tmpi = 0 To 1
        TxtWh(tmpi) = ""
Next
DTPicker1 = Now()
CboStock = ""
cbotrade = ""
LblAdm = ""
LblStock = ""
Tgl1 = "99/99/9999"
cboNG = ""
End Sub
Sub DataGrid()
Dim kode As String, Sta As String
Dim strSQL As String
Dim ix As Integer

On Error Resume Next
Dim PosS As Integer
PosS = grid.FindRow("S", 0, bteColSelect, False)
If PosS > 0 Then
    kode = Trim$(grid.TextMatrix(PosS, bteColWHCode))
            Dim UseEnd As String
            
            If Tgl1 = "99/99/9999" Then
                UseEnd = "99999999"
            Else
                UseEnd = Format(Tgl1, "YYYYMMDD")
            End If
    strSQL = "update WareHouse_master set WH_Name='" & TxtWh(1) & "', " & vbLf & _
                 " Adm_Group='" & cbotrade & "', " & vbLf & _
                 " StockControl_Cls='" & CboStock.List(CboStock.ListIndex, 0) & "' ," & vbLf & _
                 " Use_EndDay='" & UseEnd & "'," & vbLf & _
                 " last_update=getdate(), " & vbLf & _
                 " last_user='" & userLogin & "'," & vbLf & _
                 " NG_Cls = '" & cboNG.List(cboNG.ListIndex, 0) & "' " & vbLf & _
                 " where WH_Code ='" & kode & "'"
    Db.Execute (strSQL)
    
    If err.number <> 0 Then
        StaErr = True
    Else
        StaErr = False
    End If
    Exit Sub
End If

nErr = 0
nOK = 0
StrWDel = ""
For ix = 1 To grid.Rows - 1
kode = Trim$(grid.TextMatrix(ix, bteColWHCode))
Sta = Trim$(grid.TextMatrix(ix, bteColSelect))
    If Sta = "D" Then
        strSQL = "delete from WareHouse_master  where WH_Code ='" & kode & "'"
        
        '#Check Code in Trade_Master
        Dim rs2 As New ADODB.Recordset
        If rs2.State = 1 Then rs2.Close
        rs2.CursorLocation = adUseClient
        rs2.Open "select * from trade_master where Subcon_WH_Code='" & Trim(kode) & "'", Db, adOpenKeyset, adLockOptimistic
        If rs2.EOF = False Then
            GoTo skip
        Else
            If strSQL <> "" Then Db.Execute strSQL
        End If
        If rs2.State = 1 Then rs2.Close
                        
        If err.number <> 0 Then
skip:
            StrWDel = StrWDel & kode & ","
            nErr = nErr + 1
            err.clear
        Else
            nOK = nOK + 1
        End If
    End If
    strSQL = ""
Next ix
If Len(StrWDel) > 1 Then StrWDel = Mid(StrWDel, 1, Len(StrWDel) - 1)

If nErr > 0 Then
    StaErr = True
End If

kode = ""
Sta = ""
strSQL = ""
Browse
End Sub

Function NamaTrade(AdmGroup As String) As String
Dim rst As Recordset
Set rst = Db.Execute("Select Trade_Name as TN from Trade_master where trade_code='" & Trim$(AdmGroup) & "'")
If Not rst.EOF Then
    NamaTrade = rst!TN
Else
    NamaTrade = ""
End If
End Function
