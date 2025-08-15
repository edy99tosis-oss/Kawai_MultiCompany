VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmProdWorkingTime 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Working Time"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProdWorkingTime.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Tag             =   ".AddItem Trim(VSFlexGrid1.TextMatrix(0, i))"
   Begin VB.TextBox txtWorkTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   14250
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "9,999"
      Top             =   2190
      Width           =   675
   End
   Begin VB.TextBox txtTotWorkTime 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2550
      TabIndex        =   0
      Text            =   "9,999"
      Top             =   2190
      Width           =   645
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1245
      Left            =   270
      TabIndex        =   18
      Top             =   870
      Width           =   14655
      Begin VB.Line Line1 
         Index           =   5
         X1              =   4920
         X2              =   6330
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblDt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4920
         TabIndex        =   30
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result Date"
         Height          =   195
         Index           =   3
         Left            =   3810
         TabIndex        =   29
         Top             =   750
         Width           =   990
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   1740
         X2              =   3090
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblLot 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1740
         TabIndex        =   28
         Top             =   750
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot No."
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   27
         Top             =   750
         Width           =   600
      End
      Begin VB.Line Line1 
         Index           =   2
         Visible         =   0   'False
         X1              =   1740
         X2              =   3510
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4920
         X2              =   10620
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label lblItemDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
         Height          =   195
         Left            =   4920
         TabIndex        =   26
         Top             =   300
         Width           =   5040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   3810
         TabIndex        =   25
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         Height          =   195
         Index           =   0
         Left            =   435
         TabIndex        =   24
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         Height          =   195
         Left            =   1740
         TabIndex        =   23
         Top             =   300
         Width           =   1155
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1740
         X2              =   3510
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result"
         Height          =   195
         Index           =   5
         Left            =   9840
         TabIndex        =   22
         Top             =   750
         Width           =   525
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   10680
         X2              =   12210
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblResultQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11175
         TabIndex        =   21
         Top             =   750
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan "
         Height          =   195
         Index           =   4
         Left            =   7020
         TabIndex        =   20
         Top             =   750
         Width           =   420
      End
      Begin VB.Label lblDailyQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7695
         TabIndex        =   19
         Top             =   750
         Width           =   1380
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   7710
         X2              =   9120
         Y1              =   930
         Y2              =   930
      End
   End
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Index           =   2
      Left            =   11310
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9900
      Width           =   1140
   End
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
      Height          =   375
      Index           =   1
      Left            =   12540
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9900
      Width           =   1140
   End
   Begin VB.TextBox txtRemarks 
      Height          =   315
      Left            =   8010
      MaxLength       =   35
      TabIndex        =   4
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      Top             =   8745
      Width           =   6780
   End
   Begin VB.TextBox txtLossTime 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6465
      TabIndex        =   3
      Text            =   "999"
      Top             =   8745
      Width           =   1350
   End
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   0
      Left            =   13770
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9900
      Width           =   1155
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Back"
      Height          =   375
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9900
      Width           =   1140
   End
   Begin VB.Frame FrameErr 
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
      Height          =   555
      Left            =   270
      TabIndex        =   11
      Top             =   9225
      Width           =   14655
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
         TabIndex        =   12
         Top             =   210
         Width           =   14430
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13080
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   285
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5565
      Left            =   270
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2595
      Width           =   14655
      _cx             =   25850
      _cy             =   9816
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
      AllowUserResizing=   3
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Index           =   3
      Left            =   8010
      TabIndex        =   33
      Top             =   8400
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Working Time (Min)"
      Height          =   195
      Index           =   7
      Left            =   270
      TabIndex        =   32
      Top             =   2250
      Width           =   2160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Real Working Time (Min)"
      Height          =   195
      Index           =   6
      Left            =   12030
      TabIndex        =   31
      Top             =   2250
      Width           =   2115
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      Height          =   195
      Left            =   2055
      TabIndex        =   17
      Top             =   8805
      Width           =   4200
   End
   Begin VB.Line Line2 
      X1              =   2055
      X2              =   6260
      Y1              =   9045
      Y2              =   9045
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loss Time (Min)"
      Height          =   195
      Index           =   2
      Left            =   6465
      TabIndex        =   16
      Top             =   8400
      Width           =   1350
   End
   Begin MSForms.ComboBox cbo 
      Height          =   330
      Left            =   420
      TabIndex        =   2
      Top             =   8730
      Width           =   765
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "1349;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "AA"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   1
      Left            =   2055
      TabIndex        =   15
      Top             =   8400
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "W.Loss Time Cls"
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   14
      Top             =   8400
      Width           =   1410
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   2
      Left            =   270
      Top             =   8325
      Width           =   14655
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   495
      Index           =   2
      Left            =   270
      Top             =   8655
      Width           =   14655
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Working Time"
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
      Height          =   390
      Left            =   270
      TabIndex        =   13
      Top             =   285
      Width           =   14655
   End
End
Attribute VB_Name = "frmProdWorkingTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim ColS As Integer, ColCls As Integer, ColDesc As Integer, ColLossTime As Integer, ColRemarks As Integer

Dim nilKosong As Boolean
Public ProdSeqNo As Double
Public VDailySeqNo As Double

Private Sub headerGrid()
    With grid
        .clear
        .ColS = 5
        .Rows = 1
        
        .TextMatrix(0, ColS) = ""
        .TextMatrix(0, ColCls) = "W.Loss Time Cls"
        .TextMatrix(0, ColDesc) = "Description"
        .TextMatrix(0, ColLossTime) = "Loss Time (Min)"
        .TextMatrix(0, ColRemarks) = "Remarks"
        
        .ColWidth(ColS) = 300
        .ColWidth(ColCls) = 2000
        .ColWidth(ColDesc) = 3800
        .ColWidth(ColLossTime) = 2000
        .ColWidth(ColRemarks) = 4500
        
        .ColAlignment(ColS) = flexAlignCenterCenter
        .ColAlignment(ColCls) = flexAlignLeftCenter
        .ColAlignment(ColDesc) = flexAlignLeftCenter
        .ColAlignment(ColLossTime) = flexAlignRightCenter
        .ColAlignment(ColRemarks) = flexAlignLeftCenter
        
        Call ClsProc.AlignHeader(grid)
        
        .RowHeightMax = 250
        .EditMaxLength = 1
    End With
End Sub


'*************************** Initial ***********************
Sub stObject(stEnable As Boolean)
    Cbo.Enabled = stEnable
    txtLossTime.Enabled = stEnable
    txtRemarks.Enabled = stEnable
End Sub

Sub Kosong(Optional stAwal As Byte)
    nilKosong = True
    Call stObject(True)
    
    If stAwal = 1 Then
        txtWorkTime = Format(0, gs_formatWorkingTime)
        txtTotWorkTime = Format(0, gs_formatWorkingTime)
    Else
        Cbo.SetFocus
    End If
    Cbo = "": LblDesc = ""
    txtLossTime = Format(0, gs_formatWorkingTime)
    txtRemarks = ""
    nilKosong = False
End Sub

Sub SetCol()
    ColS = 0: ColCls = 1: ColDesc = 2: ColLossTime = 3: ColRemarks = 4
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    nilKosong = True
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

    HakU = hakUpdate(Me.Name)

    Call isiCbo(Cbo, "WorkingLossTime_Cls", "WorkingLossTime_Cls", "Description", 35, 100, "WorkingLossTime_Cls")
    Call SetCol
    
    Call Kosong(1): Call headerGrid
    nilKosong = False
End Sub

'***********************************************************

'*************************** View Dt ***********************
Sub ViewDt(Optional stAwal As Byte)
Dim rsView As New ADODB.Recordset
    sql = "Select R.Item_Code, I.Item_Name, R.SuratJalan_No, R.Receipt_Date, DQty = D.Qty, RQty = R.Qty " & _
        "From Part_Receipt R, Daily_Production D, Item_Master I " & _
        "Where R.DailySeq_No = D.Seq_No And R.Item_Code = I.Item_Code " & _
            "And R.Seq_No = " & ProdSeqNo
    Set rsView = Db.Execute(sql)
    With rsView
        If rsView.EOF Then
            Call Kosong(stAwal): Call headerGrid
        Else
            lblitem = Trim(!Item_Code)
            lblItemDesc = Trim(!item_name)
            lblLot = Trim(!SuratJalan_No)
            lblDt = Format(!Receipt_Date, "dd MMM yyyy")
            lblDailyQty = Format(!DQty, gs_formatQty)
            lblResultQty = Format(!Rqty, gs_formatQty)
            Call IsiGrid
        End If
    End With
    Set rsView = Nothing
End Sub

Sub IsiGrid()
Dim rsGrid As New ADODB.Recordset

With grid
    Call headerGrid

    sql = "Select WM.*, WD.*, WT.Description from WorkingTime_Master WM " & _
            "left outer join WorkingTime_Detail WD on WD.ProductionSeq_No = WM.ProductionSeq_No " & _
            "left outer join WorkingLossTime_Cls WT on WT.WorkingLossTime_Cls = WD.WorkingLossTime_Cls " & _
        "Where WM.ProductionSeq_No = " & ProdSeqNo & _
        " Order by WT.WorkingLossTime_Cls"
    Set rsGrid = Db.Execute(sql)

    i = 1
    Do While Not rsGrid.EOF
        If i = 1 Then
            txtWorkTime = Format(rsGrid!Working_Time, gs_formatWorkingTime)
            txtTotWorkTime = Format(IIf(IsNull(rsGrid!TotalWorking_Time), 0, rsGrid!TotalWorking_Time), gs_formatWorkingTime)
        End If
        
        If IsNull(rsGrid!WorkingLossTime_Cls) = False Then
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, i, ColS) = vbWhite
            .TextMatrix(i, ColCls) = Trim(rsGrid!WorkingLossTime_Cls)
            .TextMatrix(i, ColDesc) = Trim(rsGrid!Description)
            .TextMatrix(i, ColLossTime) = Format(rsGrid!Loss_Time, gs_formatWorkingTime)
            .TextMatrix(i, ColRemarks) = Trim(rsGrid!Remarks)
            i = i + 1
        End If
        rsGrid.MoveNext
    Loop
    Set rsGrid = Nothing
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True: Exit Sub
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
With grid
    If .Col = 0 Then
        If KeyAscii = Asc(".") Then KeyAscii = 0: Exit Sub
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And _
            KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And _
            KeyAscii <> vbKeyReturn Then _
            KeyAscii = 0: Exit Sub
    End If
End With
End Sub

Private Sub kosongColGrid(Optional strSD As String)
Dim i As Integer
Dim jmlD As Double

With grid
    .Col = 0

    jmlD = 0
    If strSD = "" Then
        stObject (True)
        .Cell(flexcpText, 1, 0, .Rows - 1, 0) = ""
    Else
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "D" Then jmlD = jmlD + 1 Else .TextMatrix(i, 0) = ""
        Next i
        If jmlD = 0 Then Call stObject(True) Else stObject (False)
    End If
End With
End Sub

Public Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim TextGrid As String

LblErrMsg = ""
With grid
    TextGrid = .Cell(flexcpText, Row, Col)

    If TextGrid = "S" Then
        Call kosongColGrid
        Cbo = Trim(.TextMatrix(Row, ColCls))
        Cbo.Enabled = False
        txtLossTime = Trim(.TextMatrix(Row, ColLossTime))
        txtRemarks = Trim(.TextMatrix(Row, ColRemarks))
    Else
        Call Kosong
        Call kosongColGrid("S")
    End If

    .TextMatrix(Row, Col) = TextGrid
End With
End Sub

'***********************************************************

'*************************** Process Dt ********************
Function chkSave(Optional chkDetail As Byte) As Boolean
Dim rsCheck As New ADODB.Recordset

chkSave = False

    If HakU = 0 Then LblErrMsg = DisplayMsg(3008): Exit Function 'You don't have an access for Update

'    If Grid.Rows = 1 And txtWorkTime = "0" Then
'        Lblerrmsg = DisplayMsg(8000) 'Please input Working Time (Min)
'        txtWorkTime.SetFocus
'    End If

    If txtTotWorkTime = "0" Then
        LblErrMsg = DisplayMsg(8000) 'Please input Working Time (Min)
        txtTotWorkTime.SetFocus
        Exit Function
    End If
    
    If CDbl(txtTotWorkTime) > gd_MaxWorkingTime Then
        LblErrMsg = DisplayMsg(8021) & " " & gd_MaxWorkingTime 'Working Time Must Be Equal or Lower than
        txtTotWorkTime.SetFocus
        Exit Function
    End If
    
    If Trim(Cbo) <> "" Then
        If Cbo.MatchFound = False Then
            LblErrMsg = DisplayMsg(8001) 'Working Loss Time Cls not found
            Cbo.SetFocus: Exit Function
        ElseIf CDbl(txtLossTime) = 0 Then
            LblErrMsg = DisplayMsg(8002) 'Please Input Loss Time (Min)
            txtLossTime.SetFocus: Exit Function
        End If
        
        If CDbl(txtLossTime) > gd_MaxWorkingTime Then
            LblErrMsg = DisplayMsg(8021) & " " & gd_MaxWorkingTime 'Working Time Must Be Equal or Lower than
            txtLossTime.SetFocus: Exit Function
        End If
        
        If Cbo.Enabled Then
            sql = "Select WorkingLossTime_Cls from WorkingTime_Detail " & _
                "Where ProductionSeq_No =  " & ProdSeqNo & " And WorkingLossTime_Cls = '" & Cbo & "'"
            Set rsCheck = Db.Execute(sql)
            If Not rsCheck.EOF Then
                LblErrMsg = DisplayMsg(8003) 'Data already exist
                Cbo.SetFocus: Exit Function
            End If
        End If
    End If
chkSave = True
End Function

Private Sub cmdprocess_Click(Index As Integer)
Dim tanya

Me.MousePointer = vbHourglass
Select Case Index
    Case 0: 'Save
        If grid.FindRow("D", , 0, 1, 1) > 0 Then  'Delete
            If chkSave Then
                tanya = MsgBox("Do You Really Want To Delete This Data ?", vbQuestion & vbYesNo, "Confirmation")
                If tanya = vbYes Then Call DeleteDt
            End If
        Else
            If chkSave(1) Then
                Call savemaster
                If Trim(Cbo) <> "" Then Call savedetail
                Call updateMaster
    
                Call Kosong: Call IsiGrid
                If grid.FindRow("S", , 0, 1, 1) < 0 Then
                    LblErrMsg = DisplayMsg(8004) 'Data saved success
                Else
                    LblErrMsg = DisplayMsg(8005) 'Update Record Success
                End If
            End If
        End If

    Case 1:  'Cancel
        Call Kosong: Call IsiGrid

    Case 2:  'Clear
        Call Kosong(1): Call headerGrid: Cbo.SetFocus
End Select
Me.MousePointer = vbDefault
End Sub

Sub savemaster()
Dim RsSave As New ADODB.Recordset

    sql = "Select * From WorkingTime_Master " & _
        "Where ProductionSeq_No = " & ProdSeqNo
    If RsSave.State = adStateOpen Then RsSave.Close
    RsSave.Open sql, Db, adOpenDynamic, adLockOptimistic
    With RsSave
        If .EOF Then .AddNew
        !ProductionSeq_No = ProdSeqNo
         !TotalWorking_Time = CDbl(txtTotWorkTime)
        '!Working_Time = CDbl(txtWorkTime)
        !Last_Update = Now
        !last_user = userLogin
        .update
    End With
    If RsSave.State = adStateOpen Then RsSave.Close
End Sub

Sub updateMaster()
    sql = "Update WorkingTime_Master " & _
        "Set TotalLoss_Time = isnull((Select Sum(Loss_Time) from WorkingTime_Detail Where ProductionSeq_No = WorkingTime_Master.ProductionSeq_No), 0), " & _
        " Working_Time = TotalWorking_Time - isnull((Select Sum(Loss_Time) from WorkingTime_Detail Where ProductionSeq_No = WorkingTime_Master.ProductionSeq_No), 0), " & _
        "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
        "Where ProductionSeq_No = " & ProdSeqNo
    Db.Execute sql
End Sub

Sub savedetail()
Dim RsSave As New ADODB.Recordset

    sql = "Select * From WorkingTime_Detail " & _
        "Where ProductionSeq_No = " & ProdSeqNo & _
            " And WorkingLossTime_Cls = '" & Cbo & "'"
    If RsSave.State = adStateOpen Then RsSave.Close
    RsSave.Open sql, Db, adOpenDynamic, adLockOptimistic
    With RsSave
        If .EOF Then .AddNew
        !ProductionSeq_No = ProdSeqNo
        !WorkingLossTime_Cls = Trim(Cbo)
        !Loss_Time = CDbl(txtLossTime)
        !Remarks = Trim(txtRemarks)
        !Last_Update = Now
        !last_user = userLogin
        .update
    End With
    If RsSave.State = adStateOpen Then RsSave.Close
End Sub

Sub DeleteDt()
Dim id As String

With grid
    For i = 1 To .Rows - 1
        If .Cell(flexcpText, i, ColS) = "D" Then
            id = Trim(.TextMatrix(i, ColCls))
            sql = "Delete WorkingTime_Detail " & _
                "Where ProductionSeq_No = '" & ProdSeqNo & "' And WorkingLossTime_Cls = " & id
            Db.Execute sql
        End If
    Next i
    Call updateMaster
    Call Kosong: Call IsiGrid
    LblErrMsg = DisplayMsg(1201) 'Delete Record Success
End With
End Sub

'***********************************************************

'*************************** Validate **********************
Private Sub cbo_Change()
    Cbo = Cbo
    If Cbo.MatchFound Then LblDesc = Cbo.Column(1) Else LblDesc = ""
End Sub

Private Sub txtLossTime_LostFocus()
If IsNumeric(txtLossTime) = False Then txtLossTime = Format(0, gs_formatWorkingTime)
End Sub

Private Sub txtTotWorkTime_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtTotWorkTime_LostFocus()
If IsNumeric(txtTotWorkTime) Then txtTotWorkTime = Format(txtTotWorkTime, gs_formatWorkingTime) Else txtTotWorkTime = Format(0, gs_formatWorkingTime)
End Sub

Private Sub txtWorkTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> vbKeyBack Then KeyAscii = 0
'    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then KeyAscii = 0
'    If IsNumeric(txtWorkTime) Then
'        If CDbl(txtWorkTime) > 999# And KeyAscii <> vbKeyBack Then KeyAscii = 0
'    End If
End Sub

Private Sub txtWorkTime_Change()
    If InStr(1, txtWorkTime, ",") = 1 Then txtWorkTime = Mid(txtWorkTime, 2, Len(txtWorkTime))
End Sub

Private Sub txtWorkTime_LostFocus()
    If IsNumeric(txtWorkTime) = False Then txtWorkTime = Format(0, gs_formatWorkingTime)
    txtWorkTime = Format(txtWorkTime, gs_formatWorkingTime)
End Sub

Private Sub txtLossTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
    
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

'***********************************************************

'*************************** Out ***************************
Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then Unload Me Else LblErrMsg.Caption = ErrMsg
End Sub

Private Sub CmdSubMenu_Click()
Dim rsCek As New ADODB.Recordset
    
    sql = "delete WorkingTime_Master where TotalWorking_Time = 0"
    Db.Execute sql
    
    sql = "Select ProductionSeq_No from WorkingTime_Master where ProductionSeq_No = '" & ProdSeqNo & "'"
    Set rsCek = Db.Execute(sql)
    If rsCek.EOF Then
        If txtWorkTime <> 0 Then
            LblErrMsg = DisplayMsg(1049) 'Please Submit first
            cmdProcess(0).SetFocus
'        Else
'            LblErrMsg = DisplayMsg(8014) 'Please input Working Time
'            txtWorkTime.SetFocus
            Exit Sub
        End If
    End If
    Set rsCek = Nothing
    
    Unload frmProdMaterialComp
    DoEvents
    With frmProdResult
        Call .kosongBwh
        Call .IsiGrid
        .command1(0).Enabled = False
        .Show
    End With
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub
'***********************************************************
