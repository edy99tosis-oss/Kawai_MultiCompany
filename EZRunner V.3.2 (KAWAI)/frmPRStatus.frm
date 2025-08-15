VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPRStatus 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Purchase Request Status"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPRStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   0
      Left            =   8467
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8730
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   577
      TabIndex        =   16
      Top             =   8070
      Width           =   9030
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
         Top             =   210
         Width           =   8835
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8730
      Width           =   1140
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   577
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8730
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1935
      Left            =   577
      TabIndex        =   11
      Top             =   870
      Width           =   9030
      Begin VB.ComboBox cboComplete 
         Height          =   315
         ItemData        =   "frmPRStatus.frx":0E42
         Left            =   2310
         List            =   "frmPRStatus.frx":0E4F
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1480
         Width           =   885
      End
      Begin VB.ComboBox CboType 
         Height          =   315
         ItemData        =   "frmPRStatus.frx":0E61
         Left            =   2325
         List            =   "frmPRStatus.frx":0E71
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   653
         Width           =   1875
      End
      Begin VB.ComboBox cboFix 
         Height          =   315
         ItemData        =   "frmPRStatus.frx":0EA1
         Left            =   2325
         List            =   "frmPRStatus.frx":0EAE
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1066
         Width           =   885
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Search"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1440
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtAkhir 
         Height          =   315
         Left            =   4410
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   141230083
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker dtAwal 
         Height          =   315
         Left            =   2340
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   141230083
         CurrentDate     =   37810
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complete Cls"
         Height          =   195
         Left            =   855
         TabIndex        =   19
         Top             =   1530
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Request Type"
         Height          =   195
         Index           =   3
         Left            =   855
         TabIndex        =   18
         Top             =   710
         Width           =   1170
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fix Cls"
         Height          =   195
         Left            =   1455
         TabIndex        =   15
         Top             =   1120
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Request Date From"
         Height          =   195
         Index           =   2
         Left            =   375
         TabIndex        =   13
         Top             =   300
         Width           =   1650
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   300
         Width           =   375
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5040
      Left            =   570
      TabIndex        =   6
      Top             =   2925
      Width           =   9030
      _cx             =   15928
      _cy             =   8890
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
      Height          =   435
      Left            =   7762
      TabIndex        =   10
      Top             =   270
      Width           =   1845
      _extentx        =   3254
      _extenty        =   767
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Request Status"
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
      Left            =   3720
      TabIndex        =   14
      Top             =   270
      Width           =   2730
   End
End
Attribute VB_Name = "frmPRStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClsProc As New ClsProc
Dim i As Long, StCek As Byte

Sub Kosong()
    dtAwal = Format(Year(Now) & "-" & Format(Month(Now), "#0") & "-01", "dd MMM yyyy")
    dtAkhir = Format(Now, "dd MMM yyyy")
    cboFix = "No"
    cboComplete.ListIndex = 0
    Call headerGrid
End Sub

Private Sub Form_Load()
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    
    CboType.AddItem strAll
    CboType.AddItem "Part/Material"
    CboType.AddItem "Sheet/Coil"
    CboType.AddItem "Other Item"
    
    Call Kosong
    CboType.ListIndex = 0
End Sub

Public Sub cmdSearch_Click()
    LblErrMsg = ""
    Call IsiGrid
End Sub

Private Sub headerGrid()
With grid
    .clear
    .ColS = 8
    .Rows = 1

    .ColWidth(0) = 1750
    .ColWidth(1) = 1500
    .ColWidth(2) = 1700
    .ColWidth(3) = 1700
    .ColWidth(4) = 800
    .ColWidth(5) = 0
    .ColWidth(6) = 1000
    .ColWidth(7) = 0

    .TextMatrix(0, 0) = "Request No"
    .TextMatrix(0, 1) = "Request Date"
    .ColDataType(1) = flexDTDate
    .TextMatrix(0, 2) = "Person In Charge"
    .TextMatrix(0, 3) = "Department"
    .TextMatrix(0, 4) = "Fix"
    '.ColHidden(4) = True
    .TextMatrix(0, 6) = "Complete"
    
    .ColAlignment(0) = flexAlignCenterCenter
    .ColAlignment(1) = flexAlignCenterCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignCenterCenter
    .ColAlignment(6) = flexAlignCenterCenter
    
    .EditMaxLength = 1
    Call ClsProc.AlignHeader(grid)
End With
End Sub

Sub IsiGrid()
Dim rsGrid As New ADODB.Recordset
Dim sqlResult As String


With grid
    Call headerGrid

    sql = " Select PORequest_Master.*, PersonInCharge_Cls.description as PD, department_cls.description as DD  " & _
            " From PORequest_Master  " & _
            " Left outer JOIN PersonInCharge_Cls on PORequest_Master.personincharge_cls = PersonInCharge_Cls.personincharge_cls " & _
            " Left outer JOIN department_cls on PORequest_Master.department_cls = department_cls.department_cls " & _
            " Where PORequest_Date >= '" & Format(dtAwal, "yyyy-MM-dd") & _
            " ' And PORequest_Date <= '" & Format(dtAkhir, "yyyy-MM-dd") & "' "
     
    'Fix
    If cboFix = "Yes" Then
        sql = sql & "And Fix_Cls = '1' "
    ElseIf cboFix = "No" Then
        sql = sql & "And (Fix_Cls = '0' Or Fix_Cls is NULL) "
    End If
    
    'Others_Cls
    If CboType.ListIndex = 1 Then
        sql = sql + " And Others_Cls = '0' And SheetCoil_Cls = '0' "
    ElseIf CboType.ListIndex = 2 Then
        sql = sql + " And Others_Cls = '0' And SheetCoil_Cls = '1' "
    ElseIf CboType.ListIndex = 3 Then
        sql = sql + " And Others_Cls = '1' "
    End If

    'Complete
    If cboComplete.ListIndex = 1 Then   'YES
        sql = sql & "And Complete_Cls = '1' "
    ElseIf cboComplete.ListIndex = 2 Then   'NO
        sql = sql & "And isnull(Complete_Cls,'0') = '0' "
    End If

    sql = sql & "Order By PORequest_Master.PORequest_No, PORequest_Master.PORequest_Date"
    Set rsGrid = Db.Execute(sql)

    i = 1
    If Not (rsGrid.EOF) Then
        Do While Not rsGrid.EOF
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = Trim(rsGrid("PORequest_No"))
            .TextMatrix(i, 1) = Format(rsGrid("PORequest_Date"), "dd MMM yyyy")
            .TextMatrix(i, 2) = IIf(IsNull(rsGrid("PD")), "", Trim(rsGrid("PD")))
            .TextMatrix(i, 3) = IIf(IsNull(rsGrid("DD")), "", Trim(rsGrid("DD")))
            
            .Cell(flexcpBackColor, i, 4) = vbWhite
            .Cell(flexcpChecked, i, 4) = IIf(rsGrid("Fix_Cls") = 1, flexChecked, flexUnchecked)
            .Cell(flexcpChecked, i, 5) = IIf(rsGrid("Fix_Cls") = 1, flexChecked, flexUnchecked)
            .Cell(flexcpBackColor, i, 6) = vbWhite
            .Cell(flexcpChecked, i, 6) = IIf(rsGrid("Complete_Cls") = 1, flexChecked, flexUnchecked)
            .Cell(flexcpChecked, i, 7) = IIf(rsGrid("Complete_Cls") = 1, flexChecked, flexUnchecked)
            i = i + 1
            rsGrid.MoveNext
        Loop
    Else
        LblErrMsg = DisplayMsg(4006)
    End If
    Set rsGrid = Nothing
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rsCek As New Recordset
    
    LblErrMsg = ""
    If Col <= 3 Then Cancel = 1
    If Col = 4 And Row > 0 Then
        StCek = 0
        sql = "select * from PurchaseOrder_Detail where PORequest_No = '" & Trim(grid.TextMatrix(Row, 0)) & "' "
        If rsCek.State <> adStateClosed Then rsCek.Close
        rsCek.Open sql, Db, adOpenKeyset, adLockOptimistic
        If Not (rsCek.BOF And rsCek.EOF) Then
            If grid.Cell(flexcpChecked, Row, 4) = flexChecked Then
                StCek = 1: Cancel = True
            End If
        End If
        Set rsCek = Nothing
    End If
End Sub

Private Sub grid_Click()
    If grid.Col = 4 And grid.Row > 0 Then
        If StCek = 1 Then LblErrMsg = "Can't Update Fix Status, Record already used in PO"
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim tanya

Me.MousePointer = vbHourglass
Select Case Index
    Case 0: 'Submit
        tanya = MsgBox("Do you really want to Process Purchase Request Status?", vbQuestion & vbYesNo, "Confirmation")
        If tanya = vbYes Then Call simpan
        
    Case 1: 'Cancel
        Call IsiGrid
End Select
Me.MousePointer = vbDefault
End Sub

Sub simpan()
Dim PORequestNo As String, StFix As Byte, StComplete As Byte

With grid
    For i = 1 To .Rows - 1
        If .Cell(flexcpChecked, i, 4) = flexChecked Then StFix = 1 Else StFix = 0
        If .Cell(flexcpChecked, i, 6) = flexChecked Then StComplete = 1 Else StComplete = 0
        PORequestNo = Trim(.TextMatrix(i, 0))
        
        If .Cell(flexcpChecked, i, 4) <> .Cell(flexcpChecked, i, 5) Or .Cell(flexcpChecked, i, 6) <> .Cell(flexcpChecked, i, 7) Then 'Cek jika tjd perubahan br Proses
            sql = "Update PORequest_Master Set Fix_Cls = '" & StFix & "', Complete_Cls = '" & StComplete & "', " & _
                "UserName = '" & userLogin & "', Last_Update = '" & Now & "' " & _
                "Where PORequest_No = '" & PORequestNo & "'"
            Db.Execute sql
        End If
    Next i
    
    Call IsiGrid
    LblErrMsg = DisplayMsg(1101)
End With
End Sub


Private Sub dtAwal_Change()
    Call headerGrid
    If Format(dtAwal, "yyyy-MM-dd") > Format(dtAkhir, "yyyy-MM-dd") Then LblErrMsg = DisplayMsg(4068) & " " & Format(dtAkhir, "dd MMM yyyy") Else LblErrMsg = ""
End Sub

Private Sub dtAkhir_Change()
    Call headerGrid
    If Format(dtAwal, "yyyy-MM-dd") > Format(dtAkhir, "yyyy-MM-dd") Then LblErrMsg = DisplayMsg(4066) & " " & Format(dtAwal, "dd MMM yyyy") Else LblErrMsg = ""
End Sub

Private Sub cboFix_Click()
    Call headerGrid
End Sub

Private Sub cboComplete_Click()
    Call headerGrid
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then Unload Me Else LblErrMsg.Caption = ErrMsg
End Sub

