VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPOCorrectionApproval 
   BackColor       =   &H00FDDFE3&
   Caption         =   "PO Correction Approval"
   ClientHeight    =   10485
   ClientLeft      =   120
   ClientTop       =   735
   ClientWidth     =   15120
   Icon            =   "frmPOCorrectionApproval.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   15120
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10020
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10020
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   345
      TabIndex        =   11
      Top             =   9330
      Width           =   14595
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
         Height          =   255
         Left            =   105
         TabIndex        =   12
         Top             =   195
         Width           =   14370
      End
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
      Left            =   345
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10020
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1695
      Left            =   345
      TabIndex        =   2
      Top             =   870
      Width           =   14595
      Begin VB.TextBox Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1215
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.#0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   262
         Width           =   2730
      End
      Begin VB.TextBox lblAddr 
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   300
         Width           =   5835
      End
      Begin VB.CommandButton CmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
         Height          =   345
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1200
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker SDate 
         Height          =   345
         Left            =   1785
         TabIndex        =   3
         Top             =   720
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
         Format          =   294256643
         CurrentDate     =   37868
      End
      Begin MSComCtl2.DTPicker EDate 
         Height          =   345
         Left            =   3720
         TabIndex        =   4
         Top             =   720
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
         Format          =   294256643
         CurrentDate     =   37868
      End
      Begin MSForms.ComboBox CboStatus 
         Height          =   315
         Left            =   1785
         TabIndex        =   19
         Top             =   1200
         Width           =   1545
         VariousPropertyBits=   746604571
         MaxLength       =   2
         DisplayStyle    =   7
         Size            =   "2725;556"
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
         Caption         =   "Status"
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
         Left            =   285
         TabIndex        =   17
         Top             =   1260
         Width           =   540
      End
      Begin MSForms.ComboBox cboSupp 
         Height          =   315
         Left            =   1785
         TabIndex        =   8
         Top             =   255
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2672;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code "
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
         Left            =   285
         TabIndex        =   7
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO Date"
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
         Left            =   285
         TabIndex        =   6
         Top             =   795
         Width           =   705
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         TabIndex        =   5
         Top             =   795
         Width           =   165
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13080
      TabIndex        =   1
      Top             =   270
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6375
      Left            =   360
      TabIndex        =   16
      Top             =   2760
      Width           =   14565
      _cx             =   25691
      _cy             =   11245
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PO Correction Approval"
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
      Left            =   315
      TabIndex        =   0
      Top             =   300
      Width           =   14640
   End
End
Attribute VB_Name = "FrmPOCorrectionApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim li_HakU As Integer, ls_Answer As String
Dim sql As String
Dim RS As New ADODB.Recordset
Dim bteColPOCorrectionNo As Byte
Dim bteColSupplier As Byte
Dim bteColPONo As Byte
Dim bteColPODate As Byte
Dim bteColDeliveryDate As Byte
Dim bteColChangeUser As Byte
Dim bteColChangeDate As Byte
Dim bteColApproved As Byte
Dim bteColApprovedUser As Byte
Dim bteColApprovedDate As Byte
Dim bteColHstatus As Byte
Dim bytSort As Byte

Private Sub cbostatus_Change()
If CboStatus.Text = "Yes" Then
    Label6.Text = "APPROVE"
ElseIf CboStatus.Text = "No" Then
    Label6.Text = "UNAPPROVE"
ElseIf CboStatus.Text = "ALL" Then
    Label6.Text = "ALL"
End If
Call Header
End Sub
 
Private Sub cboStatus_Click()
cbostatus_Change
End Sub
 
Private Sub cmdClear_Click()
Call blank
Call Header
End Sub

Private Sub cmdSearch_Click()
Call up_GridSearch
If cboSupp.Text = "" Then
cboSupp.SetFocus
LblErr.Caption = DisplayMsg("1054")
End If
End Sub

Private Sub CmdSubmit_Click()
Call update
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CmdMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub Form_Load()
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
Call adtocombo
    SDate.Value = Format(Now, "dd MMM YYYY")
    EDate.Value = Format(Now, "dd MMM YYYY")
    Call Header
    Call blank
End Sub

Private Sub Grid_DblClick()
    If grid.Row = 1 Then
        If bytSort = 0 Then
            grid.Sort = flexSortGenericDescending
            bytSort = 1
        Else
            grid.Sort = flexSortGenericAscending
            bytSort = 0
        End If
    End If
End Sub

Sub blank()
cboSupp.Value = "ALL"
CboStatus.Text = "ALL"
'txtName.Text = ""
SDate.Value = Format(Now, "dd MMM yyyy")
EDate.Value = Format(Now, "dd MMM yyyy")
End Sub

Sub Header()
    bteColPOCorrectionNo = 0
    bteColSupplier = 1
    bteColPONo = 2
    bteColPODate = 3
    bteColDeliveryDate = 4
    bteColChangeUser = 5
    bteColChangeDate = 6
    bteColApproved = 7
    bteColApprovedUser = 8
    bteColApprovedDate = 9
    bteColHstatus = 10
    
    With grid
        .ColS = 11
        .Rows = 1
         
        .TextMatrix(0, bteColPOCorrectionNo) = "PO Correction Approval"
        .TextMatrix(0, bteColSupplier) = "Supplier"
        .TextMatrix(0, bteColPONo) = "PO NO"
        .TextMatrix(0, bteColPODate) = "PO Date"
        .TextMatrix(0, bteColDeliveryDate) = "Delivery Date"
        .TextMatrix(0, bteColChangeUser) = "Change User"
        .TextMatrix(0, bteColChangeDate) = "Change Date"
        .TextMatrix(0, bteColApproved) = "Approved"
        .TextMatrix(0, bteColApprovedUser) = "Approved User"
        .TextMatrix(0, bteColApprovedDate) = "Approved Date"
        .TextMatrix(0, bteColHstatus) = ""
        
        .ColWidth(bteColPOCorrectionNo) = 1900
        .ColWidth(bteColSupplier) = 2700
        .ColWidth(bteColPONo) = 2500
        .ColWidth(bteColPODate) = 1400
        .ColWidth(bteColDeliveryDate) = 1400
        .ColWidth(bteColChangeUser) = 1400
        .ColWidth(bteColChangeDate) = 1800
        .ColWidth(bteColApproved) = 1200
        .ColWidth(bteColApprovedUser) = 1500
        .ColWidth(bteColApprovedDate) = 1500
        .ColHidden(bteColHstatus) = True
        .ColHidden(bteColPONo) = False
    End With
End Sub

Private Sub up_CheckHeader()
Dim iRow As Long
Dim iCol As Long
Dim CheckAll As Boolean
    With grid
        For iCol = 7 To 7
            CheckAll = True
            For iRow = 1 To .Rows - 1
                If .Cell(flexcpChecked, iRow, iCol) = flexUnchecked Then
                    CheckAll = False
                    Exit For
                End If
            Next
            If CheckAll Then
                .Cell(flexcpChecked, 0, iCol) = flexChecked
            Else
                .Cell(flexcpChecked, 0, iCol) = flexUnchecked
            End If
        Next
    End With
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Call up_CheckHeader
Dim rsSimpan As New ADODB.Recordset
Dim sqlB As String
Dim intacc As Integer
Dim intupd As Integer
Dim SSql As String
Dim li_Row As Integer
Dim rsCek As New ADODB.Recordset
    With grid
        For i = 1 To grid.Rows - 1
            If .Cell(flexcpChecked, i, bteColApproved) <> .Cell(flexcpChecked, i, bteColHstatus) Then
        sql = " select * from (select * From PurchaseOrder_Detail_History where PO_Correction_no = '" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "' ) podh " & vbCrLf & _
            " inner join InvoiceSupplier_Detail isd on podh.Item_Code=isd.Item_Code and podh.po_no=isd.po_no " & vbCrLf & _
            ""
            If rsCek.State <> adStateClosed Then rsCek.Close
                rsCek.Open sql, Db, adOpenStatic, adLockReadOnly
                If Not rsCek.EOF Then
                    .Cell(flexcpBackColor, i, bteColPOCorrectionNo, i, bteColHstatus) = vbRed
                End If
            End If
        Next i
    End With
End Sub

Private Sub up_SetCheck(Col As Long, Checked As VSFlex8Ctl.CellCheckedSettings)
Dim iRow As Long
    With grid
        For iRow = 1 To .Rows - 1
            .Cell(flexcpChecked, iRow, Col) = Checked
        Next
    End With
End Sub
 
Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = 7) Then
        Cancel = True
    ElseIf Row = 0 Then
        With grid
            If .Cell(flexcpChecked, Row, Col) = flexUnchecked Then
                up_SetCheck Col, flexChecked
            Else
                up_SetCheck Col, flexUnchecked
            End If
        End With
        up_CheckHeader
    End If
End Sub

Private Sub up_GridSearch()
    Dim ls_sql As String
    Dim RS As New ADODB.Recordset
    Dim li_Row As Integer
    Dim a As String
    Header
    
    ls_sql = uf_SQLSearch
    If RS.State = adStateOpen Then RS.Close
    Set RS = Db.Execute(ls_sql)
    
    With grid
    
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
            
            .TextMatrix(li_Row, bteColPOCorrectionNo) = Trim(RS!PO_Correction_no)
            
            .TextMatrix(li_Row, bteColSupplier) = Trim(RS!Supplier_Code) & "-" & Trim(RS!trade_name)
            .TextMatrix(li_Row, bteColPONo) = Trim(RS!po_no)
            .TextMatrix(li_Row, bteColPODate) = Format(RS!po_date, "dd MMM yyyy")
            
            .TextMatrix(li_Row, bteColDeliveryDate) = Format(RS!delivery_Date, "dd MMM yyyy")
            .TextMatrix(li_Row, bteColChangeUser) = Trim(RS!last_user)
            .TextMatrix(li_Row, bteColChangeDate) = Format(RS!Register_Date, "dd MMM yyyy hh:ss")
            .Cell(flexcpChecked, li_Row, bteColApproved) = IIf(Val(RS.Fields("Approved_Cls") & "") = 1, flexChecked, flexUnchecked)
            
            .TextMatrix(li_Row, bteColApprovedDate) = Trim(IIf(IsNull(RS!Approved_date), "", Format(RS!Approved_date, "dd MMM yyyy")))
            .TextMatrix(li_Row, bteColApprovedUser) = Trim(IIf(IsNull(RS!Approved_User), "", RS!Approved_User))
            .Cell(flexcpChecked, li_Row, bteColHstatus) = IIf(Val(RS.Fields("Approved_Cls") & "") = 1, flexChecked, flexUnchecked)
'            .TextMatrix(li_Row, bteColHstatus) = Trim(rs!item_name)
            .Cell(flexcpPictureAlignment, li_Row, bteColApproved, li_Row, bteColApproved) = flexAlignCenterCenter
            
'           If Trim(.TextMatrix(li_Row, bteColPOCorrectionNo)) = "POC.2017.05.0006" Then
'            .RowHidden(li_Row) = True
'            Else
'            .RowHidden(li_Row) = False
'        End If
            RS.MoveNext
        Wend
        up_CheckHeader
        RS.Close
        a = .Rows - 1
        If a <= 0 Then
        LblErr.Caption = DisplayMsg("0013")
        Else
        LblErr.Caption = ""
        End If
    End With
End Sub

Sub adtocombo()

sql = "SELECT Trade_Code, Trade_Name FROM Trade_Master where trade_cls='2' or trade_cls='3'"
Set RS = Db.Execute(sql)

With cboSupp
 
.clear
.columnCount = 2
.ColumnWidths = "80 pt;300 pt"
.ListWidth = 380
.ListRows = 15
.AddItem ""
.List(0, 0) = strAll
.List(0, 1) = strAll
i = 1
Do Until RS.EOF
    .AddItem ""
    .List(i, 0) = Trim(RS!Trade_Code)
    .List(i, 1) = Trim(RS!trade_name)
    i = i + 1
    RS.MoveNext
Loop
.ListIndex = 0
End With

    With CboStatus
        .AddItem ("ALL")
        .AddItem ("Yes")
        .AddItem ("No")
    End With

End Sub

Private Sub CboSupp_Change()
If cboSupp.MatchFound = True Then
    txtName.Text = Trim(cboSupp.Column(1))
Else
    txtName.Text = ""
End If
Call Header
End Sub

Private Sub CboSupp_Click()
Call CboSupp_Change
End Sub

Sub update()
Dim rsSimpan As New ADODB.Recordset
Dim sqlB As String
Dim intacc As Integer
Dim intupd As Integer
Dim SSql As String
Dim li_Row As Integer
Dim rsCek As New ADODB.Recordset
    With grid
        For i = 1 To grid.Rows - 1
            If .Cell(flexcpChecked, i, bteColApproved) <> .Cell(flexcpChecked, i, bteColHstatus) Then
                     
        sql = " select * from (select * From PurchaseOrder_Detail_History where PO_Correction_no = '" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "' ) podh " & vbCrLf & _
            " inner join InvoiceSupplier_Detail isd on podh.Item_Code=isd.Item_Code and podh.po_no='aaaaaaa' --and podh.po_no=isd.po_no" & vbCrLf & _
            ""
            If rsCek.State <> adStateClosed Then rsCek.Close
                rsCek.Open sql, Db, adOpenStatic, adLockReadOnly
                If rsCek.EOF Then
                    sql = "update PurchaseOrder_Master_History " & vbCrLf & _
                        "   set Approved_Cls ='" & IIf(.Cell(flexcpChecked, i, bteColApproved) = flexUnchecked, "0", "1") & "', " & _
                        "   Approved_user ='" & IIf(.Cell(flexcpChecked, i, bteColApproved) = flexUnchecked, Null, userLogin) & "', " & vbCrLf & _
                        "   Approved_date =" & IIf(.Cell(flexcpChecked, i, bteColApproved) = flexUnchecked, "Null", " GETDATE()") & " " & vbCrLf & _
                        "   where PO_Correction_no = '" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "' " & vbCrLf
                    Db.Execute sql
                    
            If .Cell(flexcpChecked, i, bteColApproved) = flexChecked Then
                sql = "insert into purchaseorder_master_history1" & vbCrLf & _
                        " select  '" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "',*,NULL,NULL,NULL,NULL From PurchaseOrder_Master " & vbCrLf & _
                        " where po_no = '" & Trim(grid.TextMatrix(i, bteColPONo)) & "' " & vbCrLf
                    'save purchaseorder_detail_history1
                sql = sql + "insert into purchaseorder_detail_history1" & vbCrLf & _
                        " select  '" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "',* From PurchaseOrder_detail " & vbCrLf & _
                        " where po_no = '" & Trim(grid.TextMatrix(i, bteColPONo)) & "' " & vbCrLf
                    'delete purchaseorder_master
                sql = sql + "delete purchaseorder_master" & vbCrLf & _
                        " where po_no = '" & Trim(grid.TextMatrix(i, bteColPONo)) & "' " & vbCrLf
                    'delete purchaseorder_detail
                sql = sql + "delete purchaseorder_detail" & vbCrLf & _
                        " where po_no = '" & Trim(grid.TextMatrix(i, bteColPONo)) & "' " & vbCrLf
                     Db.Execute sql
                    'save purchaseorder_master
                sql = sql + " insert into purchaseorder_master" & vbCrLf & _
                        " select PO_No,Supplier_Code,Period,PO_Date,Delivery_Date,WHTo,Deliver_To,PriceCondition_Cls,PaymentTerm_Cls,POPacking_Cls," & vbCrLf & _
                        " Insurance_Cls,PO_LOT,Transportation_Cls,POMarking1,POMarking2,POMarking3,POMarking4,POMarking5,POMarking6,Remarks,Amount," & vbCrLf & _
                        " PPN,PPH,Total_Amount,Fix_Cls,SheetCoil_Cls,Revise_No,Others_Cls,Discount,Last_Update,Last_User,Register_Date,POSet_Code,POSet_SeqNo" & vbCrLf & _
                        " From PurchaseOrder_Master_history " & vbCrLf & _
                        " where PO_Correction_no = '" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "' " & vbCrLf
                    'save purchaseorder_detail
                sql = sql + "insert into purchaseorder_detail" & vbCrLf & _
                        " select Seq_No,PO_No,Item_Code,Item_Name,PORequest_No,POReq_SeqNo,Delivery_Date,Price,Price_Service,Currency_Code,Unit_Cls," & vbCrLf & _
                        " Qty,Amount,Amount_Service,Complete_Cls,Last_Update,Last_User,Register_Date,Price_Adj" & vbCrLf & _
                        " From PurchaseOrder_Detail_history " & vbCrLf & _
                        " where PO_Correction_no = '" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "' " & vbCrLf
                    Db.Execute sql
                    
                    
               sql = "update Part_Receipt" & vbCrLf & _
                     "set Price=B.Price,Amount=B.Price*A.Qty" & vbCrLf & _
                     "From Part_Receipt A" & vbCrLf & _
                     "INNER JOIN (select * From PurchaseOrder_Detail_history where PO_Correction_no='" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "') B ON A.PO_No=B.PO_No AND A.Item_Code=B.Item_Code" & vbCrLf & _
                     "where  A.Price<>B.Price " & vbCrLf
                     
                Db.Execute sql
               
                    
            ElseIf .Cell(flexcpChecked, i, bteColApproved) = flexUnchecked Then
                    'delete purchaseorder_master
                sql = sql + "delete purchaseorder_master" & vbCrLf & _
                        " where po_no = '" & Trim(grid.TextMatrix(i, bteColPONo)) & "' " & vbCrLf
                    'delete purchaseorder_detail
                sql = sql + "delete purchaseorder_detail" & vbCrLf & _
                        " where po_no = '" & Trim(grid.TextMatrix(i, bteColPONo)) & "' " & vbCrLf
                    'save purchaseorder_master
                sql = sql + " insert into purchaseorder_master" & vbCrLf & _
                        " select PO_No,Supplier_Code,Period,PO_Date,Delivery_Date,WHTo,Deliver_To,PriceCondition_Cls,PaymentTerm_Cls,POPacking_Cls," & vbCrLf & _
                        " Insurance_Cls,PO_LOT,Transportation_Cls,POMarking1,POMarking2,POMarking3,POMarking4,POMarking5,POMarking6,Remarks,Amount," & vbCrLf & _
                        " PPN,PPH,Total_Amount,Fix_Cls,SheetCoil_Cls,Revise_No,Others_Cls,Discount,Last_Update,Last_User,Register_Date,POSet_Code,POSet_SeqNo" & vbCrLf & _
                        " From PurchaseOrder_Master_history1 " & vbCrLf & _
                        " where PO_Correction_no = '" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "' " & vbCrLf
                        Db.Execute sql
                    'save purchaseorder_detail
                sql = sql + "insert into purchaseorder_detail" & vbCrLf & _
                        " select Seq_No,PO_No,Item_Code,Item_Name,PORequest_No,POReq_SeqNo,Delivery_Date,Price,Price_Service,Currency_Code,Unit_Cls," & vbCrLf & _
                        " Qty,Amount,Amount_Service,Complete_Cls,Last_Update,Last_User,Register_Date,Price_Adj" & vbCrLf & _
                        " From PurchaseOrder_Detail_history1 " & vbCrLf & _
                        " where PO_Correction_no = '" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "' " & vbCrLf
                    'delete purchaseorder_master_history1
                sql = sql + "delete purchaseorder_master_history1" & vbCrLf & _
                        " where PO_Correction_no = '" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "' " & vbCrLf
                    'delete purchaseorder_detail_history1
                sql = sql + "delete purchaseorder_detail_history1" & vbCrLf & _
                        " where PO_Correction_no = '" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "' " & vbCrLf
                    Db.Execute sql
                    

               sql = "update Part_Receipt" & vbCrLf & _
                     "set Price=B.Price" & vbCrLf & _
                     "From Part_Receipt A" & vbCrLf & _
                     "INNER JOIN (select * From PurchaseOrder_Detail_history1 where PO_Correction_no='" & Trim(grid.TextMatrix(i, bteColPOCorrectionNo)) & "') B ON A.PO_No=B.PO_No AND A.Item_Code=B.Item_Code" & vbCrLf & _
                     "where  A.Price<>B.Price " & vbCrLf
                     
                Db.Execute sql
                    
            End If
            Call up_GridSearch
               Else
                .Cell(flexcpBackColor, i, bteColPOCorrectionNo, i, bteColHstatus) = vbRed
               LblErr.Caption = DisplayMsg("4110")
            End If
            LblErr.Caption = DisplayMsg("1101")
            End If
        Next i
    End With
End Sub

Private Function uf_SQLSearch() As String
Dim status As String
If CboStatus.Text = "Yes" Then
    status = 1
    ElseIf CboStatus.Text = "No" Then
    status = 0
End If

uf_SQLSearch = " Select * From (select * from PurchaseOrder_Master_History " & vbCrLf & _
                " where PO_Date >= '" & Format(SDate.Value, "yyyy-mm-dd") & "' and PO_Date <= '" & Format(EDate.Value, "yyyy-mm-dd") & "' " & vbCrLf & _
                IIf(cboSupp.Text = "ALL", "", " and supplier_code='" & Trim(cboSupp.Text) & "' ") & vbCrLf & _
                IIf(CboStatus.Text = "ALL", "", "and Coalesce(Approved_Cls,'0')='" & status & "'  ") & vbCrLf & _
                " ) pomh " & vbCrLf & _
                " inner join trade_master tm on pomh.supplier_code=tm.trade_code " & vbCrLf & _
                ""
End Function

Private Sub SDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
Call Header
End Sub

Private Sub sdate_Change()
Call Header
End Sub
