VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmApList 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Payment Amount Entry"
   ClientHeight    =   10950
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   15120
   Icon            =   "FrmAPList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Index           =   1
      Left            =   13366
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9795
      Width           =   1140
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
      Left            =   10876
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9795
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
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
      Left            =   12136
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9795
      Width           =   1140
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "FrmAPList.frx":0E42
      Left            =   886
      List            =   "FrmAPList.frx":0E44
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2775
      Width           =   1485
   End
   Begin VB.CommandButton BtnCreate 
      BackColor       =   &H0080FFFF&
      Caption         =   "Create"
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
      Left            =   11026
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2745
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1305
      Left            =   639
      TabIndex        =   17
      Top             =   1245
      Width           =   13875
      Begin VB.TextBox txt_name 
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
         Height          =   225
         Left            =   3735
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "txt_name"
         Top             =   360
         Width           =   4650
      End
      Begin MSComCtl2.DTPicker APDateEnd 
         Height          =   315
         Left            =   3690
         TabIndex        =   2
         Top             =   765
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
         Format          =   151453699
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker APDateSt 
         Height          =   315
         Left            =   1530
         TabIndex        =   1
         Top             =   765
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
         Format          =   151453699
         CurrentDate     =   37810
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
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
         Left            =   270
         TabIndex        =   22
         Top             =   825
         Width           =   1110
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   795
         Width           =   375
      End
      Begin VB.Line Line2 
         X1              =   3720
         X2              =   8370
         Y1              =   600
         Y2              =   600
      End
      Begin MSForms.ComboBox CboSupp 
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   345
         Width           =   2010
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3545;556"
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
         Caption         =   "Supplier CD"
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
         Left            =   270
         TabIndex        =   20
         Top             =   375
         Width           =   1035
      End
      Begin VB.Label lblfix 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   11910
         TabIndex        =   19
         Top             =   900
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   570
      Left            =   639
      TabIndex        =   14
      Top             =   9105
      Width           =   13875
      Begin VB.Label lblerror 
         Alignment       =   2  'Center
         BackColor       =   &H00FDDFE3&
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
         TabIndex        =   15
         Top             =   195
         Width           =   13665
      End
   End
   Begin VB.CommandButton cmdSubMenu 
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
      Left            =   646
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9795
      Width           =   1290
   End
   Begin MSComCtl2.DTPicker DateVou 
      Height          =   315
      Left            =   8955
      TabIndex        =   5
      Top             =   2775
      Width           =   1500
      _ExtentX        =   2646
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
      Format          =   151388163
      CurrentDate     =   37798
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   4140
      Left            =   645
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3315
      Width           =   13875
      _cx             =   24474
      _cy             =   7302
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
      FocusRect       =   5
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
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
      WordWrap        =   -1  'True
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
   Begin VSFlex8Ctl.VSFlexGrid Grid2 
      Height          =   1500
      Left            =   4215
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7560
      Width           =   10305
      _cx             =   18177
      _cy             =   2646
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   10932991
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16637923
      BackColorAlternate=   16777215
      GridColor       =   12582912
      GridColorFixed  =   12582912
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   5
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
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
      WordWrap        =   -1  'True
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
      Left            =   12630
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   375
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   714
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Voucher No"
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
      Index           =   1
      Left            =   2640
      TabIndex        =   24
      Top             =   2805
      Width           =   1560
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Date"
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
      Index           =   2
      Left            =   7455
      TabIndex        =   23
      Top             =   2805
      Width           =   1455
   End
   Begin MSForms.ComboBox CboApNo 
      Height          =   345
      Left            =   4335
      TabIndex        =   4
      Top             =   2760
      Width           =   2175
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3836;609"
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Amount Entry"
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
      Index           =   0
      Left            =   645
      TabIndex        =   16
      Top             =   375
      Width           =   13875
   End
End
Attribute VB_Name = "FrmApList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim CtrCreate As Integer
Dim rsSup As New ADODB.Recordset
Dim sql As String
Dim ColCls, ColInvDate, ColInvNo, ColCurr, ColAmountInv, ColPPnInv, ColAmountPay, ColPPnPay, ColAmountRem, ColPPnRem, ColDueDate, colpaid As Long
Dim ColBantuCurr As Integer
Dim ColCurrTT, ColTotalPay, ColGrandTotalInv, ColGrandTotalPay, ColGrandToTalRem, ColPay As Integer
Dim TampInv, TampPPn As Double

Private Sub APDateEnd_Change()
If combo1.ListIndex <> 0 Then
    NoAp
End If
End Sub

Private Sub APDateSt_Change()
If combo1.ListIndex <> 0 Then
    NoAp
End If
End Sub

Private Sub BtnCreate_Click()
Dim RS As New Recordset

lblerror = ""
If hakUpdate(Me.Name) = 0 Then lblerror = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub

'SUPPLIER CHECK
If cboSupp.Text = "" Then
    lblerror = DisplayMsg(1054)
    Exit Sub
Else
     For i = 0 To cboSupp.ListCount - 1
        If Trim(cboSupp.List(i, 0)) = Trim(cboSupp.Text) Then
            GoTo Okay
        End If
    Next
    MousePointer = Default
    Exit Sub
End If

Okay:

If combo1.ListIndex = 0 Then
'AP No
If cboapno.Text = "" Then
    lblerror = DisplayMsg(4087)
    Exit Sub
End If

    'create
    CtrCreate = 1
    'Get No dimatikan krn dikontrol manual , 20110307
    'GetNo
    sql = "Select AP_No FROM AP_Master WHERE AP_No = '" & Trim(cboapno.Text) & "'"
    If RS.State = 1 Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If Not RS.EOF Then
        lblerror.Caption = DisplayMsg(1023)
        Exit Sub
    End If
    
    Db.BeginTrans
    sql = " INSERT INTO AP_Master (Supplier_Code, AP_No, AP_Date, Last_Update, Last_User) " & _
            "VALUES('" & cboSupp & "', '" & cboapno & "', '" & Format(DateVou.Value, "yyyy-mm-dd") & "', getdate(), '" & userLogin & "')"
    Db.Execute sql
    Db.CommitTrans
    
    combo1.ListIndex = 1
    CboApNo_Change
    
    FillGrid
    HitungTotalGrid
    
    CtrCreate = 0
ElseIf combo1.ListIndex = 1 Then

'AP No
If cboapno.Text = "" Then
    lblerror = DisplayMsg(4087)
    Exit Sub
Else
    For i = 0 To cboapno.ListCount - 1
       If Trim(cboapno.List(i, 0)) = Trim(cboapno.Text) Then
           GoTo Okay2
        End If
    Next
    GoTo Okay2
    MousePointer = Default
    Exit Sub
End If
Okay2:

    'update
    FillGrid
    HitungTotalGrid
End If
End Sub

Private Sub CboApNo_Change()
Dim RsApR As New ADODB.Recordset
If combo1.ListIndex = 1 Then
    Header

    RsApR.Open "SELECT ap_date from ap_master where ap_no = '" & cboapno & "' and supplier_code = '" & cboSupp & "'", Db, adOpenKeyset, adLockOptimistic
    If Not RsApR.EOF And Not RsApR.BOF Then
        DateVou.Value = RsApR.Fields("Ap_date")
    End If
End If
End Sub

Private Sub CboApNo_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then
KeyAscii = 0
End If
End Sub

Private Sub CboSupp_Change()
    If CtrCreate = 0 Then
        cboapno.Text = ""
    End If
    txt_name.Text = ""
    ComboSuppUbah
    combo1.ListIndex = 1
End Sub

Private Sub CboSupp_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then
KeyAscii = 0
End If
End Sub

Private Sub cmdCancel_Click()
Dim suppno, apno As String
Dim dt1, dt2 As Date
suppno = cboSupp
apno = cboapno
dt1 = APDateSt
dt2 = APDateEnd
cmdClear_Click
cboSupp = suppno
APDateSt = dt1
APDateEnd = dt2
cboapno = apno
BtnCreate_Click
End Sub

Private Sub cmdClear_Click()
Form_Load
End Sub

Private Sub CmdSubMenu_Click()
    'PROG : DELETE WAKTU KELUAR BILA GAK ADA DETAIL
    Db.BeginTrans
    Dim RsDel As New ADODB.Recordset
    RsDel.Open "select * from ap_detail where ap_no = '" & cboapno & "' and supplier_code = '" & cboSupp & "'", Db, adOpenKeyset, adLockOptimistic
    If RsDel.EOF And RsDel.BOF Then
        sql = "delete ap_master where ap_no = '" & cboapno & "' and supplier_code = '" & cboSupp & "'"
        Db.Execute sql
    End If
    Db.CommitTrans
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub CmdSubmit_Click(Index As Integer)
Dim RsCu As New ADODB.Recordset
Dim EX As Double

If hakUpdate(Me.Name) = 0 Then lblerror = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub

If grid.Rows <= 2 Then
    Exit Sub
End If

lblerror = ""
Db.BeginTrans
Me.MousePointer = vbHourglass

'hapus ap detail
        sql = "delete from ap_detail where ap_no = '" & cboapno & "' and supplier_code = '" & cboSupp & "'"
        Db.Execute sql

For i = 2 To grid.Rows - 1
    If grid.Cell(flexcpChecked, i, 0) = flexChecked Then

        If grid.TextMatrix(i, ColBantuCurr) <> "03" Then
            If RsCu.State = 1 Then RsCu.Close
            'cari nilai tukar
            RsCu.Open "SELECT daily_exchangerate from daily_exchangerate where exchangerate_date " & _
                    " = '" & Format(DateVou, "yyyy-mm-dd") & "' and currency_code " & _
                    " = '" & grid.TextMatrix(i, ColBantuCurr) & "'", Db, adOpenKeyset, adLockOptimistic
            If RsCu.EOF And RsCu.BOF Then
                EX = 0
            Else
                EX = RsCu.Fields("daily_exchangerate")
            End If
            RsCu.Close

        Else
            'mata uang indo sendiri
            EX = 1
        End If
        
        sql = " INSERT INTO AP_Detail (Supplier_Code, AP_No, Invoice_No, Currency_Code, Amount, PPN, Exchange_Rate, Exchange_Amount, Last_Update, Last_User) " & _
                "VALUES ('" & cboSupp & "', '" & cboapno & "', '" & grid.TextMatrix(i, ColInvNo) & "', '" & grid.TextMatrix(i, ColBantuCurr) & "', " & _
                CDbl(grid.TextMatrix(i, ColAmountPay)) * 1 & ", " & CDbl(grid.TextMatrix(i, ColPPnPay)) * 1 & ", " & CDbl(EX) & ", " & (CDbl(grid.TextMatrix(i, ColAmountPay)) * EX) & ", " & _
                "getdate(), '" & userLogin & "')"
        Db.Execute sql
    End If
Next

sql = " UPDATE AP_Master " & _
        "SET AP_Date = '" & Format(DateVou, "yyyy-mm-dd") & "', " & _
        "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
        "WHERE ap_no = '" & cboapno & "' and supplier_code = '" & cboSupp & "' "
Db.Execute sql

Db.CommitTrans
Me.MousePointer = vbDefault
lblerror = DisplayMsg(1000)
HitungTotalGrid
End Sub

Private Sub Combo1_Click()
If combo1.ListIndex = 1 Then
    BtnCreate.Caption = "Update"
    If CtrCreate = 0 Then
        InitSave
    End If
    CboSupp_Change
    cboapno.locked = False
ElseIf combo1.ListIndex = 0 Then
    BtnCreate.Caption = "Create"
    'CboApNo.locked = True
    InitSave
    cboapno.Text = ""
    cboapno.clear
    'GetNo dimatikan berdasarkan permintaan Pak Teguh, krn dikontrol manual 20110307
    'GetNo
    Header
End If
End Sub

Sub GetNo()
Dim rsno As New ADODB.Recordset
If rsno.State = 1 Then rsno.Close
' Kawai Format need count Voucher per Month with Format --> P/B/MM/999
'rsno.Open "Select isnull(max(right(ap_no,3)),0) as ap_no from AP_Master where year(ap_date) = '" & year(DateVou.Value) & "'", Db, adOpenKeyset, adLockOptimistic

rsno.Open "Select isnull(max(right(ap_no,3)),0) as ap_no from AP_Master where year(ap_date) = '" & Year(DateVou.Value) & _
                "' And Month(ap_Date)= '" & Month(DateVou.Value) & "'", Db, adOpenKeyset, adLockOptimistic
                
If Not rsno.EOF And Not rsno.BOF Then
    rsno.MoveFirst
    'cboapno.Text = Format(DateVou.Value, "yyyymmdd") & Format(CDbl(rsno.Fields("ap_no")) + 1, "0000")
    cboapno.Text = "P/B/" & Format(DateVou.Value, "mm") & "/" & Format(CDbl(rsno.Fields("ap_no")) + 1, "000")
Else
    'cboapno.Text = Format(DateVou.Value, "yyyymmdd") & "001"
    cboapno.Text = "P/B/" & Format(DateVou.Value, "mm") & "/001"
End If
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    lblerror.Caption = ErrMsg
End If
End Sub

Private Sub DateVou_Change()
If combo1.ListIndex = 0 Then
    GetNo
End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
HakU = hakUpdate(Me.Name)
CtrlMenu1.FormName = Me.Name
Me.Caption = lblJudul(0).Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

    lblerror = ""
    'setting col
    ColCls = 0
    ColInvDate = 1
    ColInvNo = 2
    ColCurr = 3
    ColAmountInv = 4
    ColPPnInv = 5
    ColAmountPay = 6
    ColPPnPay = 7
    ColAmountRem = 8
    ColPPnRem = 9
    ColDueDate = 10
    ColBantuCurr = 11
    colpaid = 12
    
    ColCurrTT = 0
    ColPay = 1
    ColGrandTotalInv = 2
    ColGrandTotalPay = 3
    ColGrandToTalRem = 4
    TampInv = 0
    TampPPn = 0
    'Conec
    adtocombo
    Init
    Header
End Sub

Sub Conec()
    Db.Open "Provider=SQLOLEDB.1;Persist Security Info=true;User ID=sa;Initial Catalog=Santelindo_real;Data Source=.;pwd=;"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

'======================================================================================================================================
Sub FillGrid()
Dim RsInv As New ADODB.Recordset
Dim RS As New ADODB.Recordset
Dim i As Long
Dim XCount As Integer
Header
sql = " Select  " & _
            " Invoice_Date,     " & _
            " Invoice_No,   " & _
            " Curr, " & _
            " AmountInv,     " & _
            " PPN,  " & _
            " AmountPayed, " & _
            " PPNPayed, " & _
            " AmountInv  as AmountRem,     " & _
            " PPN  as PPNRem, " & _
            " exchange_amount,     "

sql = sql + " Due_Date,    " & _
            " Paid " & _
            " from( " & _
            "   SELECT " & _
            "   (select invoice_date from invoicesupplier_master where invoice_no = ta.invoice_no) as Invoice_Date, " & _
            "   Invoice_no, " & _
            "   Curr, " & _
            "   (select AirFreight_Amount+ Total_Amount from invoicesupplier_master where invoice_no = ta.invoice_no) as AmountInv, " & _
            "   (select PPN from invoicesupplier_master where invoice_no = ta.invoice_no) as PPN, " & _
            "   AmountPayed, " & _
            "   PPNPayed, " & _
            "   exchange_amount, "

sql = sql + "   (select Due_date from invoicesupplier_master where invoice_no = ta.invoice_no) as Due_Date, " & _
            "   (select paid_cls from invoicesupplier_master where invoice_no = ta.invoice_no) as paid " & _
            "   from( " & _
            "       SELECT  " & _
            "       Invoice_no, " & _
            "       Currency_Code as Curr, " & _
            "       Amount as AmountPayed, " & _
            "       PPN as PPNPayed, " & _
            "       exchange_amount " & _
            "       from Ap_Detail " & _
            "       Where ap_no = '" & cboapno & "' " & _
            "   )ta " & _
            " )TAX "

sql = sql + "  " & _
            " UNION " & _
            "  " & _
            " SELECT  " & _
            " Invoice_Date,     " & _
            " Invoice_No,   " & _
            " Curr, " & _
            " AmountInv,     " & _
            " PPN,  " & _
            " AmountPayed, " & _
            " PPNPayed, "

sql = sql + " AmountInv  as AmountRem,     " & _
            " PPN as PPNRem, " & _
            " exchange_amount,     " & _
            " Due_Date,    " & _
            " 0 as Paid " & _
            " FROM " & _
            " ( " & _
            "   SELECT     " & _
            "   Invoice_Date,     " & _
            "   Invoice_No,     " & _
            "   (Select top 1 Currency_Code from invoiceSupplier_detail where supplier_code = '" & cboSupp & "' and invoice_no= tb.invoice_no) as Curr,     " & _
            "   AmountInv,     "

sql = sql + "   PPN,  " & _
            "   0 as AmountPayed, " & _
            "   0 as PPNPayed, " & _
            "   exchange_amount,     " & _
            "   Due_Date     " & _
            "   from  " & _
            "   (     " & _
            "       SELECT        " & _
            "       InvoiceSupplier_Master.Invoice_Date,      " & _
            "       InvoiceSupplier_Master.Invoice_No,       " & _
            "       max(invoiceSupplier_Master.AirFreight_Amount)+ max(InvoiceSupplier_Master.Total_Amount) as AmountInv,       "

sql = sql + " max(InvoiceSupplier_Master.PPN) PPN," & _
            " max(InvoiceSupplier_Master.exchange_amount) exchange_amount,      " & _
            " InvoiceSupplier_Master.Due_Date, sum(isnull(ap_detail.Amount,0)) apamount       " & _
            " FROM InvoiceSupplier_Master Left Join ap_detail on InvoiceSupplier_Master.invoice_no = ap_detail.invoice_no and InvoiceSupplier_Master.supplier_Code=ap_detail.supplier_Code " & _
            " " & _
            " group by InvoiceSupplier_Master.fix_cls,invoiceSupplier_Master.Invoice_Date, InvoiceSupplier_Master.Invoice_No,  InvoiceSupplier_Master.Due_Date,InvoiceSupplier_Master.supplier_code, InvoiceSupplier_Master.paid_cls " & _
            " having InvoiceSupplier_Master.fix_cls = '1' and InvoiceSupplier_Master.supplier_code = '" & cboSupp & "' " & _
            " and  InvoiceSupplier_Master.invoice_no+InvoiceSupplier_Master.supplier_Code IN " & _
            " (select invoice_no + supplier_code from invoicesupplier_detail)  " & _
            " and InvoiceSupplier_Master.invoice_date >= '" & Format(APDateSt, "yyyy-mm-dd") & "'" & _
            " and InvoiceSupplier_Master.invoice_date <= '" & Format(APDateEnd, "yyyy-mm-dd") & "'" & _
            " and InvoiceSupplier_Master.invoice_no + InvoiceSupplier_Master.supplier_code NOT IN " & _
            " (select invoice_no+supplier_Code from ap_detail where supplier_code = '" & cboSupp & "' )/* and ap_no = '" & cboapno & "')*/ " & _
            " and (InvoiceSupplier_Master.paid_cls is null or InvoiceSupplier_Master.paid_cls = 0)  and (max(InvoiceSupplier_Master.AirFreight_Amount)+ max(InvoiceSupplier_Master.Total_Amount))- isnull(sum(ap_detail.amount),0) > 0"

sql = sql + "   )Tb   " & _
            " )TBX "

XCount = 0
RsInv.Open sql, Db, adOpenKeyset, adLockOptimistic
If Not RsInv.EOF Or Not RsInv.BOF Then
    i = 2
    With grid
        While Not RsInv.EOF
            .Rows = .Rows + 1
            If RsInv.Fields("AmountPayed") <> 0 Or RsInv.Fields("PPNPayed") <> 0 Then
                .Cell(flexcpChecked, i, ColCls) = flexChecked
                XCount = 1
            Else
                .Cell(flexcpChecked, i, ColCls) = flexUnchecked
            End If
            .TextMatrix(i, ColInvDate) = Format(RsInv.Fields("invoice_date"), "dd MMM yyyy")
            .TextMatrix(i, ColInvNo) = RsInv.Fields("invoice_no")
            .TextMatrix(i, ColCurr) = uf_GetCurrencyDescription(Trim(RsInv.Fields("Curr")))
            .TextMatrix(i, ColAmountInv) = Format(RsInv.Fields("AmountInv"), gs_formatAmount)
            .TextMatrix(i, ColPPnInv) = Format(RsInv.Fields("ppn"), gs_formatAmount)
            .TextMatrix(i, ColAmountPay) = Format(RsInv.Fields("AmountPayed"), gs_formatAmount)
            .TextMatrix(i, ColPPnPay) = Format(RsInv.Fields("PPNPayed"), gs_formatAmount)
            
            If RS.State = 1 Then RS.Close
            RS.Open "select isnull(sum(amount),0) as sumAm, isnull(sum(ppn),0) as sumPPN from ap_detail where invoice_no = '" & RsInv.Fields("invoice_no") & "'", Db, adOpenKeyset, adLockOptimistic
            .TextMatrix(i, ColAmountRem) = Format(RsInv.Fields("AmountRem") - RS.Fields("sumAm"), gs_formatAmount)
            .TextMatrix(i, ColPPnRem) = Format(RsInv.Fields("PPNREM") - RS.Fields("sumPPN"), gs_formatAmount)
            .TextMatrix(i, ColDueDate) = Format(RsInv.Fields("Due_date"), "dd MMM yyyy")
            .TextMatrix(i, ColBantuCurr) = RsInv.Fields("Curr")
            .TextMatrix(i, colpaid) = IIf(IsNull(RsInv.Fields("Paid")), "0", RsInv.Fields("Paid"))
            
            AloEdit (i)
            i = i + 1
            RsInv.MoveNext
        Wend
    End With
    If XCount = 1 Then
        Header2
        HitungTotalGrid
    End If
End If
End Sub

Sub AloEdit(X As Long)
        grid.Cell(flexcpBackColor, X, ColCls) = &HFFFFFF
        grid.Cell(flexcpBackColor, X, ColAmountPay) = &HFFFFFF
        grid.Cell(flexcpBackColor, X, ColPPnPay) = &HFFFFFF
End Sub

Sub ComboSuppUbah()
    Dim i As Long
    MousePointer = vbHourglass
    i = 0
    rsSup.Requery
    rsSup.Find "supp_code ='" & cboSupp & "'"
    If Not rsSup.EOF Then
        txt_name.Text = Trim(rsSup!supp_name)
        rsSup.Requery
    End If
    
    NoAp
    
    MousePointer = Default
End Sub

Sub NoAp()
    Dim i As Long
    Dim RsAp As New ADODB.Recordset
    sql = "select AP_no from ap_master where supplier_code = '" & cboSupp.Text & "' " & _
          " and  ap_date >= '" & Format(APDateSt.Value, "yyyy-mm-dd") & "' " & _
      " and ap_date <= '" & Format(APDateEnd.Value, "yyyy-mm-dd") & "' "
    RsAp.Open sql, Db, adOpenKeyset, adLockOptimistic
    If RsAp.RecordCount > 0 Then
        RsAp.MoveFirst
    End If
    cboapno.clear
    cboapno.columnCount = 1
    cboapno.ColumnWidths = "100pt"
    cboapno.ListWidth = 110
    cboapno.ListRows = 15
    While Not RsAp.EOF
        cboapno.AddItem
        cboapno.List(i, 0) = Trim(RsAp!ap_no)
        i = i + 1
        RsAp.MoveNext
    Wend
End Sub

Sub adtocombo()
    sql = "SELECT  rtrim(Trade_Master.trade_Code) supp_code, rtrim(Trade_Master.Trade_Name) supp_name, " & _
        "rtrim(Trade_Master.Address1) address, country_Cls From Trade_Master where trade_cls in ('2','3')"
    Set rsSup = New Recordset
    rsSup.Open sql, Db, adOpenKeyset, adLockOptimistic
    With cboSupp
        .clear
        .columnCount = 3
        .ColumnWidths = "80 pt;280 pt; 0 pt; 0 pt"
        .ListWidth = 360
        .ListRows = 15
        i = 0
        rsSup.Requery
        If Not rsSup.EOF And Not rsSup.BOF Then
            Do Until rsSup.EOF
                .AddItem ""
                .List(i, 0) = Trim(rsSup!supp_code)
                .List(i, 1) = Trim(rsSup!supp_name)
                .List(i, 2) = Trim(rsSup!Address) & " "
                .List(i, 3) = Trim(rsSup!country_cls)
                i = i + 1
                rsSup.MoveNext
            Loop
        End If
    End With
End Sub

Sub Init()
    CtrCreate = 0
    APDateSt.Value = Now
    APDateEnd.Value = Now
    combo1.clear
    combo1.AddItem "Create"
    combo1.AddItem "Update"
    combo1.ListIndex = 1
    cboapno.Text = ""
    cboapno.clear
    txt_name.Text = ""
    InitSave
End Sub

Sub InitSave()
    DateVou.Value = Now
End Sub

Sub Header()
  With grid
    .clear

    .Rows = 2
    .ColS = 13

    .ColWidth(ColCls) = 300
    .ColWidth(ColInvDate) = 1300
    .ColWidth(ColInvNo) = 1500
    .ColWidth(ColCurr) = 650
    .ColWidth(ColAmountInv) = 1400
    .ColWidth(ColPPnInv) = 1400
    .ColWidth(ColAmountPay) = 1400
    .ColWidth(ColPPnPay) = 1400
    .ColWidth(ColAmountRem) = 1400
    .ColWidth(ColPPnRem) = 1400
    .ColWidth(ColDueDate) = 1300
    .ColWidth(ColBantuCurr) = 0
    .ColWidth(colpaid) = 0

    .TextMatrix(0, ColCls) = " "
        .TextMatrix(1, ColCls) = " "
    .TextMatrix(0, ColInvDate) = "Invoice Date"
        .TextMatrix(1, ColInvDate) = "Invoice Date"
    .TextMatrix(0, ColInvNo) = "Invoice No"
        .TextMatrix(1, ColInvNo) = "Invoice No"
    .TextMatrix(0, ColCurr) = "Curr"
        .TextMatrix(1, ColCurr) = "Curr"
    .TextMatrix(0, ColAmountInv) = "Invoice"
    .TextMatrix(0, ColPPnInv) = "Invoice"
        .TextMatrix(1, ColAmountInv) = "Amount"
        .TextMatrix(1, ColPPnInv) = "PPN (IDR)"
    .TextMatrix(0, ColAmountPay) = "Payment"
    .TextMatrix(0, ColPPnPay) = "Payment"
        .TextMatrix(1, ColAmountPay) = "Amount"
        .TextMatrix(1, ColPPnPay) = "PPN (IDR)"
    .TextMatrix(0, ColAmountRem) = "Remaining"
    .TextMatrix(0, ColPPnRem) = "Remaining"
        .TextMatrix(1, ColAmountRem) = "Amount"
        .TextMatrix(1, ColPPnRem) = "PPN (IDR)"
    .TextMatrix(0, ColDueDate) = "Due Date"
        .TextMatrix(1, ColDueDate) = "Due Date"
    .TextMatrix(0, ColBantuCurr) = "Currbantu"
        .TextMatrix(1, ColBantuCurr) = "Currbantu"
    
    .Cell(flexcpAlignment, 0, 0, 1, 10) = flexAlignCenterCenter
    .ColAlignment(0) = flexAlignCenterCenter
    .ColAlignment(3) = flexAlignCenterCenter
    
    .RowHeight(0) = 250
   
    .MergeRow(0) = True
    .MergeRow(1) = True
    For i = 0 To 10
        .MergeCol(i) = True
    Next i
    .MergeCells = flexMergeFixedOnly
  End With
  
  Header2
End Sub

Sub Header2()
With Grid2
    .clear

    .Rows = 2
    .ColS = 5

    .ColWidth(ColCurrTT) = 1000
    .ColWidth(ColPay) = 2100
    .ColWidth(ColGrandTotalInv) = 2100
    .ColWidth(ColGrandTotalPay) = 2100
    .ColWidth(ColGrandToTalRem) = 2100

    .TextMatrix(0, ColCurrTT) = "Curr"
        .TextMatrix(1, ColCurrTT) = "Curr"
    .TextMatrix(0, ColPay) = "Total Payment"
        .TextMatrix(1, ColPay) = "Total Payment"
    .TextMatrix(0, ColGrandTotalInv) = "Grand Total Invoice"
        .TextMatrix(1, ColGrandTotalInv) = "Grand Total Invoice"
    .TextMatrix(0, ColGrandTotalPay) = "Grand Total Payment"
        .TextMatrix(1, ColGrandTotalPay) = "Grand Total Payment"
    .TextMatrix(0, ColGrandToTalRem) = "Grand Total Remaining"
        .TextMatrix(1, ColGrandToTalRem) = "Grand Total Remaining"
    
    .Cell(flexcpAlignment, 0, 0, 0, 4) = flexAlignCenterCenter

    .RowHeight(0) = 250
    
    .MergeRow(0) = True
    .MergeRow(1) = True
    For i = 0 To 4
        .MergeCol(i) = True
    Next i
    .MergeCells = flexMergeFixedOnly
  End With
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not IsNumeric(grid.TextMatrix(Row, ColAmountPay)) Then grid.TextMatrix(Row, ColAmountPay) = Format(0, gs_formatAmount)
If Not IsNumeric(grid.TextMatrix(Row, ColPPnPay)) Then grid.TextMatrix(Row, ColPPnPay) = Format(0, gs_formatAmount)
If Col = 0 Then
    grid.TextMatrix(Row, ColAmountPay) = Format(0, gs_formatAmount)
    grid.TextMatrix(Row, ColPPnPay) = Format(0, gs_formatAmount)
    GridOk Row
ElseIf Col = ColAmountPay Or Col = ColPPnPay Then
    GridOk Row
End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
lblerror = ""
If (Col <> ColCls And Col <> ColAmountPay) And Col <> ColPPnPay Then
    Cancel = True
Else
    If grid.TextMatrix(Row, colpaid) = "0" Then
        TampInv = grid.TextMatrix(Row, ColAmountPay)
        TampPPn = grid.TextMatrix(Row, ColPPnPay)
        If Col = ColAmountPay Or Col = ColPPnPay Then
            If grid.Cell(flexcpChecked, Row, 0) = flexUnchecked Then
                Cancel = True
                Exit Sub
            End If
        End If
    Else
        lblerror = DisplayMsg(4090)
        Cancel = True
    End If
End If
End Sub

Private Sub grid_Click()
    If grid.Col = ColCls Then
        grid.FocusRect = flexFocusInset
    ElseIf grid.Col = ColAmountPay Or grid.Col = ColPPnPay Then
        grid.FocusRect = flexFocusInset
    Else
        grid.FocusRect = flexFocusNone
    End If
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
lblerror = ""
  If grid.Col = ColAmountPay Or grid.Col = ColPPnPay Then
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
  End If
End Sub

Sub GridOk(lngRow)
    If Not IsNumeric(grid.TextMatrix(lngRow, ColAmountPay)) Then grid.TextMatrix(lngRow, ColAmountPay) = Format(0, gs_formatAmount)
    If Not IsNumeric(grid.TextMatrix(lngRow, ColPPnPay)) Then grid.TextMatrix(lngRow, ColPPnPay) = Format(0, gs_formatAmount)
    grid.TextMatrix(lngRow, ColAmountPay) = Format(grid.TextMatrix(lngRow, ColAmountPay), gs_formatAmount)
    grid.TextMatrix(lngRow, ColPPnPay) = Format(grid.TextMatrix(lngRow, ColPPnPay), gs_formatAmount)
    grid.TextMatrix(lngRow, ColAmountRem) = Format((CDbl(grid.TextMatrix(lngRow, ColAmountRem)) + TampInv) - CDbl(grid.TextMatrix(lngRow, ColAmountPay)), gs_formatAmount)
    grid.TextMatrix(lngRow, ColPPnRem) = Format((CDbl(grid.TextMatrix(lngRow, ColPPnRem)) + TampPPn) - CDbl(grid.TextMatrix(lngRow, ColPPnPay)), gs_formatAmount)
    If CDbl(grid.TextMatrix(lngRow, ColAmountRem)) < 0 Then grid.TextMatrix(lngRow, ColAmountRem) = "0"
    If CDbl(grid.TextMatrix(lngRow, ColPPnRem)) < 0 Then grid.TextMatrix(lngRow, ColPPnRem) = "0"
    HitungTotalGrid
End Sub

Sub HitungTotalGrid()
Dim CtrAdd As Integer, j As Integer
Dim RsGrand As New ADODB.Recordset

If grid.Rows > 2 Then
    Grid2.clear
    Header2
    
    If RsGrand.State = 1 Then RsGrand.Close
    sql = "select * from curr_cls"
    RsGrand.Open sql, Db, adOpenKeyset, adLockOptimistic
    While RsGrand.EOF = False
        Grid2.Rows = Grid2.Rows + 1
        Grid2.TextMatrix(Grid2.Rows - 1, ColCurrTT) = Trim(RsGrand!Description)
        Grid2.TextMatrix(Grid2.Rows - 1, ColPay) = Format(0, gs_formatAmount)
        Grid2.TextMatrix(Grid2.Rows - 1, ColGrandTotalInv) = Format(0, gs_formatAmount)
        Grid2.TextMatrix(Grid2.Rows - 1, ColGrandTotalPay) = Format(0, gs_formatAmount)
        Grid2.TextMatrix(Grid2.Rows - 1, ColGrandToTalRem) = Format(0, gs_formatAmount)
        Grid2.RowHeight(Grid2.Rows - 1) = 0
        RsGrand.MoveNext
    Wend
    If RsGrand.State = 1 Then RsGrand.Close
       
    For i = 2 To grid.Rows - 1
        CtrAdd = 1      'if 1 then add to grid karena lum ada
            For j = 2 To Grid2.Rows - 1
                Dim rowtentu As Integer
                If Trim(grid.TextMatrix(i, ColCurr)) = Trim(Grid2.TextMatrix(j, ColCurrTT)) Then
                    rowtentu = j
                    Grid2.RowHeight(j) = 250
                    CtrAdd = 0
                    GoTo Jumping
                End If
            Next
Jumping:
             If CtrAdd = 1 Then
                'add amount
                Grid2.TextMatrix(rowtentu, ColCurrTT) = grid.TextMatrix(i, ColCurr)
                Grid2.TextMatrix(rowtentu, ColGrandTotalInv) = Format(grid.TextMatrix(i, ColAmountInv), gs_formatAmount)
                Grid2.TextMatrix(rowtentu, ColPay) = Format(grid.TextMatrix(i, ColAmountPay), gs_formatAmount)
                If RsGrand.State = 1 Then RsGrand.Close
                sql = "select isnull(sum(amount),0) as GrandTotal from ap_detail where invoice_no = '" & grid.TextMatrix(i, ColInvNo) & "' and supplier_code = '" & cboSupp & "'"
                RsGrand.Open sql, Db, adOpenKeyset, adLockOptimistic
                Grid2.TextMatrix(rowtentu, ColGrandTotalPay) = Format(RsGrand.Fields("GrandTotal"), gs_formatAmount)
                Grid2.TextMatrix(rowtentu, ColGrandToTalRem) = Format(grid.TextMatrix(i, ColAmountRem), gs_formatAmount)
            Else
                'tambah saja
                Grid2.TextMatrix(rowtentu, ColGrandTotalInv) = Format(CDbl(Grid2.TextMatrix(rowtentu, ColGrandTotalInv)) + CDbl(grid.TextMatrix(i, ColAmountInv)), gs_formatAmount)
                Grid2.TextMatrix(rowtentu, ColPay) = Format(CDbl(Grid2.TextMatrix(rowtentu, ColPay)) + CDbl(grid.TextMatrix(i, ColAmountPay)), gs_formatAmount)
                If RsGrand.State = 1 Then RsGrand.Close
                sql = "select isnull(sum(amount),0) as GrandTotal from ap_detail where invoice_no = '" & grid.TextMatrix(i, ColInvNo) & "' and supplier_code = '" & cboSupp & "'"
                RsGrand.Open sql, Db, adOpenKeyset, adLockOptimistic
                Grid2.TextMatrix(rowtentu, ColGrandTotalPay) = Format(CDbl(Grid2.TextMatrix(rowtentu, ColGrandTotalPay)) + CDbl(RsGrand.Fields("GrandTotal")), gs_formatAmount)
                Grid2.TextMatrix(rowtentu, ColGrandToTalRem) = Format(CDbl(Grid2.TextMatrix(rowtentu, ColGrandToTalRem)) + CDbl(grid.TextMatrix(i, ColAmountRem)), gs_formatAmount)
            End If
            Grid2.Cell(flexcpAlignment, 0, 0, Grid2.Rows - 1, 0) = flexAlignCenterCenter
    Next
    
    CtrAdd = 1
    Dim PPNttl As Double
    
    For i = 2 To grid.Rows - 1
        PPNttl = 0
            If CtrAdd = 1 Then
                'add ppn
                CtrAdd = 0
                Grid2.Rows = Grid2.Rows + 1
                Grid2.TextMatrix(Grid2.Rows - 1, ColCurrTT) = "PPN (IDR)"
                Grid2.TextMatrix(Grid2.Rows - 1, ColPay) = Format(grid.TextMatrix(i, ColPPnPay), gs_formatAmount)
                Grid2.TextMatrix(Grid2.Rows - 1, ColGrandTotalInv) = Format(grid.TextMatrix(i, ColPPnInv), gs_formatAmount)
                
                If RsGrand.State = 1 Then RsGrand.Close
                sql = "select isnull(sum(ppn),0) as GrandTotal from ap_detail where invoice_no = '" & grid.TextMatrix(i, ColInvNo) & "' and supplier_code = '" & cboSupp & "'"
                If RsGrand.State = 1 Then RsGrand.Close
                RsGrand.Open sql, Db, adOpenKeyset, adLockOptimistic
                If Not RsGrand.EOF And Not RsGrand.BOF Then
                    PPNttl = CDbl(RsGrand.Fields("grandtotal"))
                Else
                    PPNttl = 0
                End If
                
                Grid2.TextMatrix(Grid2.Rows - 1, ColGrandTotalPay) = Format(PPNttl, gs_formatAmount)
                Grid2.TextMatrix(Grid2.Rows - 1, ColGrandToTalRem) = Format(grid.TextMatrix(i, ColPPnRem), gs_formatAmount)
            Else
                'Add PPn
                Grid2.TextMatrix(Grid2.Rows - 1, ColPay) = Format(CDbl(Grid2.TextMatrix(Grid2.Rows - 1, ColPay)) + CDbl(grid.TextMatrix(i, ColPPnPay)), gs_formatAmount)
                Grid2.TextMatrix(Grid2.Rows - 1, ColGrandTotalInv) = Format(CDbl(Grid2.TextMatrix(Grid2.Rows - 1, ColGrandTotalInv)) + CDbl(grid.TextMatrix(i, ColPPnInv)), gs_formatAmount)
                
                If RsGrand.State = 1 Then RsGrand.Close
                sql = "select isnull(sum(ppn),0) as GrandTotal from ap_detail where invoice_no = '" & grid.TextMatrix(i, ColInvNo) & "' and supplier_code = '" & cboSupp & "'"
                If RsGrand.State = 1 Then RsGrand.Close
                RsGrand.Open sql, Db, adOpenKeyset, adLockOptimistic
                If Not RsGrand.EOF And Not RsGrand.BOF Then
                    PPNttl = CDbl(RsGrand.Fields("grandtotal"))
                Else
                    PPNttl = 0
                End If
                
                Grid2.TextMatrix(Grid2.Rows - 1, ColGrandTotalPay) = Format(CDbl(Grid2.TextMatrix(Grid2.Rows - 1, ColGrandTotalPay)) + CDbl((PPNttl)), gs_formatAmount)
                Grid2.TextMatrix(Grid2.Rows - 1, ColGrandToTalRem) = Format(CDbl(Grid2.TextMatrix(Grid2.Rows - 1, ColGrandToTalRem)) + grid.TextMatrix(i, ColPPnRem), gs_formatAmount)
            End If
            Grid2.Cell(flexcpAlignment, 0, 0, Grid2.Rows - 1, 0) = flexAlignCenterCenter
    Next
    
End If
End Sub

Private Sub Grid2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Cancel = True
End Sub

Private Sub Grid2_Click()
grid.FocusRect = flexFocusNone
End Sub
