VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPart_RecStatus 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Order Entry Status"
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   15060
   Icon            =   "FrmPart_RecStatus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleMode       =   0  'User
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsubmenu 
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "TFFT*/"
      Top             =   9840
      Width           =   1140
   End
   Begin VB.CommandButton cmdClear 
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
      Index           =   2
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "FFTT*/"
      Top             =   9840
      Width           =   1140
   End
   Begin VB.CommandButton cmdSubmit 
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "FFTT*/"
      Top             =   9840
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   240
      TabIndex        =   13
      Tag             =   "TFTT*/"
      Top             =   9120
      Width           =   14565
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
         TabIndex        =   14
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   14355
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1245
      Left            =   240
      TabIndex        =   2
      Tag             =   "TTTF*/"
      Top             =   1080
      Width           =   14580
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "TTFF*/"
         Text            =   "txt_name"
         Top             =   300
         Width           =   4650
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   675
         Width           =   1155
      End
      Begin MSComCtl2.DTPicker DTTo 
         Height          =   315
         Left            =   3615
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   705
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
         Format          =   68026371
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker DTFrom 
         Height          =   315
         Left            =   1635
         TabIndex        =   6
         Tag             =   "TTFF*/"
         Top             =   720
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
         Format          =   68026371
         CurrentDate     =   37810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
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
         Left            =   135
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   315
         Width           =   1215
      End
      Begin MSForms.ComboBox cboSupplier 
         Height          =   315
         Left            =   1635
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   255
         Width           =   1425
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2514;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         X1              =   3240
         X2              =   7890
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label lbldesc 
         BackStyle       =   0  'Transparent
         Caption         =   "lbldesc"
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
         Left            =   10560
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   840
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3210
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   735
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Date"
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
         Left            =   135
         TabIndex        =   7
         Tag             =   "TTFF*/"
         Top             =   765
         Width           =   1200
      End
   End
   Begin EZRunnerv3.Anchor Anchor1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   12960
      TabIndex        =   0
      Tag             =   "FTTF*/"
      Top             =   240
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin EZRunnerv3.Anchor Anchor2 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   6495
      Left            =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   2520
      Width           =   14595
      _cx             =   25744
      _cy             =   11456
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
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
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
      Editable        =   2
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part (Material) Receipt Status"
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
      Left            =   360
      TabIndex        =   1
      Tag             =   "TTTF*/"
      Top             =   480
      Width           =   14685
   End
End
Attribute VB_Name = "FrmPart_RecStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim dbTransfer As New ADODB.Connection
Dim HakU As Integer
Dim dealerCD As String, PONO As String, SJNo As String
Dim blnFix As Integer, thnFix As Integer
Dim statusKlik As Integer
Dim newCls As New clsMRP

Dim bteColSJNo As Byte
Dim bteColSuppCode As Byte
Dim bteColSuppName As Byte
Dim bteColRecDate As Byte
Dim bteColRecQty As Byte
Dim bteColBctype As Byte
Dim bteColBcNo As Byte
Dim bteColRecStatus As Byte
Dim bteColFix As Byte
Dim bteColFixCls As Byte

Dim bteHakPrice As Byte

Private Sub headerGrid()
    Dim i As Long
    
    bteColSJNo = 0
    bteColSuppCode = 1
    bteColSuppName = 2
    bteColRecDate = 3
    bteColRecQty = 4
    bteColBctype = 5
    bteColBcNo = 6
    bteColRecStatus = 7
    bteColFix = 8
    bteColFixCls = 9
    
    With grid
        .clear
        .ColS = 10
        .Rows = 1
        
        .TextMatrix(0, bteColSJNo) = "Surat Jalan No."
        .TextMatrix(0, bteColSuppCode) = "Supp. Code"
        .TextMatrix(0, bteColSuppName) = "Supp. Name"
        .TextMatrix(0, bteColRecDate) = "Receipt Date"
        .TextMatrix(0, bteColRecQty) = "Receipt Qty"
        .TextMatrix(0, bteColBctype) = "BC Type"
        .TextMatrix(0, bteColBcNo) = "BC No."
        .TextMatrix(0, bteColRecStatus) = "Receipt Status"
        .TextMatrix(0, bteColFix) = "Fix"
        
        .ColWidth(bteColSJNo) = 2100
        .ColWidth(bteColSuppCode) = 1300
        .ColWidth(bteColSuppName) = 3700
        .ColWidth(bteColRecDate) = 1300
        .ColWidth(bteColRecQty) = 1500
        .ColWidth(bteColBctype) = 1000
        .ColWidth(bteColBcNo) = 1000
        .ColWidth(bteColRecStatus) = 1500
        
        .ColWidth(bteColFix) = 800
        
        .ColHidden(bteColFixCls) = True
        
        .Cell(flexcpAlignment, 0, bteColSJNo, 0, bteColFix) = flexAlignCenterCenter
        
        .EditMaxLength = 1
    End With
End Sub

'******** Combo **********
Sub isiCboCust() 'Isi Combo Dealer CD dr Customer Master
Dim RsCust As New ADODB.Recordset 'Data Customer

With cboSupplier
    .clear
    .columnCount = 3
    .TextColumn = 1
    
    '******** Ambil dr Customer Master utk Combo Dealer CD
    sql = "select Trade_Code,Trade_Name,Trade_Cls from Trade_Master " & _
        "where (Trade_Cls = 2  or Trade_Cls = 3) order by Trade_Code"
    Set RsCust = Db.Execute(sql)
    
    .AddItem ""
    .List(0, 0) = strAll
    .List(0, 1) = strAll
    .List(0, 2) = strAll
    
    i = 1
    Do While Not (RsCust.EOF)
        .AddItem ""
        .List(i, 0) = Trim(RsCust(0))
        .List(i, 1) = Trim(RsCust(1))
        .List(i, 2) = Trim(RsCust(2))
        i = i + 1
        RsCust.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 350
    .ColumnWidths = "50 pt;300 pt; 0 pt"
    .ListIndex = 0
    txt_name.Text = strAll
    
    Set RsCust = Nothing
End With
End Sub

Private Sub CboSupplier_Change()
    cboSupplier = cboSupplier
    dealerCD = cboSupplier
    If cboSupplier.MatchFound Then
        txt_name = cboSupplier.Column(1)
        LblErrMsg = ""
    Else
        txt_name = ""
        LblErrMsg = DisplayMsg(4011)
    End If
    Call headerGrid
End Sub

Private Sub cmdClear_Click(Index As Integer)
    cboSupplier.ListIndex = 0
    DTFrom = Date - (Day(Date) - 1)
    DTTo = Date
    headerGrid
End Sub

Private Sub cmdSearch_Click(Index As Integer)
LblErrMsg = ""
    Call IsiGrid
End Sub

Private Sub CmdSubmit_Click(Index As Integer)
Dim tanya

    tanya = vbYes 'MsgBox("Do you really want to Process Surat Jalan Status?", vbQuestion & vbYesNo, "Confirmation")
        If tanya = vbYes Then
            Me.MousePointer = vbHourglass
            
            If cboSupplier.Text = "" Then
                LblErrMsg = DisplayMsg(1033)
                cboSupplier.SetFocus
            ElseIf cboSupplier <> dealerCD Then
                LblErrMsg = DisplayMsg(1034)
                cboSupplier.SetFocus
            Else
                cboSupplier = cboSupplier
                If cboSupplier.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4011)
                    cboSupplier.SetFocus
                Else
                    
                    LblErrMsg.Caption = ""

                    dbTransfer.ConnectionTimeout = 0
                    dbTransfer.CommandTimeout = 0
                    dbTransfer.Open Db.ConnectionString
                    dbTransfer.BeginTrans

                    With grid
                        For i = 1 To .Rows - 1

                            'Melakukan perubahan atau tidak
                            statusKlik = IIf(.Cell(flexcpChecked, i, bteColFix) = flexChecked, 1, 0)
                            If .TextMatrix(i, bteColFixCls) <> statusKlik Then
                                SJNo = .TextMatrix(i, bteColSJNo)
                                
'                                '**** Update Stock****
                                sql = "EXEC SP_PartReceiptStatus_Submit '" & DTFrom.Value & "', '" & DTTo.Value & "', '" & SJNo & "', '" & statusKlik & "', '" & userLogin & "' "
                                dbTransfer.Execute sql
                            End If
                        Next i
                    End With

                    dbTransfer.CommitTrans
                    dbTransfer.Close
                    LblErrMsg = DisplayMsg(1101)
                End If
            End If
            
            IsiGrid
            
            Me.MousePointer = vbDefault
        End If
End Sub

'******************
Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
HakU = hakUpdate(Me.Name)
bteHakPrice = hakPrice(Me.Name)
Call isiCboCust
DTFrom = Date - (Day(Date) - 1)
DTTo = Date
Call headerGrid

With Anchor1
      .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
      .DoInit
End With

End Sub

Function stInvoice(noDO As String) As Integer
Dim sql As String
Dim rsSt As New ADODB.Recordset
    sql = "select DO_No from Invoice_Detail where  DO_NO = '" & noDO & "' "
    Set rsSt = Db.Execute(sql)
    If rsSt.EOF Then
        stInvoice = 0
    Else
        stInvoice = 1
    End If
End Function

Sub IsiGrid()
Dim RsPartReceipt As New ADODB.Recordset

Call headerGrid
With grid

    sql = "EXEC dbo.sp_PartReceiptStatus_GridLoad '" & cboSupplier.Text & "',  '" & DTFrom.Value & "', '" & DTTo.Value & "' "

    Set RsPartReceipt = Db.Execute(sql)

    If Not (RsPartReceipt.EOF) Then
        i = 1
        Do While Not RsPartReceipt.EOF
            .Rows = .Rows + 1
            .TextMatrix(i, bteColSJNo) = Trim(RsPartReceipt("SuratJalan_No"))
            .TextMatrix(i, bteColSuppCode) = Trim(RsPartReceipt("Supplier_Code"))
            .TextMatrix(i, bteColSuppName) = Trim(RsPartReceipt("Supplier_Name"))
            .TextMatrix(i, bteColRecDate) = Format(Trim(RsPartReceipt("Receipt_Date")), "dd MMM yyyy")
            .TextMatrix(i, bteColRecQty) = Format(RsPartReceipt("Qty"), gs_formatQty)
            .TextMatrix(i, bteColRecStatus) = RsPartReceipt("Receipt_Status")
            .TextMatrix(i, bteColBctype) = RsPartReceipt("BC_Type")
            .TextMatrix(i, bteColBcNo) = RsPartReceipt("BC_No")
            .Cell(flexcpChecked, i, bteColFix) = IIf(RsPartReceipt("Receipt_StatusFix") = 1, flexChecked, flexUnchecked)
            .Cell(flexcpBackColor, i, bteColFix) = vbWhite
            .TextMatrix(i, bteColFixCls) = RsPartReceipt("Receipt_StatusFix")
            
            .Cell(flexcpAlignment, i, bteColSJNo) = flexAlignLeftCenter
            .Cell(flexcpAlignment, i, bteColBctype) = flexAlignLeftCenter
            .Cell(flexcpAlignment, i, bteColBcNo) = flexAlignLeftCenter
            .Cell(flexcpAlignment, i, bteColRecQty) = flexAlignRightCenter
            
            i = i + 1
            RsPartReceipt.MoveNext
        Loop
    Else
        LblErrMsg = DisplayMsg(4006)
    End If
    Set RsPartReceipt = Nothing
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim pesanDTFrom As String, pesanDTTo As String

With grid
    LblErrMsg = ""
    If Col < bteColFix Then
        Cancel = 1
    Else
        pesanDTFrom = up_ValidateDateRange(Format(.TextMatrix(Row, bteColRecDate), "yyyy-MM-dd"), True)
        pesanDTTo = up_ValidateDateRange(Format(.TextMatrix(Row, bteColRecDate), "yyyy-MM-dd"), True)
        If Trim(Get_Record("select invoice_no from invoice_master where invoice_no='" & Trim(.TextMatrix(Row, bteColSJNo)) & "'  and Fix_Cls=1 ")) <> "" And .Cell(flexChecked, Row, bteColFix) = 0 Then
            LblErrMsg = DisplayMsg(4110)
            Cancel = 1
'        ElseIf Trim(Get_Record("select top 1 DN_No from packing_master pm inner join Packing_Detail pd on pd.Packing_No=pm.Packing_No where pd.DN_No='" & Trim(.TextMatrix(Row, bteColSJNo)) & "'  and pm.Fix_Cls=1 ")) <> "" And .Cell(flexChecked, Row, bteColFix) = 0 Then
'            LblErrMsg = "[8120] You Cannot Modify this Status !Packing List has Already Fixed !"
'            Cancel = 1
        Else
            If pesanDTFrom <> "" Or pesanDTTo <> "" Then
                LblErrMsg = IIf(pesanDTFrom = "", pesanDTTo, pesanDTFrom)
                Cancel = 1
            End If
        End If
    End If
End With
End Sub

Private Sub dtFrom_Change()
    LblErrMsg = ""
    If Format(DTFrom, "yyyy-MM-dd") > Format(CDate(DTTo), "yyyy-MM-dd") Then LblErrMsg = DisplayMsg("4068"): Exit Sub
    Call headerGrid
End Sub

Private Sub DTFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub dtTo_Change()
    LblErrMsg = ""
    If Format(DTFrom, "yyyy-MM-dd") > Format(CDate(DTTo), "yyyy-MM-dd") Then LblErrMsg = DisplayMsg("4066"): Exit Sub
    Call headerGrid
End Sub

Private Sub DTTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
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
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Public Function GetSupplierCode(strTrade As String)

Dim RsSuppCode As New Recordset

    sql = " SELECT Subcon_WH_Code FROM dbo.Trade_Master WHERE Trade_Code = '" & strTrade & "'  "
            
    If RsSuppCode.State <> adStateClosed Then RsSuppCode.Close
    Set RsSuppCode = Db.Execute(sql)
    
    If Not RsSuppCode.EOF Then
        GetSupplierCode = Trim(RsSuppCode!adm_group)
    Else
        GetSupplierCode = ""
    End If
    
    RsSuppCode.Close
    Set RsSuppCode = Nothing
    
End Function


