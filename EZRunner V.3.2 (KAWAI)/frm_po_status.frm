VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_po_status 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Purchase Order Status"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frm_po_status.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   840
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   14700
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
         Left            =   11970
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1155
      End
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
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "txt_name"
         Top             =   300
         Width           =   3885
      End
      Begin MSComCtl2.DTPicker dodate2 
         Height          =   315
         Left            =   10260
         TabIndex        =   2
         Top             =   255
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
         Format          =   151388163
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker dodate1 
         Height          =   315
         Left            =   8280
         TabIndex        =   1
         Top             =   255
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         TabIndex        =   15
         Top             =   315
         Width           =   705
      End
      Begin MSForms.ComboBox cbodealer 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   255
         Width           =   1965
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3466;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         X1              =   3045
         X2              =   7020
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
         Left            =   7380
         TabIndex        =   14
         Top             =   585
         Visible         =   0   'False
         Width           =   2550
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
         Left            =   9855
         TabIndex        =   13
         Top             =   285
         Width           =   375
      End
      Begin VB.Label Label7 
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
         Left            =   7380
         TabIndex        =   12
         Top             =   315
         Width           =   720
      End
   End
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
      TabIndex        =   7
      Top             =   9750
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Left            =   12555
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9750
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   525
      Left            =   240
      TabIndex        =   8
      Top             =   9030
      Width           =   14700
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
         TabIndex        =   9
         Top             =   180
         Width           =   14415
      End
   End
   Begin VB.CommandButton Command1 
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
      Left            =   13785
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9750
      Width           =   1140
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   6390
      Left            =   225
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2595
      Width           =   14715
      _cx             =   25956
      _cy             =   11271
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
      ScrollTrack     =   -1  'True
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13080
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   660
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Status"
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
      Left            =   285
      TabIndex        =   16
      Top             =   675
      Width           =   14640
   End
End
Attribute VB_Name = "frm_po_status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS As New ADODB.Recordset
Dim sql As String
Dim ubah As Boolean
Dim f_out As Boolean

Dim bteSuppCode As Byte
Dim bteSuppName As Byte
Dim btePONo As Byte
Dim btePODate As Byte
Dim btePODelDate As Byte
Dim btePOCurr As Byte
Dim btePOAmount As Byte
Dim btePOPPn As Byte
Dim btePOTotal As Byte
Dim btePOFix As Byte
Dim bteStatus As Byte

Dim bteHakPrice As Byte

Sub Header()
    
    bteSuppCode = 0
    bteSuppName = 1
    btePONo = 2
    btePODate = 3
    btePODelDate = 4
    btePOCurr = 5
    btePOAmount = 6
    btePOPPn = 7
    btePOTotal = 8
    btePOFix = 9
    bteStatus = 10
    
    With grid
        .Rows = 1
        .ColS = 11
        
        .TextMatrix(0, bteSuppCode) = "Supplier Code"
        .TextMatrix(0, bteSuppName) = "Supplier Name"
        .TextMatrix(0, btePONo) = "PO No (Ref No.)"
        .TextMatrix(0, btePODate) = "PO Date"
        .TextMatrix(0, btePODelDate) = "Delivery Date"
        .TextMatrix(0, btePOCurr) = "Currency"
        .TextMatrix(0, btePOAmount) = "Amount"
        .TextMatrix(0, btePOPPn) = "PPn"
        .TextMatrix(0, btePOTotal) = "Total Amount"
        .TextMatrix(0, btePOFix) = "Fix"
        .TextMatrix(0, bteStatus) = "Status"
        
        .ColAlignment(bteSuppCode) = flexAlignLeftCenter
        .ColAlignment(bteSuppName) = flexAlignLeftCenter
        .ColAlignment(btePONo) = flexAlignLeftCenter
        .ColAlignment(btePODate) = flexAlignCenterCenter
        .ColAlignment(btePODelDate) = flexAlignCenterCenter
        .ColAlignment(btePOCurr) = flexAlignLeftCenter
        .ColAlignment(btePOAmount) = flexAlignRightCenter
        .ColAlignment(btePOPPn) = flexAlignRightCenter
        .ColAlignment(btePOTotal) = flexAlignRightCenter
        .ColAlignment(btePOFix) = flexAlignCenterCenter
        
        .ColDataType(btePODate) = flexDTDate
        .ColDataType(btePODelDate) = flexDTDate
        
        .Cell(flexcpAlignment, 0, 0, 0, btePOFix) = flexAlignCenterCenter
        
        .ColWidth(bteSuppCode) = 1000
        .ColWidth(bteSuppName) = 2500
        .ColWidth(btePONo) = 2500
        .ColWidth(btePODate) = 1300
        .ColWidth(btePODelDate) = 1300
        .ColWidth(btePOCurr) = 1000
        .ColWidth(btePOAmount) = 1900
        .ColWidth(btePOPPn) = 1500
        .ColWidth(btePOTotal) = 1900
        .ColWidth(btePOFix) = 400
        
        .ColHidden(btePOCurr) = (bteHakPrice = 0)
        .ColHidden(btePOAmount) = (bteHakPrice = 0)
        .ColHidden(btePOPPn) = (bteHakPrice = 0)
        .ColHidden(btePOTotal) = (bteHakPrice = 0)
        .ColHidden(bteStatus) = True
        .ColHidden(bteSuppCode) = True
        
        If cbodealer.Text = strAll Then
            .ColHidden(bteSuppName) = False
        Else
            .ColHidden(bteSuppName) = True
        End If
        
    End With
End Sub

Sub adtocombo()
Dim RsCust As New ADODB.Recordset
Dim sqlcust As String

sqlcust = "SELECT * from trade_master where trade_Cls='2' or trade_Cls='3'"
          
    Set RsCust = Db.Execute(sqlcust)
    
    With cbodealer
        .clear
        .columnCount = 2
        .ColumnWidths = "80 pt;320 pt"
        .ListWidth = 400
        .ListRows = 15
        
        .AddItem ""
        .List(0, 0) = strAll
        .List(0, 1) = strAll
        
    i = 1
    Do Until RsCust.EOF
        .AddItem ""
        .List(i, 0) = Trim(RsCust!Trade_Code)
        .List(i, 1) = Trim(RsCust!trade_name)
        i = i + 1
        RsCust.MoveNext
    Loop
    End With
    cbodealer.ListIndex = 0
End Sub

Sub Kosong()
    ubah = False
    cbodealer.ListIndex = 0
    txt_name = ""
    dodate1.Value = Format(Now, "dd MMM yyyy")
    dodate2.Value = Format(Now, "dd MMM yyyy")
    LblErrMsg.Caption = ""
    Header
End Sub

Sub Browse()
Dim i As Integer, j As Integer
    
    LblErrMsg.Caption = ""
        
    Header
    i = 1

    sql = "select distinct pm.*, cc.description curr_desc,tm.Trade_Name Supplier_Name " & _
            "from purchaseorder_master pm " & _
            "inner join purchaseorder_detail pd on pm.po_no = pd.po_no " & _
            "inner join curr_cls cc on pd.currency_code = cc.curr_cls " & _
            "inner join Trade_Master tm on pm.Supplier_code = tm.Trade_Code "
    
    If cbodealer.ListIndex > 0 Then
            sql = sql & " where pm.supplier_code='" & Trim(cbodealer) & "' "
    End If
    sql = sql & _
            "and pm.po_date >='" & Format(dodate1.Value, "yyyy-MM-dd") & "' " & _
            "and pm.po_date <='" & Format(dodate2.Value, "yyyy-MM-dd") & "' "

    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    With grid
    If Not (RS.BOF And RS.EOF) Then
        
      RS.MoveFirst
      Do While Not RS.EOF

        .Rows = .Rows + 1
        
        .TextMatrix(i, bteSuppCode) = Trim(RS!Supplier_Code)
        .TextMatrix(i, bteSuppName) = Trim(RS!Supplier_Name)
        .TextMatrix(i, btePONo) = Trim(RS!po_no)
        .TextMatrix(i, btePODate) = Format(RS!po_date, "dd MMM yyyy")
        .TextMatrix(i, btePODelDate) = Format(RS!delivery_Date, "dd MMM yyyy")
        .TextMatrix(i, btePOCurr) = Trim(RS!curr_desc & "")
        .TextMatrix(i, btePOAmount) = Format(Val(RS!Amount & ""), gs_formatAmount)
        .TextMatrix(i, btePOPPn) = Format(Val(RS!ppn & ""), gs_formatAmount)
        .TextMatrix(i, btePOTotal) = Format(Val(RS!total_amount & ""), gs_formatAmount)
        
        If RS("Fix_Cls") = 1 Then
            .Cell(flexcpChecked, i, btePOFix) = flexChecked
            .TextMatrix(i, bteStatus) = flexChecked
        Else
            .Cell(flexcpChecked, i, btePOFix) = flexUnchecked
            .TextMatrix(i, bteStatus) = flexUnchecked
        End If
        
        i = i + 1
        RS.MoveNext
      Loop
    End If
    .ColHidden(bteSuppCode) = True
    End With
    
    For j = 1 To grid.Rows - 1
      grid.Cell(flexcpBackColor, j, btePOFix) = &HFFFFFF
    Next j

End Sub

Private Sub cbodealer_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cmdSearch_Click(Index As Integer)
If cbodealer.Text <> "" Then
    Call cbodealer_Click
    Browse
Else
    LblErrMsg = DisplayMsg(1033) '"Please insert customer code !"
End If
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    bteHakPrice = hakPrice(Me.Name)
    Header
    adtocombo
    Kosong
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    f_out = False
End Sub

Sub clearGrid()
    grid.clear
    grid.Rows = 1
    Call Header
End Sub

Private Sub cbodealer_Click()
Dim j As Integer
Call clearGrid
  If cbodealer.ListIndex <> -1 Then
    txt_name = cbodealer.Column(1)
    j = 1
 Else
 j = 0
  For i = 0 To cbodealer.ListCount - 1
     If UCase(Trim(cbodealer.Text)) = UCase(Trim(cbodealer.List(i, 0))) Then
        cbodealer = cbodealer.List(i, 0)
        txt_name = cbodealer.List(i, 1): j = 1: Exit For
    End If
  Next

 End If
  If j = 0 Then
    LblErrMsg = DisplayMsg(4050) '"Supplier data not found !"
    txt_name = ""
    Else
    LblErrMsg = ""
 
 End If
End Sub

Private Sub cbodealer_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then cbodealer_Click
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then txt_name = ""
End Sub

Private Sub dodate1_Change()
   If CDate(dodate1) > CDate(dodate2) Then
      LblErrMsg.Caption = DisplayMsg(1021) '"PO Date must be lower than " & Format(dodate2, "dd MMM yyyy")
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
   
Call clearGrid
End Sub

Private Sub dodate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then dodate1_Change
End Sub

Private Sub dodate2_Change()
   If CDate(dodate2) < CDate(dodate1) Then
      LblErrMsg.Caption = DisplayMsg(1021) '"PO Date must be higher than " & Format(dodate1, "dd MMM yyyy")
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
   
Call clearGrid
End Sub

Private Sub dodate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then dodate2_Change
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    ubah = True
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Col <> btePOFix Then Cancel = True
 If grid.Col = btePOFix Then
    Dim rsPO As New Recordset
     
      sql = "select * from invoicesupplier_detail where po_no='" & grid.TextMatrix(Row, btePONo) & "'"
      If rsPO.State <> adStateClosed Then rsPO.Close
      rsPO.Open sql, Db, adOpenStatic, adLockOptimistic
      If Not rsPO.EOF Then
         LblErrMsg = DisplayMsg("0056")
         Cancel = True
      Else
         LblErrMsg = ""
         Cancel = False
      End If
      
      '#20070627 Herfin - Tidak bisa unfix PO jika receipt sudah ada
      If gb_AllowInputWithoutFix_PartReceiptSchedule = False Then
            sql = "select * from part_receipt where po_no='" & grid.TextMatrix(Row, btePONo) & "'"
            If rsPO.State <> adStateClosed Then rsPO.Close
            rsPO.Open sql, Db, adOpenStatic, adLockOptimistic
            If Not rsPO.EOF Then
               LblErrMsg = DisplayMsg("8101")
               Cancel = True
            Else
               LblErrMsg = ""
               Cancel = False
            End If
      End If
      
   End If
   
End Sub

Private Sub CmdSubMenu_Click()
f_out = True
    Call Command1_Click(1)
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub Command1_Click(Index As Integer)
Dim sqlGrid As String
Dim rsGrid As New ADODB.Recordset

Select Case Index

Case 1:
            If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Exit Sub
        
        'If ubah Then
            sqlGrid = "select * from purchaseorder_master"
            If rsGrid.State <> adStateClosed Then rsGrid.Close
            rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
            
            With grid
            
            If f_out = False Then
                If .Rows = 1 Then
                    LblErrMsg = DisplayMsg(5012) '"There is no data to submit ! "
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
             
                For i = 1 To .Rows - 1
                    If .Cell(flexcpChecked, i, btePOFix) <> .TextMatrix(i, bteStatus) Then
                        rsGrid.filter = " po_no='" & .TextMatrix(i, btePONo) & "' "
                        If .Cell(flexcpChecked, i, btePOFix) = flexChecked Then
                            rsGrid("Fix_Cls") = 1
                        Else
                            rsGrid("Fix_Cls") = 0
                        End If
                        rsGrid("Last_Update") = Now
                        rsGrid("Last_User") = userLogin
                        rsGrid.update
                    End If
                Next i
            End With
            
            LblErrMsg.Caption = DisplayMsg(1101) '"Update Data Success !"
            
            ubah = False

Case 2: Kosong
        cbodealer.SetFocus
        
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If RS.State <> adStateClosed Then RS.Close
End Sub

