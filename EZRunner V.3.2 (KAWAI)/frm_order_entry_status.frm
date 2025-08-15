VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_order_entry_status 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Order Entry Status"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frm_order_entry_status.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdExcel 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Excel"
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
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9720
      Width           =   1155
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13080
      TabIndex        =   17
      Top             =   540
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
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
      Left            =   13815
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9705
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   270
      TabIndex        =   7
      Top             =   8985
      Width           =   14685
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
         TabIndex        =   8
         Top             =   180
         Width           =   14355
      End
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
      Left            =   12585
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9705
      Width           =   1140
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
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9705
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1245
      Left            =   270
      TabIndex        =   9
      Top             =   1425
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
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   675
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
         Left            =   3015
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "txt_name"
         Top             =   300
         Width           =   4650
      End
      Begin MSComCtl2.DTPicker dodate2 
         Height          =   315
         Left            =   3375
         TabIndex        =   2
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
         Format          =   137101315
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker dodate1 
         Height          =   315
         Left            =   1395
         TabIndex        =   1
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
         Format          =   137101315
         CurrentDate     =   37810
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "SI/PO Date"
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
         TabIndex        =   15
         Top             =   765
         Width           =   720
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
         Left            =   2970
         TabIndex        =   12
         Top             =   735
         Width           =   375
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
         Left            =   7785
         TabIndex        =   11
         Top             =   900
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Line Line2 
         X1              =   3000
         X2              =   7650
         Y1              =   540
         Y2              =   540
      End
      Begin MSForms.ComboBox cbodealer 
         Height          =   315
         Left            =   1395
         TabIndex        =   0
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer CD"
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
         TabIndex        =   10
         Top             =   315
         Width           =   1170
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   6150
      Left            =   255
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2790
      Width           =   14715
      _cx             =   25956
      _cy             =   10848
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
      Caption         =   "Order Entry Status"
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
      Left            =   6495
      TabIndex        =   13
      Top             =   540
      Width           =   2145
   End
End
Attribute VB_Name = "frm_order_entry_status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS As New ADODB.Recordset
Dim sql As String
Dim ubah As Boolean
Dim f_out As Boolean

Dim bteColPONo As Byte
Dim bteColPODate As Byte
Dim bteColLocation As Byte
Dim bteColContact As Byte
Dim bteColTotalPO As Byte
Dim bteColFix As Byte

Dim bteHakPrice As Byte

Sub Header()
    With grid
        bteColPONo = 0
        bteColPODate = 1
        bteColLocation = 2
        bteColContact = 3
        bteColTotalPO = 4
        bteColFix = 5
        
        .Rows = 1
        .ColS = 6
        
        .TextMatrix(0, bteColPONo) = "SI/PO No (Ref No.)"
        .TextMatrix(0, bteColPODate) = "SI/PO Date"
        .TextMatrix(0, bteColLocation) = "Location Code"
        .TextMatrix(0, bteColContact) = "Contact Person"
        .TextMatrix(0, bteColTotalPO) = "Total SI/PO Amount"
        .TextMatrix(0, bteColFix) = "Complete"
        
        .ColAlignment(bteColPONo) = flexAlignLeftCenter
        .ColAlignment(bteColPODate) = flexAlignLeftCenter
        .ColAlignment(bteColLocation) = flexAlignLeftCenter
        .ColAlignment(bteColContact) = flexAlignLeftCenter
        .ColAlignment(bteColTotalPO) = flexAlignRightCenter
        .ColAlignment(bteColFix) = flexAlignCenterCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
        
        .ColWidth(bteColPONo) = 2150
        .ColWidth(bteColPODate) = 1700
        .ColWidth(bteColLocation) = 1700
        .ColWidth(bteColContact) = 2200
        .ColWidth(bteColTotalPO) = 1900
        .ColWidth(bteColFix) = 1200
        
        .ColHidden(bteColTotalPO) = (bteHakPrice = 0)
        
    End With
End Sub

Sub adtocombo()
Dim RsCust As New ADODB.Recordset
Dim sqlcust As String

    sqlcust = "select trade_code, trade_name, address1 from trade_master " & _
                vbLf & " where (trade_cls = '2' or trade_cls='3') "
    '---
    
    Set RsCust = Db.Execute(sqlcust)
    
    With cbodealer
        .clear
        .columnCount = 2
        .ColumnWidths = "50 pt;350 pt"
        .ListWidth = 400
        .ListRows = 15
        
    i = 0
    Do Until RsCust.EOF
        .AddItem ""
        .List(i, 0) = Trim(RsCust!Trade_Code)
        .List(i, 1) = Trim(RsCust!trade_name)
        i = i + 1
        RsCust.MoveNext
    Loop
    End With
End Sub

Sub clearGrid()
grid.clear
f_out = True
Call Header

End Sub

Sub Kosong()
    ubah = False
    cbodealer.Text = ""
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
    
    sql = "select orderentry_master.cust_code, " & _
            "orderentry_master.po_no,orderentry_master.po_date,orderentry_master.location_code, " & _
            "orderentry_master.contact_person,orderentry_master.fix_cls, orderentry_detail.calculate_cls, orderentry_detail.generate_cls, sum(orderentry_detail.amount)as amount " & _
            "from orderentry_Master join orderentry_detail on orderentry_detail.po_no = " & _
            "orderentry_master.po_no " & _
            "where orderentry_master.cust_code='" & Trim(cbodealer) & "' and orderentry_master.po_date >='" & Format(dodate1.Value, "yyyy-MM-dd") & "' and  " & _
            "orderentry_master.po_date <='" & Format(dodate2.Value, "yyyy-MM-dd") & "' " & _
            "group by orderentry_master.cust_code, " & _
            "orderentry_master.po_no,orderentry_master.po_date,orderentry_master.location_code, " & _
            "orderentry_master.contact_person,orderentry_master.fix_cls, orderentry_detail.calculate_cls, orderentry_detail.generate_cls "
          
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    With grid
    If Not (RS.BOF And RS.EOF) Then
        f_out = False
      RS.MoveFirst
      Do While Not RS.EOF
        .Rows = .Rows + 1
        .TextMatrix(i, bteColPONo) = Trim(RS!po_no)
        .TextMatrix(i, bteColPODate) = Format(RS!po_date, "dd MMM yyyy")
        .TextMatrix(i, bteColLocation) = IIf(IsNull(RS!location_code), "", Trim(RS!location_code))
        .TextMatrix(i, bteColContact) = Trim(RS!contact_person)
        .TextMatrix(i, bteColTotalPO) = IIf(IsNull(RS!Amount), 0, Format(Trim(RS!Amount), gs_formatAmountIDR))
        
        If RS("Fix_Cls") = 1 Then
          .Cell(flexcpChecked, i, bteColFix) = flexChecked
        Else
          .Cell(flexcpChecked, i, bteColFix) = flexUnchecked
        End If
                
        i = i + 1
        RS.MoveNext
      Loop
    Else
        f_out = True
    End If
    End With
    
    For j = 1 To grid.Rows - 1
      grid.Cell(flexcpBackColor, j, bteColFix) = &HFFFFFF
    Next j

End Sub

Private Sub cbodealer_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub CmdExcel_Click()
Dim xlapp As New Excel.application


Dim i As Long
Dim j As Long
Dim xstr As String

If grid.Rows > 1 Then

        LblErrMsg = ""
        Me.MousePointer = vbHourglass

        Dim Idx As Integer

        With xlapp
            .Workbooks.Add
            
            .Range("A1:F2").Merge
            .Range("A1", "F2").Columns.Font.Name = "Arial"
            .Range("A1", "F2").Columns.Font.Size = "16"
            .Range("A1") = "Order Entry Status"
            .Range("A1").horizontalAlignment = xlCenter
            .Range("A1").verticalAlignment = xlCenter
            .Range("A1").Font.Bold = True
                       
            
            .Range("A4") = "Customer :"
            .Range("B4") = " " & Trim(cbodealer)
            .Range("A4").horizontalAlignment = xlLeft
            .Range("A4").Font.Bold = True
            
            .Range("C4") = " " & Trim(txt_name.Text)
            
            .Range("A5") = "SI / PO :"
            .Range("B5") = " " & Format(dodate1, "dd-MMM-yyyy")
            .Range("A5", "B6").horizontalAlignment = xlLeft
            .Range("A5").Font.Bold = True
            
            .Range("D5") = "To :"
            .Range("E5") = " " & Format(dodate2, "dd-MMM-yyyy")
            .Range("D5", "F6").horizontalAlignment = xlLeft
            .Range("D5").Font.Bold = True
            
            
            'Header
                
            .Range("A8") = "SI / PO No (Ref. No)"
            .Range("B8") = "SI / PO Date"
            .Range("C8") = "Location Code"
            .Range("D8") = "Contact Person"
            .Range("E8") = "Total SI / PO Amount"
            .Range("F8") = "Complete"
            
            
                j = 9
                
                For i = 1 To grid.Rows - 1
                    j = j + 1
                    LblErrMsg = " Transfering data ... (record " & i & ")"
                    DoEvents
                    If Trim(grid.TextMatrix(i, bteColPONo)) <> "" Then
                       
                        .Range("A" & Trim(Str(j))) = Trim(grid.TextMatrix(i, bteColPONo))
                        .Range("B" & Trim(Str(j))) = Trim(grid.TextMatrix(i, bteColPODate))
                        .Range("C" & Trim(Str(j))) = Trim(grid.TextMatrix(i, bteColLocation))
                        .Range("D" & Trim(Str(j))) = Trim(grid.TextMatrix(i, bteColContact))
                        .Range("E" & Trim(Str(j))) = Trim(grid.TextMatrix(i, bteColTotalPO))
                     
                    End If
                    If grid.Cell(flexcpChecked, i, bteColFix) = flexChecked Then
                        .Range("F" & j) = "Yes"
                    Else
                        .Range("F" & j) = "No"
                    End If
'
                Next i
                
                 LblErrMsg = " Transfering data complete. "
                .Visible = True
                .Columns("A:F").Columns.AutoFit
                .WindowState = xlMaximized
                .ActiveWindow.Zoom = 80
                
                .Range("A8", "F8").Columns.Font.Bold = True
                .Range("B8", "F8").horizontalAlignment = xlCenter
                .Range("A8").horizontalAlignment = xlLeft
            
                End With
    
            Me.MousePointer = vbDefault
            
  Else
  
      LblErrMsg = "No Data to display."
      
  End If
     
End Sub

Private Sub cmdSearch_Click(Index As Integer)
If cbodealer.Text <> "" Then
    Call cbodealer_Click
    If LblErrMsg <> "" Then Exit Sub
    Browse
    LblErrMsg = ""
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
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    bteHakPrice = hakPrice(Me.Name)
    Header
    adtocombo
    Kosong
    f_out = True
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
    LblErrMsg = DisplayMsg(4011) '"Customer data not found !"
    txt_name = ""
  Else
    LblErrMsg = ""
 End If
End Sub

Private Sub cbodealer_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Call clearGrid
  If KeyCode = 13 Then cbodealer_Click
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then txt_name = ""
End Sub

Private Sub dodate1_Change()
Call clearGrid
   If CDate(dodate1) > CDate(dodate2) Then
      LblErrMsg.Caption = DisplayMsg(1021) '"PO Date must be lower than " & Format(dodate2, "dd MMM yyyy")
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If

End Sub

Private Sub dodate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then dodate1_Change
End Sub

Private Sub dodate2_Change()
Call clearGrid
   If CDate(dodate2) < CDate(dodate1) Then
      LblErrMsg.Caption = DisplayMsg(1021) '"PO Date must be higher than " & Format(dodate1, "dd MMM yyyy")
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
End Sub

Private Sub dodate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then dodate2_Change
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    ubah = True
End Sub

Private Sub grid_Click()
With grid
    If .Row = 1 And .Col <> bteColFix Then
      If .Col = bteColPODate Or .Col = bteColLocation Or .Col = bteColContact Then
        
          If .ColSort(.Col) = flexSortStringAscending Then
             .ColSort(.Col) = flexSortStringDescending
          Else
             .ColSort(.Col) = flexSortStringAscending
          End If
     Else
        
          If .ColSort(.Col) = flexSortNumericAscending Then
             .ColSort(.Col) = flexSortNumericDescending
          Else
             .ColSort(.Col) = flexSortNumericAscending
          End If
      
      End If
       .Sort = .ColSort(.Col)
    End If
End With
LblErrMsg = ""
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Col < bteColFix Then Cancel = True
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
            
            sqlGrid = "select * from orderentry_master"
            If rsGrid.State <> adStateClosed Then rsGrid.Close
            rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
            
            With grid
            
            If f_out = False Then
            Else
                LblErrMsg = DisplayMsg(5012) '"There is no data to submit ! "
                Exit Sub
            End If
             
                For i = 1 To .Rows - 1
                    rsGrid.filter = " po_no='" & .TextMatrix(i, bteColPONo) & "' "
                    If .Cell(flexcpChecked, i, bteColFix) = flexChecked Then
                        rsGrid("Fix_Cls") = 1
                    Else
                        rsGrid("Fix_Cls") = 0
                    End If
                    rsGrid("Last_Update") = Now
                    rsGrid("Last_User") = userLogin
                    rsGrid.update
                Next i
            End With
            
            LblErrMsg.Caption = DisplayMsg(1101) '"Update Data Success !"
            ubah = False
Case 2:
    Kosong
    cbodealer.SetFocus
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If RS.State <> adStateClosed Then RS.Close
End Sub




