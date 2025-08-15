VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpajak_status 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Faktur Pajak Update Status"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frmpajak_status.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13058
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   540
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
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
      Left            =   13793
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9855
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   615
      Left            =   308
      TabIndex        =   7
      Top             =   9135
      Width           =   14625
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
         Top             =   240
         Width           =   14385
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
      Left            =   12563
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9855
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
      Left            =   308
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9855
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   840
      Left            =   338
      TabIndex        =   9
      Top             =   1410
      Width           =   14580
      Begin VB.TextBox lbldesc 
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   2985
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   315
         Left            =   8640
         TabIndex        =   2
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
         Format          =   150208515
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   315
         Left            =   6660
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
         Format          =   150208515
         CurrentDate     =   37810
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
         Left            =   8235
         TabIndex        =   12
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   6060
         TabIndex        =   11
         Top             =   270
         Width           =   720
      End
      Begin VB.Line Line2 
         X1              =   3000
         X2              =   6000
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
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6585
      Left            =   308
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2400
      Width           =   14595
      _cx             =   25744
      _cy             =   11615
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
      BackColorFixed  =   12640511
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Faktur Pajak Update Status"
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
      Left            =   338
      TabIndex        =   13
      Top             =   540
      Width           =   14565
   End
End
Attribute VB_Name = "frmpajak_status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS As New ADODB.Recordset
Dim sql As String
Dim ubah As Boolean
Dim f_out As Boolean

Dim bteColFPNo As Byte
Dim bteColFPDate As Byte
Dim bteColAmount As Byte
Dim bteColPPn As Byte
Dim bteColTotal As Byte
Dim bteColFix As Byte

Dim bteHakPrice As Byte

Sub Header()
    bteColFPNo = 0
    bteColFPDate = 1
    bteColAmount = 2
    bteColPPn = 3
    bteColTotal = 4
    bteColFix = 5
    
    With grid
        .Rows = 1
        .ColS = 6
        .TextMatrix(0, bteColFPNo) = "Faktur Pajak No"
        .TextMatrix(0, bteColFPDate) = "Faktur Date"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColPPn) = "PPN"
        .TextMatrix(0, bteColTotal) = "Total Amount"
        .TextMatrix(0, bteColFix) = "Fix"
        
        .ColAlignment(bteColFPNo) = flexAlignLeftCenter
        .ColAlignment(bteColFPDate) = flexAlignLeftCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
        .ColAlignment(bteColPPn) = flexAlignRightCenter
        .ColAlignment(bteColTotal) = flexAlignRightCenter
        .ColAlignment(bteColFix) = flexAlignCenterCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, bteColFix) = flexAlignCenterCenter
        
        .ColWidth(bteColFPNo) = 2150
        .ColWidth(bteColFPDate) = 1700
        .ColWidth(bteColAmount) = 1700
        .ColWidth(bteColPPn) = 2200
        .ColWidth(bteColTotal) = 1900
        .ColWidth(bteColFix) = 400
        
        .ColHidden(bteColAmount) = (bteHakPrice = 0)
        .ColHidden(bteColPPn) = (bteHakPrice = 0)
        .ColHidden(bteColTotal) = (bteHakPrice = 0)
    End With
End Sub

Sub adtocombo()
Dim RsCust As New ADODB.Recordset
Dim sqlcust As String

sqlcust = "Select rtrim(trade_code) as TC,Trade_name as TN, Address1 as A from trade_master where trade_cls='4' order by trade_code"
          
    Set RsCust = Db.Execute(sqlcust)
    
    With cbodealer
        .clear
        .columnCount = 2
        .ColumnWidths = "50 pt;300 pt"
        .ListWidth = 350
        .ListRows = 15
        
    i = 0
    Do Until RsCust.EOF
        .AddItem ""
        .List(i, 0) = Trim(RsCust!TC)
        .List(i, 1) = Trim(RsCust!TN)
        i = i + 1
        RsCust.MoveNext
    Loop
    End With
End Sub

Sub Kosong()
    ubah = False
    cbodealer.Text = ""
    LblDesc = ""
    Tgl1.Value = Format(Now, "dd MMM yyyy")
    Tgl2.Value = Format(Now, "dd MMM yyyy")
    LblErrMsg.Caption = ""
    Header
End Sub

Sub Browse()
Dim i As Integer, j As Integer
    
    LblErrMsg.Caption = ""
        
    Header
    i = 1
    
    sql = "select * from fakturpajak_master where cust_code='" & cbodealer & "' and (fakturpajak_date >='" & Tgl1 & "' and fakturpajak_date <='" & Tgl2 & "')"
    sql = sql & " order by fakturpajak_no"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    With grid
    If Not (RS.BOF And RS.EOF) Then
        
      RS.MoveFirst
      Do While Not RS.EOF
        .Rows = .Rows + 1

        .TextMatrix(i, bteColFPNo) = Trim(RS!fakturpajak_no)
        .TextMatrix(i, bteColFPDate) = Format(RS!fakturpajak_date, "dd MMM yyyy")
        If InStr(1, Trim(RS!Amount), ".") Then
            .TextMatrix(i, bteColAmount) = Format(Trim(RS!Amount), gs_formatAmount)
        Else
            .TextMatrix(i, bteColAmount) = Format(Trim(RS!Amount), gs_formatAmount)
        End If
        If InStr(1, Trim(RS!ppn), ".") Then
            .TextMatrix(i, bteColPPn) = Format(Trim(RS!ppn), gs_formatAmount)
        Else
            .TextMatrix(i, bteColPPn) = Format(Trim(RS!ppn), gs_formatAmount)
        End If
        If InStr(1, Trim(RS!total_amount), ".") Then
            .TextMatrix(i, bteColTotal) = Format(Trim(RS!total_amount), gs_formatAmount)
        Else
            .TextMatrix(i, bteColTotal) = Format(Trim(RS!total_amount), gs_formatAmount)
        End If
        
        If RS("Fix_Cls") = 1 Then
          .Cell(flexcpChecked, i, bteColFix) = flexChecked
        Else
          .Cell(flexcpChecked, i, bteColFix) = flexUnchecked
        End If
        
        i = i + 1
        RS.MoveNext
      Loop
    End If
    End With
    
    For j = 1 To grid.Rows - 1
      grid.Cell(flexcpBackColor, j, bteColFix) = &HFFFFFF
    Next j

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
    bteHakPrice = (hakPrice(Me.Name))
    Header
    adtocombo
    Kosong
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    f_out = False
End Sub

Private Sub cbodealer_Click()
Dim j As Integer
  If cbodealer.ListIndex <> -1 Then
    LblDesc = cbodealer.Column(1)
    j = 1
 Else
 j = 0
  For i = 0 To cbodealer.ListCount - 1
     If UCase(Trim(cbodealer.Text)) = UCase(Trim(cbodealer.List(i, 0))) Then
        cbodealer = cbodealer.List(i, 0)
        LblDesc = cbodealer.List(i, 1): j = 1: Exit For
    End If
  Next

 End If
  Browse
  If j = 0 Then
    LblErrMsg = DisplayMsg(4072)
    LblDesc = ""
    Else
    LblErrMsg = ""
 
 End If
End Sub

Private Sub cbodealer_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then cbodealer_Click
End Sub

Private Sub tgl1_Change()
   If CDate(Tgl1) > CDate(Tgl2) Then
      LblErrMsg.Caption = DisplayMsg(4068)
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
   
   If cbodealer.Text <> "" Then Browse
End Sub

Private Sub tgl1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then tgl1_Change
End Sub

Private Sub Tgl2_Change()
   If CDate(Tgl2) < CDate(Tgl1) Then
      LblErrMsg.Caption = DisplayMsg(4066)
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
   
   If cbodealer.Text <> "" Then Browse
End Sub

Private Sub Tgl2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then Tgl2_Change
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    ubah = True
End Sub

Private Sub grid_Click()
With grid
    If .Row = 1 And .Col <> bteColFix Then
      If .Col = bteColFPDate Or .Col = bteColAmount Or .Col = bteColPPn Then
        
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
            sqlGrid = "select * from fakturpajak_master"
            If rsGrid.State <> adStateClosed Then rsGrid.Close
            rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
            
            With grid
            
            If f_out = False Then
                If .Rows = 1 Then LblErrMsg = DisplayMsg(5012): Exit Sub
            Else
                Exit Sub
            End If
                For i = 1 To .Rows - 1
                    rsGrid.filter = " fakturpajak_no='" & .TextMatrix(i, bteColFPNo) & "' "
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
            LblErrMsg.Caption = DisplayMsg(1101)
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
