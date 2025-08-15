VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcostminutemaster 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Cost/Minute Master"
   ClientHeight    =   9735
   ClientLeft      =   1725
   ClientTop       =   855
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmcostminutemaster.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtlinename 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   350
      Left            =   3390
      MaxLength       =   25
      TabIndex        =   22
      Tag             =   "TFFT*/"
      Top             =   7965
      Width           =   3330
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1050
      Left            =   465
      TabIndex        =   16
      Tag             =   "TTTF*/"
      Top             =   600
      Width           =   9885
      Begin MSComCtl2.DTPicker mYear 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   1455
         TabIndex        =   17
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   1275
         _ExtentX        =   2249
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
         CustomFormat    =   "MMM yyyy"
         Format          =   141230083
         UpDown          =   -1  'True
         CurrentDate     =   37868
      End
      Begin MSForms.ComboBox cbofact 
         Height          =   345
         Left            =   1440
         TabIndex        =   21
         Tag             =   "TTFF*/"
         Top             =   165
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
      Begin VB.Label lblFact 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2865
         TabIndex        =   20
         Tag             =   "TTFF*/"
         Top             =   300
         Width           =   6015
      End
      Begin VB.Label LblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory Code"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Tag             =   "TTTF*/"
         Top             =   240
         Width           =   1140
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   2865
         X2              =   8865
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label LblTax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   390
      End
   End
   Begin VB.CommandButton Cmd_save 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Left            =   9255
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "FFTT*/"
      Top             =   9150
      Width           =   1125
   End
   Begin VB.TextBox txtCost 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   6825
      MaxLength       =   25
      TabIndex        =   4
      Tag             =   "TFFT*/"
      Top             =   7965
      Width           =   1890
   End
   Begin VB.CommandButton Cmd_SubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   495
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "TFFT*/"
      Top             =   9150
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Clear 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Left            =   8025
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "FFTT*/"
      Top             =   9150
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   525
      Left            =   495
      TabIndex        =   0
      Tag             =   "TFTT*/"
      Top             =   8475
      Width           =   9885
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
         Height          =   330
         Left            =   90
         TabIndex        =   1
         Tag             =   "TFTT*/"
         Top             =   180
         Width           =   9645
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5490
      Left            =   465
      TabIndex        =   6
      Tag             =   "TTTT*/"
      Top             =   1740
      Width           =   9885
      _cx             =   17436
      _cy             =   9684
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
      Left            =   8505
      TabIndex        =   7
      Tag             =   "FTTF*/"
      Top             =   135
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin MSComCtl2.DTPicker MDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd MMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   345
      Left            =   720
      TabIndex        =   15
      Tag             =   "TFFT*/"
      Top             =   7965
      Width           =   1245
      _ExtentX        =   2196
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
      CustomFormat    =   "MMM yyyy"
      Format          =   141230083
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   2
      Left            =   3540
      TabIndex        =   23
      Tag             =   "TFFT*/"
      Top             =   7590
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line Code"
      Height          =   195
      Index           =   5
      Left            =   1995
      TabIndex        =   14
      Tag             =   "TFFT*/"
      Top             =   7605
      Width           =   915
   End
   Begin MSForms.ComboBox CboLine 
      Height          =   345
      Left            =   2040
      TabIndex        =   13
      Tag             =   "TFFT*/"
      Top             =   7965
      Width           =   1350
      VariousPropertyBits=   746604571
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "2381;609"
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
      Caption         =   "Group Cls"
      Height          =   195
      Index           =   4
      Left            =   2025
      TabIndex        =   12
      Tag             =   "TFFT*/"
      Top             =   7605
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   495
      Tag             =   "TFTT*/"
      Top             =   7875
      Width           =   9885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cost/Minute Master"
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
      Left            =   480
      TabIndex        =   11
      Tag             =   "TTTT*/"
      Top             =   120
      Width           =   9885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month/Year"
      Height          =   195
      Index           =   1
      Left            =   765
      TabIndex        =   10
      Tag             =   "TFFT*/"
      Top             =   7590
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost/Minute (USD)"
      Height          =   195
      Index           =   3
      Left            =   6825
      TabIndex        =   9
      Tag             =   "TFFT*/"
      Top             =   7590
      Width           =   1665
   End
   Begin MSForms.ComboBox cboGroup 
      Height          =   345
      Left            =   2040
      TabIndex        =   8
      Tag             =   "TFFT*/"
      Top             =   7965
      Visible         =   0   'False
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   495
      Tag             =   "TFTT*/"
      Top             =   7515
      Width           =   9885
   End
End
Attribute VB_Name = "frmcostminutemaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim RSTR As ADODB.Recordset
Dim rswh As ADODB.Recordset
Dim baru As Boolean
Dim Pos As Integer, jml As Integer
Dim StaErr As Boolean, nErr As Integer, StrWDel As String, nOK As Integer
Dim HakU As Integer
Dim Updet, Delet As Integer, Start1 As String, Finis As String, NO As Integer
Dim Test As String, dtDate As Boolean
Dim LDay As String, ADay As String
Dim tgl_sb As String
Dim Tgl_sb2 As String
Dim StLDay As Boolean
Dim PrevWeekEmpty As Boolean
Dim TglMP As String
Dim gs_status As String

Dim bteColSelect As Byte
Dim bteColFactoryCode As Byte
Dim bteColDateYear As Byte
Dim bteColDate As Byte
Dim bteColGroupCls As Byte
Dim bteColGroup As Byte
Dim bteColLine As Byte
Dim bteColLineDesc As Byte
Dim bteColCostMinute As Byte
Dim btecolfactoryname As Byte

Sub Header()
    Dim C As Byte
    
    bteColSelect = 0
    bteColFactoryCode = 1
    bteColDateYear = 2
    bteColDate = 3
    bteColGroupCls = 4
    bteColGroup = 5
    bteColLine = 6
    bteColLineDesc = 7
    bteColCostMinute = 8
    btecolfactoryname = 9
    With grid
        .ColS = 10
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColFactoryCode) = "Factory Code"
        .TextMatrix(0, bteColDateYear) = "tahun"
        .TextMatrix(0, bteColDate) = "Period"
        .TextMatrix(0, bteColGroupCls) = "Group"
        .TextMatrix(0, bteColGroup) = "Group Description"
        .TextMatrix(0, bteColLine) = "Line"
        .TextMatrix(0, bteColLineDesc) = "Description"
        .TextMatrix(0, bteColCostMinute) = "Cost/Minute (USD"
        .TextMatrix(0, btecolfactoryname) = "Name Factory"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColFactoryCode) = 0
        .ColWidth(bteColDateYear) = 0
        .ColWidth(bteColDate) = 700
        .ColWidth(bteColGroupCls) = 0
        .ColWidth(bteColGroup) = 1700
        .ColWidth(bteColLine) = 800
        .ColWidth(bteColLineDesc) = 2000
        .ColWidth(bteColCostMinute) = 1700
        .ColWidth(btecolfactoryname) = 0

'        .ColAlignment(bteColDate) = flexAlignCenterCenter
'        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        
        .ColHidden(bteColGroupCls) = True
        .ColHidden(bteColGroup) = True
        
        .EditMaxLength = 1
    End With
End Sub


Private Sub CmdMenu_Click()
frmMainMenu.Show
Unload Me
DoEvents
End Sub

Private Sub Combobox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbofact_Change()
    If cbofact.MatchFound Then
        lblFact.Caption = Trim(cbofact.List(cbofact.ListIndex, 1))
        Call Browse
        Else
        lblFact = ""
    End If
End Sub

Private Sub cboGroup_Change()

Call up_FillCombo(CboLine, "manufacture_line", "line_code,line_name", "")
CboLine.ColumnWidths = "50;180"
CboLine.ListWidth = 230
CboLine.ListRows = 5
'call up_FillCombo(cbofact,"trade_master","trade_code"
'Call up_FillCombo(cboGroup, "Group_cls", "Group_Cls,", "")

End Sub
Private Sub up_fillcomboFactory()
Dim sql As String
Dim RS As Recordset
Dim i As Long

sql = " select * from trade_master where trade_Cls=1"
Set RS = New Recordset
Set RS = Db.Execute(sql)

cbofact.clear

If Not RS.EOF Then
    With cbofact
        .columnCount = 2
        i = 0
        Do Until RS.EOF
            .AddItem ""
            .List(i, 0) = RS!Trade_Code
            .List(i, 1) = RS!trade_name
           i = i + 1
            RS.MoveNext
        Loop
        .ColumnWidths = "50pt;180pt"
        .ListWidth = 230
        .ListRows = 5
    End With
End If
End Sub
Private Sub up_fillcomboGroupCls()
Dim sql As String
Dim RS As Recordset
Dim i As Integer

sql = " select * from Group_Cls"
Set RS = New Recordset
Set RS = Db.Execute(sql)

cboGroup.clear

If Not RS.EOF Then
    With cboGroup
        .columnCount = 2
        i = 0
        Do Until RS.EOF
            .AddItem ""
            .List(i, 0) = RS!group_cls
            .List(i, 1) = RS!Description
           i = i + 1
            RS.MoveNext
        Loop
        .ColumnWidths = "40;100"
        .ListWidth = 140
        .ListRows = 5
    End With
End If
End Sub
Private Sub cboLine_Change()
CboLine.Text = Trim(CboLine.Text)
If CboLine.MatchFound Then
    txtlinename.Text = CboLine.Column(1)
Else
    txtlinename.Text = ""
End If

End Sub
Sub DataGrid()
Dim kode As String, kode2 As String, kode3 As String, kode4 As String, Sta As String
Dim strSQL As String
Dim ix As Integer

On Error Resume Next
Dim PosS As Integer
PosS = grid.FindRow("S", 0, bteColSelect, False)
If PosS > 0 Then
    kode = Trim$(grid.TextMatrix(PosS, bteColDate))
            Dim UseEnd As String
            

         strSQL = " Update costminute_master set Line_Code='" & CboLine & "', " & vbCrLf & _
                  " Cost_Minute='" & txtCost & "', " & vbCrLf & _
                  " Currency_Code='02', " & vbCrLf & _
                  " Last_Update=GetDate() ," & vbCrLf & _
                  " Last_User='" & userLogin & "' ," & vbCrLf & _
                  " Register_Date=GetDate() " & vbCrLf & _
                  " From costminute_master " & vbCrLf & _
                  " where Factory_Code= '" & cbofact & "'" & vbCrLf & _
                  " And costMinute_Year='" & Year(mYear) & "'" & vbCrLf & _
                  " And costMinute_Month='" & Format(Month(MDate), "0#") & "'" & vbCrLf & _
                  " And Line_Code='" & Trim$(grid.TextMatrix(PosS, bteColLine)) & "'"

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
kode = Trim$(grid.TextMatrix(ix, bteColFactoryCode))
kode2 = Trim$(grid.TextMatrix(ix, bteColDateYear))
kode3 = Trim$(grid.TextMatrix(ix, bteColDate))
kode4 = Trim$(grid.TextMatrix(ix, bteColLine))
Sta = Trim$(grid.TextMatrix(ix, bteColSelect))
    If Sta = "D" Then
        strSQL = "delete from costminute_master  where Factory_Code ='" & kode & "' and costminute_year='" & kode2 & "'" & vbCrLf & _
                 " And Costminute_Month='" & kode3 & "' And Line_Code='" & kode4 & "'"
        '#Check Code in costminute_master
        Dim rs2 As New ADODB.Recordset
        If rs2.State = 1 Then rs2.Close
        rs2.CursorLocation = adUseClient
        rs2.Open uf_select, Db, adOpenKeyset, adLockOptimistic
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
kode2 = ""
kode3 = ""
kode4 = ""
Sta = ""
strSQL = ""
Browse
End Sub
Private Function uf_select()
Dim li As Integer
With grid
 uf_select = "select * from costminute_master where Factory_Code='" & grid.TextMatrix(li, bteColFactoryCode) & "' and costminute_year='" & grid.TextMatrix(li, bteColDateYear) & "' And Costminute_Month='" & grid.TextMatrix(li, bteColDate) & "' And Line_Code='" & grid.TextMatrix(li, bteColLine) & "'"
End With
End Function

Sub pakai1(stat As Boolean)
    cbofact.Enabled = stat
    mYear.Enabled = stat
    MDate.Enabled = stat
    cboGroup.Enabled = stat
    CboLine.Enabled = stat
    txtCost.Enabled = stat
End Sub

Private Sub cmd_clear_Click()
    Kosong
    Browse
    pakai1 True
    LblErrMsg = ""
    baru = True
    Dim IK As Long
    For IK = 1 To grid.Rows - 1
        grid.TextMatrix(IK, bteColSelect) = ""
    Next
    cbofact.SetFocus
End Sub

Private Sub Cmd_Save_Click()
 Dim strS As Integer, Jawab As Integer, CekS As Boolean, CekD As Boolean
    Dim strD As Integer, ie As Integer
    
    CekS = False
    CekD = False
    StaErr = False
    
    If HakU = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
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
                LblErrMsg = DisplayMsg(1101)
                baru = True
                pakai1 True
                Browse
                Kosong
                Dim IK As Long
                For IK = 1 To grid.Rows - 1
                    grid.TextMatrix(IK, bteColSelect) = ""
                Next
                cbofact.SetFocus
            Else
                LblErrMsg = DisplayMsg(1102)
            End If
        End If
    
        Dim PRec As Integer
        If CekD Then
            If Jawab = vbYes Then
                If StaErr = False Then
                    LblErrMsg = DisplayMsg(1201)
                    baru = True
                    cbofact.SetFocus
                Else
                    LblErrMsg = DisplayMsg(1000)
                    If Trim$(StrWDel) <> "" Then
                        For ie = 0 To nErr - 1
                            PRec = grid.FindRow(Trim$(Split(StrWDel, ",")(ie)), 0, 1, False)
                            grid.TextMatrix(PRec, bteColSelect) = "D"
                        Next ie
                    End If
                End If
            ElseIf Jawab = vbNo Then
                LblErrMsg = ""
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
        Dim SqlU As String, PosRec As Integer, PosRec2 As Integer, PosRec3 As Integer, PosRec4 As Integer
        If cek Then
        
            '#Check Code in costminute_master
            Dim rs2 As New ADODB.Recordset
            If rs2.State = 1 Then rs2.Close
            rs2.CursorLocation = adUseClient
            rs2.Open "select * from costminute_master where Factory_code='" & cbofact & "' and Costminute_year='" & Year(mYear) & "' and costminute_Month='" & Format(Month(MDate), "0#") & "' and Line_Code='" & CboLine & "'", Db, adOpenKeyset, adLockOptimistic
            If rs2.EOF = False Then
                LblErrMsg = DisplayMsg(8129)
                If rs2.State = 1 Then rs2.Close
                Me.MousePointer = vbDefault: Exit Sub
            End If
            If rs2.State = 1 Then rs2.Close
        

            SqlU = "insert into costminute_master (Factory_Code, costminute_year, costminute_month, Line_Code,currency_Code, cost_minute, Last_Update, Last_User,Register_Date) " & _
                " values ('" & Trim(cbofact) & "','" & Year(mYear) & "','" & Format(Month(MDate), "0#") & "','" & CboLine & "','02', '" & txtCost.Text & "',getdate(),'" & userLogin & "',getdate())"
            PosRec = grid.FindRow(Trim$(cbofact), 0, bteColDate, False)
            If PosRec < 0 Then
                Db.Execute SqlU
                LblErrMsg = DisplayMsg(1000)
                Kosong
                cbofact.SetFocus
            Else
                LblErrMsg = "Cost Minute Master " & DisplayMsg(8129)
                cbofact.SetFocus
            End If
        End If
        baru = True
        Browse
        SqlU = ""
    End If

End Sub

Private Sub Cmd_SubMenu_Click()
CmdMenu_Click
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
Sub pakai11(stat As Boolean)
    cbofact.Enabled = stat
    mYear.Enabled = stat
    MDate.Enabled = stat
    cboGroup.Enabled = stat
    CboLine.Enabled = stat
    txtCost.Enabled = stat
End Sub
Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
Dim RS As Recordset, ir As Integer

CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
HakU = hakUpdate(Me.Name)
StrWDel = ""
Header
Browse
baru = True
Call Kosong
mYear.Value = Format(Date, "MMM yyyy")
MDate.Value = Format(Date, "MMM yyyy")
Call up_fillcomboFactory
Call up_fillcomboGroupCls
Call cboGroup_Change

'With Anchor1
'    .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
'    .DoInit
'End With
'
End Sub

Sub Browse()

sql = " select CM.Factory_Code,TM.Trade_Name,CM.Costminute_Year,CM.Costminute_Month,--GC.Group_Cls,GC.Description," & vbCrLf & _
            "   ML.Line_Code,ML.Line_Name,CM.Cost_Minute " & vbCrLf & _
            " From costminute_master CM  " & vbCrLf & _
            " Left Join Manufacture_Line ML  ON CM.Line_Code=ML.Line_Code " & vbCrLf & _
            " --Left Join Group_Cls GC  ON ML.Group_Cls=GC.Group_Cls " & vbCrLf & _
            " Left Join Trade_Master TM ON CM.Factory_Code=TM.Trade_Code where factory_Code='" & Trim(cbofact.Text) & "' and costminute_year ='" & mYear.Year & "' and costminute_month ='" & Format(mYear.Month, "00") & "'"

Set rswh = New ADODB.Recordset
rswh.Open sql, Db, adOpenKeyset, adLockOptimistic
Dim RSA As Recordset
i = 0
Header

While Not rswh.EOF
        i = i + 1
        grid.AddItem ""
        grid.TextMatrix(i, bteColFactoryCode) = Trim$(rswh!Factory_Code)
        grid.TextMatrix(i, bteColDateYear) = Trim$(rswh!costminute_year)
        grid.TextMatrix(i, bteColDate) = Trim$(rswh!costMinute_Month)
        'Grid.TextMatrix(i, bteColGroupCls) = Trim$(rswh!group_cls)
        'Grid.TextMatrix(i, bteColGroup) = Trim$(rswh!Description)
        grid.TextMatrix(i, bteColLine) = Trim$(rswh!line_code)
        grid.TextMatrix(i, bteColLineDesc) = Trim$(rswh!line_name)
        grid.TextMatrix(i, bteColCostMinute) = Trim$(rswh!cost_minute)
        grid.TextMatrix(i, btecolfactoryname) = Trim$(rswh!trade_name)
        grid.Cell(flexcpBackColor, i, bteColFactoryCode, i, bteColCostMinute) = &HDFFFFF
        grid.Cell(flexcpBackColor, i, bteColSelect) = vbWhite

   rswh.MoveNext
Wend
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrGrid As String
Dim AdaS As Boolean, brs As Integer, id As Integer
StrGrid = grid.Text
AdaS = False
pakai1 False
brs = 0

If StrGrid = "S" Then
    For id = 1 To grid.Rows - 1
        If id <> Row Then grid.TextMatrix(id, bteColSelect) = ""
    Next id
    cbofact = grid.TextMatrix(grid.Row, bteColFactoryCode)
    mYear = grid.TextMatrix(grid.Row, bteColDate) & "-01-" & grid.TextMatrix(grid.Row, bteColDateYear)
    MDate.Value = grid.TextMatrix(grid.Row, bteColDate) & "-01-" & grid.TextMatrix(grid.Row, bteColDateYear)
    cboGroup = grid.TextMatrix(grid.Row, bteColGroupCls)
    CboLine = grid.TextMatrix(grid.Row, bteColLine)
    txtCost = grid.TextMatrix(grid.Row, bteColCostMinute)
    lblFact.Caption = grid.TextMatrix(grid.Row, btecolfactoryname)
    
    cbofact.Enabled = False
    mYear.Enabled = False
    MDate.Enabled = False
    cboGroup.Enabled = False
    CboLine.Enabled = False
'    lblFact.Enabled = False
    txtCost.Enabled = True
    txtCost.SetFocus
    
    baru = False
    LblErrMsg = ""
ElseIf StrGrid = "D" Then
    pakai1 True
    For id = 1 To grid.Rows - 1
        'Jika ada S maka , hapus yg S
        If grid.TextMatrix(id, bteColSelect) = "S" Then grid.TextMatrix(id, bteColSelect) = "": Exit For
    Next id
    baru = False
    LblErrMsg = ""
Else
    pakai1 True
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
End If
End Sub

Function cek() As Boolean
Dim tmpi As Byte, rsc As Recordset
cek = False
If cbofact.Enabled = False And cboGroup.Enabled = False And CboLine.Enabled = False Then

   If Trim(txtCost.Text) = "" Then
        LblErrMsg = DisplayMsg("8128")
        txtCost.SetFocus
    Else
        cek = True
    End If
Else
    If Not cbofact.MatchFound Then
        LblErrMsg = DisplayMsg("8123")
        cbofact.SetFocus
'    ElseIf Not cboGroup.MatchFound Then
'        LblErrMsg = DisplayMsg("8125")
''        cboGroup.SetFocus
    ElseIf Not CboLine.MatchFound Then
        LblErrMsg = DisplayMsg("8127")
        CboLine.SetFocus
    ElseIf Trim(txtCost.Text) = "" Then
        LblErrMsg = DisplayMsg("8128")
        txtCost.SetFocus
    Else
        cek = True
    End If
End If
End Function

Sub Kosong()
'cbofact = ""
'mYear.Value = Format(Date, "MMM yyyy")
'MDate.Value = Format(Date, "MMM yyyy")
'cboGroup = ""
CboLine = ""
txtlinename = ""
txtCost = ""
'lblFact = ""
End Sub

Private Sub TxtRate_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
End If
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub mYear_Change()
Call Browse
End Sub

Private Sub txtcost_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
End If
If KeyAscii = 13 Then SendKeys vbTab
End Sub

