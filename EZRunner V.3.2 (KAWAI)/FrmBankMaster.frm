VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmBankMaster 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Bank Master"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBankMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
      Height          =   375
      Index           =   4
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7200
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Index           =   2
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7200
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   8910
      MaxLength       =   30
      TabIndex        =   6
      Top             =   5910
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   4620
      MaxLength       =   70
      TabIndex        =   5
      Top             =   5910
      Width           =   4245
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1770
      MaxLength       =   35
      TabIndex        =   4
      Top             =   5910
      Width           =   2805
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   750
      MaxLength       =   5
      TabIndex        =   3
      Top             =   5910
      Width           =   945
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Index           =   0
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7200
      Width           =   1140
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   1
      Left            =   10350
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7200
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   450
      TabIndex        =   16
      Top             =   6420
      Width           =   11010
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
         Top             =   180
         Width           =   10725
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1155
      Left            =   450
      TabIndex        =   11
      Top             =   930
      Width           =   10965
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
         Height          =   375
         Index           =   3
         Left            =   3270
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   660
         Width           =   1140
      End
      Begin MSForms.ComboBox cbotrade 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   270
         Width           =   1185
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2090;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbocurr 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   690
         Width           =   1455
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2566;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   14
         Top             =   315
         Width           =   960
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2760
         X2              =   6480
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   13
         Top             =   750
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Code"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   12
         Top             =   345
         Width           =   1005
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   3255
      Left            =   450
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2130
      Width           =   10965
      _cx             =   19341
      _cy             =   5741
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   9570
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   150
      Width           =   1875
      _extentx        =   3307
      _extenty        =   714
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account No"
      Height          =   195
      Index           =   3
      Left            =   8940
      TabIndex        =   22
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Index           =   2
      Left            =   4650
      TabIndex        =   21
      Top             =   5520
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name "
      Height          =   195
      Index           =   1
      Left            =   1770
      TabIndex        =   19
      Top             =   5520
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Code"
      Height          =   195
      Index           =   0
      Left            =   750
      TabIndex        =   18
      Top             =   5520
      Width           =   945
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      Height          =   855
      Index           =   0
      Left            =   450
      Top             =   5460
      Width           =   10980
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Master"
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
      Left            =   4755
      TabIndex        =   15
      Top             =   300
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   330
      Index           =   1
      Left            =   450
      Top             =   5460
      Width           =   10980
   End
End
Attribute VB_Name = "FrmBankMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Idx As Long, HakU As Byte
Dim stopmsg As Boolean, X As Long, temp(10000) As String

Dim bteColSelect As Byte
Dim bteColBankCode As Byte
Dim bteColBankName As Byte
Dim bteColAddress As Byte
Dim bteColAccountNo As Byte

Sub Header()
    
    bteColSelect = 0
    bteColBankCode = 1
    bteColBankName = 2
    bteColAddress = 3
    bteColAccountNo = 4
    
    With grid
        .ColS = 5
        .Rows = 1
        .EditMaxLength = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColBankCode) = "Bank Code"
        .TextMatrix(0, bteColBankName) = "Bank Name"
        .TextMatrix(0, bteColAddress) = "Address"
        .TextMatrix(0, bteColAccountNo) = "Bank Account No."
        
        .ColWidth(bteColSelect) = 250
        .ColWidth(bteColBankCode) = 1200
        .ColWidth(bteColBankName) = 2750
        .ColWidth(bteColAddress) = 4000
        .ColWidth(bteColAccountNo) = 2500
        
        .Cell(flexcpAlignment, 0, 0, 0, 4) = flexAlignCenterCenter
    End With
End Sub

Private Sub cbocurr_Click()
cbocurr.Tag = cbocurr.Column(1)
Header
End Sub

Private Sub cbotrade_Change()
Header
cbotrade = Trim(cbotrade)
If cbotrade.MatchFound Then
    LblDescription(0) = cbotrade.Column(1)
End If
End Sub

Private Sub cbotrade_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cmdAction_Click(Index As Integer)
Select Case Index
    Case 0
        Unload Me
        frmMainMenu.Show
    Case 1
        'sumbit
        If HakU = 0 Then _
            LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
                
        If Not validateheader Then Exit Sub
        
                    
        If boldelete = False Then
            If Not Validinput Then Exit Sub
            insertupdate
            stopmsg = False
        Else
            locked (False)
            delete
            Text1(0).SetFocus
        End If
        If stopmsg = False Then
            Call DisplayData(cbotrade, cbocurr.Tag)
            clearinput
        Else
            Call DisplayData(cbotrade, cbocurr.Tag)
        End If
        
    Case 2
        'clear
        clear
        clearinput
        Header
    Case 3
        If Not validateheader Then Exit Sub
        Call DisplayData(cbotrade, cbocurr.Tag)
    Case 4
        clearmark ("S")
        clearmark ("D")
        clearinput
End Select
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Sub adtocombo()
Dim rstcust As Recordset
sql = "SELECT  rtrim(Trade_Master.trade_Code) cust_code, rtrim(Trade_Master.Trade_Name) cust_name, " & _
    "rtrim(Trade_Master.Address1) address, country_Cls From Trade_Master where trade_cls in ('2','3')"
Set rstcust = New Recordset
rstcust.Open sql, Db, adOpenKeyset, adLockOptimistic
With cbotrade
    .clear
    .columnCount = 2
    .ColumnWidths = "50 pt;280 pt; 0 pt"
    .ListWidth = 350
    .ListRows = 15
    
Idx = 0
Do Until rstcust.EOF
    .AddItem ""
    .List(Idx, 0) = Trim(rstcust!Cust_CodE)
    .List(Idx, 1) = Trim(rstcust!Cust_Name)
    .List(Idx, 2) = Trim(rstcust!Address) & " "
    Idx = Idx + 1
    rstcust.MoveNext
Loop
End With

Call up_FillCombo(cbocurr, "curr_cls", "description, curr_cls")
With cbocurr
    .ListWidth = 60
    .ColumnWidths = "60 pt;0 pt"
End With

End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
adtocombo
clear
Header
HakU = hakUpdate(Me.Name)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
End Sub

Sub clear()
cbotrade = ""
cbocurr = ""
grid.Rows = 1
LblDescription(0) = ""
locked (False)
LblErrMsg = ""
End Sub

Sub clearinput()
For Idx = 0 To 3
    Text1(Idx).Text = ""
Next
LblErrMsg = ""
locked (False)
End Sub

Function validateheader() As Boolean
validateheader = True

cbotrade = Trim(cbotrade)
If Not cbotrade.MatchFound Then
    validateheader = False
    LblErrMsg = DisplayMsg(4013)
    Exit Function
End If

cbocurr = Trim(cbocurr)
If Not cbocurr.MatchFound Then
    validateheader = False
    LblErrMsg = DisplayMsg(4005)
    Exit Function
End If

End Function

Function Validinput() As Boolean
Validinput = True
For Idx = 0 To 3
    If Trim(Text1(Idx)) = "" Then
        LblErrMsg = DisplayMsg(1) & " " & Label2(Idx) & " !"
        Validinput = False
        Exit For
    End If
Next
End Function

Sub DisplayData(TradeCode As String, Curr As String)
Dim rstdata As New Recordset, Y As Long, k As Long

sql = "select * from bank_master where trade_code = '" & TradeCode & "' and currency_code = '" & Curr & "'"
If rstdata.State <> adStateClosed Then rstdata.Close
rstdata.Open sql, Db, adOpenKeyset, adLockOptimistic
Idx = 1
With grid
.Rows = rstdata.RecordCount + 1
If Not rstdata.EOF Then
    Do While Not rstdata.EOF
        .Cell(flexcpBackColor, Idx, bteColSelect) = vbWhite
        .TextMatrix(Idx, bteColSelect) = ""
        .TextMatrix(Idx, bteColBankCode) = rstdata!bank_Code
        .TextMatrix(Idx, bteColBankName) = rstdata!bank_name
        .TextMatrix(Idx, bteColAddress) = IIf(IsNull(rstdata!bank_Address), "", Trim(rstdata!bank_Address))
        .TextMatrix(Idx, bteColAccountNo) = rstdata!bank_Account
        Idx = Idx + 1
        rstdata.MoveNext
    Loop
Else
    LblErrMsg = DisplayMsg(4006)
End If
End With

    If stopmsg Then
    Y = 0
    For Idx = 0 To X
        For k = Y + 1 To grid.Rows - 1
            If grid.TextMatrix(k, bteColBankCode) = temp(Idx) Then
                grid.TextMatrix(k, bteColSelect) = "D"
                Y = k
            Else
                grid.TextMatrix(k, bteColSelect) = ""
            End If
        Next
    Next
    stopmsg = False
    locked (True)
    End If
End Sub

Function boldelete() As Boolean
boldelete = False
With grid
    For Idx = 0 To .Rows - 1
        If Trim(.TextMatrix(Idx, bteColSelect)) = "D" Then
            boldelete = True
        End If
    Next
End With
End Function


Sub delete()
On Error Resume Next

If (MsgBox("Are you sure want to delete?", vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmation") = vbYes) Then
    stopmsg = False
    X = 0
    For Idx = 1 To grid.Rows - 1
    
        If grid.TextMatrix(Idx, bteColSelect) = "D" Then
            sql = "delete  from bank_master where bank_Code  ='" & grid.TextMatrix(Idx, bteColBankCode) & "'"
            Db.Execute sql
            ' kalo dah di pake
            If err.number = "-2147217873" Then
                LblErrMsg.Caption = DisplayMsg(8099) 'udah dipake ama yg laen
                temp(X) = grid.TextMatrix(Idx, bteColBankCode)
                stopmsg = True
                X = X + 1
                err.clear
            Else
                If stopmsg = False Then LblErrMsg.Caption = DisplayMsg(1201)
            End If
        End If
    Next
End If
End Sub


Sub insertupdate()
Dim rstbank As New Recordset
sql = "select * from bank_master where bank_Code = '" & Trim(Text1(0)) & "' and Trade_Code = '" & Trim(cbotrade.Text) & "'"
If rstbank.State <> adStateClosed Then rstbank.Close
rstbank.Open sql, Db, adOpenStatic, adLockOptimistic
With rstbank
    If .EOF Then
       .AddNew
       !bank_Code = Text1(0)
       !Trade_Code = cbotrade.Text
       !currency_code = cbocurr.Tag
       !bank_name = Text1(1)
       !bank_Address = Text1(2)
       !bank_Account = Text1(3)
       !Last_Update = Now
       !last_user = userLogin
       .update
       LblErrMsg = DisplayMsg(1000)
    Else
        If grid.Text <> "S" Then
            If MsgBox("Record already exist! Do you want to update?", vbQuestion + vbYesNo) = vbNo Then .filter = "": .Requery: Exit Sub
        End If
       !bank_Code = Text1(0)
       !Trade_Code = cbotrade.Text
       !currency_code = cbocurr.Tag
       !bank_name = Text1(1)
       !bank_Address = Text1(2)
       !bank_Account = Text1(3)
       !Last_Update = Now
       !last_user = userLogin
       .update
       LblErrMsg = DisplayMsg(1101)
    End If
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = True
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim temp As String

temp = grid.Text
If temp = "S" Then
clearmark (temp)
For Idx = 1 To 4
    Text1(Idx - 1).Text = Trim(grid.TextMatrix(Row, Idx))
Next
Text1(0).locked = True
Text1(0).BackColor = &HFDDFE3
ElseIf temp = "D" Then
    clearinput
    clearmark (temp)
    locked (False)
End If
grid.TextMatrix(Row, Col) = temp
grid.Col = Col
grid.Row = Row
LblErrMsg.Caption = ""
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
If KeyAscii = Asc(".") Then KeyAscii = 0
End If
End Sub

Private Sub clearmark(nilai$)
Dim k As Long
If nilai = "D" Then
For k = 1 To grid.Rows - 1
    If grid.TextMatrix(k, bteColSelect) = "S" Then
        grid.TextMatrix(k, bteColSelect) = ""
    End If
Next
Else
For k = 1 To grid.Rows - 1
        grid.TextMatrix(k, bteColSelect) = ""
Next
End If
End Sub

Sub locked(bol As Boolean)
For Idx = 0 To 3
    Text1(Idx).locked = bol
    If bol Then
        Text1(Idx).BackColor = &HFDDFE3
    Else
        Text1(Idx).BackColor = &H80000005
    End If
Next
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 13 Then SendKeys vbTab
End Sub
