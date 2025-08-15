VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmHSmaster 
   BackColor       =   &H00FDDFE3&
   Caption         =   "HS Master"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "FrmHSmaster.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   420
      TabIndex        =   10
      Top             =   9030
      Width           =   14325
      Begin VB.Label LblerrMsg 
         Alignment       =   2  'Center
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
         Left            =   105
         TabIndex        =   11
         Top             =   195
         Width           =   14040
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2760
      Top             =   9750
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=NIC;Data Source=SERVER"
      OLEDBString     =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=NIC;Data Source=SERVER"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   615
      MaxLength       =   15
      TabIndex        =   0
      Top             =   8475
      Width           =   2025
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Index           =   1
      Left            =   2790
      MaxLength       =   25
      TabIndex        =   1
      Top             =   8475
      Width           =   915
   End
   Begin VB.CommandButton CmdAction 
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
      Index           =   4
      Left            =   13620
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9735
      Width           =   1125
   End
   Begin VB.CommandButton CmdAction 
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
      Index           =   0
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9720
      Width           =   1125
   End
   Begin VB.CommandButton CmdAction 
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
      Index           =   3
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9735
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12840
      TabIndex        =   8
      Top             =   180
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6825
      Left            =   420
      TabIndex        =   9
      Top             =   1080
      Width           =   14295
      _cx             =   25215
      _cy             =   12039
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
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   420
      Top             =   8370
      Width           =   3495
   End
   Begin VB.Label label 
      BackStyle       =   0  'Transparent
      Caption         =   "Duty (%)"
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
      Left            =   2820
      TabIndex        =   7
      Top             =   8100
      Width           =   825
   End
   Begin VB.Label label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HS Code"
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
      Left            =   615
      TabIndex        =   6
      Top             =   8100
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   420
      Top             =   8010
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HS Master"
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
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   14205
   End
End
Attribute VB_Name = "FrmHSmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public sql As String

Dim RS As Recordset
Dim stopmsg As Boolean, temp(100000) As String
Dim X As Integer, xx As Boolean, blntab As Boolean
Dim HakU As Integer
Dim i As Long, k As Long, LblInput As String

Dim bteColSelect As Byte
Dim bteColCode As Byte
Dim bteColTax As Byte
Dim bteRegisterDate As Byte
Dim bteLastUpdateDate As Byte
Dim bteUserID As Byte

Sub Header()
With grid
    bteColSelect = 0
    bteColCode = 1
    bteColTax = 2
    bteRegisterDate = 3
    bteLastUpdateDate = 4
    bteUserID = 5
    
    .ColS = 6
    .Rows = 1
    
    .TextMatrix(0, bteColSelect) = ""
    .TextMatrix(0, bteColCode) = "HS Code"
    .TextMatrix(0, bteColTax) = "Duty(%)"
    .TextMatrix(0, bteRegisterDate) = "Register Date"
    .TextMatrix(0, bteLastUpdateDate) = "Last Update"
    .TextMatrix(0, bteUserID) = "User ID"
    
    .ColWidth(bteColSelect) = 250
    .ColWidth(bteColCode) = 2000
    .ColWidth(bteColTax) = 1200
    .ColWidth(bteRegisterDate) = 2500
    .ColWidth(bteLastUpdateDate) = 2500
    .ColWidth(bteUserID) = 1200

    .EditMaxLength = 1
End With
End Sub

Private Sub cmdAction_Click(Index As Integer)
Select Case Index
    Case 0
        If RS.State <> adStateClosed Then RS.Close
        Unload Me
        frmMainMenu.Show
    Case 3
        blank
        Me.CtrlMenu1.MenuText = ""
        RS.filter = ""
        display
        LblErrMsg.Caption = ""
    Case 4
        If HakU = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
        LblErrMsg = ""
        If valid Then
            If xx = False Then
                If uf_validasi = False Then Exit Sub
                insertupdate
                stopmsg = False
            Else
                delete
                locked (False)
                Text1(0).SetFocus
            End If
        If stopmsg = False Then
            blank
            display
        Else
            display
            Text1(0).Enabled = False
            'Text1(1).locked = True
        End If
        End If
End Select

End Sub
Private Function uf_validasi()
uf_validasi = False
If IsNumeric(Text1(1)) = False Then
    LblErrMsg = DisplayMsg(8065)
    Exit Function
End If
If CDbl(Text1(1)) > gd_MaxPercentage Then
    LblErrMsg = DisplayMsg(8064) & " " & gd_MaxPercentage
    Exit Function
End If
If Text1(0).Enabled = True Then
    Dim RS As New ADODB.Recordset
    Dim ls_sql As String
    If RS.State <> adStateClosed Then RS.Close
    RS.Open " select * from hs_master where hs_code='" & Trim(Text1(0)) & "'", Db, adOpenKeyset, adLockOptimistic
    
    If RS.EOF = False Then
        LblErrMsg = DisplayMsg(8063)
        Exit Function
    End If
End If
uf_validasi = True
End Function
Private Sub Form_Activate()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
Text1(0).SetFocus
End Sub

Private Sub Form_Load()

Header
display
Label(0).Caption = "Code"
HakU = hakUpdate(Me.Name)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
End Sub

Sub display()
sql = "Select * from hs_master order by hs_code"
Set RS = New Recordset
If RS.State <> adStateClosed Then RS.Close
RS.Open sql, Db, adOpenKeyset, adLockOptimistic
If Not RS.EOF Then
    With grid
        .Rows = RS.RecordCount + 1
        For i = 1 To RS.RecordCount
            .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
            .TextMatrix(i, bteColCode) = RS.Fields!HS_Code
            .TextMatrix(i, bteColTax) = Format(RS.Fields!tax, gs_formatPercentage)
            .TextMatrix(i, bteRegisterDate) = Format(RS.Fields!Register_Date, "dd MMM yyyy hh:mm:ss")
            .TextMatrix(i, bteLastUpdateDate) = Format(RS.Fields!Last_Update, "dd MMM yyyy hh:mm:ss")
            .TextMatrix(i, bteUserID) = RS.Fields!last_user
            .Cell(flexcpAlignment, i, 1, i, 1) = flexAlignLeftCenter
            RS.MoveNext
        Next
    End With
    
    Dim Y As Integer

    If stopmsg Then
    Y = 0
    For i = 0 To X
        For k = Y + 1 To grid.Rows - 1
            If grid.TextMatrix(k, bteColCode) = temp(i) Then
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
Else
    Header
    LblErrMsg = DisplayMsg(8012) '" No Data Found"
End If
If RS.State <> adStateClosed Then RS.Close
End Sub

Sub insertupdate()
Dim RS As New ADODB.Recordset
Dim ls_sql As String
If RS.State <> adStateClosed Then RS.Close
RS.Open " select * from hs_master where hs_code='" & Trim(Text1(0)) & "'", Db, adOpenKeyset, adLockOptimistic

If RS.EOF = False Then
    ls_sql = " update hs_master set tax= " & CDbl(Text1(1)) & ", last_update=getdate(), last_user='" & userLogin & "' where hs_code='" & Trim(Text1(0)) & "'"
    Db.Execute ls_sql
    LblErrMsg = DisplayMsg(1101)
Else
   ls_sql = " INSERT INTO [HS_Master] " & _
                  "            ([HS_Code] " & _
                  "            ,[Tax] " & _
                  "            ,[Last_Update] " & _
                  "            ,[Last_User] " & _
                  "            ,[Register_Date]) " & _
                  "      VALUES " & _
                  "            ('" & Trim(Text1(0)) & "' " & _
                  "            ," & CDbl(Text1(1)) & " " & _
                  "            ,getdate() " & _
                  "            ,'" & userLogin & "' "

    ls_sql = ls_sql + "            ,getdate()) "
    Db.Execute ls_sql
   LblErrMsg = DisplayMsg(1000)
End If
If RS.State <> adStateClosed Then RS.Close

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim temp As String

temp = grid.Text
If temp = "S" Then
clearmark (temp)
Text1(0).Text = Trim(grid.TextMatrix(Row, bteColCode))
Text1(1).Text = Trim(grid.TextMatrix(Row, bteColTax))
locked (False)
Text1(0).Enabled = False
'Text1(0).locked = True
'Text1(0).BackColor = &HE0E0E0   '&HFDDFE3
ElseIf temp = "D" Then
    Text1(0) = ""
    Text1(1) = ""
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

Sub locked(bol As Boolean)
For i = 0 To 0
    Text1(0).Enabled = Not bol
    'Text1(i).locked = bol
'    If bol Then
'        Text1(i).BackColor = &HFDDFE3
'    Else
'        Text1(i).BackColor = &H80000005
'    End If
Next
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

Sub delete()
On Error Resume Next
If (MsgBox("Are you sure want to delete?", vbQuestion + vbYesNo, "Confirmation") = vbYes) Then
    stopmsg = False
    X = 0
    For i = 1 To grid.Rows - 1
    
        If grid.TextMatrix(i, bteColSelect) = "D" Then
                sql = "delete hs_master where hs_code='" & Trim(grid.TextMatrix(i, bteColCode)) & "'"
                Db.Execute sql
                ' kalo dah di pake
                If err.number = "-2147217873" Then
    
                    LblErrMsg.Caption = DisplayMsg(1204) ' "Can't delete this record! This record used by another application" 'udah dipake ama yg laen
                    RS.Requery
                    temp(X) = grid.TextMatrix(i, bteColCode)
                    stopmsg = True
                    X = X + 1
                    err.clear
                Else
                    RS.Requery
                    If stopmsg = False Then LblErrMsg.Caption = DisplayMsg(1201)
                End If
        End If
    Next
End If
RS.Requery
End Sub

Private Sub blank()
For i = 0 To 1
    Text1(i).Text = ""
    Text1(i).Enabled = True
Next
Header
locked (False)
'Text1(0).locked = False
Text1(0).Enabled = True

'Text1(0).BackColor = vbWhite
Text1(0).SetFocus

End Sub

Function valid() As Boolean
valid = True
xx = False
For i = 1 To grid.Rows - 1
    If grid.TextMatrix(i, bteColSelect) = "D" Then
        xx = True
        Exit For
    End If
Next
If xx Then Exit Function
For i = 0 To 0
    If Trim(Text1(i).Text) = "" Then
        LblErrMsg.Caption = DisplayMsg(8060) ' "Please input '" & label(i).Caption & "'"
        valid = False
        Text1(i).SetFocus
        Exit Function
    End If
Next

End Function

Function check() As Boolean
If grid.Rows = 1 Then check = True: LblErrMsg = "": Exit Function
For i = 1 To grid.Rows - 1
    If Trim(grid.TextMatrix(i, bteColSelect)) <> "" Then
        check = False
        Exit For
    Else
        check = True
    End If
Next
If Trim(Text1(0)) <> "" Or Trim(Text1(1)) <> "" Then check = False
LblErrMsg.Caption = ""
End Function


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = 0
If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub
