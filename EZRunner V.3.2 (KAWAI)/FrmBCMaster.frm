VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmBCMaster 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BC Master"
   ClientHeight    =   9930
   ClientLeft      =   285
   ClientTop       =   270
   ClientWidth     =   15075
   Icon            =   "FrmBCMaster.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   15075
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      Left            =   510
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10080
      Width           =   1125
   End
   Begin VB.TextBox Text1 
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
      Left            =   2490
      TabIndex        =   1
      Top             =   8655
      Width           =   3795
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
      Left            =   705
      TabIndex        =   0
      Top             =   8655
      Width           =   1725
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   510
      TabIndex        =   6
      Top             =   9270
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
         Left            =   60
         TabIndex        =   7
         Top             =   210
         Width           =   14190
      End
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
      Left            =   13710
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10080
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10080
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6435
      Left            =   510
      TabIndex        =   5
      Top             =   1560
      Width           =   14205
      _cx             =   25056
      _cy             =   11351
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
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
      Left            =   12840
      TabIndex        =   12
      Top             =   330
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BC Master"
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
      Height          =   405
      Left            =   3120
      TabIndex        =   13
      Top             =   330
      Width           =   8325
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   2520
      TabIndex        =   11
      Top             =   8250
      Width           =   1545
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BC Type"
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
      Left            =   780
      TabIndex        =   10
      Top             =   8250
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   510
      Top             =   8190
      Width           =   5865
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
      Left            =   705
      TabIndex        =   9
      Top             =   8280
      Width           =   750
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
      Left            =   2910
      TabIndex        =   8
      Top             =   8280
      Width           =   825
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   510
      Top             =   8550
      Width           =   5865
   End
End
Attribute VB_Name = "FrmBCMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS As New ADODB.Recordset
Dim hakUI As Integer
Public sql As String
Dim X As Integer, xx As Boolean, blntab As Boolean
Dim stopmsg As Boolean, temp(100000) As String
Dim i As Long, k As Long, LblInput As String

Dim bteColSelect As Byte
Dim btecoltype As Byte
Dim bteColDesc As Byte
Dim bteColLastUpdate As Byte
Dim btecoluser As Byte

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If grid.Col = bteColSelect Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
   End If
If KeyAscii = Asc(".") Then KeyAscii = 0
End If
End Sub
Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 If grid.Col > bteColSelect Then Cancel = True
End Sub

Sub Header()
With grid

    bteColSelect = 0
    btecoltype = 1
    bteColDesc = 2
    bteColLastUpdate = 3
    btecoluser = 4
    
    .ColS = 5
    .Rows = 1

    .TextMatrix(0, bteColSelect) = ""
    .TextMatrix(0, btecoltype) = "BC Type"
    .TextMatrix(0, bteColDesc) = "Description"
    .TextMatrix(0, bteColLastUpdate) = "Last Update"
    .TextMatrix(0, btecoluser) = "User"
    
    
    .ColWidth(bteColSelect) = "300"
    .ColWidth(btecoltype) = "2000"
    .ColWidth(bteColDesc) = "3200"
    .ColWidth(bteColLastUpdate) = "2100"
    .ColWidth(btecoluser) = "1500"
    .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
    
    .EditMaxLength = 1
End With
End Sub
Sub insertupdate()
Dim RS As New ADODB.Recordset
Dim ls_sql As String

If RS.State <> adStateClosed Then RS.Close
RS.Open "Select * from bc_master where bc_type='" & Trim(Text1(0)) & "'", Db, adOpenKeyset, adLockOptimistic

    If RS.EOF = False Then
    ls_sql = "update bc_master set description = '" & Trim(Text1(1)) & "', last_update=getdate(),last_user='" & userLogin & "' where bc_type ='" & Trim(Text1(0)) & "'"
    Db.Execute ls_sql
    LblErrMsg = DisplayMsg(1101)
Else
    ls_sql = " INSERT INTO BC_Master " & vbCrLf & _
                  "         ( Bc_Type , " & vbCrLf & _
                  "           Description , " & vbCrLf & _
                  "           Last_UPDATE , " & vbCrLf & _
                  "           Last_USER " & vbCrLf & _
                  "         ) " & vbCrLf & _
                  " VALUES  ( '" & Trim(Text1(0)) & "' , " & vbCrLf & _
                  "           '" & Trim(Text1(1)) & "' ,  " & vbCrLf & _
                  "           getdate() , " & vbCrLf & _
                  "           '" & userLogin & "' " & vbCrLf & _
                  "         ) "
Db.Execute ls_sql
LblErrMsg = DisplayMsg(1000)
End If
If RS.State <> adStateClosed Then RS.Close

Me.MousePointer = vbDefault
End Sub
Sub display()
Set RS = New Recordset

sql = "SELECT * FROM BC_Master ORDER BY Bc_Type"

If RS.State <> adStateClosed Then RS.Close
RS.Open sql, Db, adOpenDynamic, adLockOptimistic
i = 0
With grid
While Not RS.EOF
        i = i + 1
        .Rows = .Rows + 1
        .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
        .TextMatrix(i, btecoltype) = RS.Fields!BC_Type
        .TextMatrix(i, bteColDesc) = RS.Fields!Description
        .TextMatrix(i, bteColLastUpdate) = Format(RS.Fields!Last_Update, "dd MMM yyyy hh:mm:ss")
        .TextMatrix(i, btecoluser) = RS.Fields!last_user
        .Cell(flexcpAlignment, i, btecoltype, i, bteColDesc) = flexAlignLeftCenter
        RS.MoveNext

Wend
End With
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
        LblErrMsg.Caption = DisplayMsg(9004) ' "Please input '" & label(i).Caption & "'"
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
Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then Cancel = 1
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
Private Sub cmdAction_Click(Index As Integer)
Select Case Index
    Case 0
If RS.State <> adStateClosed Then RS.Close
    Unload Me
    frmMainMenu.Show
    
    Case 3
    Kosong
    Me.CtrlMenu1.MenuText = ""
    RS.filter = ""
    display
    LblErrMsg.Caption = ""
    
    Case 4
    If hakUI = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
        LblErrMsg = ""
        If valid Then
        If xx = False Then
                If uf_validasi = False Then Exit Sub
                insertupdate
                stopmsg = False
        Else
                delete
                kunci (False)
                Text1(0).SetFocus
            End If
        If stopmsg = False Then
            Kosong
            display
        Else
            display
            Text1(0).Enabled = False
        End If
        End If
End Select
End Sub
Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim temp As String

temp = grid.Text
If temp = "S" Then
    clearmark (temp)
    Text1(0).Text = Trim(grid.TextMatrix(Row, btecoltype))
    Text1(1).Text = Trim(grid.TextMatrix(Row, bteColDesc))
    kunci (False)
    Text1(0).Enabled = False
ElseIf temp = "D" Then
    Text1(0) = ""
    Text1(1) = ""
    clearmark (temp)
    kunci (False)
End If
    grid.TextMatrix(Row, Col) = temp
    grid.Col = Col
    grid.Row = Row
    LblErrMsg.Caption = ""
End Sub
Sub delete()
On Error Resume Next
If (MsgBox("Are you sure want to delete?", vbQuestion + vbYesNo, "Confirmation") = vbYes) Then
    stopmsg = False
    X = 0
    For i = 1 To grid.Rows - 1
    
        If grid.TextMatrix(i, bteColSelect) = "D" Then
                sql = "delete BC_master where bc_type ='" & Trim(grid.TextMatrix(i, btecoltype)) & "'"
                Db.Execute sql
                ' kalo dah di pake
                If err.number = "-2147217873" Then
    
                    LblErrMsg.Caption = DisplayMsg(1204) ' "Can't delete this record! This record used by another application" 'udah dipake ama yg laen
                    RS.Requery
                    temp(X) = grid.TextMatrix(i, btecoltype)
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
Private Function uf_validasi()
uf_validasi = False

If Text1(0).Text = "" Then
LblErrMsg = DisplayMsg(9904)
Exit Function
End If

If Text1(1).Text = "" Then
LblErrMsg = DisplayMsg(9905)
Exit Function
End If

If Text1(0).Enabled = True Then
    Dim RS As New ADODB.Recordset
    Dim ls_sql As String
    If RS.State <> adStateClosed Then RS.Close
    RS.Open " select * from Bc_master where bc_type='" & Trim(Text1(0)) & "'", Db, adOpenKeyset, adLockOptimistic
    
    If RS.EOF = False Then
        LblErrMsg = DisplayMsg(8110)
        Exit Function
    End If
End If
uf_validasi = True
End Function
Sub Kosong()
For i = 0 To 1
    Text1(i).Text = ""
    Text1(i).Enabled = True
Next
    Header
    kunci (False)
Text1(0).Enabled = True
Text1(0).SetFocus

End Sub
Sub kunci(bol As Boolean)
    For i = 0 To 0
Text1(0).Enabled = Not bol
    Next
End Sub
Private Sub Command1_Click()
  frmMainMenu.Show
  Unload Me
End Sub

Private Sub Form_Load()
Call Header
Call display
Call SetMaxLength
hakUI = hakUpdate(Me.Name)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
End Sub

Private Sub SetMaxLength()
'Di maxlenght txt nya
Dim rsLen As New ADODB.Recordset
    
    Set rsLen = Db.Execute("SELECT TOP 1 * FROM BC_Master")
    
    Text1(0).MaxLength = rsLen!BC_Type.DefinedSize
    Text1(1).MaxLength = rsLen!Description.DefinedSize
    
    Set rsLen = Nothing
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
End Sub

