VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form F_MachineMaster 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Master"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
   Icon            =   "F_MachineMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Aktif"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10755
      TabIndex        =   22
      Top             =   8370
      Width           =   1095
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   1  'Right Justify
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
      Height          =   240
      Left            =   11610
      MaxLength       =   1
      TabIndex        =   2
      Top             =   8415
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox txtDescription 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   2610
      MaxLength       =   50
      TabIndex        =   1
      Top             =   8340
      Width           =   7815
   End
   Begin VB.CommandButton command1 
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
      Index           =   1
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9630
      Width           =   1125
   End
   Begin VB.CommandButton command2 
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
      Left            =   405
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9660
      Width           =   1125
   End
   Begin VB.CommandButton command1 
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
      Index           =   0
      Left            =   13710
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9630
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   405
      TabIndex        =   5
      Top             =   8850
      Width           =   14430
      Begin VB.Label lblErrMsg 
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
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   14175
      End
   End
   Begin VB.TextBox txtCode 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   540
      MaxLength       =   10
      TabIndex        =   0
      Top             =   8340
      Width           =   1995
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   6540
      Left            =   405
      TabIndex        =   7
      Top             =   1170
      Width           =   14430
      _cx             =   25453
      _cy             =   11536
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
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   0
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
      Height          =   420
      Left            =   12975
      TabIndex        =   20
      Top             =   330
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10905
      TabIndex        =   21
      Top             =   7950
      Width           =   555
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2610
      TabIndex        =   19
      Top             =   7950
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   540
      TabIndex        =   18
      Top             =   7950
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   405
      Top             =   7860
      Width           =   14430
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Machine Master"
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
      Left            =   405
      TabIndex        =   17
      Top             =   360
      Width           =   14430
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
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
      Left            =   10200
      TabIndex        =   16
      Top             =   8010
      Width           =   330
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   8160
      TabIndex        =   15
      Top             =   8010
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Currency"
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
      Left            =   6825
      TabIndex        =   14
      Top             =   8010
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
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
      Left            =   11070
      TabIndex        =   13
      Top             =   8010
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
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
      Index           =   1
      Left            =   12675
      TabIndex        =   12
      Top             =   8010
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Priority"
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
      Index           =   1
      Left            =   5520
      TabIndex        =   11
      Top             =   8010
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   10
      Top             =   8010
      Width           =   960
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   405
      Top             =   8220
      Width           =   14430
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
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
      Index           =   2
      Left            =   14025
      TabIndex        =   9
      Top             =   8010
      Width           =   630
   End
End
Attribute VB_Name = "F_MachineMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ubah As Boolean
Dim hapus As Boolean

Dim bteColSelect As Byte
Dim bteColCode As Byte
Dim bteColDescription As Byte
Dim bteColStatus As Byte
Dim bteColLastUpdate As Byte
Dim bteColLastuser As Byte
Dim bteColLastRegisterDate As Byte

Sub Header()
    bteColSelect = 0
    bteColCode = 1
    bteColDescription = 2
    bteColStatus = 3
    bteColLastuser = 4
    bteColLastUpdate = 5
    bteColLastRegisterDate = 6
    
    With Grid
        .clear
        .Rows = 1
        .ColS = 7
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColCode) = "Code"
        .TextMatrix(0, bteColDescription) = "Description"
        .TextMatrix(0, bteColStatus) = "Status"
        .TextMatrix(0, bteColLastUpdate) = "Last Update"
        .TextMatrix(0, bteColLastuser) = "Last User"
        .TextMatrix(0, bteColLastRegisterDate) = "Register Date"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColCode) = 1800
        .ColWidth(bteColDescription) = 5000
        .ColWidth(bteColStatus) = 800
        
        .ColAlignment(bteColSelect) = flexAlignLeftCenter
        .ColAlignment(bteColCode) = flexAlignLeftCenter
        .ColAlignment(bteColDescription) = flexAlignLeftCenter
        .ColAlignment(bteColStatus) = flexAlignLeftCenter
        
        .ColHidden(bteColLastUpdate) = True
        .ColHidden(bteColLastuser) = True
        .ColHidden(bteColLastRegisterDate) = True
        
        .EditMaxLength = 1
    End With
End Sub

Sub Kosong()
    LblErrMsg.Caption = ""
    txtCode = ""
    txtDescription = ""
    txtStatus = ""
    txtCode.Alignment = 0
    txtDescription.Alignment = 0
    txtStatus.Alignment = 0
    txtStatus.MaxLength = 1
End Sub

Sub Browse()
    
    Dim RS As New Recordset
    Dim rsnama As New Recordset
    
    Dim i As Integer
        
    sql = "select * FROM machine_master"
    RS.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    i = 1
    If Not (RS.BOF And RS.EOF) Then

        With Grid
        .Rows = 1
            Do While Not RS.EOF
                .Rows = .Rows + 1
                .TextMatrix(i, bteColCode) = Trim(RS("Machine_Code"))
                .TextMatrix(i, bteColDescription) = RS("Machine_Name")
                .TextMatrix(i, bteColStatus) = Trim(RS("status"))
                .TextMatrix(i, bteColLastUpdate) = IIf(IsNull(RS("Last_Update")), "", Trim(RS("Last_Update")))
                .TextMatrix(i, bteColLastuser) = IIf(IsNull(RS("Last_User")), "", Trim(RS("Last_User")))
                .TextMatrix(i, bteColLastRegisterDate) = IIf(IsNull(RS("Register_Date")), "", Trim(RS("Register_Date")))
                .Cell(flexcpBackColor, i, bteColSelect) = &HFFFFFF
                RS.MoveNext
                i = i + 1
            Loop
            
        End With
    Else
        
        Header
        
    End If
    RS.Close
    
    
    
    Set RS = Nothing
    
End Sub

Private Sub cmdReport_Click()

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    txtStatus = 1
Else
    txtStatus = 0
End If
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    Header
    Browse
    Kosong
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim TextGrid As String
Dim k As Boolean
Dim j As Integer

k = False
With Grid
    TextGrid = Grid.Text

    If TextGrid = "S" Then
        
        
        txtCode.Text = Trim(.TextMatrix(Row, bteColCode))
        
        txtDescription = Trim(.TextMatrix(Row, bteColDescription))
        
        txtStatus = Trim(.TextMatrix(Row, bteColStatus))
        ubah = True
       Call kosongColGrid
    ElseIf TextGrid = "D" Then
       Call kosongColGrid("S")
    End If
    
    .TextMatrix(Row, Col) = TextGrid
        
    For j = 1 To .Rows - 1
        If .TextMatrix(j, bteColSelect) <> "" Then
            k = True
        End If
    Next j

End With

End Sub

Private Sub kosongColGrid(Optional Kolom As String)
    Dim i As Integer
    
    With Grid
        .Col = bteColSelect
    
        If Kolom <> "" Then
           For i = 1 To .Rows - 1

              If .Text = Kolom Then .Text = ""
              If .TextMatrix(i, bteColSelect) <> "D" Then .TextMatrix(i, bteColSelect) = ""
           Next i
'           kosonggrid
        Else
           For i = 1 To .Rows - 1

              If .TextMatrix(i, bteColSelect) <> "" Then .TextMatrix(i, bteColSelect) = ""

           Next i
        End If
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  If Grid.Col <> bteColSelect Then Cancel = True
End Sub


Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
  If Grid.Col = bteColSelect Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> Asc("S") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
  End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim sql1 As String, tanya
Dim RS As New Recordset

hapus = False
Select Case Index
Case 0:
        If hakUpdate(Me.Name) = 0 Then _
            LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
  
          'HapusData
           With Grid
           Dim LblInput As String
           For i = 1 To .Rows - 1
             If .TextMatrix(i, 0) = "D" Then
                   LblInput = MsgBox("Do you really want to delete ?", _
                    vbYesNo + vbQuestion, "Confirmation")
                    If LblInput = vbYes Then
                        Db.Execute "DELETE FROM Machine_master WHERE Machine_Code='" & .TextMatrix(i, bteColCode) & "'"
                        LblErrMsg = DisplayMsg(1201)
                    End If
             
           Browse
           Exit Sub
             End If
           Next
           End With
           
          If txtCode.Text = "" Then
            txtCode.SetFocus
            LblErrMsg = DisplayMsg(1024)
            Exit Sub
          ElseIf txtDescription.Text = "" Then
            txtDescription.SetFocus
            LblErrMsg.Caption = DisplayMsg(1006)
            Exit Sub

          ElseIf txtStatus = "" Then
            txtStatus.SetFocus
            LblErrMsg.Caption = DisplayMsg(1) & " Status"
            Exit Sub
          End If
          
          'ISI Data...
          Set RS = Nothing
          sql1 = "select * FROM machine_master WHERE Machine_Code='" & txtCode & "'"
          
          RS.Open sql1, Db, 1, 3
          If RS.EOF Then
           RS.AddNew
           RS!Register_Date = Now()
           RS!Machine_code = Trim(txtCode)
          End If
          RS!Machine_Name = txtDescription
          RS!status = Trim(txtStatus.Text)
          RS!Last_Update = Now()
          RS!last_user = Trim(userLogin)
          RS.update
          'tampil data yg baru di update
          Browse
          Kosong
          LblErrMsg = DisplayMsg(1101)
Case 1
    Kosong
    
End Select
Set RS = Nothing
txtCode.SetFocus
End Sub

Private Sub command2_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub txtPrice_LostFocus()
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
End If
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub


Private Sub txtCode_Change()
If Len(txtCode) > 10 Then txtCode = Left(txtCode, 10)
End Sub

Private Sub txtStatus_Change()
If Len(txtStatus) > 1 Then txtStatus = Right(txtStatus, 1)
Check1 = Val(txtStatus)
End Sub
