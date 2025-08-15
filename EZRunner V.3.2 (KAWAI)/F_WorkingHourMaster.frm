VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form F_WorkingHourMaster 
   Caption         =   "Working Hour Master"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "F_WorkingHourMaster.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtShift3 
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
      Left            =   8190
      MaxLength       =   1
      TabIndex        =   21
      Top             =   7860
      Width           =   1095
   End
   Begin VB.TextBox txtShift2 
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
      Left            =   6840
      MaxLength       =   1
      TabIndex        =   20
      Top             =   7860
      Width           =   1095
   End
   Begin VB.CommandButton cmdreport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
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
      Left            =   11565
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9435
      Width           =   1125
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
      Left            =   330
      MaxLength       =   10
      TabIndex        =   7
      Top             =   7860
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   690
      Left            =   240
      TabIndex        =   5
      Top             =   8370
      Width           =   14865
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
         Height          =   405
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Width           =   14640
      End
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
      Left            =   13980
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9435
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
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9465
      Width           =   1125
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
      Left            =   12765
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9435
      Width           =   1125
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
      Left            =   1470
      MaxLength       =   50
      TabIndex        =   1
      Top             =   7860
      Width           =   3885
   End
   Begin VB.TextBox txtShift1 
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
      Left            =   5580
      MaxLength       =   1
      TabIndex        =   0
      Top             =   7860
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5925
      Left            =   210
      TabIndex        =   9
      Top             =   1245
      Width           =   14865
      _cx             =   26220
      _cy             =   10451
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
      Height          =   405
      Left            =   13200
      TabIndex        =   10
      Top             =   345
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   240
      Top             =   7755
      Width           =   14865
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
      Left            =   1470
      TabIndex        =   19
      Top             =   7440
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shift1#"
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
      Left            =   5640
      TabIndex        =   18
      Top             =   7440
      Width           =   630
   End
   Begin VB.Label Label9 
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
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   7440
      Width           =   450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shift2#"
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
      Left            =   6870
      TabIndex        =   16
      Top             =   7440
      Width           =   630
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shift3#"
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
      Left            =   8250
      TabIndex        =   15
      Top             =   7440
      Width           =   630
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Working Hour Master"
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
      Left            =   0
      TabIndex        =   14
      Top             =   360
      Width           =   14865
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   240
      Top             =   7305
      Width           =   14865
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   330
      TabIndex        =   13
      Top             =   7395
      Width           =   555
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2970
      TabIndex        =   12
      Top             =   7395
      Width           =   1035
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   10770
      TabIndex        =   11
      Top             =   7395
      Width           =   705
   End
End
Attribute VB_Name = "F_WorkingHourMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim ubah As Boolean, hapus As Boolean, gavalid As Boolean, ubahedate As Boolean
Dim SDate, EDate, sdateawal, edateakhir
Dim tcode As String, priority As String

Const isiPart = "Finish Good,Parts/wip/material"
Const isiPrice = "Purchase,Sales,Supply,Inventory,Service"
Dim bteColSelect As Byte
Dim bteColCode As Byte
Dim bteColDescription As Byte
Dim bteColSift1 As Byte, bteColSift2 As Byte, bteColSift3 As Byte


Sub Header()
    
    bteColSelect = 0
    bteColCode = 1
    bteColDescription = 2
    bteColSift1 = 3
    bteColSift2 = 4
        bteColSift3 = 5
        
    
    
    With Grid
        .clear
        .Rows = 2
        .FixedRows = 2
        .ColS = 6
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True: .MergeRow(1) = True:
        .MergeCol(bteColSelect) = True
        .Cell(flexcpText, 0, bteColSelect, 1, bteColSelect) = "."
        
        
        .MergeCol(bteColCode) = True
        .Cell(flexcpText, 0, bteColCode, 1, bteColCode) = "Code"
        
        .MergeCol(bteColDescription) = True
        .TextMatrix(0, bteColDescription) = "Description"
        .Cell(flexcpText, 0, bteColDescription, 1, bteColDescription) = "Description"
         
         
        
        .ColAlignment(bteColSift1) = flexAlignCenterTop
        .ColAlignment(bteColSift2) = flexAlignCenterTop
        .ColAlignment(bteColSift3) = flexAlignCenterTop
        .Cell(flexcpText, 0, bteColSift1, 0, bteColSift3) = "Working Hour"
        .TextMatrix(1, bteColSift1) = "Sifht1#"
        .TextMatrix(1, bteColSift2) = "Sifht2#"
        .TextMatrix(1, bteColSift3) = "Sifht3#"
        

        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColCode) = 1000
        .ColWidth(bteColDescription) = 3500
        .ColWidth(bteColSift1) = 1500
        .ColWidth(bteColSift2) = 1500
        .ColWidth(bteColSift3) = 1500
        
        .EditMaxLength = 1
    End With
End Sub

Sub Kosong()
    LblErrMsg.Caption = ""
    txtCode = ""
    txtDescription = ""
    txtShift1 = ""
    txtShift2 = ""
    txtShift3 = ""
    txtCode.Alignment = 0
    txtDescription.Alignment = 0
    txtShift1.Alignment = 0
    txtShift2.Alignment = 0
    txtShift3.Alignment = 0

End Sub



Sub Browse()
    
    Dim RS As New Recordset
    Dim rsnama As New Recordset
    
    Dim i As Integer
        
    sql = "select * FROm dbo.WorkingHour_Master"
    RS.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    i = 2
    If Not (RS.BOF And RS.EOF) Then

        With Grid
        .Rows = 2
            Do While Not RS.EOF
                .Rows = .Rows + 1
                .TextMatrix(i, bteColCode) = Trim(RS("Code"))
                .TextMatrix(i, bteColDescription) = RS("Description")
                .TextMatrix(i, bteColSift1) = Trim(RS("Shift1"))
                .TextMatrix(i, bteColSift2) = Trim(RS("Shift2"))
                .TextMatrix(i, bteColSift3) = Trim(RS("Shift3"))
                
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




Private Sub Form_Load()
    
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    Header
    Browse
    Kosong
    
    
    
    
        
    
    
End Sub

Private Sub mask_LostFocus()
    
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
        txtShift1 = Trim(.TextMatrix(Row, bteColSift1))
        txtShift2 = Trim(.TextMatrix(Row, bteColSift2))
        txtShift3 = Trim(.TextMatrix(Row, bteColSift3))
        
        
        'txtStatus = Trim(.TextMatrix(Row, bteColStatus))
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
                   LblInput = MsgBox("Do you really to delete  ?", _
                    vbYesNo + vbQuestion, "Confirmation")
                    If LblInput = vbYes Then
                        Db.Execute "DELETE FROM WorkingHour_Master WHERE Code='" & .TextMatrix(i, bteColCode) & "'"
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

          ElseIf txtShift1 = "" Then
            txtShift1.SetFocus
            LblErrMsg.Caption = DisplayMsg(1) & " Status"
            Exit Sub
            
          ElseIf txtShift2 = "" Then
            txtShift2.SetFocus
            LblErrMsg.Caption = DisplayMsg(1) & " Status"
            Exit Sub
        
          ElseIf txtShift3 = "" Then
            txtShift3.SetFocus
            LblErrMsg.Caption = DisplayMsg(1) & " Status"
            Exit Sub
        
        
        
          End If
        
           
           

          
          
          'ISI Data...
          Set RS = Nothing
          sql1 = "select * FROM WorkingHour_Master WHERE Code='" & txtCode & "'"
          
          RS.Open sql1, Db, 1, 3
          If RS.EOF Then
           RS.AddNew
           RS!Register_Date = Now()
           RS!code = Trim(txtCode)
          End If
          RS!Description = txtDescription
          RS!Shift1 = Trim(txtShift1.Text)
          RS!Shift2 = Trim(txtShift2.Text)
          RS!Shift3 = Trim(txtShift3.Text)
          'rs!last_date= Now()
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


End Sub


