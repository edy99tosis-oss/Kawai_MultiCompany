VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form F_TAX 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Tax Exchange Rate"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "F_TAX.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   390
      TabIndex        =   17
      Top             =   9210
      Width           =   14430
      Begin VB.Label LblErrMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lbl_pesan"
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
         TabIndex        =   18
         Top             =   180
         Width           =   14175
      End
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   540
      MaxLength       =   20
      TabIndex        =   0
      Top             =   8670
      Width           =   975
   End
   Begin VB.CommandButton Cmd_Save 
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
      Left            =   13710
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9885
      Width           =   1125
   End
   Begin VB.TextBox TxtNilai 
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
      Height          =   350
      Left            =   5010
      MaxLength       =   6
      TabIndex        =   3
      Text            =   "0"
      Top             =   8670
      Width           =   975
   End
   Begin VB.TextBox TxtName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1575
      MaxLength       =   20
      TabIndex        =   2
      Top             =   8670
      Width           =   3375
   End
   Begin VB.CommandButton Cmd_SubMenu 
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
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9885
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Clear 
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
      Left            =   12465
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9885
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker dtsdate 
      Height          =   345
      Left            =   6045
      TabIndex        =   4
      Top             =   8670
      Width           =   1545
      _ExtentX        =   2725
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
      CustomFormat    =   "dd MMM yyyy"
      Format          =   151650307
      CurrentDate     =   37798
   End
   Begin MSMask.MaskEdBox MEDate 
      Height          =   345
      Left            =   7650
      TabIndex        =   5
      Top             =   8670
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker dtedate 
      Height          =   345
      Left            =   7650
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8670
      Width           =   1545
      _ExtentX        =   2725
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
      CustomFormat    =   "dd MMM yyyy"
      Format          =   151650307
      CurrentDate     =   37798
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6525
      Left            =   405
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1545
      Width           =   14445
      _cx             =   25479
      _cy             =   11509
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
      Left            =   13005
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   555
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label LblCode 
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
      Left            =   795
      TabIndex        =   16
      Top             =   8265
      Width           =   450
   End
   Begin VB.Label Label3 
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
      Left            =   7650
      TabIndex        =   15
      Top             =   8265
      Width           =   780
   End
   Begin VB.Label Label2 
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
      Left            =   6045
      TabIndex        =   14
      Top             =   8265
      Width           =   885
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   405
      Top             =   8565
      Width           =   8895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate %"
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
      Left            =   5190
      TabIndex        =   13
      Top             =   8265
      Width           =   630
   End
   Begin VB.Label LblName 
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
      Left            =   1605
      TabIndex        =   12
      Top             =   8265
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Classification"
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
      Left            =   405
      TabIndex        =   11
      Top             =   585
      Width           =   14445
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   405
      Top             =   8205
      Width           =   8895
   End
End
Attribute VB_Name = "F_TAX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim RS As New Recordset
Dim ubah As Boolean, hapus As Boolean, gavalid As Boolean, ubahedate As Boolean
Dim SDate, EDate, sdateawal, edateakhir
Dim i As Integer

Dim bteColSelect As Byte
Dim bteColCode As Byte
Dim bteColDesc As Byte
Dim bteColRate As Byte
Dim bteColDateStart As Byte
Dim bteColDateEnd As Byte

Sub Kosong()
    txtName.Text = ""
    TxtNilai.Text = ""
    dtsdate.Value = Format(Now, "dd MMM yyyy")
    dtedate.Value = Format(Now, "dd MMM yyyy")
    MEDate.Text = "99/99/9999"
    LblErrMsg.Caption = ""
    ubah = False
End Sub

Sub Header()
    With Grid
        bteColSelect = 0
        bteColCode = 1
        bteColDesc = 2
        bteColRate = 3
        bteColDateStart = 4
        bteColDateEnd = 5
                
        .clear
        .Rows = 1
        .ColS = 6
        
        .TextMatrix(0, bteColCode) = "Code"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColRate) = "Rate %"
        .TextMatrix(0, bteColDateStart) = "Start Date"
        .TextMatrix(0, bteColDateEnd) = "End Date"
        
        .ColWidth(bteColSelect) = 250
        .ColWidth(bteColCode) = 1000
        .ColWidth(bteColDesc) = 3500
        .ColWidth(bteColRate) = 1000
        .ColWidth(bteColDateStart) = 1350
        .ColWidth(bteColDateEnd) = 1350
                
        .ColDataType(bteColDateStart) = flexDTDate
        .ColDataType(bteColDateEnd) = flexDTDate
        
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColCode) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColRate) = flexAlignRightCenter
        .ColAlignment(bteColDateStart) = flexAlignCenterCenter
        .ColAlignment(bteColDateEnd) = flexAlignCenterCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
        .EditMaxLength = 1
    End With
End Sub

Sub Browse()
    Dim tglAwal As String
    Dim tglAkhir As String
    RS.filter = ""
    RS.Requery
    i = 1
    If Not (RS.BOF And RS.EOF) Then
        With Grid
            Do While Not RS.EOF
                .Rows = .Rows + 1
                
                tglAwal = Mid(RS("start_date"), 5, 2) & "/" & Right(RS("start_date"), 2) & "/" & Left(RS("start_date"), 4)
                tglAkhir = IIf(IsNull(RS("End_date")), "99/99/9999", Mid(RS("end_date"), 5, 2) & "/" & Right(RS("end_date"), 2) & "/" & Left(RS("end_date"), 4))
                
                .TextMatrix(i, bteColCode) = Trim(RS("tax_code"))
                .TextMatrix(i, bteColDesc) = IIf(IsNull(RS("tax_name")), "", Trim(RS("tax_name")))
                .TextMatrix(i, bteColRate) = IIf(IsNull(RS("rate")), 0, Format(Trim(RS("rate")), gs_formatExchangeRate))
                .TextMatrix(i, bteColDateStart) = Format(tglAwal, "dd MMM yyyy")
                .TextMatrix(i, bteColDateEnd) = Format(tglAkhir, "dd MMM yyyy")
                .Cell(flexcpBackColor, i, bteColSelect) = &HFFFFFF
                
                RS.MoveNext
                i = i + 1
            Loop
        End With
    Else
        Kosong
        Header
    End If
End Sub

Sub cektgl()
    Dim rs2 As New Recordset
    Dim rs3 As New Recordset
    Dim Tgl As String
    Dim TempDate As String
    
    gavalid = False
    ubahedate = False
    
    If hapus Then
        sql = "select * from tax_cls where tax_code='" & txtCode.Text & "' and " & _
        "start_date<'" & SDate & "' order by start_date, end_date"
        If rs2.State <> adStateClosed Then rs2.Close
        rs2.Open sql, Db, adOpenKeyset, adLockOptimistic
        
        sql = "select * from tax_cls where tax_code='" & txtCode.Text & "' and " & _
        "start_date>'" & SDate & "' order by start_date, end_date"
        If rs3.State <> adStateClosed Then rs3.Close
        rs3.Open sql, Db, adOpenKeyset, adLockOptimistic
    
        If Not (rs2.BOF And rs2.EOF) Then
            rs2.MoveLast
            If Not (rs3.BOF And rs3.EOF) Then
                rs3.MoveFirst
                Tgl = Mid(rs3("start_date"), 5, 2) & "/" & Right(rs3("start_date"), 2) & "/" & Left(rs3("start_date"), 4)
                TempDate = Format(DateAdd("d", -1, CDate(Tgl)), "yyyymmdd")
                sql = "update tax_cls set end_date='" & TempDate & "', Last_Update = getdate(), Last_User = '" & userLogin & "' where tax_code='" & _
                rs2!tax_code & "' and start_date='" & rs2!Start_Date & "' "
                Db.Execute sql
            Else
                sql = "update tax_cls set end_date='99999999', Last_Update = getdate(), Last_User = '" & userLogin & "'  where tax_code='" & _
                rs2!tax_code & "' and start_date='" & rs2!Start_Date & "' "
                Db.Execute sql
            End If
        End If
        Exit Sub
    End If
    
    If ubah = False Then
        SDate = Format(dtsdate.Value, "yyyymmdd")
        EDate = Format(MEDate.Text, "yyyymmdd")
    
        sql = "select * from tax_cls where tax_code='" & txtCode.Text & "' and " & _
        "start_date<'" & SDate & "' order by start_date,end_date"
        If rs2.State <> adStateClosed Then rs2.Close
        rs2.Open sql, Db, adOpenKeyset, adLockOptimistic
    
        sql = "select * from tax_cls where tax_code='" & txtCode.Text & "' and " & _
        "start_date>'" & SDate & "' order by start_date, end_date"
        If rs3.State <> adStateClosed Then rs3.Close
        rs3.Open sql, Db, adOpenKeyset, adLockOptimistic
    
        If Not (rs3.BOF And rs3.EOF) Then
            rs3.MoveFirst
            Tgl = Mid(rs3("start_date"), 5, 2) & "/" & Right(rs3("start_date"), 2) & "/" & Left(rs3("start_date"), 4)
            TempDate = Format(CDate(Tgl), "yyyymmdd")
            If EDate = "99/99/9999" Then
                ubahedate = True
                edateakhir = Format(DateAdd("d", -1, CDate(Tgl)), "yyyymmdd")
            Else
                If (EDate >= TempDate) Then
                    LblErrMsg.Caption = DisplayMsg(8054) & " " & Format(CDate(Tgl), "dd MMM yyyy") '"End Date must be lower than "
                    gavalid = True
                    dtedate.SetFocus
                    MEDate.SetFocus
                    Exit Sub
                End If
            End If
        End If
        
        If Not (rs2.BOF And rs2.EOF) Then
            rs2.MoveLast
            TempDate = Format(DateAdd("d", -1, CDate(dtsdate.Value)), "yyyymmdd")
            sql = "update tax_cls set end_date='" & TempDate & "', Last_Update = getdate(), Last_User = '" & userLogin & "' where tax_code='" & _
            rs2!tax_code & "' and start_date='" & rs2!Start_Date & "' "
            Db.Execute sql
        End If
        Exit Sub
    Else
        SDate = Format(dtsdate.Value, "yyyymmdd")
        EDate = Format(MEDate.Text, "yyyymmdd")
    
        sql = "select * from tax_cls where tax_code='" & txtCode.Text & "' and " & _
        "start_date<'" & sdateawal & "' order by start_date,end_date"
        If rs2.State <> adStateClosed Then rs2.Close
        rs2.Open sql, Db, adOpenKeyset, adLockOptimistic
    
        sql = "select * from tax_cls where tax_code='" & txtCode.Text & "' and " & _
        "start_date>'" & sdateawal & "' order by start_date, end_date"
        If rs3.State <> adStateClosed Then rs3.Close
        rs3.Open sql, Db, adOpenKeyset, adLockOptimistic
    
        If Not (rs3.BOF And rs3.EOF) Then
            rs3.MoveFirst
            Tgl = Mid(rs3("start_date"), 5, 2) & "/" & Right(rs3("start_date"), 2) & "/" & Left(rs3("start_date"), 4)
            TempDate = Format(CDate(Tgl), "yyyymmdd")
            If EDate = "99/99/9999" Then
                ubahedate = True
                edateakhir = Format(DateAdd("d", -1, CDate(Tgl)), "yyyymmdd")
            Else
                If (EDate >= TempDate) Then
                    LblErrMsg.Caption = DisplayMsg(8054) & " " & Format(CDate(Tgl), "dd MMM yyyy") '"End Date must be lower than "
                    gavalid = True
                    dtedate.SetFocus
                    MEDate.SetFocus
                    Exit Sub
                End If
            End If
        End If
    
        If Not (rs2.BOF And rs2.EOF) Then
            rs2.MoveLast
            Tgl = Mid(rs2("start_date"), 5, 2) & "/" & Right(rs2("start_date"), 2) & "/" & Left(rs2("start_date"), 4)
            TempDate = Format(CDate(Tgl), "yyyymmdd")
            If (SDate <= TempDate) Then
                LblErrMsg.Caption = DisplayMsg(8055) & " " & Format(CDate(Tgl), "dd MMM yyyy") '"Start Date must be higher than "
                gavalid = True
                dtsdate.SetFocus
                Exit Sub
            Else
                TempDate = Format(DateAdd("d", -1, CDate(dtsdate.Value)), "yyyymmdd")
                sql = "update tax_cls set end_date='" & TempDate & "', Last_Update = getdate(), Last_User = '" & userLogin & "' where tax_code='" & _
                rs2!tax_code & "' and start_date='" & rs2!Start_Date & "' "
                Db.Execute sql
            End If
        End If
        Exit Sub
    End If

End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    Kosong
    Header
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    sql = "select * from tax_cls order by tax_code, start_date, end_date"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    Browse
End Sub

Private Sub dtsdate_Change()
    If MEDate.Text <> "99/99/9999" Then
        If CDate(dtsdate) > CDate(dtedate) Then
            LblErrMsg.Caption = DisplayMsg(4068)
            dtsdate.SetFocus
            Exit Sub
        Else
            LblErrMsg.Caption = ""
        End If
    End If
End Sub

Private Sub dtsdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtsdate_Change
End Sub

Private Sub dtedate_Change()
    MEDate.Text = Format(dtedate, "MM/dd/yyyy")
    If CDate(dtedate) < CDate(dtsdate) Then
        LblErrMsg.Caption = DisplayMsg(4066)
        MEDate.SetFocus
        Exit Sub
    Else
        LblErrMsg.Caption = ""
    End If
End Sub

Private Sub dtedate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtedate_Change
End Sub

Private Sub medate_LostFocus()
    If IsDate(MEDate.Text) = True Then dtedate.Value = CDate(MEDate.Text)
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim TextGrid As String
    Dim k As Boolean
    Dim j As Integer
    
    k = False
    With Grid
        TextGrid = Grid.Text
        If TextGrid = "S" Then
            txtCode.Text = .TextMatrix(Row, bteColCode)
            txtName.Text = .TextMatrix(Row, bteColDesc)
            TxtNilai.Text = Format(CDbl(.TextMatrix(Row, bteColRate)), gs_formatExchangeRate)
            dtsdate.Value = Format(.TextMatrix(Row, bteColDateStart), "mm/dd/yyyy")
            sdateawal = Format(.TextMatrix(Row, bteColDateStart), "yyyymmdd")
            MEDate.Text = Format(.TextMatrix(Row, bteColDateEnd), "mm/dd/yyyy")
            If .TextMatrix(Row, bteColDateEnd) <> "99/99/9999" Then dtedate = Format(.TextMatrix(Row, bteColDateEnd), "mm/dd/yyyy")
            ubah = True
            Call kosongColGrid
        ElseIf TextGrid = "D" Then
            Call kosongColGrid("S")
        End If
        
        .TextMatrix(Row, Col) = TextGrid
        For j = 1 To .Rows - 1
            If .TextMatrix(j, bteColSelect) <> "" Then k = True
        Next j
        If k = False Then Kosong
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
            Kosong
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

Private Sub Cmd_Save_Click()
    Dim sql1 As String
    Dim tanya
    
    hapus = False
    If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
    With Grid
        For i = 1 To .Rows - 1
            If .TextMatrix(i, bteColSelect) = "D" Then
                If IsEmpty(tanya) Then tanya = MsgBox("Do You Really Want To Delete This Data ?", vbQuestion & vbYesNo, "Confirmation")
                If tanya = vbYes Then
                    sql1 = "delete from tax_cls where tax_Code='" & .TextMatrix(i, bteColCode) & "' and " & _
                    "start_date='" & Format(.TextMatrix(i, bteColDateStart), "yyyymmdd") & "'"
                    Db.Execute sql1
    
                    hapus = True
                    SDate = Format(.TextMatrix(i, bteColDateStart), "yyyymmdd")
                    EDate = Format(.TextMatrix(i, bteColDateEnd), "yyyymmdd")
                    cektgl
                Else
                    Exit For
                End If
            End If
        Next i
        
        If (hapus) Then Kosong: Header: Browse: LblErrMsg = DisplayMsg(1201): Exit Sub
    
        If txtName.Text = "" Then
            txtName.SetFocus
            LblErrMsg = DisplayMsg(1050)
            Exit Sub
        ElseIf TxtNilai.Text = "" Then
            TxtNilai.SetFocus
            LblErrMsg = DisplayMsg(1051)
            Exit Sub
        Else
            If MEDate.Text <> "99/99/9999" Then
                If IsDate(MEDate.Text) = False Then
                    LblErrMsg.Caption = DisplayMsg(4065)  '"End Date is not valid"
                    MEDate.SetFocus
                    Exit Sub
                End If
                If CDate(dtsdate) > CDate(dtedate) Then
                    LblErrMsg.Caption = "Start Date must be lower than " & Format(dtedate, "dd MMM yyyy")     '"Start Date must be lower than "
                    dtsdate.SetFocus
                    Exit Sub
                End If
            End If
            If ubah = False Then
                RS.filter = "tax_Code='" & txtCode.Text & "' and start_date='" & _
                Format(dtsdate.Value, "yyyymmdd") & "' "
                If Not (RS.EOF And RS.BOF) Then
                    LblErrMsg = DisplayMsg(1023): dtsdate.SetFocus: Exit Sub
                Else
                    cektgl
                    If gavalid Then Exit Sub
                    RS.AddNew
                    RS("tax_Code") = txtCode.Text
                End If
            Else
                RS.filter = "tax_Code='" & txtCode.Text & "' and start_date='" & _
                    sdateawal & "' "
            End If
            cektgl
            If gavalid Then Exit Sub
    
            RS("Tax_name") = txtName.Text
            RS("rate") = TxtNilai.Text
            RS("start_date") = Format(dtsdate.Value, "yyyymmdd")
    
            If MEDate.Text = "99/99/9999" Then
                If ubahedate = True Then
                    RS("End_date") = edateakhir
                Else
                    RS("end_date") = "99999999"
                End If
            Else
                RS("end_date") = Format(MEDate.Text, "yyyymmdd")
            End If
            RS("last_update") = Now
            RS("last_user") = userLogin
            RS.update
            RS.Requery
            RS.filter = ""
    
            Kosong
            Header
            Browse
    
            LblErrMsg = DisplayMsg(IIf((ubah = False), 1000, 1101))
            ubah = False
        End If
    End With
End Sub

Private Sub cmd_clear_Click()
    Kosong
    Header
    Browse
    txtName.SetFocus
End Sub

Private Sub Cmd_SubMenu_Click()
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

Private Sub txtnilai_LostFocus()
    Dim z As Double
    If TxtNilai.Text <> "" Then
        z = CDbl(TxtNilai.Text)
        If z > gd_MaxExchangeRate Then TxtNilai.Text = Left(z, 3)
    End If
    TxtNilai.Text = Format(TxtNilai.Text, gs_formatExchangeRate)
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
    If (TxtNilai.Text & Chr(KeyAscii)) > gd_MaxExchangeRate And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RS.State <> adStateClosed Then RS.Close
End Sub


