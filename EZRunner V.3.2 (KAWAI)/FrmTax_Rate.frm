VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTax_Rate 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Exchange Rate"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "FrmTax_Rate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   377
      TabIndex        =   19
      Top             =   5700
      Width           =   7275
      Begin VB.Label LblErr 
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
         Left            =   105
         TabIndex        =   20
         Top             =   195
         Width           =   7050
      End
   End
   Begin VB.CommandButton CmdMenu 
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
      Left            =   377
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1140
      Left            =   377
      TabIndex        =   11
      Top             =   1350
      Width           =   7275
      Begin VB.Frame FrTgl1 
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1290
         TabIndex        =   12
         Top             =   660
         Width           =   1515
         Begin MSComCtl2.DTPicker Tgl1 
            Height          =   330
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   582
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
            Format          =   294191107
            CurrentDate     =   37868
         End
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   330
         Left            =   3990
         TabIndex        =   3
         Top             =   660
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
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
         Format          =   294191107
         CurrentDate     =   37868
      End
      Begin MSComCtl2.DTPicker Mperiod 
         Height          =   330
         Left            =   1290
         TabIndex        =   0
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
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
         Format          =   294191107
         UpDown          =   -1  'True
         CurrentDate     =   37868
      End
      Begin VB.Label LblTax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
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
         Left            =   180
         TabIndex        =   16
         Top             =   315
         Width           =   540
      End
      Begin VB.Label LblTax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week"
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
         Left            =   3090
         TabIndex        =   15
         Top             =   315
         Width           =   480
      End
      Begin VB.Label LblTax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start date"
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
         Left            =   180
         TabIndex        =   14
         Top             =   735
         Width           =   855
      End
      Begin VB.Label LblTax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End date"
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
         Index           =   4
         Left            =   3090
         TabIndex        =   13
         Top             =   735
         Width           =   750
      End
      Begin MSForms.ComboBox CboWeek 
         Height          =   330
         Left            =   3990
         TabIndex        =   2
         Top             =   240
         Width           =   765
         VariousPropertyBits=   746604571
         MaxLength       =   1
         DisplayStyle    =   3
         Size            =   "1349;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton CmdSubmit 
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
      Left            =   6497
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   1155
   End
   Begin VB.TextBox TxtRate 
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
      Left            =   2293
      TabIndex        =   5
      Top             =   5280
      Width           =   2115
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   5805
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2205
      Left            =   390
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2565
      Width           =   7275
      _cx             =   12832
      _cy             =   3889
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
   Begin VB.Label LblTax 
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
      Index           =   3
      Left            =   510
      TabIndex        =   21
      Top             =   4920
      Width           =   795
   End
   Begin MSForms.ComboBox CboCurr 
      Height          =   315
      Left            =   510
      TabIndex        =   4
      Top             =   5280
      Width           =   795
      VariousPropertyBits=   746604571
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "1402;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      Height          =   510
      Index           =   0
      Left            =   390
      Top             =   5175
      Width           =   5685
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Exchange Rate"
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
      Left            =   390
      TabIndex        =   18
      Top             =   540
      Width           =   7275
   End
   Begin VB.Line Line1 
      X1              =   1395
      X2              =   2175
      Y1              =   5580
      Y2              =   5580
   End
   Begin VB.Label LblCurr 
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
      Height          =   255
      Left            =   1395
      TabIndex        =   9
      Top             =   5310
      Width           =   795
   End
   Begin VB.Label LblTax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exchange Rate"
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
      Index           =   5
      Left            =   2295
      TabIndex        =   8
      Top             =   4920
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   330
      Index           =   2
      Left            =   390
      Top             =   4860
      Width           =   5685
   End
End
Attribute VB_Name = "FrmTax_Rate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RSTR As ADODB.Recordset
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

Dim bteColSelect As Byte
Dim bteColCurr As Byte
Dim bteColExRate As Byte
Dim bteColCurrCode As Byte
Dim bteColWeekCode As Byte
Dim bteColStartDate As Byte
Dim bteColEndDate As Byte

Sub Header()
    
    bteColSelect = 0
    bteColCurr = 1
    bteColExRate = 2
    bteColCurrCode = 3
    bteColWeekCode = 4
    bteColStartDate = 5
    bteColEndDate = 6
    
    With grid
        .ColS = 7
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColCurr) = "Currency Code"
        .TextMatrix(0, bteColExRate) = "Exchange Rate"
        .TextMatrix(0, bteColCurrCode) = "CurrCode"
        .TextMatrix(0, bteColWeekCode) = "WeekCode"
        .TextMatrix(0, bteColStartDate) = "StartDate"
        .TextMatrix(0, bteColEndDate) = "EndDate"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColCurr) = 1500
        .ColWidth(bteColExRate) = 1800
        
        .ColHidden(bteColCurrCode) = True
        .ColHidden(bteColWeekCode) = True
        .ColHidden(bteColStartDate) = True
        .ColHidden(bteColEndDate) = True
        
        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        .ColAlignment(bteColExRate) = flexAlignCenterCenter
        
        .EditMaxLength = 1
    End With
    
End Sub

Private Sub cbocurr_Change()
LblCurr = ""
cbocurr = cbocurr
If cbocurr.MatchFound Then LblCurr = cbocurr.List(cbocurr.ListIndex, 1)
TxtRate = Format(0, gs_formatExchangeRate)
TxtRate.SetFocus
End Sub

Private Sub cbocurr_Click()
LblCurr = ""
cbocurr = cbocurr
If cbocurr.MatchFound Then LblCurr = cbocurr.List(cbocurr.ListIndex, 1)
TxtRate = Format(0, gs_formatExchangeRate)
TxtRate.SetFocus
End Sub

Private Sub CboWeek_Change()
Dim RsCT As Recordset, Sqlc As String, TglBaru As String, tglAkhir As String
'Dim AM As Boolean
Set RsCT = New Recordset
'If AM = False Then

CboWeek = CboWeek
If CboWeek.MatchFound = False Then Exit Sub

Dim RsLM As Recordset
Set RsLM = Db.Execute("Select Week_code,End_date from tax_exchangerate where exch_year='" & Mperiod.Year & "' and exch_month='" & Format(Mperiod.Month, "00") & "' and End_date='" & Format(LDay, "YYYYMMDD") & "'")
If Not RsLM.EOF Then
  If CboWeek > Val(RsLM!week_code) Then
    LblErr = DisplayMsg("0066")
    CboWeek.Text = Val(RsLM!week_code)
    RsLM.Close
    Set RsLM = Nothing
    Exit Sub
  End If
End If
LblErr = ""
Sqlc = "Select * from Tax_exchangerate where exch_year='" & Mperiod.Year & "' and exch_month='" & Format(Mperiod.Month, "00") & "' "
Sqlc = Sqlc & " and week_code='" & Val(CboWeek) & "'"
CboWeek = CboWeek
Header
RsCT.Open Sqlc, Db, adOpenDynamic, adLockOptimistic
 
FrTgl1.Enabled = True
If Not RsCT.EOF Then
     FrTgl1.Enabled = True
     If CboWeek > 1 Then FrTgl1.Enabled = False
        
     Tgl1 = Format(Left(RsCT!Start_Date, 4) & "/" & Mid(RsCT!Start_Date, 5, 2) & "/" & Right(RsCT!Start_Date, 2), "DD MMM YYYY")
     Tgl2 = Format(Left(RsCT!end_date, 4) & "/" & Mid(RsCT!end_date, 5, 2) & "/" & Right(RsCT!end_date, 2), "DD MMM YYYY")
     PrevWeekEmpty = False
     Browse
Else
    If CboWeek > 1 Then
        Dim RsK As Recordset
        Sqlc = "Select * from Tax_exchangerate where exch_year='" & Mperiod.Year & "' and exch_month='" & Format(Mperiod.Month, "00") & "' and week_code='" & CboWeek - 1 & "'"
        FrTgl1.Enabled = False
        Set RsK = Db.Execute(Sqlc)
         If RsK.EOF Then
            LblErr = DisplayMsg("0067")
            PrevWeekEmpty = True
            Exit Sub
         Else
            
            PrevWeekEmpty = False
            TglBaru = RsK!end_date + 1
            tglAkhir = DateSerial(Left(TglBaru, 4), Mid(TglBaru, 5, 2) + 1, 0)
            
            
            FrTgl1.Enabled = False
            If TglBaru > Format(tglAkhir, "YYYYMMDD") Then
                Tgl1 = Format(Left(RsK!end_date, 4) & "/" & Mid(RsK!end_date, 5, 2) & "/" & Right(RsK!end_date, 2), "DD MMM YYYY")
                Tgl2 = Tgl1
            Else
                Tgl1 = Format(Left(TglBaru, 4) & "/" & Mid(TglBaru, 5, 2) & "/" & Right(TglBaru, 2), "DD MMM YYYY")
                Tgl2 = Tgl1
            End If
            If RsK!end_date = Format(tglAkhir, "YYYYMMDD") Then LblErr = DisplayMsg("0066"): Exit Sub
            Browse
        End If
    Else
        PrevWeekEmpty = False
    End If
End If
tgl_sb = Tgl2
End Sub

Private Sub CboWeek_Click()
CboWeek_Change
End Sub

Private Sub CmdMenu_Click()
frmMainMenu.Show
Unload Me
DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub
Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErr.Caption = ErrMsg
End If
End Sub

Private Sub CmdSubmit_Click()
Dim strS As Integer, Jawab As Integer, CekS As Boolean, CekD As Boolean
Dim strD As Integer, ie As Integer, RSCA As Recordset
Dim TglA As String, TglAk As String
CekS = False
CekD = False
StaErr = False

If HakU = 0 Then _
LblErr = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub

If Tgl1 > Tgl2 Then
   LblErr = DisplayMsg(4066)
   Exit Sub
End If

If Not IsNumeric(TxtRate) Then TxtRate = Format(0, gs_formatExchangeRate)
If CDbl(TxtRate) > gd_MaxExchangeRate Then
   LblErr = DisplayMsg("0068") & " " & Format(gd_MaxExchangeRate, gs_formatExchangeRate)
   Exit Sub
End If

If PrevWeekEmpty Then: Exit Sub

If CboWeek > 1 And CboWeek <= 5 Then
    If Mperiod.Year = Tgl1.Year Then
        If Mperiod.Month <> Tgl1.Month Then LblErr = DisplayMsg("0069") & " (" & Format(Mperiod, "MMM YYYY") & ")": Exit Sub
    Else
        LblErr = DisplayMsg("0069") & " (" & Format(Mperiod, "MMM YYYY") & ")": Exit Sub
    End If
End If

If Mperiod.Year = Tgl2.Year Then
    If Mperiod.Month > Tgl2.Month Then LblErr = DisplayMsg("0070") & " (" & Format(Mperiod, "MMM YYYY") & ")": Tgl2 = tgl_sb: Exit Sub
    If Tgl2.Month - Mperiod.Month > 1 Then LblErr = DisplayMsg("0070") & " (" & Format(Mperiod, "MMM YYYY") & ")": Tgl2 = tgl_sb: Exit Sub
ElseIf Tgl2.Year - Mperiod.Year = 1 Then
    If Tgl2.Month + 12 - Mperiod.Month > 1 Then
        LblErr = DisplayMsg("0070") & " (" & Format(Mperiod, "MMM YYYY") & "," & Format(DateSerial(Mperiod.Year + 1, 1, 1), "MMM YYYY") & ")": Tgl2 = tgl_sb: Exit Sub
    End If
Else
     LblErr = DisplayMsg("0070") & " range": Tgl2 = tgl_sb: Exit Sub
End If

'tgl1_Change

Set RSCA = Db.Execute("Select Week_code from tax_exchangerate where exch_year='" & Mperiod.Year & "' and exch_month='" & Mperiod.Month & "' and end_date='" & LDay & "'")
If Not RSCA.EOF Then LblErr = DisplayMsg("0066"): Exit Sub

    RSCA.Close
    Set RSCA = Nothing

    strS = 0
    strD = 0
    If baru = False Then
        strS = grid.FindRow("S", 0, bteColSelect, False)
        strD = grid.FindRow("D", 0, bteColSelect, False, False)
        If strD > 0 Then CekD = True: Jawab = MsgBox("Do you really want to Delete this Record", vbInformation + vbYesNo + vbDefaultButton2, "Confirmation")
                
        If Jawab = vbYes Then DataGrid
        If strS > 0 And cek Then
            '## Cek apakah Awal Week lebih dari 1 dan Awal Bulan berikutnya
            If CboWeek > 1 Then
                If Mperiod.Month = 12 Then
                    If CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year + 1) <> "" Then
                        TglA = Left(Split(CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year + 1), "|")(0), 4) & "-" & Mid(Split(CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year + 1), "|")(0), 5, 2) & "-" & Right(Split(CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year + 1), "|")(0), 2)
                    Else
                        TglA = Format(DateSerial(Mperiod.Year + 1, Mperiod.Month + 1, 1), "yyyy-mm-dd")
                    End If
                Else
                    If CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year) <> "" Then
                        TglA = Left(Split(CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year), "|")(0), 4) & "-" & Mid(Split(CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year), "|")(0), 5, 2) & "-" & Right(Split(CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year), "|")(0), 2)
                    Else
                        TglA = Format(DateSerial(Mperiod.Year, Mperiod.Month + 1, 1), "yyyy-mm-dd")
                    End If
                End If
                TglAk = Left(Split(CariTglAkhir(Mperiod.Month, Mperiod.Year), "|")(1), 4) & "-" & Mid(Split(CariTglAkhir(Mperiod.Month, Mperiod.Year), "|")(1), 5, 2) & "-" & Right(Split(CariTglAkhir(Mperiod.Month, Mperiod.Year), "|")(1), 2)
            Else
                TglA = Left(Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(0), 4) & "-" & Mid(Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(0), 5, 2) & "-" & Right(Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(0), 2)
                
                    If Mperiod.Month = 1 Then
                        If CariTglAkhir(Mperiod.Month - 1, Mperiod.Year - 1) <> "" Then
                            TglAk = Left(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year - 1), "|")(1), 4) & "-" & Mid(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year - 1), "|")(1), 5, 2) & "-" & Right(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year - 1), "|")(1), 2)
                        Else
                            TglAk = Format(DateSerial(Mperiod.Year - 1, Mperiod.Month - 1, 1), "yyyy-mm-dd")
                        End If
                    Else
                        If CariTglAkhir(Mperiod.Month - 1, Mperiod.Year) <> "" Then
                            TglAk = Left(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year), "|")(1), 4) & "-" & Mid(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year), "|")(1), 5, 2) & "-" & Right(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year), "|")(1), 2)
                        Else
                            TglAk = Format(DateSerial(Mperiod.Year - 1, Mperiod.Month, 1), "yyyy-mm-dd")
                        End If
                    End If
                
            End If
            
             If Trim(CboWeek) > 5 Then
                If Tgl2 >= CDate(TglA) Then
                    LblErr = DisplayMsg("0071"): Exit Sub
                End If
            ElseIf Trim(CboWeek) = 1 Then
                If Tgl2 <= CDate(TglA) Then
                    LblErr = DisplayMsg("0071"): Exit Sub
                End If
            End If
            
            sql = "update tax_exchangerate set Start_date='" & Format(Tgl1, "YYYYMMDD") & "', End_date='" & Format(Tgl2, "YYYYMMDD") & "'," & _
                       " Last_Update=getdate(), Last_User='" & userLogin & "'" & _
                       " where Exch_year ='" & Mperiod.Year & "' and Exch_month='" & Format(Mperiod.Month, "00") & "' " & _
                       " and Week_code='" & CboWeek & "'"
            Db.Execute sql
            
            DataGrid
        
            If StaErr = False Then
                LblErr = DisplayMsg(1101)
                baru = True
                Browse
                Kosong
                Dim IK As Long
                For IK = 1 To grid.Rows - 1
                    grid.TextMatrix(IK, bteColSelect) = ""
                Next
                
            Else
                LblErr = DisplayMsg(1102)
            End If
        End If
        Dim PRec As Integer
        If CekD Then
            If Jawab = vbYes Then
                If StaErr = False Then
                    LblErr = DisplayMsg(1201)
                    baru = True
                Else
                    LblErr = DisplayMsg(1202)
                    If Trim$(StrWDel) <> "" Then
                        For ie = 0 To nErr - 1
                            PRec = grid.FindRow(Trim$(Split(StrWDel, ",")(ie)), 0, bteColCurr, False)
                            grid.TextMatrix(PRec, bteColSelect) = "D"
                        Next ie
                    End If
                End If
            ElseIf Jawab = vbNo Then
                LblErr = ""
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
        Dim SqlU As String, PosRec As Integer
        
        If cek Then
        
            Dim RE As Recordset
            Set RE = New Recordset
            RE.Open "Select * from Tax_Exchangerate where Currency_code='" & cbocurr & "' and Exch_year='" & Mperiod.Year & "' and exch_month='" & Format(Mperiod.Month, "00") & "'", Db, adOpenDynamic, adLockOptimistic
            If RE.EOF Then
                GoTo Tambah
            Else
                GoTo Tambah
            End If
            
            Exit Sub
Tambah:

'## Cek apakah Awal Week lebih dari 1 dan Awal Bulan berikutnya
If CboWeek > 1 Then
    If Mperiod.Month = 12 Then
        If CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year + 1) <> "" Then
            TglA = Left(Split(CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year + 1), "|")(0), 4) & "-" & Mid(Split(CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year + 1), "|")(0), 5, 2) & "-" & Right(Split(CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year + 1), "|")(0), 2)
        Else
            TglA = Format(DateSerial(Mperiod.Year + 1, Mperiod.Month + 1, 1), "yyyy-mm-dd")
        End If
    Else
        If CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year) <> "" Then
            TglA = Left(Split(CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year), "|")(0), 4) & "-" & Mid(Split(CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year), "|")(0), 5, 2) & "-" & Right(Split(CekAkhirBulan(Mperiod.Month + 1, Mperiod.Year), "|")(0), 2)
        Else
            TglA = Format(DateSerial(Mperiod.Year, Mperiod.Month + 1, 1), "yyyy-mm-dd")
        End If
    End If
    TglAk = Left(Split(CariTglAkhir(Mperiod.Month, Mperiod.Year), "|")(1), 4) & "-" & Mid(Split(CariTglAkhir(Mperiod.Month, Mperiod.Year), "|")(1), 5, 2) & "-" & Right(Split(CariTglAkhir(Mperiod.Month, Mperiod.Year), "|")(1), 2)
Else
    If CariTglAwal(Mperiod.Month, Mperiod.Year) <> "" Then
        TglA = Left(Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(0), 4) & "-" & Mid(Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(0), 5, 2) & "-" & Right(Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(0), 2)
        If Mperiod.Month = 1 Then
            If CariTglAkhir(Mperiod.Month - 1, Mperiod.Year - 1) <> "" Then
                TglAk = Left(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year - 1), "|")(1), 4) & "-" & Mid(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year - 1), "|")(1), 5, 2) & "-" & Right(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year - 1), "|")(1), 2)
            Else
                TglAk = Format(DateSerial(Mperiod.Year - 1, Mperiod.Month - 1, 1), "yyyy-mm-dd")
            End If
        Else
            If CariTglAkhir(Mperiod.Month - 1, Mperiod.Year) <> "" Then
                TglAk = Left(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year), "|")(1), 4) & "-" & Mid(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year), "|")(1), 5, 2) & "-" & Right(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year), "|")(1), 2)
            Else
                TglAk = Format(DateSerial(Mperiod.Year - 1, Mperiod.Month, 1), "yyyy-mm-dd")
            End If
        End If
    End If
End If

    If Trim(CboWeek) > 5 Then
        If Tgl2 >= CDate(TglA) Then
            LblErr = DisplayMsg("0071"): Exit Sub
        End If
    ElseIf Trim(CboWeek) = 1 And TglA <> "" Then
        If Tgl2 <= CDate(TglA) Then
            LblErr = DisplayMsg("0071"): Exit Sub
        End If
    
    End If
            SqlU = "insert into Tax_exchangeRate(Exch_Year, Exch_Month, Week_Code, Currency_Code, Tax_ExchangeRate, Start_Date, End_Date, Last_Update, Last_User) " & _
            "values ('" & Mperiod.Year & "','" & Format(Mperiod.Month, "00") & "','" & CboWeek & "','" & cbocurr & "','" & CDbl(TxtRate) & "','" & Format(Tgl1, "YYYYMMDD") & "','" & Format(Tgl2, "YYYYMMDD") & "', getdate(),'" & userLogin & "')"
            PosRec = grid.FindRow(Trim$(cbocurr), 0, bteColCurrCode, False)
            'Db.Execute SqlU
            If PosRec < 0 Then
                Db.Execute SqlU
                
                SqlU = "update tax_exchangerate set Start_date='" & Format(Tgl1, "YYYYMMDD") & "',End_date='" & Format(Tgl2, "YYYYMMDD") & "'," & _
                            " Last_Update=getdate(), Last_User='" & userLogin & "'" & _
                            " where Exch_year ='" & Mperiod.Year & "' and Exch_month='" & Format(Mperiod.Month, "00") & "' " & _
                            " and Week_code='" & CboWeek & "'"
                Db.Execute SqlU
                LblErr = DisplayMsg(1000)
                Kosong
                
            Else
                LblErr = DisplayMsg(1023)
            End If
        End If
        baru = True
        Browse
        SqlU = ""
    'End If

End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
Dim RS As Recordset, ir As Integer


On Error GoTo handler

CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
HakU = hakUpdate(Me.Name)
StrWDel = ""
Header
'Browse
CboWeek.clear
CboWeek.AddItem "1"
CboWeek.AddItem "2"
CboWeek.AddItem "3"
CboWeek.AddItem "4"
CboWeek.AddItem "5"

baru = True
Tgl1 = Now()

Tgl1 = Format(Now, "MM/dd/YYYY")
Tgl2 = Format(Now, "MM/dd/YYYY")
Mperiod = Format(Now(), "MMM YYYY")
Call up_FillCombo(cbocurr, "curr_cls")
dtDate = True
CboWeek.Text = 1
cbocurr.Text = ""
LblCurr.Caption = ""
Tgl1.Enabled = True
'# Munculkan Week terakhir  dari period tertentu
Dim RSWT As Recordset
Set RSWT = Db.Execute("select top 1 week_code from tax_exchangerate where exch_year='" & Mperiod.Year & "' and exch_month='" & Mperiod.Month & "' order by week_code desc")
If Not RSWT.EOF Then
    CboWeek = RSWT!week_code
Else
    CboWeek = 1
End If
RSWT.Close
Set RSWT = Nothing
LblErr.Caption = "Lday"
LDay = Format(DateSerial(Mperiod.Year, Mperiod.Month + 1, 0), "YYYY-MM-DD")

LblErr.Caption = "Browse"
Browse
MPeriod_Change
tgl_sb = Tgl2
Tgl_sb2 = Tgl1
LblErr.Caption = ""
Exit Sub
handler:
LblErr.Caption = LblErr.Caption & err.Description

End Sub

Sub Browse()
Dim RSTR As Recordset
sql = " select Exch_year EY,Exch_month EM,Week_code WC, Currency_code CC, Tax_Exchangerate TX,Start_date SD,End_Date ED " & _
      " from Tax_exchangerate where " & _
      " Exch_year='" & Mperiod.Year & "' and Exch_month='" & Format(Mperiod.Month, "00") & "' and Week_code='" & CboWeek & "'"
Set RSTR = New ADODB.Recordset
RSTR.Open sql, Db, adOpenKeyset, adLockOptimistic
i = 0
Header
While Not RSTR.EOF
        i = i + 1
        grid.AddItem ""
        grid.TextMatrix(i, bteColCurr) = uf_GetCurrencyDescription(RSTR!CC)
        grid.Cell(flexcpAlignment, i, bteColCurr) = flexAlignLeftCenter
        grid.TextMatrix(i, bteColExRate) = Format(RSTR!TX, gs_formatExchangeRate)
        grid.Cell(flexcpAlignment, i, bteColExRate) = flexAlignRightCenter
        grid.TextMatrix(i, bteColCurrCode) = RSTR!CC
        grid.TextMatrix(i, bteColWeekCode) = RSTR!wC
        grid.TextMatrix(i, bteColStartDate) = Mid(RSTR!sd, 5, 2) & "/" & Right(RSTR!sd, 2) & "/" & Left(RSTR!sd, 4)
        grid.TextMatrix(i, bteColEndDate) = Mid(RSTR!ed, 5, 2) & "/" & Right(RSTR!ed, 2) & "/" & Left(RSTR!ed, 4)
        grid.Cell(flexcpBackColor, i, bteColCurr, i, bteColExRate) = &HDFFFFF
        grid.Cell(flexcpBackColor, i, bteColSelect) = vbWhite
        
   RSTR.MoveNext
Wend
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrGrid As String
Dim AdaS As Boolean, brs As Integer, id As Integer
StrGrid = grid.Text
AdaS = False
brs = 0

If StrGrid = "S" Then
    For id = 1 To grid.Rows - 1
        If id <> Row Then grid.TextMatrix(id, bteColSelect) = ""
    Next id
    Tgl1 = Format(grid.TextMatrix(grid.Row, bteColStartDate), "DD MMM YYYY")
    Tgl2 = Format(grid.TextMatrix(grid.Row, bteColEndDate), "DD MMM YYYY")
    cbocurr = grid.TextMatrix(grid.Row, bteColCurrCode)
    TxtRate = Format(grid.TextMatrix(grid.Row, bteColExRate), gs_formatExchangeRate)
    baru = False
    LblErr = ""
ElseIf StrGrid = "D" Then
    For id = 1 To grid.Rows - 1
        'Jika ada S maka , hapus yg S
        If grid.TextMatrix(id, bteColSelect) = "S" Then grid.TextMatrix(id, bteColSelect) = "": Exit For
    Next id
    baru = False
    LblErr = ""
Else
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
cek = True
CboWeek = CboWeek
If CboWeek.MatchFound = False Then CboWeek.SetFocus: LblErr = DisplayMsg("0072"): cek = False: Exit Function
cbocurr = cbocurr
If cbocurr.MatchFound = False Then cbocurr.SetFocus: LblErr = DisplayMsg(1028): cek = False: Exit Function
If CDbl(TxtRate) = 0 Then TxtRate.SetFocus: LblErr = DisplayMsg("0073"): cek = False: Exit Function
End Function

Sub Kosong()
cbocurr = ""
TxtRate = Format(0, gs_formatExchangeRate)
End Sub

Sub DataGrid()
Dim kode As String, Sta As String
Dim strSQL As String
Dim ix As Integer

On Error Resume Next
Dim PosS As Integer
PosS = grid.FindRow("S", 0, bteColSelect, False)
If PosS > 0 Then
    kode = Trim$(grid.TextMatrix(PosS, bteColCurr))
    strSQL = "update tax_exchangerate set Tax_exchangerate='" & CDbl(TxtRate) & "'," & _
                    " Last_Update=getdate(), Last_User='" & userLogin & "'" & _
                    " where Exch_year ='" & Mperiod.Year & "' and Exch_month='" & Format(Mperiod.Month, "00") & "' " & _
                    " and Week_code='" & CboWeek & "' and currency_code='" & cbocurr & "' and start_date='" & Format(grid.TextMatrix(PosS, bteColStartDate), "YYYYMMDD") & "'"
    Db.Execute (strSQL)
    
    strSQL = "update tax_exchangerate set Start_date='" & Format(Tgl1, "YYYYMMDD") & "',End_date='" & Format(Tgl2, "YYYYMMDD") & "'," & _
                    " Last_Update=getdate(), Last_User='" & userLogin & "'" & _
                    " where Exch_year ='" & Mperiod.Year & "' and Exch_month='" & Format(Mperiod.Month, "00") & "' " & _
                    " and Week_code='" & CboWeek & "'"
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
'DbX.BeginTrans
For ix = 1 To grid.Rows - 1
kode = Trim$(grid.TextMatrix(ix, bteColCurr))
Sta = Trim$(grid.TextMatrix(ix, bteColSelect))
    If Sta = "D" Then
        strSQL = "delete from Tax_exchangeRate  " & _
        " where Exch_year ='" & Mperiod.Year & "' and Exch_month='" & Format(Mperiod.Month, "00") & "' " & _
        " and  Week_code='" & CboWeek & "' and currency_code='" & grid.TextMatrix(ix, bteColCurrCode) & "' and start_date='" & Right(grid.TextMatrix(ix, bteColStartDate), 4) & Left(grid.TextMatrix(ix, bteColStartDate), 2) & Mid(grid.TextMatrix(ix, bteColStartDate), 4, 2) & "'"

        If strSQL <> "" Then Db.Execute strSQL

        If err.number <> 0 Then
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
Sta = ""
strSQL = ""
Browse
'Exit Sub
'errx:
'StaErr = True
End Sub

Private Sub MPeriod_Change()
'Mperiod.value = Format(Mperiod.value, "MMM yyyy")
MPeriod_Click
TglMP = Mperiod.Month


If CariTglAwal(Mperiod.Month, Mperiod.Year) = "" Then
    Tgl1 = Format(Mperiod.Year & "-" & Format(Mperiod.Month, "00") & "-01", "YYYY-MM-DD")
    CboWeek.Text = 1
    Tgl2 = Tgl1
Else
    Tgl1.Day = Right(Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(0), 2)
    Tgl1.Month = Mid(Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(0), 5, 2)
    Tgl1.Year = Left(Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(0), 4)
    
    Tgl2.Day = Right(Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(1), 2)
    Tgl2.Month = Mid(Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(1), 5, 2)
    Tgl2.Year = Left(Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(1), 4)
    CboWeek.Text = Split(CariTglAwal(Mperiod.Month, Mperiod.Year), "|")(2)
End If

If CariTglAwal(Mperiod.Month + 1, Mperiod.Year) = "" Then
    LDay = Format(DateSerial(Mperiod.Year, Mperiod.Month + 1, 0), "YYYY-MM-DD")
    StLDay = False
Else
    LDay = Left(Split(CariTglAwal(Mperiod.Month + 1, Mperiod.Year), "|")(0), 4) & "-" & Mid(Split(CariTglAwal(Mperiod.Month + 1, Mperiod.Year), "|")(0), 5, 2) & "-" & Right(Split(CariTglAwal(Mperiod.Month + 1, Mperiod.Year), "|")(0), 2)
    StLDay = True
End If

If CariTglAkhir(Mperiod.Month - 1, Mperiod.Year) = "" Then
    ADay = Format(DateSerial(Mperiod.Year, Mperiod.Month - 1, 0), "YYYY-MM-DD")
Else
    ADay = Left(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year), "|")(1), 4) & "-" & Mid(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year), "|")(1), 5, 2) & "-" & Right(Split(CariTglAkhir(Mperiod.Month - 1, Mperiod.Year), "|")(1), 2)
End If

tgl1_Click

Browse
End Sub

Private Sub MPeriod_Click()
LDay = Format(DateSerial(Mperiod.Year, Mperiod.Month + 1, 0), "YYYY-MM-DD")
If Mperiod.Month = 1 And Val(TglMP) = 12 Then Mperiod.Year = Mperiod.Year + 1
If Mperiod.Month = 12 And Val(TglMP) = 1 Then Mperiod.Year = Mperiod.Year - 1
Browse
End Sub

Private Sub tgl1_Change()
LblErr = ""

If CboWeek > 1 And CboWeek <= 5 Then
    If Mperiod.Year = Tgl1.Year Then
        If Mperiod.Month <> Tgl1.Month Then LblErr = DisplayMsg("0069") & " (" & Format(Mperiod, "MMM YYYY") & ")": Exit Sub
    Else
        LblErr = DisplayMsg("0069") & " (" & Format(Mperiod, "MMM YYYY") & ")": Exit Sub
    End If
ElseIf CboWeek = 1 Then
    If Mperiod.Month - Tgl1.Month >= 2 Then
        LblErr = DisplayMsg("0069") & " ( " & Format(DateSerial(Mperiod.Year, Mperiod.Month, 0), "MMM") & "," & Format(DateSerial(Mperiod.Year, Mperiod.Month, 1), "MMM") & " )"
        Exit Sub
    ElseIf Mperiod - Tgl1.Month < 0 Then
        LblErr = DisplayMsg("0069") & " ( " & Format(DateSerial(Mperiod.Year, Mperiod.Month, 0), "MMM") & "," & Format(DateSerial(Mperiod.Year, Mperiod.Month, 1), "MMM") & " )"
        Exit Sub
    Else
        If Format(Tgl1, "YYYYMMDD") <= Format(ADay, "YYYYMMDD") Then
            LblErr = DisplayMsg("0075") & " ( " & Format(CDate(ADay) + 1, "dd MMM YYYY") & ")"
            Tgl1 = Format(CDate(ADay) + 1, "YYYY-MM-DD")
            Exit Sub
        End If
    End If
End If
    
If Tgl1 > Tgl2 Then
   LblErr = DisplayMsg(4068)
   Tgl2 = Tgl1
   Exit Sub
End If
Tgl_sb2 = Tgl1
End Sub

Private Sub tgl1_Click()
tgl1_Change
End Sub

Private Sub Tgl2_Change()
LblErr = ""
If Mperiod.Year = Tgl2.Year Then
    If Mperiod.Month > Tgl2.Month Then LblErr = DisplayMsg("0070") & " (" & Format(Mperiod, "MMM YYYY") & ")": Tgl2 = tgl_sb: Exit Sub
    If Tgl2.Month - Mperiod.Month > 1 Then LblErr = DisplayMsg("8054") & " (" & Format(DateAdd("m", 1, Mperiod), "MMM YYYY") & ")": Tgl2 = tgl_sb: Exit Sub
ElseIf Tgl2.Year - Mperiod.Year = 1 Then
    If Tgl2.Month + 12 - Mperiod.Month > 1 Then
        LblErr = DisplayMsg("0070") & " (" & Format(Mperiod, "MMM YYYY") & "," & Format(DateSerial(Mperiod.Year + 1, 1, 1), "MMM YYYY") & ")": Tgl2 = tgl_sb: Exit Sub
    End If
Else
     LblErr = DisplayMsg("0070") & " range":
     If tgl_sb <> "" Then Tgl2 = tgl_sb
     Exit Sub
End If
If Tgl1 > Tgl2 Then
   If tgl_sb <> "" Then Tgl2 = tgl_sb
   LblErr = DisplayMsg(4066)
   Exit Sub
End If

Dim RsCS As Recordset
Set RsCS = Db.Execute("Select Week_code,Start_date,End_date from tax_exchangerate where exch_year='" & Mperiod.Year & "' and exch_month='" & Format(Mperiod.Month, "00") & "' and start_date='" & Format(Tgl2, "YYYYMMDD") & "'")
If Not RsCS.EOF Then
    LblErr = DisplayMsg("0074")
    If Format(Tgl2, "YYYYMMDD") <= Format(ADay, "YYYYMMDD") Then
            LblErr = DisplayMsg("0076") & " ( " & Format(CDate(ADay) + 1, "dd MMM YYYY") & ")"
            Tgl2 = Format(CDate(ADay) + 1, "YYYY-MM-DD")
            Exit Sub
    End If
    Exit Sub
Else
    Dim TglSesudah As String
    If CboWeek = 5 Then
        TglSesudah = DateAdd("m", "1", Mperiod)
        Set RsCS = Db.Execute("Select Week_code,Start_date,End_date from tax_exchangerate where exch_year='" & Year(TglSesudah) & "' and exch_month='" & Format(Month(TglSesudah), "00") & "' and week_code = '1'")
    Else
        Set RsCS = Db.Execute("Select Week_code,Start_date,End_date from tax_exchangerate where exch_year='" & Mperiod.Year & "' and exch_month='" & Format(Mperiod.Month, "00") & "' and week_code > '" & CboWeek & "'")
    End If
    
    If Not RsCS.EOF Then
        If RsCS!Start_Date - Format(Tgl2, "YYYYMMDD") > 1 Or RsCS!Start_Date - Format(Tgl2, "YYYYMMDD") < 1 Then
            LblErr = DisplayMsg("0074")
            Tgl2 = Format(DateAdd("d", -1, Left(RsCS!Start_Date, 4) & "-" & Mid(RsCS!Start_Date, 5, 2) & "-" & Right(RsCS!Start_Date, 2)), "dd MMM yyyy")
            tgl_sb = Tgl2
            Exit Sub
        End If
    End If
End If
If Format(Tgl2, "YYYY-MM-DD") = LDay Then
    If StLDay = True Then
        If Format(Tgl2, "YYYYMMDD") <= Format(ADay, "YYYYMMDD") Then
            LblErr = DisplayMsg("0076") & " ( " & Format(CDate(ADay) + 1, "dd MMM YYYY") & ")"
            If tgl_sb <> "" Then Tgl2 = tgl_sb
            Exit Sub
        End If
        LblErr = DisplayMsg("0074"): If tgl_sb <> "" Then Tgl2 = tgl_sb
        Exit Sub
    End If
End If
tgl_sb = Tgl2

End Sub

Private Sub Tgl2_Click()
Tgl2_Change
End Sub

Private Sub TxtRate_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
End If
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtRate_LostFocus()
On Error GoTo errhndl

If IsNumeric(TxtRate) = False Then TxtRate = Format(0, gs_formatExchangeRate)
If Round(CDbl(TxtRate)) / CDbl(TxtRate) = 1 Then
    TxtRate = Format(CDbl(TxtRate), gs_formatExchangeRate)
Else
    TxtRate = Format(CDbl(TxtRate), gs_formatExchangeRate)
End If
Exit Sub
errhndl:
DoEvents
End Sub

Function CariTglAwal(bulan As String, tahun As String) As String
Dim RsCT As Recordset
   
Set RsCT = Db.Execute("select top 1 start_date,End_date,Week_code  from tax_exchangeRate where exch_year='" & tahun & "' and Exch_month='" & bulan & "' order by start_date asc")
If Not RsCT.EOF Then
    CariTglAwal = RsCT!Start_Date & "|" & RsCT!end_date & "|" & RsCT!week_code
Else
    CariTglAwal = "" 'tahun & Format(bulan, "00") & "01" & "|" & tahun & Format(bulan, "00") & "01" & "|" & "1"
End If
End Function
Function CariTglAkhir(bulan As String, tahun As String) As String
Dim RsCT As Recordset
   
Set RsCT = Db.Execute("select top 1 start_date,End_date,Week_code  from tax_exchangeRate where exch_year='" & tahun & "' and Exch_month='" & bulan & "' order by start_date desc")
If Not RsCT.EOF Then
    CariTglAkhir = RsCT!Start_Date & "|" & RsCT!end_date & "|" & RsCT!week_code
Else
    CariTglAkhir = tahun & Format(bulan, "00") & "01" & "|" & tahun & Format(bulan, "00") & "01" & "|" & "1"
End If
End Function
Function CekAkhirBulan(bulan As String, tahun As String) As String
Dim RsCB As Recordset
   
Set RsCB = Db.Execute("select top 1 start_date,End_date,Week_code  from tax_exchangeRate where exch_year='" & tahun & "' and Exch_month='" & bulan & "' order by start_date Asc")
If Not RsCB.EOF Then
    CekAkhirBulan = RsCB!Start_Date & "|" & RsCB!end_date & "|" & RsCB!week_code
Else
    CekAkhirBulan = "" 'tahun & Format(bulan, "00") & "01" & "|" & tahun & Format(bulan, "00") & "01" & "|" & "1"
End If

End Function


