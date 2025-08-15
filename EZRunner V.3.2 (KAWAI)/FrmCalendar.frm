VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCalendar 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar Master"
   ClientHeight    =   6975
   ClientLeft      =   765
   ClientTop       =   1800
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCalendar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   5775
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   315
      Width           =   1845
      _extentx        =   3254
      _extenty        =   767
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1125
      Left            =   341
      TabIndex        =   11
      Top             =   1065
      Width           =   7275
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Calendar Cls"
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   675
         Width           =   1755
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   285
         Width           =   4185
      End
      Begin MSComCtl2.DTPicker MYDate 
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   660
         Width           =   1305
         _ExtentX        =   2302
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
         CustomFormat    =   "MMM yyyy"
         Format          =   150929411
         UpDown          =   -1  'True
         CurrentDate     =   37802
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory Code"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   1140
      End
      Begin MSForms.ComboBox cbodealer 
         Height          =   315
         Left            =   1470
         TabIndex        =   0
         Top             =   240
         Width           =   1305
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2302;556"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2880
         X2              =   7080
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month / Year"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Index           =   1
      Left            =   4091
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6045
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   341
      TabIndex        =   9
      Top             =   5370
      Width           =   7275
      Begin VB.Label LblerrMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LblErrMsg"
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
         TabIndex        =   10
         Top             =   195
         Width           =   7050
      End
   End
   Begin VB.CommandButton CmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
      Height          =   375
      Index           =   3
      Left            =   5291
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6045
      Width           =   1125
   End
   Begin VB.CommandButton CmdAction 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Index           =   0
      Left            =   341
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6045
      Width           =   1125
   End
   Begin VB.CommandButton CmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   4
      Left            =   6491
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6045
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   3045
      Left            =   345
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2295
      Width           =   7275
      _cx             =   12832
      _cy             =   5371
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Calendar Master"
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
      Left            =   345
      TabIndex        =   8
      Top             =   345
      Width           =   7275
   End
End
Attribute VB_Name = "FrmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstcust As Recordset, rstcal As Recordset, rsCek As Recordset
Dim i As Integer, Firstday As String, lastday As Byte, z As Integer
Dim edited As Boolean, HakU As Integer
Dim firstflag As Boolean, tempx As Byte, thn_sb As String * 4, tgl_sb As Byte

Dim bteColSun As Byte
Dim bteColMon As Byte
Dim bteColTue As Byte
Dim bteColWed As Byte
Dim bteColThu As Byte
Dim bteColFri As Byte
Dim bteColSat As Byte
Dim bteColId As Byte

Sub Header()
    
    bteColSun = 0
    bteColMon = 1
    bteColTue = 2
    bteColWed = 3
    bteColThu = 4
    bteColFri = 5
    bteColSat = 6
    bteColId = 7
    
    With grid
        .ColS = 8
        .Rows = 1
        
        .TextMatrix(0, bteColSun) = "Sunday"
        .TextMatrix(0, bteColMon) = "Monday"
        .TextMatrix(0, bteColTue) = "Tuesday"
        .TextMatrix(0, bteColWed) = "Wednesday"
        .TextMatrix(0, bteColThu) = "Thursday"
        .TextMatrix(0, bteColFri) = "Friday"
        .TextMatrix(0, bteColSat) = "Saturday"
        .TextMatrix(0, bteColId) = "x"
        
        .ColWidth(bteColSun) = 1000
        .ColWidth(bteColMon) = 1000
        .ColWidth(bteColTue) = 1000
        .ColWidth(bteColWed) = 1050
        .ColWidth(bteColThu) = 1000
        .ColWidth(bteColFri) = 1000
        .ColWidth(bteColSat) = 1000
        
        .ColHidden(bteColId) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, 6) = flexAlignCenterCenter
        .EditMaxLength = 1
    End With
End Sub

Sub adtocombo()
sql = "select *, trade_code from trade_master where trade_code in (select distinct manufacture_code from manufacture_line)"
Set rstcust = New Recordset
rstcust.Open sql, Db, adOpenKeyset, adLockOptimistic
With cbodealer
    .clear
    .columnCount = 2
    .ColumnWidths = "50 pt;280 pt"
    .ListWidth = 350
    .ListRows = 15
i = 0
Do Until rstcust.EOF
    .AddItem ""
    .List(i, 0) = Trim(rstcust!Trade_Code)
    .List(i, 1) = Trim(rstcust!trade_name)
    i = i + 1
    rstcust.MoveNext
Loop
End With
End Sub

Private Sub cbodealer_Change()
LblErrMsg = ""
cbodealer = Trim(cbodealer)
If Not cbodealer.MatchFound Then
    Text4 = ""
Else
    rstcust.Requery
    rstcust.Find "trade_code = '" & cbodealer.Text & "'"
    If Not rstcust.EOF Then Text4 = Trim(rstcust!trade_name): Browse
End If
End Sub

Private Sub cbodealer_Click()
cbodealer = Trim(cbodealer)
If cbodealer.MatchFound Then Browse
End Sub

Private Sub cbodealer_GotFocus()
If edited Then LblErrMsg = DisplayMsg(1049): Frame1.Enabled = False
End Sub

Private Sub Check1_Click()
LblErrMsg = ""
End Sub

Private Sub cmdAction_Click(Index As Integer)
Select Case Index
    Case 0
        If edited Then LblErrMsg = DisplayMsg(1049): Exit Sub
        frmMainMenu.Show
        Unload Me
    Case 1
        If edited Then LblErrMsg = DisplayMsg(1049): Exit Sub
        cbodealer.ListIndex = -1
        Text4 = ""
        LblErrMsg = ""
        Browse
        edited = False
    Case 3
        Frame1.Enabled = True
        edited = False
        Browse
        LblErrMsg = ""
    Case 4
        If Not checkfactory Then Exit Sub
        If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
        If edited Then
            updatedata
            edited = False
            Browse
        End If
        If Check1.DataChanged Then
            sql = "update calendar_master set cal_cls= '" & Check1.Value & "', Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                "where factory_code ='" & cbodealer & "' and month(cal_date) = '" & Month(MYDate) & "' and year(cal_date) ='" & Year(MYDate) & "'"
            Db.Execute (sql)
            LblErrMsg = DisplayMsg(1101)
        End If
        
End Select
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub Form_Activate()
  
cbodealer.SetFocus
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
LblErrMsg = ""
Header
adtocombo
MYDate = Format(Month(Now) & "/01/" & Year(Now), "MMM YYYY")
Browse
edited = False
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
HakU = hakUpdate(Me.Name)
Text4 = ""
tgl_sb = Month(Now)
thn_sb = Year(Now)
End Sub

Sub Browse()

Firstday = Weekday(MYDate) - 1
lastday = DateDiff("d", MYDate, DateAdd("m", 1, MYDate))
With grid
If InStr(1, (lastday + Firstday) / 7, ".") Then
    .Rows = ((lastday + Firstday) \ 7) + 2
Else
    .Rows = (lastday + Firstday) / 7 + 1
End If
firstflag = True
tempx = 1
For i = 1 To grid.Rows - 1
    For z = 0 To 6
        If tempx > lastday Then Exit For
        If firstflag Then z = Firstday
        .TextMatrix(i, z) = tempx
            sql = "select * from calendar_master where factory_code = '" & cbodealer & "' and cal_date = '" & Year(MYDate) & "-" + Format(Month(MYDate), "00") & "-" & grid.TextMatrix(i, z) & "' "
            Set rstcal = New Recordset
            rstcal.Open sql, Db, adOpenDynamic, adLockOptimistic
            If Not rstcal.EOF Then
                .Cell(flexcpBackColor, i, z) = vbRed
                If IsNull(rstcal!cal_cls) Or rstcal!cal_cls = 0 Then
                    Check1.Value = 0
                Else
                    Check1.Value = 1
                End If
            Else
                .Cell(flexcpBackColor, i, z) = &H80000018
            End If
        firstflag = False
        tempx = tempx + 1
    Next z
Next
End With

rstcal.Close
Set rstcal = Nothing
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then Cancel = 1
 rstcust.Close
 Set rstcust = Nothing
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 Cancel = True
End Sub

Private Sub grid_Click()
With grid
    If Val(.TextMatrix(.Row, .Col)) = 0 Then Exit Sub
    If Not checkfactory Then Exit Sub
    
    'cek date in receip or daily
    sql = "select * from " & _
        "(select link='b', receipt_date from part_receipt where receipt_cls='P1' and receipt_date = '" & Year(MYDate) & "-" & Format(Month(MYDate), "00") & "-" & Format(.TextMatrix(.Row, .Col), "00") & "'" & _
        " Union " & _
        "select link='b' , schedule_date from Daily_production where schedule_date ='" & Year(MYDate) & "-" & Format(Month(MYDate), "00") & "-" & Format(.TextMatrix(.Row, .Col), "00") & "' " & _
         ") xx"
         
    Set rsCek = New Recordset
    rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic
    If Not rsCek.EOF Then LblErrMsg = DisplayMsg("0038") & " transactions !": rsCek.Close: Set rsCek = Nothing: Exit Sub
    LblErrMsg = ""
    
    If .Cell(flexcpBackColor, .Row, .Col) = vbRed Then
        .Cell(flexcpBackColor, .Row, .Col) = &H80000018
    Else
        .Cell(flexcpBackColor, .Row, .Col) = vbRed
    End If

    edited = True
    
    rsCek.Close
    Set rsCek = Nothing
    
End With
End Sub

Private Sub MYDate_Change()
If edited Then LblErrMsg = DisplayMsg(1049): Exit Sub
MYDate = Format(Month(MYDate) & "/01/" & Year(MYDate), "MMM YYYY")
MYDate_Click
tgl_sb = MYDate.Month
thn_sb = MYDate.Year
Header
Browse
End Sub

Sub updatedata()
Dim TempDate As String
If HakU = 0 Then LblErrMsg = ""
tempx = 0
firstflag = True
For i = 1 To grid.Rows - 1
    For z = 0 To 6
        If tempx > lastday Then Exit For
        If firstflag Then z = Firstday
        If grid.Cell(flexcpBackColor, i, z) = vbRed Then
            TempDate = TempDate & "'" & grid.TextMatrix(i, z) & "',"
            sql = "select * from calendar_master where factory_code = '" & cbodealer & "' and cal_date = '" & Year(MYDate) & "-" + Format(Month(MYDate), "00") & "-" & grid.TextMatrix(i, z) & "' "
            Set rstcal = New Recordset
            rstcal.Open sql, Db, adOpenDynamic, adLockOptimistic
            If rstcal.EOF Then
                sql = "select * from " & _
                    "(select link='b', receipt_date from part_receipt where receipt_cls='P1' and receipt_date = '" & Year(MYDate) & "-" & Format(Month(MYDate), "00") & "-" & Format(grid.TextMatrix(i, z), "00") & "'" & _
                    " Union " & _
                    "select link='b' , schedule_date from Daily_production where schedule_date ='" & Year(MYDate) & "-" & Format(Month(MYDate), "00") & "-" & Format(grid.TextMatrix(i, z), "00") & "' " & _
                     ") xx"
                Set rsCek = New Recordset
                rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic
                If rsCek.EOF Then
                    sql = "insert into calendar_master (Factory_Code, Cal_Date, Cal_Cls, Last_Update, Last_User) " & _
                            "values ('" & cbodealer & "', '" & Year(MYDate) & "-" & Format(Month(MYDate), "00") & "-" & Format(grid.TextMatrix(i, z), "00") & "','" & Check1.Value & "', getdate(), '" & userLogin & "')"
                    Db.Execute sql
                End If
                rsCek.Close
                Set rsCek = Nothing
            End If
            rstcal.Close
            Set rstcal = Nothing
        End If
        
        firstflag = False
        tempx = tempx + 1
    Next z
Next

If TempDate = "" Then
    TempDate = "''"
Else
    TempDate = Left(TempDate, Len(TempDate) - 1)
End If

sql = " delete from calendar_master " & _
    "where factory_code = '" & cbodealer.Text & "' " & _
    "and day(cal_date) not in (" & TempDate & ") " & _
    "and cal_date not in " & _
    "(select distinct receipt_date from part_receipt " & _
    "    where receipt_cls='P1' " & _
    "    and year(receipt_date) = '" & Year(MYDate) & "' and month(receipt_date) ='" & Month(MYDate) & "' " & _
    "    and day(receipt_date) " & _
    "        not in (" & TempDate & ") " & _
    " Union " & _
    " select distinct schedule_date " & _
    " From Daily_production " & _
    "where year(schedule_date) = '" & Year(MYDate) & "' and month(schedule_date) ='" & Month(MYDate) & "' " & _
    "and day(schedule_date) " & _
    "        not in (" & TempDate & ") " & _
    ") " & _
    "and year(cal_date) = '" & Year(MYDate) & "' and month(cal_date) ='" & Month(MYDate) & "' "
Db.Execute sql
LblErrMsg = DisplayMsg(1101)
End Sub

Function checkfactory() As Boolean
checkfactory = True
If Trim(cbodealer.Text) = "" Then
    LblErrMsg = DisplayMsg(1060)
    checkfactory = False
    Exit Function
Else
    cbodealer = Trim(cbodealer)
    If cbodealer.MatchFound = False Then
        LblErrMsg = DisplayMsg(4060)
        checkfactory = False
        Exit Function
    End If
End If
End Function

Private Sub MYDate_Click()
If MYDate.Month = 1 And Val(tgl_sb) = 12 Then MYDate.Year = MYDate.Year + 1
If MYDate.Month = 12 And Val(tgl_sb) = 1 Then MYDate.Year = MYDate.Year - 1
End Sub

Private Sub MYDate_GotFocus()
If edited Then Frame1.Enabled = False
End Sub
