VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptDailySchedule 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Production Result Report"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRptDailySchedule.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1770
      Left            =   600
      TabIndex        =   10
      Top             =   1380
      Width           =   8895
      Begin MSComCtl2.DTPicker dtAwal 
         Height          =   330
         Left            =   1890
         TabIndex        =   2
         Top             =   1245
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   138149891
         CurrentDate     =   37860
      End
      Begin MSComCtl2.DTPicker dtAkhir 
         Height          =   330
         Left            =   4020
         TabIndex        =   3
         Top             =   1245
         Width           =   1665
         _ExtentX        =   2937
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
         Format          =   138149891
         CurrentDate     =   37799
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   3300
         X2              =   8550
         Y1              =   1065
         Y2              =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machine"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   16
         Top             =   825
         Width           =   690
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   3300
         TabIndex        =   15
         Top             =   825
         Width           =   960
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   1
         Left            =   1860
         TabIndex        =   1
         Top             =   765
         Width           =   1335
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2355;556"
         ListRows        =   15
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   1860
         TabIndex        =   0
         Top             =   330
         Width           =   1335
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2355;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   195
         Index           =   2
         Left            =   3660
         TabIndex        =   14
         Top             =   1320
         Width           =   165
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   3300
         TabIndex        =   13
         Top             =   390
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   12
         Top             =   390
         Width           =   630
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   3300
         X2              =   8550
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Production Date"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   11
         Top             =   1320
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   600
      TabIndex        =   8
      Top             =   3270
      Width           =   8895
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
         TabIndex        =   9
         Top             =   225
         Width           =   8985
      End
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   8310
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1185
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   7650
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   540
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Production Schedule Report"
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
      Left            =   3135
      TabIndex        =   7
      Top             =   540
      Width           =   3825
   End
End
Attribute VB_Name = "frmRptDailySchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim sql As String, i As Integer, sqldays As String, sqlholiday As String
Dim lastday As Integer, Tgl As Date
Dim bolrpt22 As Boolean

'******** Combo Factory Code **********
Sub isiCboCust() 'Isi Combo Dealer CD dr Customer Master
Dim RsCust As New ADODB.Recordset 'Data Customer

With cbo(0)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    '******** Ambil dr Customer Master utk Combo Dealer CD
    sql = "select distinct(manufacture_code), trade_name from manufacture_line inner join trade_master " & _
          "on trade_master.trade_code = manufacture_line.manufacture_code order by manufacture_code"
    Set RsCust = Db.Execute(sql)
    
    i = 0
    Do While Not (RsCust.EOF)
        .AddItem ""
        .List(i, 0) = Trim(RsCust(0))
        .List(i, 1) = Trim(RsCust(1))
        i = i + 1
        RsCust.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 350
    .ColumnWidths = "50 pt;300 pt"
    
    Set RsCust = Nothing
End With
End Sub

Sub isiCboLine()
Dim RsLine As New ADODB.Recordset

With cbo(1)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    sql = "select * from manufacture_line where manufacture_code='" & cbo(0) & "'"
    Set RsLine = Db.Execute(sql)
    
    i = 0
    Do While Not (RsLine.EOF)
        .AddItem ""
        .List(i, 0) = Trim(RsLine("Line_Code"))
        .List(i, 1) = Trim(RsLine("Line_Name"))
        i = i + 1
        RsLine.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 350
    .ColumnWidths = "50 pt;300 pt"
    
    Set RsLine = Nothing
End With
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    cbo(1) = ""
    lblNm(0) = ""
    lblNm(1) = ""
    dtAwal = Format(Now, "mmm 01 yyyy")
    dtAkhir = Now
    Call isiCboCust
End Sub

Private Sub cbo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbo_Click(Index)
End Sub

Public Sub cbo_Click(Index As Integer)
    cbo(Index) = cbo(Index)
    If cbo(Index).MatchFound Then
        lblNm(Index) = cbo(Index).Column(1)
        LblErrMsg = ""
        If Index = 0 Then isiCboLine
    Else
        lblNm(Index) = ""
        If Index = 0 Then LblErrMsg = DisplayMsg(4016)
        If Index = 1 Then LblErrMsg = DisplayMsg(4017)
    End If
End Sub

Private Sub cbo_Change(Index As Integer)
    lblNm(Index) = ""
    LblErrMsg = ""
End Sub

Private Sub dtAwal_Change()
    LblErrMsg = ""
    If Month(dtAwal) <> Month(dtAkhir) Or Year(dtAwal) <> Year(dtAkhir) Then _
        dtAkhir.Value = Format(DateAdd("d", -1, Format(DateAdd("m", 1, dtAwal), "yyyy-MM-01")), "dd MMM yyyy")
    If Format(dtAwal, "yyyy-MM-dd") > _
        Format(CDate(dtAkhir), "yyyy-MM-dd") Then _
    LblErrMsg = DisplayMsg(4068): Exit Sub
End Sub

Private Sub dtAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub dtAkhir_Change()
    LblErrMsg = ""
    If Month(dtAwal) <> Month(dtAkhir) Or Year(dtAwal) <> Year(dtAkhir) Then _
        dtAwal.Value = Format(dtAkhir, "yyyy-MM-01")
        
    If Format(dtAwal, "yyyy-MM-dd") > _
        Format(CDate(dtAkhir), "yyyy-MM-dd") Then _
        LblErrMsg = DisplayMsg(4066): Exit Sub
End Sub

Private Sub dtAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub cmdReport_Click()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3

    Me.MousePointer = vbHourglass
    
    If cbo(0) = "" Then
        LblErrMsg = DisplayMsg(1040)
        cbo(0).SetFocus
    Else
        cbo(0) = cbo(0)
        cbo(1) = cbo(1)
        If cbo(0).MatchFound = False Then
            LblErrMsg = DisplayMsg(4016)
            cbo(0).SetFocus
        ElseIf cbo(1).MatchFound = False Then
            LblErrMsg = DisplayMsg(4017)
            cbo(1).SetFocus
        Else
            LblErrMsg = ""

            sql = "select  distinct DP.Factory_Code, (select trade_name from trade_master where trade_code = '" & cbo(0) & "') factory_name, " & vbCrLf & _
                    " DP.Line_Code , rtrim(DP.Item_Code) Item_Code, rtrim(IM.makeritem_code) makeritem_code, rtrim(IM.Item_Name) item_name, " & vbCrLf & _
                    "    isnull((Select sum(qty) from daily_production   where item_code = dp.item_code and schedule_date  >='" & Format(dtAwal, "YYYY-MM-DD") & "' and schedule_date <='" & Format(dtAkhir, "YYYY-MM-DD") & "' and Factory_code ='" & cbo(0) & "' and line_code ='" & cbo(1) & "'),0) totalP," & vbCrLf & _
                    " (select isnull(sum(Qty),0) from part_receipt    where dailyseq_no in (select Seq_No from daily_Production where item_code =dp.item_code and Factory_code ='" & cbo(0) & "' and line_code ='" & cbo(1) & "')and productionResult_cls = '1' and receipt_date >='" & Format(dtAwal, "YYYY-MM-DD") & "' and receipt_Cls ='p1' and receipt_date<='" & Format(dtAkhir, "YYYY-MM-DD") & "') TotalR " & vbCrLf

            Call cekdatadaily
            
            sql = sql & sqldays & sqlholiday & " from    daily_production DP inner join item_master IM on IM.Item_Code = DP.Item_Code " & vbCrLf & _
                        " where   DP.Factory_Code = '" & cbo(0).Text & "' and " & vbCrLf & _
                        "DP.Line_Code = '" & cbo(1).Text & "' order by  rtrim(DP.Item_Code)  " & vbCrLf
                        
            If rsRpt.State <> adStateClosed Then rsRpt.Close
            rsRpt.CursorLocation = adUseClient
            Text1.Text = sql
            rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
            
            If rsRpt.EOF Then
                LblErrMsg.Caption = DisplayMsg(4006)
            Else
                sqlprint = sql
                
                printorient = 2
                If bolrpt22 Then
                    Set report = application.OpenReport(App.path & "\Reports\MatrixDaily22.rpt")
                    report.Database.Tables(1).SetDataSource rsRpt
                    tglAwalRptPrint = Format(dtAwal, "dd MMMM yyyy") & " to " & Format(dtAkhir, "dd MMMM yyyy")
                    report.FormulaFields(1).Text = "'" & Format(dtAwal, "dd MMMM yyyy") & " to " & Format(dtAkhir, "dd MMMM yyyy") & "'"
                    For i = 0 To 21
                        report.FormulaFields(i + 2).Text = "'" & xdays(i) & "'"
                    Next
                    reportcode = "MatrixDaily22"
                    
                    '#####################################################################
                    '# Qty Digit and decimal
                    report.FormulaFields(24).Text = "" & gi_decimalDigitQty & ""
                    report.FormulaFields(25).Text = "" & gi_decimalDigitQty & ""
                    '#####################################################################
                    
                Else
                    Set report = application.OpenReport(App.path & "\Reports\MatrixDaily.rpt")
                    report.Database.Tables(1).SetDataSource rsRpt
                    tglAwalRptPrint = Format(dtAwal, "dd MMMM yyyy") & " to " & Format(dtAkhir, "dd MMMM yyyy")
                    report.FormulaFields(1).Text = "'" & Format(dtAwal, "dd MMMM yyyy") & " to " & Format(dtAkhir, "dd MMMM yyyy") & "'"
                    For i = 0 To 30
                        report.FormulaFields(i + 2).Text = "'" & xdays(i) & "'"
                    Next
                    reportcode = "MatrixDaily"
                    '#####################################################################
                    '# Qty Digit and decimal
                    report.FormulaFields(33).Text = "" & gi_decimalDigitQty & ""
                    report.FormulaFields(34).Text = "" & gi_decimalDigitQty & ""
                    '#####################################################################
                End If
                
                Rpt.CRViewer1.ReportSource = report
                Rpt.CRViewer1.ViewReport
                Rpt.CRViewer1.Zoom 1
                
                Rpt.WindowState = 2
                Rpt.Show 1
            End If
            Set rsRpt = Nothing
        End If
    End If
    
    Me.MousePointer = vbDefault
    
End Sub

'************ Unload **********
Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
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
'**************

Sub cekdatadaily()
Dim rstcekdata As Recordset
Dim jumdata As Long, sqlx As String
Dim sqldummy As String, sqlP As String
Dim rscal As Recordset

Set rscal = New Recordset
sqldummy = "select top 1 cal_cls from calendar_master where month(cal_date) = '" & Month(dtAwal) & "' and year(cal_date) = '" & Year(dtAkhir) & "' and factory_code ='" & cbo(0) & "'"
rscal.CursorLocation = adUseClient
rscal.Open sqldummy, Db, adOpenKeyset, adLockOptimistic

If Not rscal.EOF Then
    If rscal!cal_cls = "0" Or IsNull(rscal!cal_cls) Then
        sqlP = " and tempday not in (select day(cal_date) from calendar_master where cal_date >= '" & Format(dtAwal, "YYYY-MM-DD") & "' and cal_date <= '" & Format(dtAkhir, "YYYY-MM-DD") & "') "
    Else
        sqlP = ""
    End If
Else
    sqlP = ""
End If

sqlx = "select tempday from tempday where tempday >= '" & Day(dtAwal) & "' and tempday <= '" & Day(dtAkhir) & "' " & _
        sqlP

Set rstcekdata = New Recordset
rstcekdata.CursorLocation = adUseClient
rstcekdata.Open sqlx, Db, adOpenKeyset, adLockOptimistic
bolrpt22 = False
If rstcekdata.RecordCount <= 22 Then
    bolrpt22 = True
    lastday = 22
    jumdata = rstcekdata.RecordCount
    For i = 0 To jumdata - 1
        Tgl = Year(dtAwal) & "-" & Month(dtAwal) & "-" & rstcekdata!tempday
        If i = 0 Then
            
            sqldays = " ,isnull(( " & _
              "         Select sum(qty)  " & _
              "         from daily_production  " & _
              "         where item_code = dp.item_code  " & _
              "         and schedule_date  ='" & Format(Tgl, "YYYY-MM-DD") & "' " & _
              "         and Factory_code ='" & cbo(0) & "' and line_code ='" & cbo(1) & "' " & _
              "     ),0) day" & i + 1 & ", " & _
              "     ( " & _
              "         select isnull(sum(Qty),0) from part_receipt  " & _
              "             where dailyseq_no in  " & _
              "             ( " & _
              "                 select Seq_No  " & _
              "                 from daily_Production " & _
              "                 where item_code =dp.item_code  " & _
              "                     and Factory_code ='" & cbo(0) & "' and line_code ='" & cbo(1) & "' " & _
              "             )  " & _
              "             and productionResult_cls = '1' and receipt_Cls ='p1' and receipt_date ='" & Format(Tgl, "YYYY-MM-DD") & "' " & _
              "     ) Result" & i + 1 & " "
            xdays(i) = Day(Tgl)
            sqlholiday = ", (select count(cal_date) Holiday from calendar_master " & _
                            " where factory_code ='" & cbo(0) & "' and cal_date ='" & Format(Tgl, "YYYY-MM-DD") & "') Holiday" & i + 1 & " "
        Else
            sqldays = sqldays & " ,isnull(( " & _
              "         Select sum(qty)  " & _
              "         from daily_production  " & _
              "         where item_code = dp.item_code  " & _
              "         and schedule_date  ='" & Format(Tgl, "YYYY-MM-DD") & "' " & _
              "         and Factory_code ='" & cbo(0) & "' and line_code ='" & cbo(1) & "' " & _
              "     ),0) day" & i + 1 & ", " & _
              "     ( " & _
              "         select isnull(sum(Qty),0) from part_receipt  " & _
              "             where dailyseq_no in  " & _
              "             ( " & _
              "                 select Seq_No  " & _
              "                 from daily_Production " & _
              "                 where item_code =dp.item_code  " & _
              "                     and Factory_code ='" & cbo(0) & "' and line_code ='" & cbo(1) & "' " & _
              "             )  " & _
              "             and productionResult_cls = '1' and receipt_Cls ='p1' and receipt_date ='" & Format(Tgl, "YYYY-MM-DD") & "' " & _
              "     ) Result" & i + 1 & " "
            xdays(i) = Day(Tgl)
            sqlholiday = sqlholiday & ", (select count(cal_date) Holiday from calendar_master " & _
                            " where factory_code ='" & cbo(0) & "' and cal_date ='" & Format(Tgl, "YYYY-MM-DD") & "') Holiday" & i + 1 & " "
        
        End If
            rstcekdata.MoveNext
    Next
    If jumdata < 22 Then
        For i = jumdata + 1 To 22
                Tgl = DateAdd("d", 1, Tgl)
                sqldays = sqldays & " , (Select 0) Day" & i & ", (Select 0) Result" & i & ""
                sqlholiday = sqlholiday & ", (Select 1) Holiday" & i & ""
            xdays(i - 1) = " "
        Next
    End If
Else

    For i = 0 To rstcekdata.RecordCount - 1
        lastday = 31
        jumdata = rstcekdata.RecordCount
        If i = 0 Then
            Tgl = Year(dtAwal) & "-" & Format(dtAwal, "MM") & "-" & rstcekdata!tempday
            sqldays = " ,isnull(( " & vbCrLf & _
              "         Select sum(qty)  " & vbCrLf & _
              "         from daily_production  " & vbCrLf & _
              "         where item_code = dp.item_code  " & vbCrLf & _
              "         and schedule_date  ='" & Format(Tgl, "YYYY-MM-DD") & "' " & vbCrLf & _
              "         and Factory_code ='" & cbo(0) & "' and line_code ='" & cbo(1) & "' " & vbCrLf & _
              "     ),0) day" & i + 1 & ", " & vbCrLf & _
              "     ( " & vbCrLf & _
              "         select isnull(sum(Qty),0) from part_receipt  " & vbCrLf & _
              "             where dailyseq_no in  " & vbCrLf & _
              "             ( " & vbCrLf & _
              "                 select Seq_No  " & vbCrLf & _
              "                 from daily_Production " & vbCrLf & _
              "                 where item_code =dp.item_code  " & vbCrLf & _
              "                     and Factory_code ='" & cbo(0) & "' and line_code ='" & cbo(1) & "' " & vbCrLf & _
              "             )  " & vbCrLf & _
              "             and productionResult_cls = '1' and receipt_Cls ='p1'  and receipt_date ='" & Format(Tgl, "YYYY-MM-DD") & "' " & vbCrLf & _
              "     ) Result" & i + 1 & " " & vbCrLf
              
            sqlholiday = ", (select count(cal_date) Holiday from calendar_master " & vbCrLf & _
                            " where factory_code ='" & cbo(0) & "' and cal_date ='" & Format(Tgl, "YYYY-MM-DD") & "') Holiday" & i + 1 & " " & vbCrLf
            xdays(i) = Day(Tgl)
        Else
            If i + 1 <= lastday Then
                Tgl = Year(dtAwal) & "-" & Format(dtAwal, "MM") & "-" & rstcekdata!tempday
                sqldays = sqldays & " ,isnull(( " & vbCrLf & _
                  "         Select sum(qty)  " & vbCrLf & _
                  "         from daily_production  " & vbCrLf & _
                  "         where item_code = dp.item_code  " & vbCrLf & _
                  "         and schedule_date  ='" & Format(Tgl, "YYYY-MM-DD") & "' " & vbCrLf & _
                  "         and Factory_code ='" & cbo(0) & "' and line_code ='" & cbo(1) & "' " & vbCrLf & _
                  "     ),0) day" & i + 1 & ", " & vbCrLf & _
                  "     ( " & vbCrLf & vbCrLf & _
                  "         select isnull(sum(Qty),0) from part_receipt  " & vbCrLf & _
                  "             where dailyseq_no in  " & vbCrLf & _
                  "             ( " & vbCrLf & _
                  "                 select Seq_No  " & vbCrLf & _
                  "                 from daily_Production " & vbCrLf & _
                  "                 where item_code =dp.item_code  " & vbCrLf & _
                  "                     and Factory_code ='" & cbo(0) & "' and line_code ='" & cbo(1) & "' " & vbCrLf & _
                  "             )  " & vbCrLf & _
                  "             and productionResult_cls = '1' and receipt_Cls ='p1' and receipt_date ='" & Format(Tgl, "YYYY-MM-DD") & "' " & vbCrLf & _
                  "     ) Result" & i + 1 & " " & vbCrLf
                  
            sqlholiday = sqlholiday & ", (select count(cal_date) Holiday from calendar_master " & vbCrLf & _
                            " where factory_code ='" & cbo(0) & "' and cal_date ='" & Format(Tgl, "YYYY-MM-DD") & "') Holiday" & i + 1 & " " & vbCrLf
            
            xdays(i) = Day(Tgl)
            End If
        End If
        rstcekdata.MoveNext
    Next
    If jumdata < 31 Then
        For i = jumdata + 1 To 31
                Tgl = DateAdd("d", 1, Tgl)
                sqldays = sqldays & " , (Select 0) Day" & i & ", (Select 0) Result" & i & ""
                sqlholiday = sqlholiday & ", (Select 0) Holiday" & i & ""
            xdays(i - 1) = " "
        Next
    End If
End If
'sqldays = sqldays & " , (Select 0.00) Day" & i + 1 & ", (Select 0.00) Result" & i + 1 & ""
'sqlholiday = sqlholiday & ", (Select 0) Holiday" & i + 1 & ""

End Sub


