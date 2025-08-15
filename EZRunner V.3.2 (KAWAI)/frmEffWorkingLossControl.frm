VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEffWorkingLossControl 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Working Loss Time Control "
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEffWorkingLossControl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Preview"
      Height          =   405
      Left            =   7012
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   697
      TabIndex        =   6
      Top             =   2760
      Width           =   7500
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
         Left            =   3690
         TabIndex        =   7
         Top             =   210
         Width           =   75
      End
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   697
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1140
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   6345
      TabIndex        =   8
      Top             =   210
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin MSComCtl2.DTPicker Period 
      Height          =   330
      Left            =   1950
      TabIndex        =   2
      Top             =   2190
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   293797891
      UpDown          =   -1  'True
      CurrentDate     =   37860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Period"
      Height          =   195
      Index           =   2
      Left            =   690
      TabIndex        =   13
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Line Line8 
      Index           =   0
      X1              =   3765
      X2              =   8000
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Factory CD"
      Height          =   195
      Index           =   0
      Left            =   690
      TabIndex        =   12
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Line Line8 
      Index           =   1
      X1              =   3765
      X2              =   6015
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblLine 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   3765
      TabIndex        =   11
      Top             =   1815
      Width           =   2265
   End
   Begin MSForms.ComboBox cboLine 
      Height          =   315
      Left            =   1950
      TabIndex        =   1
      Top             =   1755
      Width           =   1695
      VariousPropertyBits=   746604571
      MaxLength       =   10
      DisplayStyle    =   3
      Size            =   "2990;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "AAA"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboFactory 
      Height          =   315
      Left            =   1950
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
      VariousPropertyBits=   746604571
      MaxLength       =   10
      DisplayStyle    =   3
      Size            =   "2990;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "AAAAAA"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Machine No"
      Height          =   195
      Index           =   1
      Left            =   690
      TabIndex        =   10
      Top             =   1785
      Width           =   1215
   End
   Begin VB.Label lblFactory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   3765
      TabIndex        =   9
      Top             =   1380
      Width           =   4530
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Working Loss Time Control"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   600
      Width           =   3120
   End
End
Attribute VB_Name = "frmEffWorkingLossControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temptgl As Byte

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

    Call isiCbo(cboFactory, "Trade_Master", "Trade_Code", "Trade_Name", 80, 200, "Trade_Code", , , " trade_code in (select distinct manufacture_code from manufacture_line) ")
    
    cboFactory.Text = "": lblFactory.Caption = ""
    CboLine.clear
    CboLine.Text = "": lblLine.Caption = ""
    Period = Format(Now, "yyyy-mm-01")
    temptgl = Period.Month
End Sub

Private Sub cboFactory_Change()
    If cboFactory.ListIndex <> -1 Then
        lblFactory.Caption = cboFactory.Column(1)
        Call isiCbo(CboLine, "Manufacture_Line", "Line_Code", "Line_Name", 80, 140, "Line_Code", , , " Manufacture_Code = '" & cboFactory & "' ", 1)
    Else
        If cboFactory.MatchFound = False Then lblFactory.Caption = "": CboLine.clear
    End If
End Sub

Private Sub cboFactory_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboFactory_Change
End Sub

Private Sub cboLine_Change()
    If CboLine.ListIndex <> -1 Then
        lblLine.Caption = CboLine.Column(1)
    Else
        If CboLine.MatchFound = False Then lblLine.Caption = ""
    End If
End Sub

Private Sub cboLine_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboLine_Change
End Sub

Private Sub period_Change()
    LblErrMsg.Caption = ""
    Call period_Click
    temptgl = Period.Month
End Sub

Private Sub period_Click()
    If Period.Month = 1 And Val(temptgl) = 12 Then Period.Year = Period.Year + 1
    If Period.Month = 12 And Val(temptgl) = 1 Then Period.Year = Period.Year - 1
End Sub

Private Sub CmdPreview_Click()
Dim RS As New ADODB.Recordset
Dim sqlFactory  As String, sqlLine As String
    
    LblErrMsg = ""
    
'    If hakUpdate(Me.Name) = 0 Then _
'    LblErrMsg = DisplayMsg(1040): Me.MousePointer = vbDefault: Exit Sub

    'VALIDATION
    If cboFactory.Text = "" Then
        LblErrMsg = DisplayMsg(1060)    '"Please Select Factory Code"
        cboFactory.SetFocus
        Exit Sub
    ElseIf cboFactory.Text <> "" Then
        If cboFactory.MatchFound = False Then
            LblErrMsg = DisplayMsg(4016)    '"Record with this Factory Code not found"
            cboFactory.SetFocus
            Exit Sub
        End If
    End If
    If CboLine.Text = "" Then
        LblErrMsg = DisplayMsg(1041)    '"Please Input Line Code"
        CboLine.SetFocus
        Exit Sub
    ElseIf CboLine.Text <> "" Then
        If CboLine.MatchFound = False Then
            LblErrMsg = DisplayMsg(4017)    '"Record with this Line Code not found"
            CboLine.SetFocus
            Exit Sub
        End If
    End If
    
    Call toGraph
'    Call toExcel
    
End Sub

Sub toGraph()
Dim sqlFactory As String, sqlLine As String, sqlFactory1 As String, sqlline1 As String
Dim rsGraph As New ADODB.Recordset

    sqlFactory = "": sqlLine = ""
    If Trim(cboFactory) <> "" Then _
        sqlFactory = " and pr.Supplier_Code = '" & Trim(cboFactory) & "' "
    If Trim(CboLine) <> "" And UCase(Trim(CboLine)) <> "ALL" Then _
        sqlLine = " and pr.po_no = '" & Trim(CboLine) & "' "
        
    sql = "Select * from " & _
        "( " & _
            "select top 7 wtd.WorkingLossTime_Cls, wtc.Description, sum(wtd.Loss_Time) Loss_Time " & _
                "from WorkingTime_Detail wtd " & _
                    "inner join (select * from Part_Receipt where receipt_Cls = 'P1') pr on pr.seq_no = wtd.ProductionSeq_No " & _
                    "left outer join WorkingLossTime_Cls wtc on wtc.WorkingLossTime_Cls = wtd.WorkingLossTime_Cls " & _
                "where year(pr.Receipt_Date) = '" & Year(Period) & "' and month(pr.Receipt_Date) = '" & Month(Period) & "' " & sqlFactory & sqlLine & _
                "group by wtd.WorkingLossTime_Cls, wtc.Description "
                
    sql = sql & _
            "Union " & _
            "select 'ZZ' WorkingLossTime_Cls, 'Others' Description, sum(wtd.Loss_Time) Loss_Time  " & _
            "from WorkingTime_Detail wtd " & _
                "inner join (select * from Part_Receipt where receipt_Cls = 'P1') pr on pr.seq_no = wtd.ProductionSeq_No " & _
                "left outer join WorkingLossTime_Cls wtc on wtc.WorkingLossTime_Cls = wtd.WorkingLossTime_Cls " & _
            "where wtd.WorkingLossTime_Cls not in " & _
                "(select distinct top 7 wtd.WorkingLossTime_Cls from WorkingTime_Detail wtd " & _
                "inner join " & _
                    "(select * from Part_Receipt where receipt_Cls = 'P1') pr on pr.seq_no = wtd.ProductionSeq_No " & _
                    "where year(pr.Receipt_Date) = '" & Year(Period) & "' and month(pr.Receipt_Date) = '" & Month(Period) & "' " & sqlFactory & sqlLine & _
                    "order by wtd.WorkingLossTime_Cls) " & _
                "and year(pr.Receipt_Date) = '" & Year(Period) & "' and month(pr.Receipt_Date) = '" & Month(Period) & "' " & sqlFactory & sqlLine & _
            "Having sum(wtd.Loss_Time) is not null  " & _
        ") dt " & _
        "order by (Case When dt.WorkingLossTime_Cls = 'zz' Then -1 Else Loss_Time End) desc"

    If rsGraph.State = adStateOpen Then rsGraph.Close
    rsGraph.CursorLocation = adUseClient
    
    rsGraph.Open sql, Db, adOpenDynamic, adLockOptimistic
    If rsGraph.EOF Then LblErrMsg = DisplayMsg(4006): Exit Sub
    
    With frmEffGraph8
        .lblJudul(0) = lblJudul
        .lblJudul(1) = "Machine No : " & CboLine & " (" & CboLine.Column(1) & ")"
        .lblJudul(2) = "Period : " & Format(Period, "MMM yyyy")
        .lblJudul(3) = ""
        .visibleSign = False
        .jmlBar = rsGraph.RecordCount
        
        i = 0
        Do While Not rsGraph.EOF
            .lblBarVal(i) = Format(rsGraph!Loss_Time, gs_formatWorkingTime)
            If i = 0 Then
                .lblBarVal(i).Tag = rsGraph!Loss_Time
            Else
                .lblBarVal(i).Tag = CDbl(.lblBarVal(i - 1).Tag) + rsGraph!Loss_Time
            End If
            .lblX(i * 4).Tag = Trim(rsGraph!Description)
            .MaxQty = .lblBarVal(i).Tag
            i = i + 1
            rsGraph.MoveNext
        Loop
        
        .lblN = "N = " & Format(.MaxQty, gs_formatNSample)
        .viewGraph
        .Show 1
    End With
    
    If rsGraph.EOF Then If rsGraph.State = adStateOpen Then rsGraph.Close
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblErrMsg.Caption = ErrMsg
    End If
End Sub

Private Sub CmdSubMenu_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

