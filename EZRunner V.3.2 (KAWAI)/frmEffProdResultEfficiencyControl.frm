VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEffProdResultEfficiencyControl 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Efficiency Control"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEffProdResultEfficiencyControl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Preview"
      Height          =   375
      Left            =   7305
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   480
      TabIndex        =   6
      Top             =   2670
      Width           =   7965
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
         Left            =   3930
         TabIndex        =   7
         Top             =   210
         Width           =   60
      End
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1140
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   6600
      TabIndex        =   8
      Top             =   240
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   330
      Left            =   2295
      TabIndex        =   2
      Top             =   2220
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
      Left            =   840
      TabIndex        =   13
      Top             =   2295
      Width           =   540
   End
   Begin VB.Line Line8 
      Index           =   0
      X1              =   4350
      X2              =   8150
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Factory CD"
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   12
      Top             =   1500
      Width           =   960
   End
   Begin VB.Line Line8 
      Index           =   1
      X1              =   4350
      X2              =   8150
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   1
      Left            =   4350
      TabIndex        =   11
      Top             =   1890
      Width           =   960
   End
   Begin MSForms.ComboBox cbo 
      Height          =   315
      Index           =   1
      Left            =   2295
      TabIndex        =   1
      Top             =   1830
      Width           =   1920
      VariousPropertyBits=   746604571
      MaxLength       =   10
      DisplayStyle    =   3
      Size            =   "3387;556"
      ListRows        =   15
      ShowDropButtonWhen=   2
      Value           =   "AAA"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cbo 
      Height          =   315
      Index           =   0
      Left            =   2295
      TabIndex        =   0
      Top             =   1440
      Width           =   1920
      VariousPropertyBits=   746604571
      MaxLength       =   10
      DisplayStyle    =   3
      Size            =   "3387;556"
      ListRows        =   15
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
      Left            =   840
      TabIndex        =   10
      Top             =   1890
      Width           =   975
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   0
      Left            =   4350
      TabIndex        =   9
      Top             =   1500
      Width           =   960
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Efficiency Control"
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
      Left            =   3435
      TabIndex        =   5
      Top             =   780
      Width           =   2055
   End
End
Attribute VB_Name = "frmEffProdResultEfficiencyControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TampungDt As Byte, kondisi As String
Dim file As String

Dim rsGraph As New ADODB.Recordset

'********************** Initial *********************
Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    kondisi = "trade_code in (select distinct manufacture_code from manufacture_line) "
    Call isiCbo(Cbo(0), "Trade_Master", "Trade_Code", "Trade_Name", 80, 200, "Trade_Code", , , kondisi)
    Cbo(1).clear
    
    Cbo(0) = "": LblDesc(0) = ""
    Cbo(1) = "": LblDesc(1) = ""
    dt = Format(Now, "yyyy-MM-01")
    TampungDt = Month(dt)
    
    file = App.path & "\Excel\ProdResultEfficiencyControl.xls"
End Sub

'****************************************************

'******************* Tampil Data ********************
Function chkSave() As Boolean

chkSave = False

'    If hakU = 0 Then LblErrMsg = DisplayMsg(3008): Exit Function 'You don't have an access for Update
   
    If Trim(Cbo(0)) = "" Then
        LblErrMsg = DisplayMsg(8007) 'Please Input Factore Code
        Cbo(0).SetFocus: Exit Function
    ElseIf Cbo(0).MatchFound = False Then
        LblErrMsg = DisplayMsg(8008) 'Record with This Factory Code not found
        Cbo(0).SetFocus: Exit Function
    ElseIf Trim(Cbo(1)) = "" Then
        LblErrMsg = DisplayMsg(1041) 'Please Input Line Code
        Cbo(1).SetFocus: Exit Function
    ElseIf Cbo(1).MatchFound = False Then
        LblErrMsg = DisplayMsg(8010) 'Record with This Line Code not found
        Cbo(1).SetFocus: Exit Function
    End If
chkSave = True
End Function

Private Sub cmdReport_Click()
    If chkSave Then Call isiChart
End Sub

Sub isiChart()
Dim strSQL As String
    
    LblErrMsg = ""
    
    If Cbo(1) <> "All" Then strSQL = ",dt.Po_No "
    sql = "Select dt.Supplier_Code, dt.Receipt_Date " & strSQL & _
            ", WkTime = case when isnull(sum(dt.TotalWorking_Time),0) = 0 then 0 else (isnull(sum(dt.Working_Time),0) / isnull(sum(dt.TotalWorking_Time),0)) * 100 End, " & _
            "LossTime = case when isnull(sum(dt.TotalWorking_Time),0) = 0 then 0 else (isnull(sum(dt.TotalLoss_Time),0) / isnull(sum(dt.TotalWorking_Time),0)) * 100 End, " & _
            "CountTime = case when isnull(sum(dt.TotalWorking_Time),0) = 0 then 0 else (isnull(sum(dt.CountTime),0) / isnull(sum(dt.TotalWorking_Time),0)) * 100 End " & _
        "From " & _
            "( " & _
                "Select R.Supplier_Code, R.Po_No, R.Receipt_Date, " & _
                    "WM.Working_Time , WM.TotalLoss_Time, WM.TotalWorking_Time, WM.CountTime " & _
                "from Part_Receipt R " & _
                    "Inner Join " & _
                        "(Select ProductionSeq_No, WM.Working_Time, WM.TotalLoss_Time, WM.TotalWorking_Time, " & _
                            "CountTime = (Select CountWD = Count(ProductionSeq_No) from WorkingTime_Detail WD " & _
                                        "Where WD.ProductionSeq_No = WM.ProductionSeq_No) " & _
                        "from WorkingTime_Master WM " & _
                        ") WM on WM.ProductionSeq_No =  R.Seq_No " & _
                "Where year(r.Receipt_Date) = " & Year(dt) & _
                    " And Month(R.Receipt_Date) = " & Month(dt) & _
                    " And R.ProductionResult_Cls = 1 " & _
            ")dt " & _
        "Where dt.Supplier_Code = '" & Cbo(0) & "' "

    If Cbo(1) <> "All" Then sql = sql & " And dt.PO_No= '" & Cbo(1) & "' "
    
    sql = sql & _
        "group by dt.Supplier_Code, dt.Receipt_Date " & strSQL & _
        " Order by dt.Receipt_Date"
    
    If rsGraph.State = adStateOpen Then rsGraph.Close
    rsGraph.CursorLocation = adUseClient
    
    rsGraph.Open sql, Db, adOpenDynamic, adLockOptimistic
    
    If rsGraph.EOF Then
        LblErrMsg = DisplayMsg(8012)
    Else
        Call toGraph
    End If
    If rsGraph.EOF Then If rsGraph.State = adStateOpen Then rsGraph.Close
    Set rsGraph = Nothing
End Sub

Sub toGraph()
    With frmEffGraph22
        .lblJudul(0) = lblJudul
        .lblJudul(1) = "Machine No : " & Cbo(1) & " (" & Cbo(1).Column(1) & ")"
        .lblJudul(2) = "Period : " & Format(dt, "MMM yyyy")
        .jmlBar = rsGraph.RecordCount: .MaxQty = 0
        
        i = 0
        .MaxQty = 0: .MaxQty2 = 0: .MaxQty3 = 0
        Do While Not rsGraph.EOF
            If i >= 22 Then Exit Do
            .lblBarVal1(i) = Format(rsGraph!WkTime, gs_formatWorkingTime)
            .lblBarVal2(i) = Format(rsGraph!LossTime, gs_formatWorkingTime)
            .lblBarVal3(i) = Format(rsGraph!CountTime, gs_formatWorkingTime)
            .lblX(i) = Day(rsGraph!Receipt_Date)
            If rsGraph!WkTime > .MaxQty Then .MaxQty = rsGraph!WkTime
            If rsGraph!LossTime > .MaxQty2 Then .MaxQty2 = rsGraph!LossTime
            If rsGraph!CountTime > .MaxQty3 Then .MaxQty3 = rsGraph!CountTime
            i = i + 1
            rsGraph.MoveNext
        Loop
        
        .viewGraph
        .Show 1
    End With
End Sub

'****************************************************

'********************** Validate ********************
Private Sub cbo_Change(Index As Integer)
    Cbo(Index) = Cbo(Index)
    If Cbo(Index).MatchFound Then
        LblDesc(Index) = Cbo(Index).Column(1)
        If Index = 0 Then
            kondisi = "Manufacture_Code = '" & Cbo(0) & "'"
            Call isiCbo(Cbo(1), "Manufacture_Line", "Line_Code", "Line_Name", 80, 200, "Line_Code", , , kondisi, 1)
        End If
    Else
        If Index = 0 Then Cbo(1).clear
        LblDesc(Index) = ""
    End If
End Sub

Private Sub dt_change()
    Call dt_Click
    TampungDt = dt.Month
End Sub

Private Sub dt_Click()
    If dt.Month = 1 And Val(TampungDt) = 12 Then dt.Year = dt.Year + 1
    If dt.Month = 12 And Val(TampungDt) = 1 Then dt.Year = dt.Year - 1
End Sub

'****************************************************

'********************** Unload **********************
Private Sub CmdSubMenu_Click()
    Call DeleteFile(file)
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then Unload Me Else LblErrMsg.Caption = ErrMsg
End Sub
'****************************************************
