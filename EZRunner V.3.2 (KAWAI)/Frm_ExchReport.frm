VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_ExchReport 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Book Keeping Exchange Rate"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "Frm_ExchReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_submit 
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
      Left            =   6079
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3180
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   379
      TabIndex        =   5
      Top             =   2520
      Width           =   6825
      Begin VB.Label LblErr 
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
         Height          =   225
         Left            =   105
         TabIndex        =   6
         Top             =   195
         Width           =   6600
      End
   End
   Begin VB.CommandButton command3 
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
      Left            =   379
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3180
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker Tgl1 
      Height          =   315
      Left            =   5415
      TabIndex        =   1
      Top             =   1800
      Width           =   945
      _ExtentX        =   1667
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
      CustomFormat    =   "yyyy"
      Format          =   149159939
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   5370
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   300
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Line Line1 
      X1              =   3705
      X2              =   4695
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label LblCurr 
      BackStyle       =   0  'Transparent
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
      Left            =   3705
      TabIndex        =   10
      Top             =   1830
      Width           =   915
   End
   Begin MSForms.ComboBox CboCurr 
      Height          =   315
      Left            =   2655
      TabIndex        =   0
      Top             =   1800
      Width           =   945
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "1667;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Currency Code"
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
      Left            =   1230
      TabIndex        =   9
      Top             =   1860
      Width           =   1305
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Report Book Keeping Exchange Rate"
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
      Height          =   390
      Left            =   390
      TabIndex        =   8
      Top             =   825
      Width           =   6825
   End
   Begin VB.Label LblCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Left            =   4875
      TabIndex        =   7
      Top             =   1860
      Width           =   390
   End
End
Attribute VB_Name = "Frm_ExchReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs1 As Recordset, rs2 As Recordset
Dim sql1 As String, sql2 As String
Dim Idx As Byte

Private Sub cbocurr_Change()
LblCurr = ""
cbocurr = cbocurr
If cbocurr.MatchFound Then LblCurr = Trim(cbocurr.List(cbocurr.ListIndex, 1))
End Sub

Private Sub cbocurr_Click()
LblCurr = ""
cbocurr = cbocurr
If cbocurr.MatchFound Then LblCurr = Trim(cbocurr.List(cbocurr.ListIndex, 1))
End Sub

Private Sub Cmd_Submit_Click()
    Dim AppRpt As New CRAXDDRT.application
    Dim Rpt2 As New CRAXDDRT.report
    'u Formula Title
    Dim Judul1 As String, CurrBook As String, CurrTax As String
    
    'U/ Formula di Book_exchrate
    Dim Jan1 As String, Feb1 As String, Mar1 As String, Apr1 As String, Mei1 As String, Jun1 As String, Jul1 As String, Aug1 As String, Sep1 As String, Okt1 As String, Nov1 As String, Des1 As String
    Dim Jan2 As String, Feb2 As String, Mar2 As String, Apr2 As String, Mei2 As String, Jun2 As String, Jul2 As String, Aug2 As String, Sep2 As String, Okt2 As String, Nov2 As String, Des2 As String
    Dim Begin1(11) As String, Ending1(11) As String
    Dim week1(11) As String, week2(11) As String, week3(11) As String, week4(11) As String, week5(11) As String
    
    'u/ di Tax_exchrate
    Dim Jan3 As String, Feb3 As String, Mar3 As String, Apr3 As String, Mei3 As String, Jun3 As String, Jul3 As String, Aug3 As String, Sep3 As String, Okt3 As String, Nov3 As String, Des3 As String
    Dim Jan4 As String, Feb4 As String, Mar4 As String, Apr4 As String, Mei4 As String, Jun4 As String, Jul4 As String, Aug4 As String, Sep4 As String, Okt4 As String, Nov4 As String, Des4 As String
    Dim Jan5 As String, Feb5 As String, Mar5 As String, Apr5 As String, Mei5 As String, Jun5 As String, Jul5 As String, Aug5 As String, Sep5 As String, Okt5 As String, Nov5 As String, Des5 As String
    Dim Jan6 As String, Feb6 As String, Mar6 As String, Apr6 As String, Mei6 As String, Jun6 As String, Jul6 As String, Aug6 As String, Sep6 As String, Okt6 As String, Nov6 As String, Des6 As String
    Dim Jan7 As String, Feb7 As String, Mar7 As String, Apr7 As String, Mei7 As String, Jun7 As String, Jul7 As String, Aug7 As String, Sep7 As String, Okt7 As String, Nov7 As String, Des7 As String
    
    cbocurr = cbocurr
    If cbocurr.MatchFound Then
        CurrBook = Trim(cbocurr.List(cbocurr.ListIndex, 1))
        CurrTax = Trim(cbocurr.List(cbocurr.ListIndex, 1))
    Else
        LblErr = DisplayMsg(1011)
        Exit Sub
    End If
    Judul1 = "Data Base of Exchange Rate " & vbCrLf & "Book Keeping in Monthly Base"
    MousePointer = vbHourglass
    Set Rpt2 = AppRpt.OpenReport(App.path & "\Reports\Rpt_Exchange.rpt")
        
    Fbulan = cbocurr
    Ftahun = Tgl1.Year
    reportcode = "bookkeepingexchrate"
        
    Set rs1 = New Recordset
    
    cbocurr = cbocurr
    If cbocurr.MatchFound = False Then LblErr = DisplayMsg(1028): MousePointer = vbDefault: Exit Sub
    
    sql1 = "select * from book_exchangerate  " & _
         " where Exch_year='" & Tgl1.Year & "' and Currency_code='" & cbocurr & "'  order by Currency_code,Term_cls"
    
    rs1.CursorLocation = adUseClient
    rs1.Open sql1, Db, adOpenDynamic, adLockBatchOptimistic
    MousePointer = vbHourglass
    Jan1 = 0: Feb1 = 0: Mar1 = 0: Apr1 = 0: Mei1 = 0: Jun1 = 0: Jul1 = 0: Aug1 = 0: Sep1 = 0: Okt1 = 0: Nov1 = 0: Des1 = 0
    Jan2 = 0: Feb2 = 0: Mar2 = 0: Apr2 = 0: Mei2 = 0: Jun2 = 0: Jul2 = 0: Aug2 = 0: Sep2 = 0: Okt2 = 0: Nov2 = 0: Des2 = 0
    For Idx = 0 To 11
        Begin1(Idx) = 0
        Ending1(Idx) = 0
    Next
    If Not rs1.EOF Then
        While Not rs1.EOF
            If Trim$(rs1!Term_cls) = "1" Then 'Beginning
                If IsNull(rs1!Exch01) Then
                    Jan1 = "-   "
                Else
                    Jan1 = Format(rs1!Exch01, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch02) Then
                    Feb1 = "-   "
                Else
                    Feb1 = Format(rs1!Exch02, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch03) Then
                    Mar1 = "-   "
                Else
                    Mar1 = Format(rs1!Exch03, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch04) Then
                    Apr1 = "-   "
                Else
                    Apr1 = Format(rs1!Exch04, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch05) Then
                    Mei1 = "-   "
                Else
                    Mei1 = Format(rs1!Exch05, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch06) Then
                    Jun1 = "-   "
                Else
                    Jun1 = Format(rs1!Exch06, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch07) Then
                    Jul1 = "-   "
                Else
                    Jul1 = Format(rs1!Exch07, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch08) Then
                    Aug1 = "-   "
                Else
                    Aug1 = Format(rs1!Exch08, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch09) Then
                    Sep1 = "-   "
                Else
                    Sep1 = Format(rs1!Exch09, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch010) Then
                    Okt1 = "-   "
                Else
                    Okt1 = Format(rs1!Exch010, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch011) Then
                    Nov1 = "-   "
                Else
                    Nov1 = Format(rs1!Exch011, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch012) Then
                    Des1 = "-   "
                Else
                    Des1 = Format(rs1!Exch012, gs_formatExchangeRate)
                End If
            ElseIf Trim$(rs1!Term_cls) = "2" Then 'Beginning
                If IsNull(rs1!Exch01) Then
                    Jan2 = "-   "
                Else
                    Jan2 = Format(rs1!Exch01, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch02) Then
                    Feb2 = "-   "
                Else
                    Feb2 = Format(rs1!Exch02, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch03) Then
                    Mar2 = "-   "
                Else
                    Mar2 = Format(rs1!Exch03, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch04) Then
                    Apr2 = "-   "
                Else
                    Apr2 = Format(rs1!Exch04, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch05) Then
                    Mei2 = "-   "
                Else
                    Mei2 = Format(rs1!Exch05, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch06) Then
                    Jun2 = "-   "
                Else
                    Jun2 = Format(rs1!Exch06, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch07) Then
                    Jul2 = "-   "
                Else
                    Jul2 = Format(rs1!Exch07, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch08) Then
                    Aug2 = "-   "
                Else
                    Aug2 = Format(rs1!Exch08, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch09) Then
                    Sep2 = "-   "
                Else
                    Sep2 = Format(rs1!Exch09, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch010) Then
                    Okt2 = "-   "
                Else
                    Okt2 = Format(rs1!Exch010, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch011) Then
                    Nov2 = "-   "
                Else
                    Nov2 = Format(rs1!Exch011, gs_formatExchangeRate)
                End If
                If IsNull(rs1!Exch012) Then
                    Des2 = "-   "
                Else
                    Des2 = Format(rs1!Exch012, gs_formatExchangeRate)
                End If
            End If
NEXTRECORD:
        rs1.MoveNext
        Wend
    End If
    
    sql2 = "Select * from Tax_exchangerate where Exch_year='" & Tgl1.Year & "' and Currency_code='" & cbocurr & "' order by Exch_Month,Week_code,Currency_code"
    
    Set rs2 = Db.Execute(sql2)
    
    Jan3 = "-   ": Feb3 = "-   ": Mar3 = "-   ": Apr3 = "-   ": Mei3 = "-   ": Jun3 = "-   ": Jul3 = "-   ": Aug3 = "-   ": Sep3 = "-   ": Okt3 = "-   ": Nov3 = "-   ": Des3 = "-   "
    Jan4 = "-   ": Feb4 = "-   ": Mar4 = "-   ": Apr4 = "-   ": Mei4 = "-   ": Jun4 = "-   ": Jul4 = "-   ": Aug4 = "-   ": Sep4 = "-   ": Okt4 = "-   ": Nov4 = "-   ": Des4 = "-   "
    Jan5 = "-   ": Feb5 = "-   ": Mar5 = "-   ": Apr5 = "-   ": Mei5 = "-   ": Jun5 = "-   ": Jul5 = "-   ": Aug5 = "-   ": Sep5 = "-   ": Okt5 = "-   ": Nov5 = "-   ": Des5 = "-   "
    Jan6 = "-   ": Feb6 = "-   ": Mar6 = "-   ": Apr6 = "-   ": Mei6 = "-   ": Jun6 = "-   ": Jul6 = "-   ": Aug6 = "-   ": Sep6 = "-   ": Okt6 = "-   ": Nov6 = "-   ": Des6 = "-   "
    Jan7 = "-   ": Feb7 = "-   ": Mar7 = "-   ": Apr7 = "-   ": Mei7 = "-   ": Jun7 = "-   ": Jul7 = "-   ": Aug7 = "-   ": Sep7 = "-   ": Okt7 = "-   ": Nov7 = "-   ": Des7 = "-   "
    
    For Idx = 0 To 11
        week1(Idx) = "-   "
        week2(Idx) = "-   "
        week3(Idx) = "-   "
        week4(Idx) = "-   "
        week5(Idx) = "-   "
    Next
    
    While Not rs2.EOF
        If rs2!week_code = "1" Then
            If rs2!exch_month = "1" Then Jan3 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "2" Then Feb3 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "3" Then Mar3 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "4" Then Apr3 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "5" Then Mei3 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "6" Then Jun3 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "7" Then Jul3 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "8" Then Aug3 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "9" Then Sep3 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "10" Then Okt3 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "11" Then Nov3 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "12" Then Des3 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
        ElseIf rs2!week_code = "2" Then
            If rs2!exch_month = "1" Then Jan4 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "2" Then Feb4 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "3" Then Mar4 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "4" Then Apr4 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "5" Then Mei4 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "6" Then Jun4 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "7" Then Jul4 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "8" Then Aug4 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "9" Then Sep4 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "10" Then Okt4 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "11" Then Nov4 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "12" Then Des4 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
        ElseIf rs2!week_code = "3" Then
            If rs2!exch_month = "1" Then Jan5 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "2" Then Feb5 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "3" Then Mar5 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "4" Then Apr5 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "5" Then Mei5 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "6" Then Jun5 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "7" Then Jul5 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "8" Then Aug5 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "9" Then Sep5 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "10" Then Okt5 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "11" Then Nov5 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "12" Then Des5 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
        ElseIf rs2!week_code = "4" Then
            If rs2!exch_month = "1" Then Jan6 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "2" Then Feb6 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "3" Then Mar6 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "4" Then Apr6 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "5" Then Mei6 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "6" Then Jun6 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "7" Then Jul6 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "8" Then Aug6 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "9" Then Sep6 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "10" Then Okt6 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "11" Then Nov6 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "12" Then Des6 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
        ElseIf rs2!week_code = "5" Then
            If rs2!exch_month = "1" Then Jan7 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "2" Then Feb7 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "3" Then Mar7 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "4" Then Apr7 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "5" Then Mei7 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "6" Then Jun7 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "7" Then Jul7 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "8" Then Aug7 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "9" Then Sep7 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "10" Then Okt7 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "11" Then Nov7 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
            If rs2!exch_month = "12" Then Des7 = Format(rs2!Tax_exchangerate, gs_formatExchangeRate)
        End If
nextrec:
    rs2.MoveNext
Wend
        
    Dim sqlcom As String
    Dim rscom As New Recordset
    sqlcom = "select company_name from company_profile"
    Set rscom = Db.Execute(sqlcom)
    
    Rpt2.Database.Tables(2).SetDataSource rscom
    
    LblErr = ""
    Rpt2.FormulaFields(2).Text = "'" & CurrBook & "'"
    Rpt2.FormulaFields(3).Text = "'" & CurrTax & "'"
    'beginning
    Rpt2.FormulaFields(4).Text = "'" & Jan1 & "'"
    Rpt2.FormulaFields(5).Text = "'" & Feb1 & "'"
    Rpt2.FormulaFields(6).Text = "'" & Mar1 & "'"
    Rpt2.FormulaFields(7).Text = "'" & Apr1 & "'"
    Rpt2.FormulaFields(8).Text = "'" & Mei1 & "'"
    Rpt2.FormulaFields(9).Text = "'" & Jun1 & "'"
    Rpt2.FormulaFields(10).Text = "'" & Jul1 & "'"
    Rpt2.FormulaFields(11).Text = "'" & Aug1 & "'"
    Rpt2.FormulaFields(12).Text = "'" & Sep1 & "'"
    Rpt2.FormulaFields(13).Text = "'" & Okt1 & "'"
    Rpt2.FormulaFields(14).Text = "'" & Nov1 & "'"
    Rpt2.FormulaFields(15).Text = "'" & Des1 & "'"
    'Ending
    Rpt2.FormulaFields(16).Text = "'" & Jan2 & "'"
    Rpt2.FormulaFields(17).Text = "'" & Feb2 & "'"
    Rpt2.FormulaFields(18).Text = "'" & Mar2 & "'"
    Rpt2.FormulaFields(19).Text = "'" & Apr2 & "'"
    Rpt2.FormulaFields(20).Text = "'" & Mei2 & "'"
    Rpt2.FormulaFields(21).Text = "'" & Jun2 & "'"
    Rpt2.FormulaFields(22).Text = "'" & Jul2 & "'"
    Rpt2.FormulaFields(23).Text = "'" & Aug2 & "'"
    Rpt2.FormulaFields(24).Text = "'" & Sep2 & "'"
    Rpt2.FormulaFields(25).Text = "'" & Okt2 & "'"
    Rpt2.FormulaFields(26).Text = "'" & Nov2 & "'"
    Rpt2.FormulaFields(27).Text = "'" & Des2 & "'"

    'Tax Rate
    'Week 1
    Rpt2.FormulaFields(28).Text = "'" & Jan3 & "'"
    Rpt2.FormulaFields(29).Text = "'" & Feb3 & "'"
    Rpt2.FormulaFields(30).Text = "'" & Mar3 & "'"
    Rpt2.FormulaFields(31).Text = "'" & Apr3 & "'"
    Rpt2.FormulaFields(32).Text = "'" & Mei3 & "'"
    Rpt2.FormulaFields(33).Text = "'" & Jun3 & "'"
    Rpt2.FormulaFields(34).Text = "'" & Jul3 & "'"
    Rpt2.FormulaFields(35).Text = "'" & Aug3 & "'"
    Rpt2.FormulaFields(36).Text = "'" & Sep3 & "'"
    Rpt2.FormulaFields(37).Text = "'" & Okt3 & "'"
    Rpt2.FormulaFields(38).Text = "'" & Nov3 & "'"
    Rpt2.FormulaFields(39).Text = "'" & Des3 & "'"
    'Week 2
    Rpt2.FormulaFields(40).Text = "'" & Jan4 & "'"
    Rpt2.FormulaFields(41).Text = "'" & Feb4 & "'"
    Rpt2.FormulaFields(42).Text = "'" & Mar4 & "'"
    Rpt2.FormulaFields(43).Text = "'" & Apr4 & "'"
    Rpt2.FormulaFields(44).Text = "'" & Mei4 & "'"
    Rpt2.FormulaFields(45).Text = "'" & Jun4 & "'"
    Rpt2.FormulaFields(46).Text = "'" & Jul4 & "'"
    Rpt2.FormulaFields(47).Text = "'" & Aug4 & "'"
    Rpt2.FormulaFields(48).Text = "'" & Sep4 & "'"
    Rpt2.FormulaFields(49).Text = "'" & Okt4 & "'"
    Rpt2.FormulaFields(50).Text = "'" & Nov4 & "'"
    Rpt2.FormulaFields(51).Text = "'" & Des4 & "'"
    'Week 3
    Rpt2.FormulaFields(52).Text = "'" & Jan5 & "'"
    Rpt2.FormulaFields(53).Text = "'" & Feb5 & "'"
    Rpt2.FormulaFields(54).Text = "'" & Mar5 & "'"
    Rpt2.FormulaFields(55).Text = "'" & Apr5 & "'"
    Rpt2.FormulaFields(56).Text = "'" & Mei5 & "'"
    Rpt2.FormulaFields(57).Text = "'" & Jun5 & "'"
    Rpt2.FormulaFields(58).Text = "'" & Jul5 & "'"
    Rpt2.FormulaFields(59).Text = "'" & Aug5 & "'"
    Rpt2.FormulaFields(60).Text = "'" & Sep5 & "'"
    Rpt2.FormulaFields(61).Text = "'" & Okt5 & "'"
    Rpt2.FormulaFields(62).Text = "'" & Nov5 & "'"
    Rpt2.FormulaFields(63).Text = "'" & Des5 & "'"
    'Week 4
    Rpt2.FormulaFields(64).Text = "'" & Jan6 & "'"
    Rpt2.FormulaFields(65).Text = "'" & Feb6 & "'"
    Rpt2.FormulaFields(66).Text = "'" & Mar6 & "'"
    Rpt2.FormulaFields(67).Text = "'" & Apr6 & "'"
    Rpt2.FormulaFields(68).Text = "'" & Mei6 & "'"
    Rpt2.FormulaFields(69).Text = "'" & Jun6 & "'"
    Rpt2.FormulaFields(70).Text = "'" & Jul6 & "'"
    Rpt2.FormulaFields(71).Text = "'" & Aug6 & "'"
    Rpt2.FormulaFields(72).Text = "'" & Sep6 & "'"
    Rpt2.FormulaFields(73).Text = "'" & Okt6 & "'"
    Rpt2.FormulaFields(74).Text = "'" & Nov6 & "'"
    Rpt2.FormulaFields(75).Text = "'" & Des6 & "'"
    'Week 5
    Rpt2.FormulaFields(76).Text = "'" & Jan7 & "'"
    Rpt2.FormulaFields(77).Text = "'" & Feb7 & "'"
    Rpt2.FormulaFields(78).Text = "'" & Mar7 & "'"
    Rpt2.FormulaFields(79).Text = "'" & Apr7 & "'"
    Rpt2.FormulaFields(80).Text = "'" & Mei7 & "'"
    Rpt2.FormulaFields(81).Text = "'" & Jun7 & "'"
    Rpt2.FormulaFields(82).Text = "'" & Jul7 & "'"
    Rpt2.FormulaFields(83).Text = "'" & Aug7 & "'"
    Rpt2.FormulaFields(84).Text = "'" & Sep7 & "'"
    Rpt2.FormulaFields(85).Text = "'" & Okt7 & "'"
    Rpt2.FormulaFields(86).Text = "'" & Nov7 & "'"
    Rpt2.FormulaFields(87).Text = "'" & Des7 & "'"
    Rpt2.FormulaFields(88).Text = "'" & Tgl1.Year & "'"
    
    Screen.MousePointer = 1
    
    With FrmRpt3
        .CRViewer1.ReportSource = Rpt2
        .CRViewer1.ViewReport
        .CRViewer1.Zoom 1
        .WindowState = 2
        .Show
    End With
    If rs1.State <> adStateClosed Then rs1.Close
    If rs2.State <> adStateClosed Then rs2.Close
MousePointer = vbDefault
End Sub

Private Sub command3_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErr.Caption = ErrMsg
End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
Tgl1 = Format(Now, "MMM YYYY")
Call up_FillCombo(cbocurr, "curr_cls")
With cbocurr
    .ListWidth = 80
    .ColumnWidths = "20 pt;60 pt"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub
