VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPurchasingPriceReport 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Purchasing Price Report"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   Icon            =   "FrmPurchasingPriceReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   525
      Left            =   390
      TabIndex        =   11
      Top             =   2880
      Width           =   8805
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
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   8535
      End
   End
   Begin VB.CommandButton cmdReport 
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
      Left            =   8055
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3630
      Width           =   1140
   End
   Begin VB.CommandButton cmdSubMenu 
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
      TabIndex        =   4
      Top             =   3630
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1755
      Left            =   390
      TabIndex        =   5
      Top             =   1050
      Width           =   8805
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Left            =   4710
         TabIndex        =   17
         Top             =   847
         Width           =   300
      End
      Begin VB.TextBox LblProduct 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
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
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   70
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   855
         Width           =   3570
      End
      Begin MSComCtl2.DTPicker Period 
         Height          =   315
         Left            =   2145
         TabIndex        =   2
         Top             =   1230
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   293601283
         UpDown          =   -1  'True
         CurrentDate     =   38716
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Finish Good Part Cls"
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
         Left            =   270
         TabIndex        =   15
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": "
         Height          =   195
         Index           =   2
         Left            =   2055
         TabIndex        =   14
         Top             =   480
         Width           =   135
      End
      Begin MSForms.ComboBox CboGroup 
         Height          =   315
         Left            =   2145
         TabIndex        =   0
         Tag             =   "Product Code"
         Top             =   450
         Width           =   2235
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "3942;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": "
         Height          =   195
         Index           =   0
         Left            =   2055
         TabIndex        =   10
         Top             =   870
         Width           =   135
      End
      Begin VB.Line Line1 
         X1              =   5160
         X2              =   8500
         Y1              =   1110
         Y2              =   1110
      End
      Begin MSForms.ComboBox CboProduct 
         Height          =   315
         Left            =   2145
         TabIndex        =   1
         Top             =   840
         Width           =   2520
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "4445;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
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
         Left            =   270
         TabIndex        =   9
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period "
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
         Left            =   270
         TabIndex        =   8
         Top             =   1260
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": "
         Height          =   195
         Index           =   1
         Left            =   2055
         TabIndex        =   7
         Top             =   1260
         Width           =   135
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   7350
      TabIndex        =   16
      Top             =   360
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchasing Price Report"
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
      Index           =   1
      Left            =   3435
      TabIndex        =   13
      Top             =   390
      Width           =   2700
   End
End
Attribute VB_Name = "FrmPurchasingPriceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dateUp As Date


Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = CboProduct.Text
 frm_BrowseItem.Show 1
 CboProduct.Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub cmdReport_Click()
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsRpt As New ADODB.Recordset
    Dim Rpt As New FrmRpt3
    Dim Kt1 As String, Kt2 As String, Kt3 As String
    
    LblErrMsg = ""
    If Trim(cboGroup.Text) = "" Then
        LblErrMsg = DisplayMsg(8057) '"Please Select Finish Good Part Cls"
        cboGroup.SetFocus
        Exit Sub
    End If
    
    cboGroup.Text = cboGroup.Text
    If cboGroup.MatchFound = False Then
        LblErrMsg = DisplayMsg(8057) ' "Record with This Finish Good Part Cls Not found"
        cboGroup.SetFocus
        Exit Sub
    End If
    
    If Trim(CboProduct.Text) = "" Then
        LblErrMsg = DisplayMsg(4003) '"Please Select Product Code"
        CboProduct.SetFocus
        Exit Sub
    End If
    
    CboProduct.Text = CboProduct.Text
    If CboProduct.MatchFound = False Then
        LblErrMsg = DisplayMsg(4003) ' "Record with This Product Code Not found"
        CboProduct.SetFocus
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    Call up_DropSQLFunctionALL
    Call up_CreateSQLFunctionALL
    sql = SqlSintaks
    Set rsRpt = Db.Execute(sql)

            
    If rsRpt.EOF Then
        LblErrMsg.Caption = DisplayMsg(4006)
    Else
        sqlprint = sql
        reportcode = "PurchasingPrice"
        printorient = 2
        
        Set report = application.OpenReport(App.path & "\Reports\RptPurchasingPrice.rpt")
        report.Database.Tables(1).SetDataSource rsRpt
        
        ginvno = "'" & ": " & cboGroup.Text & "'"
        Fbulan = "'" & ": " & CboProduct.Text & " / " & CboProduct.Column(1) & "'"
        Ftahun = "'" & ": " & Format(Period, "MMM yyyy") & " '"

        report.FormulaFields(1).Text = ginvno
        report.FormulaFields(2).Text = Fbulan
        report.FormulaFields(3).Text = Ftahun
        
        report.FormulaFields(4).Text = gi_decimalDigitQty
        report.FormulaFields(5).Text = gi_decimalDigitPrice
        report.FormulaFields(6).Text = gi_decimalDigitAmount
        
        Rpt.CRViewer1.ReportSource = report
        Rpt.CRViewer1.ViewReport
        Rpt.CRViewer1.Zoom 1
        
        Rpt.WindowState = 2
        Rpt.Show 1
    End If
    Set rsRpt = Nothing
        Call up_DropSQLFunctionALL
    Me.MousePointer = vbDefault
    
End Sub

Function SqlSintaks() As String
    Dim SqlS As String
    Dim SqlGroup As String, sqlProduct As String


    SqlS = "": SqlGroup = "": sqlProduct = ""

    If UCase(Trim(cboGroup.Text)) <> strAll Then _
        SqlGroup = " and IM.FinishGoodPart_Cls = '" & Trim(cboGroup.Column(0)) & "' "
    If UCase(Trim(CboProduct.Text)) <> strAll Then _
        sqlProduct = " and IM.Item_Code = '" & Trim(CboProduct.Column(0)) & "' "
    
    SqlS = "Select PR.Item_Code Product_Code, IM.Item_Name Product_Name, PR.Receipt_Date, PR.Supplier_Code, PR.Qty, " & _
            "(Select Description From Unit_Cls Where Unit_Cls = PR.Unit_Cls) Unit, " & _
            "(Select Description From Curr_Cls Where Curr_Cls = PR.Currency_Code) Currency, " & _
            "PR.Price, " & _
            "PR.Price *  dbo.UF_GetBookExchangeRate(Year(PR.Receipt_Date), Month(PR.Receipt_Date), PR.Currency_Code)  As RpPrice, "
    
    SqlS = SqlS + _
            "PR.Qty * " & _
            " (PR.Price *  dbo.UF_GetBookExchangeRate(Year(PR.Receipt_Date), Month(PR.Receipt_Date), PR.Currency_Code)) As Cost " & _
            "From Part_Receipt PR " & _
            "Left Join Item_Master IM On IM.Item_Code = PR.Item_Code " & _
            "Where PR.PO_NO <> '0' And PR.Receipt_Cls = 'R' " & _
            "And Year(PR.Receipt_Date) = '" & Year(Period) & "' And Month(PR.Receipt_Date) = '" & Month(Period) & "' "
    SqlS = SqlS + SqlGroup + sqlProduct
    SqlS = SqlS + _
                "Order By Product_Code, Receipt_Date, PR.Supplier_Code, Unit, Currency "
    SqlSintaks = SqlS

End Function

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    
    IsiCboGroup
    IsiCboProduct
    
    Period = Format(Now, "MMM yyyy")
    dateUp = Period.Value
End Sub

Sub IsiCboGroup()
    With cboGroup
        .clear
        .columnCount = 2
        .ColumnWidths = "0;200pt"
        .ListWidth = 200
        .ListRows = 3
        
        .AddItem
        .List(0, 0) = strAll
        .List(0, 1) = strAll
        .AddItem
        .List(1, 0) = "01"
        .List(1, 1) = "Finish Good"
        .AddItem
        .List(2, 0) = "02"
        .List(2, 1) = "Parts/WIP/Material"
        
    End With
End Sub

Sub IsiCboProduct()
    Dim RsProduct As Recordset
    Dim sqlProduct As String
    
    sqlProduct = "Select Item_Code, Item_Name From Item_Master where use_endday >= convert(char(8), getdate(), 112) "
    
    Set RsProduct = Db.Execute(sqlProduct)
    
    With CboProduct
        .clear
        .columnCount = 2
        .ColumnWidths = "130pt;200pt"
        .ListWidth = 330
        .ListRows = 15
        
        .AddItem
        .List(0, 0) = strAll
        .List(0, 1) = strAll
        
        i = 1
        Do While Not RsProduct.EOF
            .AddItem
            .List(i, 0) = Trim(RsProduct("Item_Code"))
            .List(i, 1) = Trim(RsProduct("Item_Name"))
                        
            RsProduct.MoveNext
            i = i + 1
        Loop
    End With
    Set RsProduct = Nothing
End Sub

Private Sub CboProduct_Change()
    LblErrMsg = ""
    lblProduct = ""
End Sub

Private Sub CboProduct_Click()
    LblErrMsg = ""
    If CboProduct.ListIndex <> -1 Then
        lblProduct = CboProduct.Column(1)
    End If
End Sub

Private Sub CboProduct_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call CboProduct_Click
End Sub

Private Sub CboProduct_LostFocus()
    Call CboProduct_Click
End Sub

Private Sub period_Change()
If Format(Period.Value, "MM") < Format(dateUp, "MM") And Val(Format(Period.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then _
            Period.Year = Period.Year + 1: GoTo pass
    If Format(Period.Value, "MM") > Format(dateUp, "MM") And Val(Format(Period.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then _
            Period.Year = Period.Year - 1
pass:
    dateUp = Format(Period.Value, "dd MMM yyyy")
End Sub


Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then Unload Me Else LblErrMsg.Caption = ErrMsg
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub
