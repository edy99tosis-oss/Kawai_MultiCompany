VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAPStatus 
   BackColor       =   &H00FDDFE3&
   Caption         =   "AP Status"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "FrmAPStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdexcel 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Excel"
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
      Index           =   3
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9720
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
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
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9720
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   2
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9720
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   0
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9750
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   840
      Left            =   390
      TabIndex        =   9
      Top             =   1125
      Width           =   14490
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
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
         Left            =   13020
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   255
         Width           =   1140
      End
      Begin VB.TextBox Text1 
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
         Height          =   225
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   4560
      End
      Begin MSComCtl2.DTPicker edate 
         Height          =   315
         Left            =   11325
         TabIndex        =   2
         Top             =   315
         Width           =   1605
         _ExtentX        =   2831
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   141230083
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker sdate 
         Height          =   315
         Left            =   9360
         TabIndex        =   1
         Top             =   315
         Width           =   1665
         _ExtentX        =   2937
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   141230083
         CurrentDate     =   37810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier CD"
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
         Left            =   135
         TabIndex        =   12
         Top             =   375
         Width           =   1035
      End
      Begin MSForms.ComboBox cbodealer 
         Height          =   315
         Left            =   1290
         TabIndex        =   0
         Top             =   315
         Width           =   1830
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3228;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         X1              =   3300
         X2              =   7775
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
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
         Left            =   8175
         TabIndex        =   11
         Top             =   375
         Width           =   1125
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   10995
         TabIndex        =   10
         Top             =   375
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   390
      TabIndex        =   7
      Top             =   9060
      Width           =   14475
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
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   165
         Width           =   14235
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13020
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   270
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6825
      Left            =   390
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2160
      Width           =   14475
      _cx             =   25532
      _cy             =   12039
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AP Status"
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
      TabIndex        =   13
      Top             =   300
      Width           =   14430
   End
End
Attribute VB_Name = "FrmAPStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstcust As Recordset
Dim rst As Recordset
Dim i As Long, Y As Integer, HakU As Integer

Dim bteColSuppCode As Byte
Dim bteColSuppName As Byte
Dim bteColInvoiceNo As Byte
Dim bteColInvoiceDate As Byte
Dim bteColDelDate As Byte
Dim bteColCurr As Byte
Dim btecolreateUSD As Byte
Dim bteColAmount As Byte
Dim bteColAmountIDR As Byte
Dim bteColBaseAmount As Byte

Dim bteColPPn As Byte
Dim bteColFreight As Byte
Dim bteColTotal As Byte
Dim bteColIssue As Byte
Dim bteColFix As Byte
Dim bteColStatus As Byte

Dim bteHakPrice As Byte

Sub Header()
With grid
    
    bteColSuppCode = 0
    bteColSuppName = 1
    bteColInvoiceNo = 2
    bteColInvoiceDate = 3
    bteColDelDate = 4
    bteColCurr = 5
    btecolreateUSD = 6
    bteColAmount = 7
    bteColAmountIDR = 8
    bteColBaseAmount = 9
    bteColPPn = 10
    bteColFreight = 11
    bteColTotal = 12
    bteColIssue = 13
    bteColFix = 14
    bteColStatus = 15
    
    .ColS = 16
    .Rows = 1
    
    .TextMatrix(0, bteColSuppCode) = "Supplier Code"
    .TextMatrix(0, bteColSuppName) = "Supplier Name"
    .TextMatrix(0, bteColInvoiceNo) = "Invoice No"
    .TextMatrix(0, bteColInvoiceDate) = "Invoice Date"
    .TextMatrix(0, bteColDelDate) = "B/L Date"
    .TextMatrix(0, bteColCurr) = "Curr"
    .TextMatrix(0, btecolreateUSD) = "Rate(USD)"
    .TextMatrix(0, bteColAmount) = "Amount"
    .TextMatrix(0, bteColAmountIDR) = "Amount (IDR)"
    .TextMatrix(0, bteColBaseAmount) = "Amount (USD)"
    .TextMatrix(0, bteColPPn) = "PPN"
    .TextMatrix(0, bteColFreight) = "Air Freight"
    .TextMatrix(0, bteColTotal) = "Total Amount"
    .TextMatrix(0, bteColIssue) = "Paid"
    .TextMatrix(0, bteColFix) = "Fix"
    .TextMatrix(0, bteColStatus) = "status"
    
    .ColWidth(bteColSuppCode) = 1500
    .ColWidth(bteColSuppName) = 3800
    .ColWidth(bteColInvoiceNo) = 1500
    .ColWidth(bteColInvoiceDate) = 1300
    .ColWidth(bteColDelDate) = 1500
    .ColWidth(btecolreateUSD) = 1500
    .ColWidth(bteColAmount) = 2200
    .ColWidth(bteColAmountIDR) = 2200
    .ColWidth(bteColBaseAmount) = 2200
    .ColWidth(bteColPPn) = 1800
    .ColWidth(bteColFreight) = 1800
    .ColWidth(bteColTotal) = 2200
    .ColWidth(bteColIssue) = 800
    .ColWidth(bteColFix) = 600
    .ColWidth(bteColStatus) = 600
    
    .ColHidden(bteColStatus) = True
    .ColHidden(bteColDelDate) = True
    .ColHidden(bteColFreight) = True

    .ColHidden(bteColAmount) = (bteHakPrice = 0)
    .ColHidden(bteColAmountIDR) = (bteHakPrice = 0)
    .ColHidden(bteColPPn) = True '(bteHakPrice = 0)
    .ColHidden(bteColFreight) = True '(bteHakPrice = 0)
    .ColHidden(bteColTotal) = True '(bteHakPrice = 0)
    
    .Cell(flexcpAlignment, 0, 0, 0, bteColFix) = flexAlignCenterCenter
    .EditMaxLength = 1
End With
End Sub

Private Sub cbodealer_Click()
Header
cbodealer = cbodealer
If cbodealer.MatchFound Then
    LblErrMsg = ""
    Text1 = cbodealer.Column(1)
Else
    Text1 = ""
    LblErrMsg = DisplayMsg(4050)
End If
End Sub

Private Sub cbodealer_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then cbodealer_Click
End Sub

Private Sub cmdAction_Click(Index As Integer)
Dim blnupdate As Boolean
Dim rsc As Recordset
Select Case Index
        Case 0
            Unload Me
            frmMainMenu.Show
        Case 2
            Me.MousePointer = vbHourglass
            blnupdate = False
            For i = 1 To grid.Rows - 1
                If grid.Cell(flexcpChecked, i, bteColFix) = flexChecked And CDbl(grid.TextMatrix(i, bteColAmountIDR)) = 0 Then
                    LblErrMsg.Caption = DisplayMsg(4085)
                    grid.Col = bteColFix
                    grid.Row = i
                    grid.SetFocus
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
            Next
            For i = 1 To grid.Rows - 1
                If grid.Cell(flexcpChecked, i, bteColFix) <> grid.TextMatrix(i, bteColStatus) Then
                    Db.BeginTrans
                    If grid.Cell(flexcpChecked, i, bteColFix) = flexChecked Then
                        sql = "update invoiceSupplier_master set fix_cls='1', Last_Update = getdate(), Last_User = '" & userLogin & "' where invoice_no = '" & grid.TextMatrix(i, bteColInvoiceNo) & "'"
                    Else
                        sql = "update invoiceSupplier_master set fix_cls= NULL, Last_Update = getdate(), Last_User = '" & userLogin & "' where invoice_no = '" & grid.TextMatrix(i, bteColInvoiceNo) & "'"
                    End If
                    blnupdate = True
                    Db.Execute sql
                    Db.CommitTrans
                End If
            Next
            display
            If blnupdate Then LblErrMsg = DisplayMsg(1101)
            Me.MousePointer = vbDefault
        Case 1
            blank
            Me.CtrlMenu1.MenuText = ""
            Header
            LblErrMsg = ""
End Select
End Sub

Private Sub CmdExcel_Click(Index As Integer)
    Dim xlapp As New Excel.application
    Dim rsCek As New Recordset, Idx As Long, tempi As String, tempcust As String
    Dim bolcust As Boolean, bolinv As Boolean
    Dim rsCompany As New Recordset, rate As Double
    
    If Trim(cbodealer) = "" Then LblErrMsg = DisplayMsg(1045): Exit Sub
    
    If Not cbodealer.MatchFound Then LblErrMsg = DisplayMsg(4072): Exit Sub
    
    
    
    sql = " select distinct a.*, c.trade_name, c.country_cls, b.currency_code, " & vbCrLf & _
            "    USD_Rate=(select Daily_ExchangeRate from Daily_ExchangeRate where Currency_Code='02' and ExchangeRate_Date=a.Invoice_Date), " & vbCrLf & _
            "    Amount_USD=(select sum(coalesce(Amount,0)) From InvoiceSupplier_Detail where Invoice_No=a.Invoice_No), " & vbCrLf & _
            "    Currency=Coalesce((Select Description From Curr_Cls where Curr_Cls=b.Currency_Code),'') " & vbCrLf & _
            " from invoiceSupplier_master a  " & vbCrLf & _
            " inner join invoiceSupplier_detail b on a.invoice_no = b.invoice_no  " & vbCrLf & _
            " inner join trade_master c on a.supplier_code = c.trade_code " & vbCrLf & _
            "  where invoice_date >='" & Format(SDate, "yyyy-MM-dd") & "' and invoice_date <= '" & Format(EDate, "yyyy-MM-dd") & "'  "
            
    If cbodealer <> strAll Then sql = sql & "and a.supplier_code = '" & Trim(cbodealer) & "' " & vbCrLf
    sql = sql & "order by a.invoice_no"
    
    If rsCek.State <> adStateClosed Then rsCek.Close
    rsCek.CursorLocation = adUseClient
    rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic
        
    If Not rsCek.EOF Then
        Screen.MousePointer = vbHourglass
        With xlapp
            
            sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax  From company_profile "
            If rsCompany.State <> adStateClosed Then rsCompany.Close
            rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
            If rsCompany.EOF Then Screen.MousePointer = vbDefault: Exit Sub
            
            .Workbooks.Add
            
            .Range("a2", "k2").Merge
            .Range("a2") = rsCompany!company_name
            .Range("a3", "k3").Merge
            .Range("a3") = rsCompany!address1 & " " & rsCompany!address2 & " " & rsCompany!City & " " & rsCompany!Province & " " & rsCompany!postal_code
            .Range("a4", "k4").Merge
            .Range("a4") = "Phone: " & rsCompany!phone1 & " " & rsCompany!phone2 & " Fax: " & rsCompany!fax
            
            .Range("a6") = "AP Status Apporval"
            .Range("b6") = ""
            .Range("a6", "b6").Merge
            .Range("a6").horizontalAlignment = xlLeft
            .Range("a7") = "Date"
            .Range("b7") = ": " & Format(Now, "dd MMMM YYYY")
            .Range("a8") = "Period"
            .Range("b8") = ": " & Format(SDate, "dd MMMM YYYY") & " to " & Format(EDate, "dd MMMM YYYY")
            
            
             .Range("a10") = "Supplier Code"
             .Range("b10") = "Supplier Name"
             .Range("c10") = "Invoice No"
             .Range("d10") = "Invoice Date"
             .Range("e10") = "Curr"
             .Range("f10") = "Rate(USD)"
             .Range("g10") = "Amount"
             .Range("h10") = "Amount (IDR)"
             .Range("i10") = "Amount (USD)"
             
            Idx = 11
           
            
            Do While Not rsCek.EOF
            
               
                Idx = Idx
                'Content
                .Range("a" & Idx) = Trim(rsCek!Supplier_Code)
                .Range("b" & Idx) = Trim(rsCek!trade_name)
                .Range("c" & Idx) = "'" & Trim(rsCek!Invoice_No)
                .Range("d" & Idx) = Format(rsCek!Invoice_Date, "dd MMM yyyy")
                .Range("e" & Idx) = "'" & Trim(rsCek!Currency)
                .Range("f" & Idx) = "'" & Format(rsCek!USD_Rate, gs_formatAmountIDR)
                .Range("g" & Idx) = "'" & Format(rsCek!total_amount, gs_formatAmountIDR)
                .Range("h" & Idx) = "'" & Format(rsCek!exchange_amount, gs_formatAmountIDR)

                .Range("i" & Idx) = Format(rsCek!exchange_amount / rsCek!USD_Rate, gs_formatAmountIDR)

                Idx = Idx + 1
                rsCek.MoveNext
            
            Loop
           
            
            .Range("a1", "k" & Idx + 3).Columns.Font.Name = "Arial"
            .Range("a1", "k" & Idx + 3).Columns.Font.Size = 8
            .Range("a2", "k2").Columns.Font.Name = "Arial"
            .Range("a2", "k2").Columns.Font.Size = "10"
            .Range("a2", "k2").Columns.Font.Bold = True
            .Range("a2", "k4").horizontalAlignment = xlCenter
            .Range("a6", "k6").Columns.Font.Bold = True
            .Range("a10", "k10").Columns.Font.Bold = True
            
            .Visible = True
            .ActiveSheet.PageSetup.PaperSize = xlPaperA4
            .Range("A:k").Columns.AutoFit
            .WindowState = xlMaximized
        
        End With
    Else
        LblErrMsg = DisplayMsg(4006)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
    Me.MousePointer = vbHourglass
    display
    Me.MousePointer = vbDefault
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub edate_Change()
Header
End Sub

Private Sub edate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
bteHakPrice = hakPrice(Me.Name)
adtocombo
SDate = Format(Now, "dd mmm yyyy")
EDate = Format(Now, "dd mmm yyyy")
HakU = hakUpdate(Me.Name)
Text1 = ""
End Sub

Sub adtocombo()
sql = "SELECT  rtrim(Trade_Master.trade_Code) cust_code, rtrim(Trade_Master.Trade_Name) cust_name, " & _
    "rtrim(Trade_Master.Address1) address From Trade_Master where trade_cls in ('2', '3')"
Set rstcust = New Recordset
rstcust.Open sql, Db, adOpenKeyset, adLockOptimistic
With cbodealer
.clear
.columnCount = 2
.ColumnWidths = "80 pt;300 pt; 0 pt"
.ListWidth = 380
.ListRows = 15
.AddItem ""
.List(0, 0) = strAll
.List(0, 1) = strAll
.List(0, 2) = strAll
i = 1
Do Until rstcust.EOF
    .AddItem ""
    .List(i, 0) = Trim(rstcust!Cust_CodE)
    .List(i, 1) = Trim(rstcust!Cust_Name)
    .List(i, 2) = Trim(rstcust!Address) & " "
    i = i + 1
    rstcust.MoveNext
Loop
.ListIndex = 0
End With
End Sub

Sub display()
Dim sqlP As String
cbodealer = cbodealer
If Not cbodealer.MatchFound Then LblErrMsg = DisplayMsg(4050): Text1.Text = "": Exit Sub

If CDate(SDate) > CDate(EDate) Then
    LblErrMsg.Caption = DisplayMsg(4068)
    Exit Sub
ElseIf CDate(EDate) < CDate(SDate) Then
    LblErrMsg.Caption = DisplayMsg(4066)
    Exit Sub
End If

Header

sql = " select distinct a.*, c.trade_name, c.country_cls, b.currency_code, " & vbCrLf & _
            "    USD_Rate=(select Daily_ExchangeRate from Daily_ExchangeRate where Currency_Code='02' and ExchangeRate_Date=a.Invoice_Date), " & vbCrLf & _
            "    Amount_USD=(select Daily_ExchangeRate from Daily_ExchangeRate where Currency_Code='02' and ExchangeRate_Date=a.Invoice_Date),  " & vbCrLf & _
            "    Currency=Coalesce((Select Description From Curr_Cls where Curr_Cls=b.Currency_Code),''), " & vbCrLf & _
            "    RateOri=case when b.Currency_Code<>'03' then  coalesce((select Daily_ExchangeRate from Daily_ExchangeRate where Currency_Code=b.Currency_Code and ExchangeRate_Date=a.Invoice_Date),0) else 1 end  " & vbCrLf & _
            " from invoiceSupplier_master a  " & vbCrLf & _
            " inner join invoiceSupplier_detail b on a.invoice_no = b.invoice_no  " & vbCrLf & _
            " inner join trade_master c on a.supplier_code = c.trade_code " & vbCrLf & _
            "  where invoice_date >='" & Format(SDate, "yyyy-MM-dd") & "' and invoice_date <= '" & Format(EDate, "yyyy-MM-dd") & "'  "
    
    
' sql = " select distinct a.*, c.trade_name, c.country_cls, b.currency_code, " & vbCrLf & _
            "    USD_Rate=(select Daily_ExchangeRate from Daily_ExchangeRate where Currency_Code='02' and ExchangeRate_Date=a.Invoice_Date), " & vbCrLf & _
            "    Amount_USD=(select sum(coalesce(Amount,0)) From InvoiceSupplier_Detail where Invoice_No=a.Invoice_No), " & vbCrLf & _
            "    Currency=Coalesce((Select Description From Curr_Cls where Curr_Cls=b.Currency_Code),'') " & vbCrLf & _
            " from invoiceSupplier_master a  " & vbCrLf & _
            " inner join invoiceSupplier_detail b on a.invoice_no = b.invoice_no  " & vbCrLf & _
            " inner join trade_master c on a.supplier_code = c.trade_code " & vbCrLf & _
            "  where invoice_date >='" & Format(SDate, "yyyy-MM-dd") & "' and invoice_date <= '" & Format(EDate, "yyyy-MM-dd") & "'  "
                

If cbodealer <> strAll Then sql = sql & "and a.supplier_code = '" & Trim(cbodealer) & "' " & vbCrLf
sql = sql & "order by a.invoice_no"

Set rst = New Recordset
rst.CursorLocation = adUseClient
rst.Open sql, Db, adOpenDynamic, adLockOptimistic
If Not rst.EOF Then
    With grid
        .Refresh
        .Rows = rst.RecordCount + 1
        .Row = .Rows - 1
    For i = 1 To rst.RecordCount
        
        .TextMatrix(i, bteColSuppCode) = Trim(rst!Supplier_Code)
        .TextMatrix(i, bteColSuppName) = Trim(rst!trade_name)
        .TextMatrix(i, bteColInvoiceNo) = Trim(rst!Invoice_No)
        .TextMatrix(i, bteColInvoiceDate) = Format(rst!Invoice_Date, "dd mmm yyyy")
        .TextMatrix(i, bteColDelDate) = Format(rst!BL_Date, "dd mmm yyyy")
        .TextMatrix(i, bteColCurr) = Trim(rst!Currency)
        .TextMatrix(i, btecolreateUSD) = Format(rst!USD_Rate, gs_formatAmountIDR)
        If Trim(rst!currency_code) = "03" Then
            .TextMatrix(i, bteColAmount) = Format(rst!total_amount, gs_formatAmountIDR)
        Else
            .TextMatrix(i, bteColAmount) = Format(rst!total_amount, gs_formatAmount)
        End If
        .TextMatrix(i, bteColAmountIDR) = Format(rst!exchange_amount, gs_formatAmountIDR)
        
        .TextMatrix(i, bteColBaseAmount) = Format(rst!exchange_amount / rst!USD_Rate, gs_formatAmount)
        If rst!country_cls = 1 Then
            .TextMatrix(i, bteColPPn) = Format(0, gs_formatAmountIDR)
        Else
            .TextMatrix(i, bteColPPn) = Format(rst!ppn, gs_formatAmountIDR)
        End If
        .TextMatrix(i, bteColFreight) = Format(rst!AirFreight_Amount, gs_formatAmountIDR)
        .TextMatrix(i, bteColTotal) = Format(rst!exchange_amount + rst!AirFreight_Amount + rst!ppn, gs_formatAmountIDR)
        If IsNull(rst!Paid_Cls) Then
            .Cell(flexcpChecked, i, bteColIssue) = flexUnchecked
        Else
            .Cell(flexcpChecked, i, bteColIssue) = 1
        End If
        
        If IsNull(rst!fix_cls) Then
            .Cell(flexcpChecked, i, bteColFix) = flexUnchecked
            .TextMatrix(i, bteColStatus) = 2
        Else
            .Cell(flexcpChecked, i, bteColFix) = flexChecked
            .TextMatrix(i, bteColStatus) = 1
        End If
        .Cell(flexcpBackColor, i, bteColFix) = vbWhite
        
       .Cell(flexcpAlignment, i, bteColInvoiceNo, i, bteColDelDate) = flexAlignLeftCenter
       .Cell(flexcpAlignment, i, bteColAmount, i, bteColTotal) = flexAlignRightCenter
       'whitecols
       
    rst.MoveNext
    Next i
    End With
    'hapus message
    LblErrMsg.Caption = ""
    'button

Else
    LblErrMsg.Caption = DisplayMsg(4006)
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
rst.Close
Set rst = Nothing
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col <> bteColFix Then
    Cancel = True
Else
    LblErrMsg = up_ValidateDateRange(CDate(grid.TextMatrix(grid.Row, bteColInvoiceDate)), True)
    If LblErrMsg <> "" Then Cancel = True
End If
End Sub

Sub blank()
SDate = Format(Now, "dd MMM YYYY")
EDate = Format(Now, "dd MMM YYYY")
cbodealer.ListIndex = 0
Text1 = ""
End Sub

Private Sub sdate_Change()
Header
End Sub

Private Sub sdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub









