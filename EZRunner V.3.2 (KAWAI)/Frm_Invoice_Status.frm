VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Invoice_Status 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Invoice Status"
   ClientHeight    =   10320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "Frm_Invoice_Status.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
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
      Left            =   12480
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
      Left            =   13770
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9750
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   840
      Left            =   480
      TabIndex        =   9
      Top             =   1125
      Width           =   14490
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0000FFFF&
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
         Left            =   12810
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   285
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
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   3825
      End
      Begin MSComCtl2.DTPicker edate 
         Height          =   315
         Left            =   10695
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
         Format          =   146538499
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker sdate 
         Height          =   315
         Left            =   8505
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
         Format          =   146538499
         CurrentDate     =   37810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust CD"
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
         Left            =   270
         TabIndex        =   12
         Top             =   375
         Width           =   720
      End
      Begin MSForms.ComboBox cbodealer 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   315
         Width           =   1425
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2514;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         X1              =   2850
         X2              =   6720
         Y1              =   645
         Y2              =   645
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
         Left            =   7185
         TabIndex        =   11
         Top             =   375
         Width           =   1260
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
         Left            =   10230
         TabIndex        =   10
         Top             =   345
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   480
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
      Left            =   13110
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   270
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6825
      Left            =   480
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
      ScrollTrack     =   -1  'True
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
      Caption         =   "Invoice Status"
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
      Left            =   480
      TabIndex        =   13
      Top             =   300
      Width           =   14430
   End
End
Attribute VB_Name = "Frm_Invoice_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim i As Integer, Y As Integer, HakU As Integer

Dim bteColInvoiceNo As Byte
Dim bteColCustCode As Byte
Dim bteColCustName As Byte
Dim bteColInvoiceDate As Byte
Dim bteColDelMonth As Byte
Dim bteColDelDate As Byte
Dim bteColCurr As Byte
Dim bteColAmount As Byte
Dim bteColAmountIDR As Byte
Dim bteColPPn As Byte
Dim bteColTotal As Byte
Dim bteColRemarks As Byte
Dim bteColIssue As Byte
Dim bteColFix As Byte
Dim bteColStatus As Byte

Dim bteHakPrice As Byte

Sub Header()
With Grid
    
    bteColInvoiceNo = 0
    bteColCustCode = 1
    bteColCustName = 2
    bteColInvoiceDate = 3
    bteColDelMonth = 4
    bteColDelDate = 5
    bteColCurr = 6
    bteColAmount = 7
    bteColAmountIDR = 8
    bteColPPn = 9
    bteColTotal = 10
    bteColRemarks = 11
    bteColIssue = 12
    bteColFix = 13
    bteColStatus = 14
    
    .ColS = 15
    .Rows = 1
    
    .TextMatrix(0, bteColInvoiceNo) = "Invoice No"
    .TextMatrix(0, bteColCustCode) = "Cust. Code"
    .TextMatrix(0, bteColCustName) = "Cust. Name"
    .TextMatrix(0, bteColInvoiceDate) = "Invoice Date"
    .TextMatrix(0, bteColDelMonth) = "Delivery Month"
    .TextMatrix(0, bteColDelDate) = "Delv. Date/ETD"
    .TextMatrix(0, bteColCurr) = "Curr"
    .TextMatrix(0, bteColAmount) = "Amount"
    .TextMatrix(0, bteColAmountIDR) = "Amount (IDR)"
    .TextMatrix(0, bteColPPn) = "PPN"
    .TextMatrix(0, bteColTotal) = "Total Amount"
    .TextMatrix(0, bteColRemarks) = "Remarks"
    .TextMatrix(0, bteColIssue) = "Issued"
    .TextMatrix(0, bteColFix) = "Fix"
    .TextMatrix(0, bteColStatus) = "status"
    
    .ColWidth(bteColInvoiceNo) = 1300
    .ColWidth(bteColCustCode) = 1100
    .ColWidth(bteColCustName) = 1200
    .ColWidth(bteColInvoiceDate) = 1300
    .ColWidth(bteColDelMonth) = 1500
    .ColWidth(bteColDelDate) = 1500
    .ColWidth(bteColCurr) = 600
    .ColWidth(bteColAmount) = 1600
    .ColWidth(bteColAmountIDR) = 1600
    .ColWidth(bteColPPn) = 1400
    .ColWidth(bteColTotal) = 1600
    .ColWidth(bteColRemarks) = 1000
    .ColWidth(bteColIssue) = 800
    .ColWidth(bteColFix) = 400
    
    .ColHidden(bteColStatus) = True
    .ColHidden(bteColRemarks) = True
    '.ColHidden(bteColDelMonth) = True
    .ColHidden(bteColDelDate) = True
    
    .ColHidden(bteColCurr) = (bteHakPrice = 0)
    .ColHidden(bteColAmount) = (bteHakPrice = 0)
    .ColHidden(bteColAmountIDR) = (bteHakPrice = 0)
    .ColHidden(bteColPPn) = (bteHakPrice = 0)
    .ColHidden(bteColTotal) = (bteHakPrice = 0)
    
    .Cell(flexcpAlignment, 0, 0, 0, bteColFix) = flexAlignCenterCenter
    .EditMaxLength = 1
End With
End Sub

Private Sub cbodealer_Change()
cbodealer = cbodealer
If cbodealer.MatchFound Then
    LblErrMsg = ""
    Text1 = cbodealer.Column(1)
Else
    Text1 = ""
    LblErrMsg = DisplayMsg(4072)
End If
Header
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
            For i = 1 To Grid.Rows - 1
                If Grid.Cell(flexcpChecked, i, bteColFix) = flexChecked And CDbl(Grid.TextMatrix(i, bteColAmountIDR)) = 0 Then
                    LblErrMsg.Caption = DisplayMsg(4085)
                    Grid.Col = bteColFix
                    Grid.Row = i
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
            Next
            For i = 1 To Grid.Rows - 1
                If Grid.TextMatrix(i, bteColStatus) <> Grid.Cell(flexcpChecked, i, bteColFix) Then
                    Db.BeginTrans
                    If Grid.Cell(flexcpChecked, i, bteColFix) = flexChecked Then
                        If Not CekFixPacking(Grid.TextMatrix(i, bteColInvoiceNo)) Then
                            Db.RollbackTrans
                            Grid.Row = i
                            Grid.Col = bteColFix
                            LblErrMsg = DisplayMsg("0078")
                            Me.MousePointer = vbDefault
                            Exit Sub
                        End If
                        sql = "update invoice_master set fix_cls='1', Last_Update = getdate(), Last_User = '" & userLogin & "' where invoice_no = '" & Grid.TextMatrix(i, bteColInvoiceNo) & "'"
                    Else
                        sql = "update invoice_master set fix_cls= NULL, Last_Update = getdate(), Last_User = '" & userLogin & "' where invoice_no = '" & Grid.TextMatrix(i, bteColInvoiceNo) & "'"
                    End If
                    blnupdate = True
                    Db.Execute sql
                    Db.CommitTrans
                End If
            Next
            If blnupdate Then LblErrMsg = DisplayMsg(1101)
            Me.MousePointer = vbDefault
        Case 1
            blank
            Me.CtrlMenu1.MenuText = ""
            LblErrMsg = ""
End Select
End Sub

Private Sub cmdSearch_Click()
On Error GoTo ErrorMesage
MousePointer = vbHourglass

display

MousePointer = vbDefault
Exit Sub
ErrorMesage:
LblErrMsg = err.number & " " & err.Description
MousePointer = vbDefault
Exit Sub

End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub edate_Change()
'header
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
SDate = Format(Date - (Day(Date) - 1), "dd mmm yyyy")
EDate = Format(Date, "dd mmm yyyy")
HakU = hakUpdate(Me.Name)
Header
End Sub

Sub adtocombo()
    Dim rstcust As Recordset
    sql = "SELECT  rtrim(Trade_Master.trade_Code) cust_code, rtrim(Trade_Master.Trade_Name) cust_name, " & _
        "rtrim(Trade_Master.Address1) address From Trade_Master where trade_cls in('2', '4')"
    
    Set rstcust = New Recordset
    rstcust.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    With cbodealer
        .clear
        .columnCount = 2
        .ColumnWidths = "50 pt;300 pt; 0 pt"
        .ListWidth = 350
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
            .List(i, 2) = IIf(IsNull(Trim(rstcust!Address)), "", Trim(rstcust!Address))
            i = i + 1
            rstcust.MoveNext
        Loop
        .ListIndex = 0
    End With
    rstcust.Close
    Set rstcust = Nothing
End Sub

Sub display()
Dim rst As Recordset
Dim sqlP As String
Me.MousePointer = vbHourglass
cbodealer = cbodealer
If Not cbodealer.MatchFound Then LblErrMsg = DisplayMsg(4072): Text1.Text = "": Me.MousePointer = vbDefault: Exit Sub

If CDate(SDate) > CDate(EDate) Then
    LblErrMsg.Caption = DisplayMsg(4068)
    Me.MousePointer = vbDefault
    Exit Sub
ElseIf CDate(EDate) < CDate(SDate) Then
    LblErrMsg.Caption = DisplayMsg(4066)
    Me.MousePointer = vbDefault
    Exit Sub
End If

'sql = "select distinct a.*, b.trade_name, b.country_cls, d.description curr_desc, " & _
'    "case when c.currency_code = '03' then 1 " & _
'    "else isnull(( " & _
'        "select Daily_Exchangerate from Daily_Exchangerate where currency_code = c.currency_code " & _
'        "and exchangeRate_Date = case when Country_Cls='0' then (select distinct delivery_date from delivery_order where do_no= a.List_Do) else (select distinct ETD from packing_master where packing_no= a.List_Do) end), 0) end DailyExchange_Rate, " & _
'    "case when Country_Cls='0' then (select distinct delivery_date from delivery_order where do_no= a.List_Do) else (select distinct ETD from packing_master where packing_no= a.List_Do) end delv_date " & _
'    "from invoice_master a " & _
'    "inner join trade_master b on a.cust_code = b.trade_code " & _
'    "inner join invoice_detail c on a.invoice_no = c.invoice_no " & _
'    "inner join curr_cls d on c.currency_code = d.curr_cls " & _
'    "where a.invoice_date >='" & sdate & "' and a.invoice_date <= '" & edate & "' "
    
sql = "select distinct a.*, b.trade_name, b.country_cls, d.description curr_desc, " & _
      vbLf & "DailyExchange_Rate = " & _
      vbLf & "( " & _
      vbLf & "case when c.currency_code = '03' then 1 else " & _
      vbLf & "(isnull(( " & _
      vbLf & " select Daily_Exchangerate from Daily_Exchangerate " & _
      vbLf & " where currency_code = c.currency_code and " & _
      vbLf & " exchangeRate_Date =  " & _
      vbLf & " Coalesce( " & _
      vbLf & " (select distinct pm.packing_date from packing_master pm where pm.packing_no = c.Packing_no), " & _
      vbLf & " (select distinct dm.DO_Date from DO_Master dm where dm.do_no= c.DO_No)" & _
      vbLf & "          )" & _
      vbLf & "), 0) )  " & _
      vbLf & " end)  " & _
      vbLf & " " & _
      vbLf & "/*, delv_date = " & _
      vbLf & "( " & _
      vbLf & "case when b.Country_Cls='0' " & _
      vbLf & "then (select distinct dm.DO_Date from DO_Master dm where dm.do_no= c.DO_No) " & _
      vbLf & "else ( " & _
      vbLf & "select distinct pm.packing_date from packing_master pm where pm.packing_no = c.Packing_no " & _
      vbLf & ") " & _
      vbLf & "End " & _
      vbLf & ") */ "

sql = sql & _
      vbLf & "from invoice_master a " & _
      vbLf & "inner join trade_master b on a.cust_code = b.trade_code " & _
      vbLf & "inner join invoice_detail c on a.invoice_no = c.invoice_no " & _
      vbLf & "inner join curr_cls d on c.currency_code = d.curr_cls " & _
      vbLf & "where a.invoice_date >='" & SDate & "' and a.invoice_date <= '" & EDate & "' "
    

If cbodealer <> strAll Then sql = sql & "and cust_Code= '" & cbodealer & "' "
sql = sql & "order by a.invoice_no"

'header
Set rst = New Recordset
rst.CursorLocation = adUseClient
rst.Open sql, Db, adOpenDynamic, adLockOptimistic
If Not rst.EOF Then
    With Grid
        .Refresh
        .Rows = rst.RecordCount + 1
        .Row = .Rows - 1
    For i = 1 To rst.RecordCount
        
        .TextMatrix(i, bteColInvoiceNo) = Trim(rst!Invoice_No)
        .TextMatrix(i, bteColCustCode) = Trim(rst!Cust_CodE)
        .TextMatrix(i, bteColCustName) = Trim(rst!trade_name)
        .TextMatrix(i, bteColInvoiceDate) = Format(rst!Invoice_Date, "dd mmm yyyy")
        .TextMatrix(i, bteColDelMonth) = MonthName(Month(rst!delivery_Date)) & " " & Year(rst!delivery_Date)
        '.TextMatrix(i, bteColDelDate) = Format(rst!delv_date, "dd MMM yyyy")
        .TextMatrix(i, bteColCurr) = Trim(rst!curr_desc)
        If InStr(1, rst!Amount, ".") Then
            .TextMatrix(i, bteColAmount) = Format(rst!Amount, gs_formatAmountIDR)
        Else
            .TextMatrix(i, bteColAmount) = Format(rst!Amount, gs_formatAmountIDR)
        End If
        If InStr(1, rst!Amount, ".") Then
            .TextMatrix(i, bteColAmountIDR) = Format(rst!Amount * rst!DailyExchange_Rate, gs_formatAmountIDR)
        Else
            .TextMatrix(i, bteColAmountIDR) = Format(rst!Amount * rst!DailyExchange_Rate, gs_formatAmountIDR)
        End If
        If rst!country_cls = 1 Then
            .TextMatrix(i, bteColPPn) = 0
            .TextMatrix(i, bteColTotal) = .TextMatrix(i, bteColAmount)
        Else
            If InStr(1, rst!ppn, ".") Then
                .TextMatrix(i, bteColPPn) = Format(rst!ppn, gs_formatAmountIDR)
            Else
                .TextMatrix(i, bteColPPn) = Format(rst!ppn, gs_formatAmountIDR)
            End If
            
            If InStr(1, rst!total_amount, ".") Then
                .TextMatrix(i, bteColTotal) = Format(rst!total_amount, gs_formatAmountIDR)
            Else
                .TextMatrix(i, bteColTotal) = Format(rst!total_amount, gs_formatAmountIDR)
            End If
        
        End If
        
        .TextMatrix(i, bteColRemarks) = Trim(rst!Remarks)
        If IsNull(rst!reissue_cls) Then
            .Cell(flexcpChecked, i, bteColIssue) = flexUnchecked
        Else
            .Cell(flexcpChecked, i, bteColIssue) = 1
        End If
        
        If IsNull(rst!fix_cls) Then
            .Cell(flexcpChecked, i, bteColFix) = flexUnchecked
            .TextMatrix(i, bteColStatus) = flexUnchecked
        Else
            .Cell(flexcpChecked, i, bteColFix) = flexChecked
            .TextMatrix(i, bteColStatus) = flexChecked
        End If
        .Cell(flexcpBackColor, i, bteColFix) = vbWhite
        
       .Cell(flexcpAlignment, i, bteColInvoiceNo, i, bteColRemarks) = flexAlignLeftCenter
       .Cell(flexcpAlignment, i, bteColAmount, i, bteColTotal) = flexAlignRightCenter
       
        rst.MoveNext
    Next i
    End With
    LblErrMsg.Caption = ""
Else
    LblErrMsg.Caption = DisplayMsg(4006)
End If
rst.Close
Set rst = Nothing
Me.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'If Grid.Col = bteColFix Then
Dim rsc As Recordset
sql = "select distinct fakturpajak_no  from fakturpajak_detail where invoice_no = '" & Grid.TextMatrix(Row, bteColInvoiceNo) & "'"
Set rsc = New Recordset
rsc.Open sql, Db, adOpenKeyset, adLockOptimistic
If Not rsc.EOF Then
    Grid.Cell(flexcpChecked, Row, Col) = flexChecked
    LblErrMsg.Caption = DisplayMsg("0046")
Else
    LblErrMsg.Caption = ""
End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col <> bteColFix Then
    Cancel = True
Else
    LblErrMsg = up_ValidateDateRange(CDate(Grid.TextMatrix(Row, bteColInvoiceDate)), True)
    If LblErrMsg <> "" Then Cancel = True
End If
End Sub

Sub blank()
SDate = Format(Date - (Day(Date) - 1), "dd MMM YYYY")
EDate = Format(Date, "dd MMM YYYY")
cbodealer.ListIndex = 0
'header
End Sub

Private Sub sdate_Change()
'header
End Sub

Private Sub sdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Function CekFixPacking(strInvoice) As Boolean
    
    Dim adoRs As New ADODB.Recordset
    
    sql = "select max(fix_cls) from(" & _
        "select max(a.fix_cls) fix_cls from packing_master a inner join invoice_detail b on b.packing_no = a.packing_no where b.invoice_no = '" & strInvoice & "' " & _
        "Union " & _
        "select max(a.fix_cls) fix_cls from do_master a inner join invoice_detail b on b.do_no = a.do_no where b.invoice_no = '" & strInvoice & "') a having max(fix_cls) = 1"
    
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    CekFixPacking = Not adoRs.EOF
    adoRs.Close
    Set adoRs = Nothing
    
End Function
