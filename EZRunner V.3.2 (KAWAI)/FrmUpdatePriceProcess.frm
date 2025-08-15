VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmUpdatePriceProcess 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Update Price Process"
   ClientHeight    =   5400
   ClientLeft      =   3360
   ClientTop       =   3720
   ClientWidth     =   8445
   Icon            =   "FrmUpdatePriceProcess.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   8445
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1740
      Top             =   4890
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   6180
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   390
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.TextBox TxtRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
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
      Height          =   795
      Left            =   270
      MultiLine       =   -1  'True
      TabIndex        =   20
      Text            =   "FrmUpdatePriceProcess.frx":0E42
      Top             =   240
      Width           =   7785
   End
   Begin VB.PictureBox PicProg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   330
      ScaleHeight     =   495
      ScaleWidth      =   7755
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3330
      Width           =   7785
      Begin MSComctlLib.ProgressBar Prg1 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   90
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1920
      Left            =   270
      TabIndex        =   5
      Top             =   1110
      Width           =   7800
      Begin MSComCtl2.DTPicker DtStart 
         Height          =   315
         Left            =   1980
         TabIndex        =   6
         Top             =   390
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   294584323
         UpDown          =   -1  'True
         CurrentDate     =   37810
      End
      Begin VB.Label lblProduct 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   3630
         TabIndex        =   13
         Top             =   1200
         Width           =   2850
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   3630
         TabIndex        =   12
         Top             =   840
         Width           =   2850
      End
      Begin VB.Line Line3 
         X1              =   3630
         X2              =   6480
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         X1              =   3630
         X2              =   6480
         Y1              =   1080
         Y2              =   1080
      End
      Begin MSForms.ComboBox CboProduct 
         Height          =   315
         Left            =   1980
         TabIndex        =   11
         Top             =   1140
         Width           =   1575
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2778;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox CboSupplier 
         Height          =   315
         Left            =   1980
         TabIndex        =   10
         Top             =   780
         Width           =   1575
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2778;556"
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
         Index           =   2
         Left            =   420
         TabIndex        =   9
         Top             =   1230
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
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
         Left            =   420
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Left            =   420
         TabIndex        =   7
         Top             =   450
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdsubmenu 
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
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4860
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   525
      Left            =   300
      TabIndex        =   2
      Top             =   4230
      Width           =   7800
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
         Left            =   240
         TabIndex        =   3
         Top             =   180
         Width           =   7455
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Process"
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
      Left            =   6990
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4890
      Width           =   1140
   End
   Begin VB.Label lblCheck2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1230
      TabIndex        =   19
      Top             =   3060
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblcheck 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checking"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   18
      Top             =   3060
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblpart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part Receipt : 0 Updated"
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
      Left            =   330
      TabIndex        =   17
      Top             =   3900
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label lblinvoice 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice : 0 Updated"
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
      Left            =   2700
      TabIndex        =   16
      Top             =   3900
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "FrmUpdatePriceProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jmlpart As Double
Dim jmlinvoice As Double
Dim waktu As Integer

Private Sub CboProduct_Change()
If CboProduct.MatchFound Then
    lblProduct.Caption = CboProduct.List(CboProduct.ListIndex, 1)
    LblErrMsg = ""
Else
    lblProduct.Caption = ""
End If
lblcheck.Visible = False
lblCheck2.Visible = False
LblPart.Visible = False
lblinvoice.Visible = False
End Sub

Private Sub CboSupplier_Change()
If cboSupplier.MatchFound Then
    lblSupplier.Caption = cboSupplier.List(cboSupplier.ListIndex, 1)
    LblErrMsg = ""
Else
    lblSupplier.Caption = ""
End If
Call AddToComboProduct
lblcheck.Visible = False
lblCheck2.Visible = False
LblPart.Visible = False
lblinvoice.Visible = False
End Sub

Private Sub CmdSubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub Command1_Click(Index As Integer)
jmlpart = 0
jmlinvoice = 0

LblPart.Caption = "Part Receipt : " & jmlpart & " Updated"
lblinvoice.Caption = "Invoice : " & jmlinvoice & " Updated"

LblErrMsg.Caption = ""
If cboSupplier.MatchFound = False Then LblErrMsg.Caption = DisplayMsg(1054): Exit Sub
If CboProduct.MatchFound = False Then LblErrMsg.Caption = DisplayMsg(8082): Exit Sub

If MsgBox("Are You Sure To Process !!", vbOKCancel + vbInformation, "Process") = vbOK Then
    lblcheck.Visible = True
    lblCheck2.Visible = True
    Me.MousePointer = vbHourglass
    'Timer1.Enabled = True
    Call Update_Part_Receipt
Else
'    lblcheck.Visible = True
'    lblCheck2.Visible = True
End If
Me.MousePointer = vbDefault
End Sub
Sub Update_Part_Receipt()
Dim rsUpdate As New ADODB.Recordset
Dim RsPartReceipt As New ADODB.Recordset
Dim sqlUpdate As String
Dim Harga As Double
Dim Curr As String
Dim a As Integer


'On Error GoTo Errhandle
lblcheck.Visible = True
lblCheck2.Visible = True
LblPart.Visible = True
lblinvoice.Visible = True

Db.BeginTrans

sql = "select * from price_master where Left(start_date,6) ='" & Format(DtStart, "yyyyMM") & "' and price_cls='01' " & _
        IIf(cboSupplier.Text = strAll, "", " and trade_code='" & Trim(cboSupplier) & "' ") & _
        IIf(CboProduct.Text = strAll, "", " and item_code='" & Trim(CboProduct) & "' ")
If rsUpdate.State <> adStateClosed Then rsUpdate.Close
rsUpdate.CursorLocation = adUseClient
rsUpdate.Open sql, Db, adOpenForwardOnly, adLockReadOnly

i = rsUpdate.RecordCount
If i = 0 Then
    Prg1.Max = 1
Else
    Prg1.Max = i
End If
i = 0
a = 0

Do While Not rsUpdate.EOF

    Harga = IIf(IsNull(rsUpdate.Fields("Price")), 0, rsUpdate.Fields("Price"))
    Curr = rsUpdate.Fields("Currency_Code")
    
    i = i + 1
    If a = 6 Then a = 0
    a = a + 1
    sqlUpdate = "select * from part_receipt where convert(char,receipt_date,112) between '" & Trim(rsUpdate.Fields("Start_Date")) & "' and '" & Trim(rsUpdate.Fields("End_Date")) & "' "
    sqlUpdate = sqlUpdate & "And Item_code='" & Trim(rsUpdate("item_code")) & "'"
    sqlUpdate = sqlUpdate & "And Supplier_code='" & Trim(rsUpdate("trade_code")) & "'"
    If RsPartReceipt.State <> adStateClosed Then RsPartReceipt.Close
    RsPartReceipt.Open sqlUpdate, Db, adOpenKeyset, adLockOptimistic

    Do While Not RsPartReceipt.EOF
        
        'Harga = Get_Record("select top 1 price from price_master where start_date<='" & Format(RsPartReceipt("receipt_date"), "yyyymmdd") & "' and price_cls='01' and trade_code='" & Trim(RsPartReceipt("supplier_code")) & "' and item_code='" & Trim(RsPartReceipt("item_code")) & "' order by start_date desc,priority_cls desc,last_update desc ")
        'curr = Trim(Get_Record("select top 1 currency_code from price_master where start_date<='" & Format(RsPartReceipt("receipt_date"), "yyyymmdd") & "' and price_cls='01' and trade_code='" & Trim(RsPartReceipt("supplier_code")) & "' and item_code='" & Trim(RsPartReceipt("item_code")) & "' order by start_date desc,priority_cls desc,last_update desc "))
        
        'check  harga terbaru pada price
        If RsPartReceipt("price") <> Harga Or RsPartReceipt("Currency_code") <> Curr Then
            
            
            'simpan ke tabel part_receipt_update_price dengan harga yang lama
            Call Insert_Part_Receipt_Process(Trim(RsPartReceipt("po_no")), Trim(RsPartReceipt("supplier_code")), Trim(RsPartReceipt("item_code")), RsPartReceipt("price"), Trim(RsPartReceipt("Currency_code")), RsPartReceipt("seq_no"))
            
            'simpan harga ke tabel part_receipt
            RsPartReceipt("price") = Harga
            RsPartReceipt("Currency_code") = Curr
            RsPartReceipt("Amount") = Harga * RsPartReceipt("qty")
            RsPartReceipt("last_update") = Now
            RsPartReceipt("last_user") = userLogin
            RsPartReceipt.update
            jmlpart = jmlpart + 1
            LblPart.Caption = "Part Receipt : " & jmlpart & " Updated"
        End If
        
        Call Update_Invoice_Supplier(RsPartReceipt("seq_no"), Harga, Curr, RsPartReceipt("item_code"), RsPartReceipt("supplier_code"))
        RsPartReceipt.MoveNext
        
    Loop
    
    Prg1.Value = i
    rsUpdate.MoveNext
    
  
Loop

Db.CommitTrans
If jmlinvoice > 0 Or jmlpart > 0 Then LblErrMsg.Caption = DisplayMsg(1000)
ErrExit:
    Set RsPartReceipt = Nothing
    Set rsUpdate = Nothing
    Prg1.Value = 0
    lblcheck.Visible = False
    lblCheck2.Visible = False
    If LblErrMsg = "" Then LblErrMsg.Caption = DisplayMsg(9003)
    Exit Sub

Errhandle:
    Db.RollbackTrans
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit

End Sub
Function Insert_Part_Receipt_Process(PONO As String, Supplier As String, Item As String, Price As Double, Curr As String, Seq_no As Double)
Dim RsInsertPartReceipt As New ADODB.Recordset
Dim SqlInsertPart As String

SqlInsertPart = "Select * from Part_Receipt_Update_Price"
If RsInsertPartReceipt.State <> adstateclose Then RsInsertPartReceipt.Close
RsInsertPartReceipt.Open SqlInsertPart, Db, adOpenKeyset, adLockOptimistic
RsInsertPartReceipt.AddNew
RsInsertPartReceipt("PO_No") = Trim(PONO)
RsInsertPartReceipt("Receiptseq_no") = Seq_no
RsInsertPartReceipt("Supplier_Code") = Trim(Supplier)
RsInsertPartReceipt("Item_Code") = Trim(Item)
RsInsertPartReceipt("Price_Old") = Price
RsInsertPartReceipt("Currency_Code_Old") = Trim(Curr)
RsInsertPartReceipt("Process_Date") = Now
RsInsertPartReceipt("User_ID") = userLogin
RsInsertPartReceipt.update

'LblErrMsg.Caption = DisplayMsg(1000)
Set RsInsertPartReceipt = Nothing
End Function
Function Update_Invoice_Supplier(Seq_no As Double, Price As Double, Curr As String, Item As String, Supplier As String)
Dim RsInvoice As New ADODB.Recordset
Dim SqlUpdateInvoice As String

SqlUpdateInvoice = "select * from invoicesupplier_detail where receiptseq_no='" & Seq_no & "' "
SqlUpdateInvoice = SqlUpdateInvoice & "And Item_code='" & Trim(Item) & "'"
SqlUpdateInvoice = SqlUpdateInvoice & "And Supplier_code='" & Trim(Supplier) & "'"
If RsInvoice.State <> adstateclose Then RsInvoice.Close
RsInvoice.Open SqlUpdateInvoice, Db, adOpenKeyset, adLockOptimistic
Do While Not RsInvoice.EOF
    If RsInvoice("price") <> Price Or RsInvoice("currency_code") <> Curr Then
        Call Insert_Invoice_Supplier_Process(Trim(RsInvoice("PO_No")), Trim(RsInvoice("Supplier_Code")), Trim(RsInvoice("Item_Code")), RsInvoice("Price"), Trim(RsInvoice("Currency_Code")), Seq_no)
        RsInvoice("price") = Price
        RsInvoice("price_Mat") = Price
        RsInvoice("currency_code") = Trim(Curr)
        RsInvoice("Amount") = Price * RsInvoice("qty")
        RsInvoice("Exchange_Amount") = Price * RsInvoice("qty")
        RsInvoice("Amount_Mat") = Price * RsInvoice("qty")
        RsInvoice("last_update") = Now
        RsInvoice("last_user") = userLogin
        RsInvoice.update
        jmlinvoice = jmlinvoice + 1
        lblinvoice.Caption = "Invoice : " & jmlinvoice & " Updated"
        'LblErrMsg.Caption = DisplayMsg(1000)
    End If
    RsInvoice.MoveNext
Loop
Set RsInvoice = Nothing
End Function
Function Insert_Invoice_Supplier_Process(PONO As String, Supplier As String, Item As String, Price As Double, Curr As String, Seq_no As Double)
Dim RsInsertInvoiceSupplier As New ADODB.Recordset
Dim SqlInsertInvoice As String

SqlInsertInvoice = "Select * from Invoice_Supplier_Update_Price"
If RsInsertInvoiceSupplier.State <> adstateclose Then RsInsertInvoiceSupplier.Close
RsInsertInvoiceSupplier.Open SqlInsertInvoice, Db, adOpenKeyset, adLockOptimistic
RsInsertInvoiceSupplier.AddNew
RsInsertInvoiceSupplier("PO_No") = Trim(PONO)
RsInsertInvoiceSupplier("receiptseq_no") = Seq_no
RsInsertInvoiceSupplier("Supplier_Code") = Trim(Supplier)
RsInsertInvoiceSupplier("Item_Code") = Trim(Item)
RsInsertInvoiceSupplier("Price_Old") = Price
RsInsertInvoiceSupplier("Currency_Code_Old") = Trim(Curr)
RsInsertInvoiceSupplier("Process_Date") = Now
RsInsertInvoiceSupplier("User_ID") = userLogin
RsInsertInvoiceSupplier.update

Set RsInsertInvoiceSupplier = Nothing
End Function
Private Sub dtStart_Change()
Call AddToComboSupplier
Call AddToComboProduct
lblcheck.Visible = False
lblCheck2.Visible = False
LblPart.Visible = False
lblinvoice.Visible = False
End Sub

Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(Me)

DtStart = Now
Call AddToComboSupplier
Call AddToComboProduct

CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

End Sub
Sub AddToComboSupplier()
Dim rsSupplier As New ADODB.Recordset

sql = "select distinct  trade_code,(select isnull(trade_name,'') as trade_name from trade_master where trade_code=price_master.trade_code) trade_name from price_master where substring(start_date,1,6)<='" & Year(DtStart) & Format(Month(DtStart), "00") & "' and substring(end_date,1,6)>='" & Year(DtStart) & Format(Month(DtStart), "00") & "' and price_cls='01' "
If rsSupplier.State <> adStateClosed Then rsSupplier.Close
rsSupplier.Open sql, Db, adOpenForwardOnly, adLockPessimistic

With cboSupplier
    .clear
    .columnCount = 2
    .ColumnWidths = "50pt;200pt"
    .ListWidth = 250
    .ListRows = 15
    
    If Not rsSupplier.EOF Then
        .AddItem
        .List(0, 0) = strAll
        .List(0, 1) = strAll
    End If
    
    i = 1
    Do While Not rsSupplier.EOF
        .AddItem
        .List(i, 0) = Trim(rsSupplier("Trade_code"))
        .List(i, 1) = Trim(rsSupplier("Trade_Name")) & ""
        rsSupplier.MoveNext
        i = i + 1
    Loop
    If .ListCount > 0 Then .ListIndex = 0
End With

Set rsSupplier = Nothing
End Sub

Sub AddToComboProduct()
Dim RsProduct As New ADODB.Recordset

sql = "select distinct item_code,(select item_name from item_master where item_code=price_master.item_code) Item_name  from price_master where substring(start_date,1,6)<='" & Year(DtStart) & Format(Month(DtStart), "00") & "' and substring(end_date,1,6)>='" & Year(DtStart) & Format(Month(DtStart), "00") & "' and price_cls='01' " & IIf(cboSupplier.Text = strAll, "", " and trade_code='" & Trim(cboSupplier) & "' ")
If RsProduct.State <> adStateClosed Then RsProduct.Close
RsProduct.Open sql, Db, adOpenForwardOnly, adLockPessimistic

With CboProduct
    .clear
    .columnCount = 2
    .ColumnWidths = "50pt;200pt"
    .ListWidth = 250
    .ListRows = 15
    
    If Not RsProduct.EOF Then
        .AddItem
        .List(0, 0) = strAll
        .List(0, 1) = strAll
    End If
    
    i = 1
    Do While Not RsProduct.EOF
        .AddItem
        .List(i, 0) = Trim(RsProduct("Item_code"))
        .List(i, 1) = Trim(RsProduct("Item_name"))
        RsProduct.MoveNext
        i = i + 1
    Loop
    If .ListCount > 0 Then .ListIndex = 0
End With

Set RsProduct = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub
'
'Private Sub Timer1_Timer()
'
'If waktu > 10 Then waktu = 0
'waktu = waktu + 1
'lblCheck2.Caption = Mid("..........", 1, waktu)
'End Sub
