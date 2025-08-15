VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmAPListReport 
   BackColor       =   &H00FDDFE3&
   Caption         =   "AP Detail List Report"
   ClientHeight    =   3765
   ClientLeft      =   1935
   ClientTop       =   3585
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmApDetailList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   60
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Excel"
      Height          =   375
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3180
      Width           =   1185
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5685
      Left            =   300
      TabIndex        =   13
      Top             =   4020
      Visible         =   0   'False
      Width           =   14295
      _cx             =   25215
      _cy             =   10028
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
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
      Editable        =   0
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   6780
      TabIndex        =   11
      Top             =   180
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1395
      Left            =   270
      TabIndex        =   8
      Top             =   900
      Width           =   8385
      Begin MSComCtl2.DTPicker dtAwal 
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Top             =   780
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   137953283
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker dtAkhir 
         Height          =   315
         Left            =   3900
         TabIndex        =   2
         Top             =   780
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   137953283
         CurrentDate     =   37810
      End
      Begin VB.Line Line1 
         X1              =   3840
         X2              =   8220
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label LblCust 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   4335
      End
      Begin MSForms.ComboBox CboSupplier 
         Height          =   315
         Left            =   1740
         TabIndex        =   0
         Top             =   315
         Width           =   1875
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   7
         Size            =   "3307;556"
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
         Caption         =   "Supplier Code"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   255
         Left            =   3390
         TabIndex        =   10
         Top             =   810
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Date"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   855
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3180
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   270
      TabIndex        =   6
      Top             =   2445
      Width           =   8385
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
         TabIndex        =   7
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Preview"
      Height          =   375
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3180
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AP Detail List Report"
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
      Left            =   2985
      TabIndex        =   5
      Top             =   210
      Width           =   2385
   End
End
Attribute VB_Name = "FrmAPListReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ClsProc As New ClsProc
Dim nilKosong As Boolean, i As Integer

Sub Kosong()
    dtAwal = Format(Year(Now) & "-" & Format(Month(Now), "#0") & "-01", "dd MMM yyyy")
    dtAkhir = Format(Now, "dd MMM yyyy")
    
    Call IsiCboSuppl
End Sub

Sub IsiCboSuppl() 'Filter Request No
Dim rscbo As New ADODB.Recordset 'Data Customer
Dim strSQL As String

With cboSupplier
    .clear
    .columnCount = 2
    .ColumnWidths = "100pt;300pt"
    .ListWidth = 400
    .ListRows = 10
    
    strSQL = "select Trade_Code,Trade_Name From Trade_Master Where Trade_Cls in ('2','3')"
    
    If rscbo.State <> adStateClosed Then rscbo.Close
    Set rscbo = Db.Execute(strSQL)
    
    .AddItem ""
    .List(0, 0) = strAll
    .List(0, 1) = strAll
        
    i = 1
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo(0))
        .List(i, 1) = Trim(rscbo(1))
        i = i + 1
        rscbo.MoveNext
    Loop
        
    Set rscbo = Nothing
    cboSupplier.ListIndex = 0
End With
End Sub

Private Sub CboSupplier_Change()
    Call cbosupplier_Click
End Sub

Private Sub cbosupplier_Click()
    If cboSupplier.ListIndex < 0 Then
        cboSupplier.ListIndex = 0
        lblcust = strAll
    Else
        lblcust = cboSupplier.Column(1)
    End If
End Sub

Private Sub Command1_Click()
    Dim AdoExcel As New ADODB.Recordset
    Dim strSQL As String
    
   ' On Error GoTo ErrExcel
    Screen.MousePointer = vbHourglass
    LblErrMsg = ""
    
'    Strsql = " Select PR.Item_Code,IM.Item_Name, TM.Trade_Name,Isnull(ISM.BL_NO,'') Voucher_No,Isnull(ISM.VoucherDesc,'') Voucher_Desc, " & vbCrLf & _
'                      "     isnull(ISD.Invoice_No,'') Invoice_NO, ISM.Invoice_Date, Ism.FakturPajak_No, PR.SuratJalan_No Do_No, PR.Receipt_Date DO_Date, " & vbCrLf & _
'                      "     PR.Qty,(select Description from Unit_Cls U where U.Unit_Cls=PR.Unit_Cls) Unit, " & vbCrLf & _
'                      "     PR.Price,(select Description from Curr_Cls C where C.Curr_Cls=PR.Currency_Code) Currency,  " & vbCrLf & _
'                      "     isnull((Select Price_Adj from PurchaseOrder_Detail POD Where POD.PO_No=PR.Po_No  " & vbCrLf & _
'                      "         And POD.Item_Code=PR.Item_Code) ,0) Adj_Price , " & vbCrLf & _
'                      "     isnull((Select Price_Adj from PurchaseOrder_Detail POD Where POD.PO_No=PR.Po_No  " & vbCrLf & _
'                      "         And POD.Item_Code=PR.Item_Code) ,PR.Price) * PR.Qty Amount_Price , " & vbCrLf & _
'                      "     pr.BC_Type, Isnull(PR.BC40_No,'') BC40, BC40_Date,PR.PO_No, " & vbCrLf & _
'                      "     Case When PR.Currency_Code='03' Then PR.Amount else Null End AmountIDR, " & vbCrLf & _
'                      "     Case When PR.Currency_Code='02' Then PR.Amount else Null End AmountUSD, " & vbCrLf
'
'    Strsql = Strsql + "     Case When PR.Currency_Code='01' Then PR.Amount else Null End AmountJPY " & vbCrLf & _
'                      " From Part_Receipt PR " & vbCrLf & _
'                      " Inner Join Item_Master IM on IM.Item_Code=PR.Item_Code " & vbCrLf & _
'                      " Inner Join Trade_Master TM on TM.Trade_Code=PR.Supplier_Code " & vbCrLf & _
'                      " Left Join InvoiceSupplier_Detail ISD " & vbCrLf & _
'                      " Left Join InvoiceSupplier_Master ISM on ISD.Invoice_No=ISM.Invoice_No  " & vbCrLf & _
'                      "     on ISD.DO_No=PR.SuratJalan_No And ISD.PO_No=PR.PO_No And ISD.Item_Code=PR.Item_Code " & vbCrLf & _
'                      "     and ISD.ReceiptSeq_No=PR.Seq_No " & vbCrLf & _
'                      " Left Join AP_Detail APD on APD.Invoice_No=ISD.Invoice_NO " & vbCrLf & _
'                      " Left Join AP_Master APM on APD.AP_No=APM.AP_No " & vbCrLf & _
'                      " where receipt_cls='R' " & IIf(Trim(cboSupplier) = strAll, "", " And PR.Supplier_Code='" & Trim(cboSupplier) & "' ") & vbCrLf & _
'                      " And PR.Receipt_Date>='" & Format(dtAwal, "dd-MMM-YYYY") & "' And Receipt_Date <='" & Format(dtAkhir, "dd-MMM-YYYY") & "' " & vbCrLf & _
'                      " order by PR.Supplier_Code, PR.PO_No,PR.SuratJalan_No "
                      
                              
'        strSQL = "  Select PR.Item_Code,IM.Item_Name, Trade_Name=ltrim(rtrim(TM.Trade_Code)) + ' - '  + ltrim(rtrim(TM.Trade_Name)),Isnull(ISM.BL_NO,'') Voucher_No,Isnull(ISM.VoucherDesc,'') Voucher_Desc,  " & vbCrLf & _
'                          "      isnull(ISD.Invoice_No,'') Invoice_NO, ISM.Invoice_Date, ISM.InvoiceReceipt_Date, ISM.Due_Date, Ism.FakturPajak_No, PR.SuratJalan_No Do_No, PR.Receipt_Date DO_Date,  " & vbCrLf & _
'                          "      PR.Qty,(select Description from Unit_Cls U where U.Unit_Cls=PR.Unit_Cls) Unit,  " & vbCrLf & _
'                          "       " & vbCrLf & _
'                          "       " & vbCrLf & _
'                          "      case when (Select Price_Adj from PurchaseOrder_Detail POD Where POD.PO_No=PR.Po_No And POD.Item_Code=PR.Item_Code) > 0 then  " & vbCrLf & _
'                          "             (Select Price_Adj from PurchaseOrder_Detail POD Where POD.PO_No=PR.Po_No  And POD.Item_Code=PR.Item_Code) " & vbCrLf & _
'                          "      else pr.Price end Price, " & vbCrLf & _
'                          "       " & vbCrLf & _
'                          "      (select Description from Curr_Cls C where C.Curr_Cls=PR.Currency_Code) Currency,   " & vbCrLf & _
'                          "      isnull((Select Price_Adj from PurchaseOrder_Detail POD Where POD.PO_No=PR.Po_No   "
'
'        strSQL = strSQL + "          And POD.Item_Code=PR.Item_Code) ,0) Adj_Price ,  " & vbCrLf & _
'                          "      case when (Select Price_Adj from PurchaseOrder_Detail POD Where POD.PO_No=PR.Po_No And POD.Item_Code=PR.Item_Code) > 0 then  " & vbCrLf & _
'                          "             (Select Price_Adj from PurchaseOrder_Detail POD Where POD.PO_No=PR.Po_No  And POD.Item_Code=PR.Item_Code) " & vbCrLf & _
'                          "      else pr.Price end  " & vbCrLf & _
'                          "  " & vbCrLf & _
'                          "           " & vbCrLf & _
'                          "           * PR.Qty Amount_Price ,  " & vbCrLf & _
'                          "      pr.BC_Type, Isnull(PR.BC40_No,'') BC40, BC40_Date,PR.PO_No,  " & vbCrLf & _
'                          "      Case When PR.Currency_Code='03' Then PR.Amount else  PR.Amount * (select Daily_ExchangeRate From Daily_ExchangeRate where ExchangeRate_Date=ISM.Invoice_Date and Currency_Code='02') End AmountIDR,  " & vbCrLf & _
'                          "      Case When PR.Currency_Code='02' Then PR.Amount else PR.Amount / (select Daily_ExchangeRate From Daily_ExchangeRate where ExchangeRate_Date=ISM.Invoice_Date and Currency_Code='02') End AmountUSD,  " & vbCrLf & _
'                          "      Case When PR.Currency_Code='01' Then PR.Amount else Null End AmountJPY  "
'
'        strSQL = strSQL + "  From Part_Receipt PR  " & vbCrLf & _
'                          "  Inner Join Item_Master IM on IM.Item_Code=PR.Item_Code  " & vbCrLf & _
'                          "  Inner Join Trade_Master TM on TM.Trade_Code=PR.Supplier_Code  " & vbCrLf & _
'                          "  Left Join InvoiceSupplier_Detail ISD  " & vbCrLf & _
'                          "  Left Join InvoiceSupplier_Master ISM on ISD.Invoice_No=ISM.Invoice_No   " & vbCrLf & _
'                          "      on ISD.DO_No=PR.SuratJalan_No And ISD.PO_No=PR.PO_No And ISD.Item_Code=PR.Item_Code  " & vbCrLf & _
'                          "      and ISD.ReceiptSeq_No=PR.Seq_No  " & vbCrLf & _
'                          "  Left Join AP_Detail APD on APD.Invoice_No=ISD.Invoice_NO  " & vbCrLf & _
'                          "  Left Join AP_Master APM on APD.AP_No=APM.AP_No  " & vbCrLf & _
'                        " where receipt_cls='R' " & IIf(Trim(cbosupplier) = strAll, "", " And PR.Supplier_Code='" & Trim(cbosupplier) & "' ") & vbCrLf & _
'                        " And PR.Receipt_Date>='" & Format(dtAwal, "yyyy-mm-dd") & "' And Receipt_Date <='" & Format(dtAkhir, "yyyy-mm-dd") & "' " & vbCrLf & _
'                        " order by PR.Supplier_Code, PR.PO_No,PR.SuratJalan_No "
                      
                      
    strSQL = "EXEC sp_IncomingMaterialReport_Sel '" & Trim(cboSupplier) & "', '" & Format(dtAwal, "yyyy-mm-dd") & "', '" & Format(dtAkhir, "yyyy-mm-dd") & "'  "

'    If rsCek.State <> adStateClosed Then rsCek.Close
'        rsCek.Ope
'        n sql, Db, adOpenForwardOnly, adLockReadOnly
        
    If AdoExcel.State <> adStateClosed Then AdoExcel.Close
    Set AdoExcel = Db.Execute(strSQL)
    
    If AdoExcel.EOF Then
        LblErrMsg = DisplayMsg(13)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Dim xlapp As New Excel.application
    Dim Baris As Integer, barisawal As Byte
    Dim Idx As Long
    Dim VSupplier As String
    Dim jumlahQty As Double
    Dim jumlahAmountIDR As Double
    Dim jumlahAmountUSD As Double
    Dim jumlahAmountJPY As Double
    
    With xlapp
        .Workbooks.Add
        '.Visible = True
       
        .Range("A1") = "PT. KAWAI INDONESIA PLANT-3"
        .Range("A1").verticalAlignment = xlCenter
        .Range("A1").Columns.Font.Name = "Arial"
        .Range("A1").Columns.Font.Size = 10
        .Range("A1").Columns.Font.Bold = True
    
        .Range("A2") = "Incoming Material"
        .Range("A2").verticalAlignment = xlCenter
        .Range("A2").Columns.Font.Name = "Arial"
        .Range("A2").Columns.Font.Size = 8
        .Range("A2").Columns.Font.Bold = True
    
        .Range("A3") = "Period : " & Format(dtAwal, "dd MMM YYYY") & " to " & Format(dtAkhir, "dd MMM YYYY")
        .Range("A3").verticalAlignment = xlCenter
        .Range("A3").Columns.Font.Name = "Arial"
        .Range("A3").Columns.Font.Size = 8
        .Range("A3").Columns.Font.Bold = True
                
        barisawal = 5
        
        .Range("A" & barisawal) = "Item Code" '0
        .Range("B" & barisawal) = "Description" '1
        .Range("C" & barisawal) = "Invoice No." ' 2
        .Range("D" & barisawal) = "Invoice Date" '3
        .Range("E" & barisawal) = "Faktur Pajak No." '4
        .Range("F" & barisawal) = "Surat Jalan No." '5
        .Range("G" & barisawal) = "Factory" '6
        .Range("H" & barisawal) = "Delivery Date" '7
        .Range("I" & barisawal) = "Faktur Pajak Date" '8
        .Range("J" & barisawal) = "Qty" '9
        .Range("K" & barisawal) = "Unit" '"Qty": .Range("K" & barisawal, "M" & barisawal).Merge '10
        .Range("L" & barisawal) = "Original Curr" '11
        .Range("M" & barisawal) = "Original Price" '12
        .Range("N" & barisawal) = "Amount" ': .Range("N" & barisawal, "K" & barisawal).Merge '13
        'Original Amoun
        .Range("O" & barisawal) = "Curr Symbol" '14
        .Range("P" & barisawal) = "Amount USD" '15
        .Range("Q" & barisawal) = "Amount IDR" '16
        .Range("R" & barisawal) = "Exchange Rate" '17
        .Range("S" & barisawal) = "PO No." '18
        .Range("T" & barisawal) = "BC No." '19
        .Range("U" & barisawal) = "Tgl BC" '20
        .Range("V" & barisawal) = "BC Type" '21
        .Range("W" & barisawal) = "Price Budget" '22
        .Range("X" & barisawal) = "Kakakusai" '23
        .Range("Y" & barisawal) = "PO Type" '24
        .Range("Z" & barisawal) = "No. Seri" '25
        
        
        .Range("AA" & barisawal) = "Commercial Cls" '26
        .Range("AB" & barisawal) = "Vouvcher No." '27
        .Range("AC" & barisawal) = "Remarks" '28
                    
        .Range("A" & barisawal, "AC" & barisawal).verticalAlignment = xlCenter
        .Range("A" & barisawal, "AC" & barisawal).horizontalAlignment = xlCenter
        
        .Range("A" & barisawal, "AC" & barisawal).Columns.Font.Name = "Arial"
        .Range("A" & barisawal, "AC" & barisawal).Columns.Font.Size = 8
        .Range("A" & barisawal, "AC" & barisawal).Columns.Font.Bold = True
                
        Baris = barisawal + 1
        Idx = 1
        
        Do While Not AdoExcel.EOF
        
            VSupplier = AdoExcel("Trade_Name")
            .Range("A" & Baris) = Trim(VSupplier)
            .Range("A" & Baris, "B" & Baris).Merge
            
            .Range("A" & Baris, "B" & Baris).Select
             .Range("a" & Baris, "AC" & Baris).Interior.ColorIndex = 37
             
    
            
            Baris = Baris + 1
            
            Do While Not AdoExcel.EOF And AdoExcel("Trade_Name") = VSupplier
    
                jumlahQty = jumlahQty + AdoExcel("Qty")
                If IsNull(jumlahAmountUSD = jumlahAmountUSD + AdoExcel("AmountUSD")) Then
                    jumlahAmountUSD = 0
                    Else
                    jumlahAmountUSD = jumlahAmountUSD + AdoExcel("AmountUSD")
                End If
                
                If IsNull(jumlahAmountIDR = jumlahAmountIDR + AdoExcel("AmountIDR")) Then
                    jumlahAmountIDR = 0
                    Else
                    jumlahAmountIDR = jumlahAmountIDR + AdoExcel("AmountIDR")
                End If
                
                If IsNull(jumlahAmountJPY = jumlahAmountJPY + AdoExcel("AmountJPY")) Then
                    jumlahAmountJPY = 0
                    Else
                    jumlahAmountJPY = jumlahAmountJPY + AdoExcel("AmountJPY")
                End If
                
                .Range("A" & Baris) = Trim(AdoExcel("Item_Code")) '0
                .Range("B" & Baris) = Trim(AdoExcel("Item_Name")) '0
                .Range("C" & Baris) = Trim(AdoExcel("Invoice_No")) '1
                .Range("D" & Baris) = AdoExcel("Invoice_Date") '2
                .Range("E" & Baris) = Trim(AdoExcel("FakturPajak_No")) '3
                                
                .Range("F" & Baris) = "'" & Trim(AdoExcel("Do_No")) '4
                .Range("G" & Baris) = Trim(AdoExcel("Factory"))
                .Range("H" & Baris) = Format(AdoExcel("Do_Date"), "dd-MMM-yyyy") '5
                .Range("I" & Baris) = Format(AdoExcel("FakturPajak_Date"), "dd-MMM-yyyy") '6
                .Range("J" & Baris) = AdoExcel("Qty") '7
                .Range("K" & Baris) = Trim(AdoExcel("Unit"))
                .Range("L" & Baris) = Trim(AdoExcel("Currency"))
                .Range("M" & Baris) = AdoExcel("Price")
                .Range("N" & Baris) = AdoExcel("Amount")
                .Range("O" & Baris) = "USD"
                
                If IsNull(.Range("P" & Baris) = AdoExcel("AmountUSD")) Then
                    .Range("P" & Baris) = 0
                Else
                    .Range("P" & Baris) = AdoExcel("AmountUSD")
                End If
                
                If IsNull(.Range("Q" & Baris) = AdoExcel("AmountIDR")) Then
                    .Range("Q" & Baris) = "0"
                 Else
                    .Range("Q" & Baris) = AdoExcel("AmountIDR")
                End If
                
'                If IsNull(.Range("R" & Baris) = AdoExcel("AmountJPY")) Then
'                    .Range("R" & Baris) = "0"
'                 Else
'                    .Range("R" & Baris) = AdoExcel("AmountJPY")
'                End If
                .Range("R" & Baris) = AdoExcel("ExchangeRate")
                .Range("S" & Baris) = Trim(AdoExcel("PO_No"))
                .Range("T" & Baris) = Trim("'" & AdoExcel("BC40") & "")
                
                If IsNull(AdoExcel("BC40_Date")) Then
                    .Range("U" & Baris) = ""
                Else
                    .Range("U" & Baris) = Format(AdoExcel("BC40_Date"), "dd-MMM-yyyy")
                End If
                
                .Range("V" & Baris) = Trim("'" & AdoExcel("BC_TYPE") & "")
                .Range("W" & Baris) = ""
                .Range("X" & Baris) = ""
                .Range("Y" & Baris) = Trim(AdoExcel("POType_Cls"))
                .Range("Z" & Baris) = ""
                .Range("AA" & Baris) = ""
                .Range("AB" & Baris) = Trim(AdoExcel("Voucher_No"))
                .Range("AC" & Baris) = Trim(AdoExcel("Voucher_Desc"))
                
               
                AdoExcel.MoveNext
                Baris = Baris + 1
                Idx = Idx + 1
                If AdoExcel.EOF Then Exit Do
            Loop
            
            .Range("A" & Baris) = "Sub Total"
'            .Range("H" & Baris).NumberFormat = "DD MMM YYYY"
'            .Range("I" & Baris).NumberFormat = "DD MMM YYYY"
            .Range("J" & Baris) = jumlahQty
            
            .Range("P" & Baris) = jumlahAmountUSD
            .Range("Q" & Baris) = jumlahAmountIDR
'            .Range("R" & Baris) = jumlahAmountJPY
        
            Baris = Baris + 1
            Idx = Idx + 1
            jumlahQty = 0
            jumlahAmountIDR = 0
            jumlahAmountUSD = 0
            jumlahAmountJPY = 0
            
            
        Loop
        .Range("A" & barisawal + 1, "C" & Baris).horizontalAlignment = xlLeft
        .Range("E" & barisawal + 1, "I" & Baris).horizontalAlignment = xlLeft
        .Range("Y" & barisawal + 1, "AC" & Baris).horizontalAlignment = xlLeft
        .Range("S" & barisawal + 1, "V" & Baris).horizontalAlignment = xlLeft
        .Range("J" & barisawal + 1, "J" & Baris).horizontalAlignment = xlRight
        .Range("M" & barisawal + 1, "R" & Baris).horizontalAlignment = xlRight
        .Range("O" & barisawal + 1, "O" & Baris).horizontalAlignment = xlLeft
'        .Range("S" & barisawal + 1, "S" & Baris).horizontalAlignment = xlLeft
        
        .WindowState = xlMaximized
        .Range("a" & barisawal, "AC" & Baris).Columns.Font.Name = "Arial"
        .Range("a" & barisawal, "AC" & Baris).Columns.Font.Size = 8

        .Range("A" & barisawal, "AC" & Baris).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("A" & barisawal, "AC" & Baris).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A" & barisawal, "AC" & Baris).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("A" & barisawal, "AC" & Baris).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("A" & barisawal, "AC" & Baris).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range("A" & barisawal, "AC" & Baris).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'
'        .Range("W" & barisawal + 2, "AC" & Baris).NumberFormat = "DD MMM YYYY"
''        .Range("E" & barisawal + 2, "J" & Baris).NumberFormat = "DD MMM YYYY"
''        .Range("W" & barisawal + 2, "X" & Baris).NumberFormat = "DD MMM YYYY"
'
'        '.Range("I" & barisawal + 2, "I" & Baris).NumberFormat = gs_formatQty
'
        .Range("N" & barisawal + 2, "S" & Baris).NumberFormat = "#,##0.00000"
'        .Range("M" & barisawal + 2, "O" & Baris).NumberFormat = "#,##0.00000"
'        .Range("M" & barisawal + 2, "O" & Baris).NumberFormat = "#,##0.00000"
        
        .Range("A" & barisawal, "AC" & Baris).Columns.AutoFit
        
        .Visible = True
        
    End With

    Screen.MousePointer = vbDefault
    Set AdoExcel = Nothing
    LblErrMsg.Caption = DisplayMsg("9008")
    Exit Sub
    
ErrExcel:
    Screen.MousePointer = vbDefault
    Set AdoExcel = Nothing
    LblErrMsg = err.Description

End Sub

Private Sub Form_Load()
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

    dtAkhir.Value = Now
    dtAwal.Value = Now
    
    Call Kosong
End Sub

Private Sub dtAwal_Change()
If CDate(dtAwal) > CDate(dtAkhir) Then
    LblErrMsg.Caption = "Start Date must be lower than " & dtAkhir.Value & " !!!"
    Exit Sub
ElseIf CDate(dtAkhir) < CDate(dtAwal) Then
    LblErrMsg.Caption = "End Date must be higher than " & dtAwal.Value & " !!!"
    Exit Sub
End If

    If Format(dtAwal, "yyyy-MM-dd") > Format(dtAkhir, "yyyy-MM-dd") Then LblErrMsg = DisplayMsg(4068) & " " & Format(dtAkhir, "dd MMM yyyy") Else LblErrMsg = ""
End Sub

Private Sub dtAkhir_Change()
If CDate(dtAwal) > CDate(dtAkhir) Then
    LblErrMsg.Caption = "Start Date must be lower than " & dtAkhir.Value & " !!!"
    Exit Sub
ElseIf CDate(dtAkhir) < CDate(dtAwal) Then
    LblErrMsg.Caption = "End Date must be higher than " & dtAwal.Value & " !!!"
    Exit Sub
End If
    
    If Format(dtAwal, "yyyy-MM-dd") > Format(dtAkhir, "yyyy-MM-dd") Then LblErrMsg = DisplayMsg(4066) & " " & Format(dtAwal, "dd MMM yyyy") Else LblErrMsg = ""
End Sub

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
    If ErrMsg = "" Then Unload Me Else LblErrMsg.Caption = ErrMsg
End Sub


