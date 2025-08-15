VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptLedgerAP 
   BackColor       =   &H00FDDFE3&
   Caption         =   "General / Sub Ledger - AP"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRptLedgerAP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1575
      Left            =   398
      TabIndex        =   10
      Top             =   1710
      Width           =   8805
      Begin VB.TextBox lblNm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   0
         Left            =   4005
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   380
         Width           =   4620
      End
      Begin MSComCtl2.DTPicker dt 
         Height          =   315
         Left            =   1965
         TabIndex        =   1
         Top             =   960
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
         Format          =   294125571
         UpDown          =   -1  'True
         CurrentDate     =   37985
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": "
         Height          =   195
         Index           =   1
         Left            =   1695
         TabIndex        =   15
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period "
         Height          =   195
         Left            =   270
         TabIndex        =   14
         Top             =   990
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   390
         Width           =   1845
      End
      Begin MSForms.ComboBox Cbo 
         Height          =   315
         Index           =   0
         Left            =   1965
         TabIndex        =   0
         Top             =   360
         Width           =   1890
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3334;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         X1              =   4005
         X2              =   8500
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": "
         Height          =   195
         Index           =   0
         Left            =   1695
         TabIndex        =   12
         Top             =   390
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdsubmenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   398
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4005
      Width           =   1140
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xcel"
      Height          =   375
      Left            =   8063
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4005
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   525
      Left            =   398
      TabIndex        =   8
      Top             =   3375
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
         TabIndex        =   9
         Top             =   180
         Width           =   8535
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   735
      Left            =   398
      TabIndex        =   7
      Top             =   885
      Width           =   8805
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FDDFE3&
         Caption         =   "General Ledger"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Sub Ledger"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   7365
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   300
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "General / Sub Ledger - AP"
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
      Index           =   1
      Left            =   405
      TabIndex        =   16
      Top             =   315
      Width           =   8805
   End
End
Attribute VB_Name = "frmRptLedgerAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim TampungDt As Byte
Dim i As Integer, j As Integer
Dim brs As Integer, Col As Integer
Dim tampungBrs As Integer

Dim rsLedger As New ADODB.Recordset
Dim excelapp As New Excel.application
Dim strCurrRate As String, totCurrRate As Integer

Dim sqlcust As String
Dim tglAkhir As String, kurang As Integer, arrCol As Integer

Dim Opening(6) As Integer, sales(5) As Integer, colFP As Integer, collection(5) As Integer, Ending(5) As Integer
Dim TotOpening As Double, TotFP As Double, TotCollection As Double, TotEnding As Double
Dim GrandTotOpening(5) As Double, GrandTotsales(5) As Double, GrandtotFP As Double, GrandTotcollection(5) As Double, GrandTotEnding(5) As Double

Dim custCD As String, nilFormat As String, Curr As Integer

'************************* Isi Cbo *****************************
Sub isiCboCust() 'Isi Combo Dealer CD dr Supplier Master
Dim RsCust As New ADODB.Recordset 'Data Supplier

With Cbo(0)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    '******** Ambil dr Supplier Master utk Combo Dealer CD
    sql = "select Trade_Code,Trade_Name from Trade_Master " & _
        "where Trade_Cls = 2 order by Trade_Code"
    Set RsCust = Db.Execute(sql)
    
    .AddItem ""
    .List(0, 0) = strAll
    .List(0, 1) = strAll
    
    i = 1
    Do While Not (RsCust.EOF)
        .AddItem ""
        .List(i, 0) = Trim(RsCust(0))
        .List(i, 1) = Trim(RsCust(1))
        i = i + 1
        RsCust.MoveNext
    Loop
    
    .ListIndex = 0
    .ListWidth = 350
    .ListRows = 20
    .ColumnWidths = "50 pt;300 pt"
    
    Set RsCust = Nothing
End With
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    lblNm(0) = ""
    dt = Format(Now, "MMM yyyy")
    TampungDt = Month(dt)
    Option1(0).Value = True
    
    Call isiCboCust
End Sub

Sub resetArr(Optional resetAll As Byte)
    For i = 0 To 5
        If resetAll = 1 Then
            Opening(i) = 0: sales(i) = 0: collection(i) = 0: colFP = 0: Ending(i) = 0
            GrandTotOpening(i) = 0: GrandTotsales(i) = 0: GrandTotcollection(i) = 0: GrandTotEnding(i) = 0: GrandtotFP = 0
        End If
        TotOpening = 0: TotFP = 0: TotCollection = 0: TotEnding = 0
    Next i
End Sub

Function sqlMaster(st As Byte) As String
Dim sqlM As String, sqlMAR As String
    
    If st = 0 Then 'Curr
        sqlM = "distinct CurrCD = ISNULL((Select Top 1 Case Currency_Code When '03' then '00' else Currency_Code End From InvoiceSupplier_Detail Where Invoice_No = INV.Invoice_No),'') "
        sqlMAR = "distinct CurrCD  = Case AD.Currency_Code When '03' then '00' else AD.Currency_Code End  "
    
    Else 'Inv
        sqlM = "InvNo = Invoice_No "
        sqlMAR = sqlM
    End If
    
    sql = "( Select " & sqlM & "From InvoiceSupplier_Master INV " & _
                "Where month(invoice_date)= " & Month(dt) & " and year(invoice_date) = " & Year(dt) & sqlcust & _
            " UNION " & _
            "Select " & sqlMAR & "From AP_Master INV, AP_Detail AD " & _
                "Where INV.AP_No = AD.AP_No And month(AP_Date)= " & Month(dt) & " and year(AP_Date) =" & Year(dt) & sqlcust & _
            " UNION " & _
            "Select " & sqlM & "From " & _
                "( " & _
                    "Select INV.invoice_no, " & _
                        "VAT = ISNULL(INV.PPN,0), " & _
                        "Total_amount = INV.Total_amount , " & _
                        "CollectionBefore = ISNULL((Select SUM(Amount) From AP_Master AM, AP_Detail AD " & _
                                            "Where AM.Supplier_Code = INV.Supplier_Code And AM.AP_No = AD.AP_No And AD.Invoice_No = INV.Invoice_No " & _
                                                "And AM.AP_Date < '" & Format(dt, "yyyy-MM-01") & "'),0), " & _
                        "CollectionPPNBefore = ISNULL((Select SUM(PPN) From AP_Master AM, AP_Detail AD " & _
                                            "Where AM.Supplier_Code = INV.Supplier_Code And AM.AP_No = AD.AP_No And AD.Invoice_No = INV.Invoice_No " & _
                                                "And AM.AP_Date < '" & Format(dt, "yyyy-MM-01") & "'),0) "

    sql = sql & _
                    "from InvoiceSupplier_Master INV " & _
                    "where invoice_date < '" & Format(dt, "yyyy-MM-01") & "' " & sqlcust & _
                        " and (INV.Paid_Date >= '" & Format(dt, "yyyy-MM-01") & "' Or INV.Paid_Date IS NULL) " & _
                ") INV Where (Total_amount - CollectionBefore > 0 Or VAT - CollectionPPNBefore > 0) "
    
    If st = 0 Then sql = sql & "UNION Select CurrCD = '00' "
    
    sql = sql & ")Master "
    sqlMaster = sql
End Function

Function currRate() As String
Dim sqlrate As String
Dim rsrate As New ADODB.Recordset

Dim monthRate As String
    If Len(Month(dt)) = 1 Then monthRate = Format(Month(dt), "0#") Else monthRate = Format(Month(dt), "00#")

    sqlrate = "Select * From " & sqlMaster(0) & " Where CurrCD <> '' order by Master.CurrCD Desc"
    Set rsrate = Db.Execute(sqlrate)
    currRate = ""

    i = 9
    Do While Not rsrate.EOF
        currRate = currRate & IIf(rsrate!currCD = "00", "03", rsrate!currCD) & ","
        i = i + 1
        rsrate.MoveNext
    Loop

    If currRate <> "" Then currRate = Left(Trim(currRate), Len(Trim(currRate)) - 1)
    Set rsrate = Nothing
End Function

Sub headerExcel()
Dim clsMRP As New clsMRP
    
    tglAkhir = Format(DateAdd("d", -1, DateAdd("m", 1, dt)), "dd MMMM yyyy")
    With excelapp
        'HEADER
        .Workbooks.Add
                
        strCurrRate = currRate
        totCurrRate = UBound(Split(strCurrRate, ",")) + 1
        
        If Option1(0).Value = True Then
            .Cells(2, 1) = "General Ledger - AP "
            kurang = 4
        Else
            .Cells(1, 1) = clsMRP.CompanyName
            .Cells(2, 1) = "Sub Ledger - AP "
            kurang = 0
        End If
        
        .Cells(3, 1) = "Period End : " & tglAkhir
        .Range("A1", "A3").Font.Bold = True
        
        .Cells(5, 7 - kurang) = "Outstanding till end of Last Month"
        .Range(.Cells(5, 7 - kurang), .Cells(5, 6 + totCurrRate - kurang)).Merge
        
        .Cells(5, 7 + totCurrRate - kurang) = "Purchase Invoice Current"
        .Range(.Cells(5, 7 + totCurrRate - kurang), .Cells(5, 6 + (totCurrRate * 2) - kurang)).Merge
        
        .Cells(5, 7 + (totCurrRate * 2) - kurang) = "FP Amount"
        
        .Cells(5, 8 + (totCurrRate * 2) - kurang) = "Payment Current Month"
        .Range(.Cells(5, 8 + (totCurrRate * 2) - kurang), .Cells(5, 7 + (totCurrRate * 3) - kurang)).Merge
        
        .Cells(5, 8 + (totCurrRate * 3) - kurang) = "Total Outstanding at End Current Month"
        .Range(.Cells(5, 8 + (totCurrRate * 3) - kurang), .Cells(5, 7 + (totCurrRate * 4) - kurang)).Merge
        
        .Cells(6, 1) = "Supplier CD"
        .Range(.Cells(5, 1), .Cells(7, 1)).Merge
        .Range(.Cells(5, 1), .Cells(7, 1)).VerticalAlignment = xlCenter
        
        .Cells(6, 2) = "Supplier Name"
        .Range(.Cells(5, 2), .Cells(7, 2)).Merge
        .Range(.Cells(5, 2), .Cells(7, 2)).VerticalAlignment = xlCenter
        
        If Option1(1).Value = True Then
            .Cells(6, 3) = "Invoice No"
            .Range(.Cells(5, 3), .Cells(7, 3)).Merge
            .Range(.Cells(5, 3), .Cells(7, 3)).VerticalAlignment = xlCenter
            
            .Cells(6, 4) = "Invoice Date"
            .Range(.Cells(5, 4), .Cells(7, 4)).Merge
            .Range(.Cells(5, 4), .Cells(7, 4)).VerticalAlignment = xlCenter
            
            .Cells(6, 5) = "Due Date"
            .Range(.Cells(5, 5), .Cells(7, 5)).Merge
            .Range(.Cells(5, 5), .Cells(7, 5)).VerticalAlignment = xlCenter
            
            .Cells(6, 6) = "Bank Name"
            .Range(.Cells(5, 6), .Cells(7, 6)).Merge
            .Range(.Cells(5, 6), .Cells(7, 6)).VerticalAlignment = xlCenter
        End If
                        
        .Cells(6, 7 - kurang) = "Opening Balance"
        .Range(.Cells(6, 7 - kurang), .Cells(6, 6 + totCurrRate - kurang)).Merge
        
        .Cells(6, 7 + totCurrRate - kurang) = "Purchase In Month"
        .Range(.Cells(6, 7 + totCurrRate - kurang), .Cells(6, 7 + (totCurrRate * 2) - kurang)).Merge
        
        .Cells(6, 8 + (totCurrRate * 2) - kurang) = "Payment"
        
        If totCurrRate = 1 Then
            .Cells(6, 8 + (totCurrRate * 3) - kurang) = "Ending Balance"
        Else
            .Cells(6, 9 + (totCurrRate * 3) - kurang) = "Ending Balance"
            .Range(.Cells(6, 8 + (totCurrRate * 2) - kurang), .Cells(6, 7 + (totCurrRate * 3) - kurang)).Merge
            .Range(.Cells(6, 8 + (totCurrRate * 3) - kurang), .Cells(6, 7 + (totCurrRate * 4) - kurang)).Merge
        End If
        
        .Range(.Cells(5, 1), .Cells(7, 7 + (totCurrRate * 4) - kurang)).HorizontalAlignment = xlCenter
        
        '************* Opening Balance ***************
        For i = 0 To 3
            For j = 0 To totCurrRate - 1
                If i <= 1 Then
                    .Cells(7, 7 + (totCurrRate * i) + j - kurang) = uf_GetCurrencyDescription(Trim(Split(strCurrRate, ",")(j)))
                Else
                    .Cells(7, 8 + (totCurrRate * i) + j - kurang) = uf_GetCurrencyDescription(Trim(Split(strCurrRate, ",")(j)))
                End If
                
                If i = 1 And j = totCurrRate - 1 Then
                    .Cells(7, 8 + (totCurrRate * i) + j - kurang) = "VAT (IDR)"
                End If
                
                If i = 0 Then
                    Opening(CInt(Split(strCurrRate, ",")(j))) = 7 + j - kurang
                ElseIf i = 1 Then
                    sales(CInt(Split(strCurrRate, ",")(j))) = 7 + (totCurrRate * i) + j - kurang
                    If j = totCurrRate - 1 Then colFP = 8 + (totCurrRate * i) + j - kurang
                ElseIf i = 2 Then
                    collection(CInt(Split(strCurrRate, ",")(j))) = 8 + (totCurrRate * i) + j - kurang
                ElseIf i = 3 Then
                    Ending(CInt(Split(strCurrRate, ",")(j))) = 8 + (totCurrRate * i) + j - kurang
                End If
            Next j
        Next i
        '*********************************************
    End With
End Sub

Sub bottom()
With excelapp
    '********************** Format Angka *******************
    .Range(.Cells(8, 7 - kurang), .Cells(brs, 7 + (totCurrRate * 4) - kurang)).NumberFormat = gs_formatAmount
    .Range(.Cells(8, Opening(3)), .Cells(brs, Opening(3))).NumberFormat = gs_formatAmount
    If sales(3) > 0 Then .Range(.Cells(8, sales(3)), .Cells(brs, sales(3))).NumberFormat = gs_formatAmount
    .Range(.Cells(8, colFP), .Cells(brs, colFP)).NumberFormat = gs_formatAmount
    .Range(.Cells(8, collection(3)), .Cells(brs, collection(3))).NumberFormat = gs_formatAmount
    .Range(.Cells(8, Ending(3)), .Cells(brs, Ending(3))).NumberFormat = gs_formatAmount
    '*******************************************************
    
    .Range(.Cells(4, 1), .Cells(brs, 7 + (totCurrRate * 4) - kurang)).Font.Name = "Arial"
    .Range(.Cells(4, 1), .Cells(brs, 7 + (totCurrRate * 4) - kurang)).Font.Size = 7
    .Range(.Cells(4, 1), .Cells(brs, 7 + (totCurrRate * 4) - kurang)).Columns.AutoFit
    .Range(.Cells(4, 1), .Cells(brs, 7 + (totCurrRate * 4) - kurang)).Columns.HorizontalAlignment = xlLeft
    .Range(.Cells(4, 7), .Cells(brs, 7 + (totCurrRate * 4))).Columns.HorizontalAlignment = xlRight
    
    .Range(.Cells(5, 1), .Cells(brs, 7 + (totCurrRate * 4) - kurang)).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(.Cells(5, 1), .Cells(brs, 7 + (totCurrRate * 4) - kurang)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range(.Cells(5, 1), .Cells(brs, 7 + (totCurrRate * 4) - kurang)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range(.Cells(5, 1), .Cells(brs, 7 + (totCurrRate * 4) - kurang)).Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range(.Cells(5, 1), .Cells(brs, 7 + (totCurrRate * 4) - kurang)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Range(.Cells(5, 1), .Cells(brs, 7 + (totCurrRate * 4) - kurang)).Borders(xlInsideVertical).LineStyle = xlContinuous
    .ActiveSheet.PageSetup.Orientation = 2
    .WindowState = xlMaximized
    .Visible = True
End With
End Sub

Sub isiZero()
With excelapp
    For i = 1 To UBound(Split(strCurrRate, ",")) + 1
        Curr = Split(strCurrRate, ",")(i - 1)
        If i = 3 Then nilFormat = gs_formatAmount Else nilFormat = gs_formatAmount
        If .Cells(brs, Opening(Curr)) = "" Then .Cells(brs, Opening(Curr)) = Format(0, nilFormat)
        If .Cells(brs, sales(Curr)) = "" Then .Cells(brs, sales(Curr)) = Format(0, nilFormat)
        If .Cells(brs, collection(Curr)) = "" Then .Cells(brs, collection(Curr)) = Format(0, nilFormat)
        If .Cells(brs, Ending(Curr)) = "" Then .Cells(brs, Ending(Curr)) = Format(0, nilFormat)
    Next i
    If .Cells(brs, colFP) = "" Then .Cells(brs, colFP) = 0
End With
End Sub

Sub isiSubTotal()
With excelapp
    For i = 1 To UBound(Split(strCurrRate, ",")) + 1
        Curr = Split(strCurrRate, ",")(i - 1)
        If i = 3 Then nilFormat = gs_formatAmount Else nilFormat = gs_formatAmount
        .Range(.Cells(brs, Opening(Curr)), .Cells(brs, Opening(Curr))).FormulaR1C1 = "=SUM(R[" & tampungBrs - brs & "]C:R[-1]C"
        .Range(.Cells(brs, sales(Curr)), .Cells(brs, sales(Curr))).FormulaR1C1 = "=SUM(R[" & tampungBrs - brs & "]C:R[-1]C"
        .Range(.Cells(brs, collection(Curr)), .Cells(brs, collection(Curr))).FormulaR1C1 = "=SUM(R[" & tampungBrs - brs & "]C:R[-1]C"
        .Range(.Cells(brs, Ending(Curr)), .Cells(brs, Ending(Curr))).FormulaR1C1 = "=SUM(R[" & tampungBrs - brs & "]C:R[-1]C"
    Next i
    .Range(.Cells(brs, colFP), .Cells(brs, colFP)).FormulaR1C1 = "=SUM(R[" & tampungBrs - brs & "]C:R[-1]C"

    .Cells(brs, 6 - kurang) = "Sub Total"
    .Range(.Cells(brs, 1), .Cells(brs, Ending(3))).Interior.ColorIndex = 36
    .Range(.Cells(brs, 1), .Cells(brs, 6 - kurang)).Merge
    .Range(.Cells(brs, 1), .Cells(brs, 6 - kurang)).HorizontalAlignment = xlRight
End With
End Sub

Sub isiGrandTotal()
With excelapp
    For i = 1 To UBound(Split(strCurrRate, ",")) + 1
        Curr = Split(strCurrRate, ",")(i - 1)
        If i = 3 Then nilFormat = gs_formatAmount Else nilFormat = gs_formatAmount
        .Cells(brs, Opening(Curr)) = Format(GrandTotOpening(Curr), nilFormat)
        .Cells(brs, sales(Curr)) = Format(GrandTotsales(Curr), nilFormat)
        .Cells(brs, collection(Curr)) = Format(GrandTotcollection(Curr), nilFormat)
        .Cells(brs, Ending(Curr)) = Format(GrandTotEnding(Curr), nilFormat)
    Next i
    .Cells(brs, colFP) = Format(GrandtotFP, gs_formatAmount)
    
    .Cells(brs, 6 - kurang) = "Grand Total"
    .Range(.Cells(brs, 1), .Cells(brs, Ending(3))).Interior.ColorIndex = 36
    .Range(.Cells(brs, 1), .Cells(brs, 6 - kurang)).Merge
    .Range(.Cells(brs, 1), .Cells(brs, 6 - kurang)).HorizontalAlignment = xlRight
End With
End Sub

Function sqlLedger(stGeneral As Byte) As String
    If stGeneral = 1 Then
        sql = "Select Supplier_Code, Trade_Name, CurrCD, " & _
            "OpeningAmount = ISNULL(SUM(OpeningAmount),0), " & _
            "OpeningPPN = ISNULL(SUM(OpeningPPN),0), " & _
            "InvAmount = ISNULL(SUM(InvAmount),0), PPNIDR = ISNULL(SUM(PPNIDR),0)," & _
            "ARAmount = ISNULL(SUM(ARAmount),0), ARPPN = ISNULL(SUM(ARPPN),0), " & _
            "RemainingAmount = ISNULL(SUM(RemainingAmount),0), " & _
            "RemainingPPN = ISNULL(SUM(RemainingPPN),0) " & _
        "From ( "
            
        sql = sql & "Select dt.Supplier_Code, dt.Trade_Name, dt.CurrCD, " & _
                        "OpeningAmount = " & _
                            "Case When dt.CurrCD = '03' Or (Month(Invoice_Date) = " & Month(dt) & " And Year(Invoice_Date) = " & Year(dt) & ") Then 0 " & _
                                "Else InvAmount - ARAmountBefore End, " & _
                        "OpeningPPN = " & _
                            "Case When (Month(Invoice_Date) = " & Month(dt) & " And Year(Invoice_Date) = " & Year(dt) & ") Then 0 " & _
                                "When dt.CurrCD = '03' Then (InvAmount + PPNIDR) - (ARAmountBefore + ARPPNBefore) " & _
                                "Else PPNIDR - ARPPNBefore End, " & _
                        "InvAmount = Case When Month(Invoice_Date) = " & Month(dt) & " And Year(Invoice_Date) = " & Year(dt) & " Then InvAmount Else 0 End, " & _
                        "PPNIDR = Case When Month(Invoice_Date) = " & Month(dt) & " And Year(Invoice_Date) = " & Year(dt) & " Then PPNIDR Else 0 End, " & _
                        "ARAmount = Case dt.CurrCD When '03' Then 0 Else ARAmount End, " & _
                        "ARPPN = Case dt.CurrCD When '03' Then ARAmount + ARPPN Else ARPPN End, " & _
                        "RemainingAmount = " & _
                            "Case When dt.CurrCD <> '03' And (Paid_Date > '" & tglAkhir & "' Or Paid_Date IS NULL) " & _
                                "Then InvAmount - ARAmountBefore - ARAmount Else 0 End, " & _
                        "RemainingPPN = " & _
                            "Case When dt.CurrCD = '03' And (Paid_Date > '" & tglAkhir & "' Or Paid_Date IS NULL) " & _
                                "Then (InvAmount + PPNIDR) - (ARAmountBefore + ARPPNBefore) - (ARAmount + ARPPN) " & _
                            "When dt.CurrCD <> '03' And (Paid_Date > '" & tglAkhir & "' Or Paid_Date IS NULL) " & _
                                "Then PPNIDR - ARPPNBefore - ARPPN Else 0 End "
                            
        sql = sql & "From " & _
                    "( " & _
                        "Select IVM.Supplier_Code, T.Trade_Name, IVM.Invoice_No, IVM.Invoice_Date, " & _
                            "IVM.Due_Date, Paid_Cls = ISNULL(IVM.Paid_Cls,0), Paid_Date," & _
                            "CurrCD = ISNULL((Select Top 1 Currency_Code From InvoiceSupplier_Detail Where Invoice_No = IVM.Invoice_No),''), " & _
                            "InvAmount = IVM.Total_Amount + + IVM.Airfreight_Amount, " & _
                            "PPNIDR = ISNULL(PPN,0), "
                                
        sql = sql & _
                            "ARAmount = ISNULL((Select ISNULL(SUM(APD.Amount),0) From AP_Master APM, AP_Detail APD " & _
                                            "Where APM.Supplier_Code = IVM.Supplier_Code And APM.AP_No = APD.AP_No And APD.Invoice_No = IVM.Invoice_No And Year(APM.AP_Date) = " & Year(dt) & " And Month(APM.AP_Date) = " & Month(dt) & "),0), " & _
                            "ARPPN = ISNULL((Select ISNULL(SUM(ARD.PPN),0) From AP_Master APM, AP_Detail ARD " & _
                                            "Where APM.Supplier_Code = IVM.Supplier_Code And APM.AP_No = ARD.AP_No And ARD.Invoice_No = IVM.Invoice_No And Year(APM.AP_Date) = " & Year(dt) & " And Month(APM.AP_Date) = " & Month(dt) & "),0), " & _
                            "ARAmountBefore = ISNULL((Select ISNULL(SUM(APD.Amount),0) From AP_Master APM, AP_Detail APD " & _
                                            "Where APM.Supplier_Code = IVM.Supplier_Code And APM.AP_No = APD.AP_No And APD.Invoice_No = IVM.Invoice_No And APM.AP_Date < '" & Format(dt, "yyyy-MM-01") & "'),0), " & _
                            "ARPPNBefore = ISNULL((Select ISNULL(SUM(APD.PPN),0) From AP_Master APM, AP_Detail APD " & _
                                            "Where APM.Supplier_Code = IVM.Supplier_Code And APM.AP_No = APD.AP_No And APD.Invoice_No = IVM.Invoice_No And APM.AP_Date < '" & Format(dt, "yyyy-MM-01") & "' " & "),0) " & _
                        "From InvoiceSupplier_Master IVM, Trade_Master T " & _
                        "Where IVM.Supplier_Code = T.Trade_Code " & _
                    ") dt, " & sqlMaster(1) & _
                "Where dt.Invoice_No = Master.InvNo And dt.CurrCD <> '' " & _
                ") data " & _
            "Group By Supplier_Code, Trade_Name, CurrCD " & _
            "Order By Supplier_Code, Trade_Name, CurrCD "
    
    Else
    
        sql = "Select dt.Supplier_Code, dt.Trade_Name, dt.Invoice_No, dt.Invoice_Date, " & _
                "dt.Due_Date, dt.Bank_Name, dt.Paid_Cls, dt.CurrCD, " & _
                "OpeningAmount = " & _
                    "Case When dt.CurrCD = '03' Or (Month(Invoice_Date) = " & Month(dt) & " And Year(Invoice_Date) = " & Year(dt) & ") Then 0 " & _
                        "Else InvAmount - ARAmountBefore End, " & _
                "OpeningPPN = " & _
                    "Case When (Month(Invoice_Date) = " & Month(dt) & " And Year(Invoice_Date) = " & Year(dt) & ") Then 0 " & _
                        "When dt.CurrCD = '03' Then (InvAmount + PPNIDR) - (ARAmountBefore + ARPPNBefore) " & _
                        "Else PPNIDR - ARPPNBefore End, " & _
                "InvAmount = Case When Month(Invoice_Date) = " & Month(dt) & " And Year(Invoice_Date) = " & Year(dt) & " Then InvAmount Else 0 End, " & _
                "PPNIDR = Case When Month(Invoice_Date) = " & Month(dt) & " And Year(Invoice_Date) = " & Year(dt) & " Then PPNIDR Else 0 End, " & _
                "ARAmount = Case dt.CurrCD When '03' Then 0 Else ARAmount End, " & _
                "ARPPN = Case dt.CurrCD When '03' Then ARAmount + ARPPN Else ARPPN End, " & _
                "RemainingAmount = " & _
                    "Case When dt.CurrCD <> '03' And (Paid_Date > '" & tglAkhir & "' Or Paid_Date IS NULL) " & _
                        "Then InvAmount - ARAmountBefore - ARAmount Else 0 End, " & _
                "RemainingPPN = " & _
                    "Case When dt.CurrCD = '03' And (Paid_Date > '" & tglAkhir & "' Or Paid_Date IS NULL) " & _
                        "Then (InvAmount + PPNIDR) - (ARAmountBefore + ARPPNBefore) - (ARAmount + ARPPN) " & _
                    "When dt.CurrCD <> '03' And (Paid_Date > '" & tglAkhir & "' Or Paid_Date IS NULL) " & _
                        "Then PPNIDR - ARPPNBefore - ARPPN Else 0 End "
        sql = sql & _
            "From " & _
                    "( " & _
                        "Select IVM.Supplier_Code, T.Trade_Name, IVM.Invoice_No, IVM.Invoice_Date, " & _
                            "IVM.Due_Date, BM.Bank_Name, " & _
                            "Paid_Cls = ISNULL(IVM.Paid_Cls,0), Paid_Date," & _
                            "CurrCD = ISNULL((Select Top 1 Currency_Code From InvoiceSupplier_Detail Where Invoice_No = IVM.Invoice_No),''), " & _
                            "InvAmount = IVM.Total_Amount + IVM.Airfreight_Amount , " & _
                            "PPNIDR = ISNULL(PPN,0), "
        sql = sql & _
                            "ARAmount = ISNULL((Select ISNULL(SUM(APD.Amount),0) From AP_Master APM, AP_Detail APD " & _
                                            "Where APM.Supplier_Code = IVM.Supplier_Code And APM.AP_No = APD.AP_No And APD.Invoice_No = IVM.Invoice_No And Year(APM.AP_Date) = " & Year(dt) & " And Month(APM.AP_Date) = " & Month(dt) & "),0), " & _
                            "ARPPN = ISNULL((Select ISNULL(SUM(APD.PPN),0) From AP_Master APM, AP_Detail APD " & _
                                            "Where APM.Supplier_Code = IVM.Supplier_Code And APM.AP_No = APD.AP_No And APD.Invoice_No = IVM.Invoice_No And Year(APM.AP_Date) = " & Year(dt) & " And Month(APM.AP_Date) = " & Month(dt) & "),0), " & _
                            "ARAmountBefore = ISNULL((Select ISNULL(SUM(APD.Amount),0) From AP_Master APM, AP_Detail APD " & _
                                            "Where APM.Supplier_Code = IVM.Supplier_Code And APM.AP_No = APD.AP_No And APD.Invoice_No = IVM.Invoice_No And APM.AP_Date < '" & Format(dt, "yyyy-MM-01") & "'),0), " & _
                            "ARPPNBefore = ISNULL((Select ISNULL(SUM(APD.PPN),0) From AP_Master APM, AP_Detail APD " & _
                                            "Where APM.Supplier_Code = IVM.Supplier_Code And APM.AP_No = APD.AP_No And APD.Invoice_No = IVM.Invoice_No And APM.AP_Date < '" & Format(dt, "yyyy-MM-01") & "' " & "),0) " & _
                        "From InvoiceSupplier_Master IVM, Trade_Master T, Bank_Master BM " & _
                        "Where IVM.Supplier_Code = T.Trade_Code " & _
                            "And IVM.Bank_Code = BM.Bank_Code " & _
                    ") dt, " & sqlMaster(1) & _
        "Where dt.Invoice_No = Master.InvNo And dt.CurrCD <> '' " & _
        "Order By Supplier_Code, CurrCD, Invoice_No, Invoice_Date"
    End If
    sqlLedger = sql
End Function

Sub ExcelLedger(stGeneral As Byte)
    
    Set rsLedger = Db.Execute(sqlLedger(stGeneral))
    
    If Not rsLedger.EOF Then
        With excelapp
            Call resetArr(1)
            Call headerExcel
                                    
            brs = 8: tampungBrs = 1: custCD = ""
            Do While Not rsLedger.EOF
                arrCol = CInt(rsLedger!currCD)
                If custCD <> Trim(rsLedger!Supplier_Code) Then
                    TotOpening = 0: TotEnding = 0
                    TotCollection = 0: TotFP = 0
                    If custCD <> "" Then
                        If stGeneral = 0 Then Call isiSubTotal
                        brs = brs + 1
                    End If
                    tampungBrs = brs
                End If
                    
                custCD = Trim(rsLedger!Supplier_Code)
                If arrCol = 3 Then nilFormat = gs_formatAmount Else nilFormat = gs_formatAmount
                
                .Cells(brs, 1) = Trim(rsLedger!Supplier_Code)
                .Cells(brs, 2) = Trim(rsLedger!trade_name)
                If stGeneral = 0 Then
                    .Cells(brs, 1) = Trim(rsLedger!Supplier_Code)
                    .Cells(brs, 2) = Trim(rsLedger!trade_name)
                    .Cells(brs, 3) = Trim(rsLedger!Invoice_No)
                    .Cells(brs, 4) = Format(rsLedger!Invoice_Date, "dd MMM yyyy")
                    .Cells(brs, 5) = IIf(IsNull(rsLedger!due_date), "", Format(rsLedger!due_date, "dd MMM yyyy"))
                    .Cells(brs, 6) = Trim(rsLedger!bank_name)
                    
                    TotOpening = rsLedger!OpeningPPN
                    TotFP = rsLedger!PPNIDR
                    TotCollection = rsLedger!ARPPN
                    TotEnding = rsLedger!RemainingPPN
                Else
                    TotOpening = TotOpening + rsLedger!OpeningPPN
                    TotFP = TotFP + rsLedger!PPNIDR
                    TotCollection = TotCollection + rsLedger!ARPPN
                    TotEnding = TotEnding + rsLedger!RemainingPPN
                End If
                
                .Cells(brs, Opening(arrCol)) = IIf(rsLedger!OpeningAmount = 0, "", Format(rsLedger!OpeningAmount, nilFormat))
                .Cells(brs, Opening(3)) = IIf(TotOpening = 0, "", Format(TotOpening, nilFormat))
                .Cells(brs, sales(arrCol)) = IIf(rsLedger!InvAmount = 0, "", Format(rsLedger!InvAmount, nilFormat))
                .Cells(brs, colFP) = IIf(TotFP = 0, "", Format(TotFP, nilFormat))
                .Cells(brs, collection(arrCol)) = IIf(rsLedger!ARAmount = 0, "", Format(rsLedger!ARAmount, nilFormat))
                .Cells(brs, collection(3)) = IIf(TotCollection = 0, "", Format(TotCollection, nilFormat))
                .Cells(brs, Ending(arrCol)) = IIf(rsLedger!RemainingAmount = 0, "", Format(rsLedger!RemainingAmount, nilFormat))
                .Cells(brs, Ending(3)) = IIf(TotEnding = 0, "", Format(TotEnding, nilFormat))
                
                '***************** Grand Total ********************
                GrandTotOpening(arrCol) = GrandTotOpening(arrCol) + rsLedger!OpeningAmount
                GrandTotsales(arrCol) = GrandTotsales(arrCol) + rsLedger!InvAmount
                GrandTotcollection(arrCol) = GrandTotcollection(arrCol) + rsLedger!ARAmount
                GrandTotEnding(arrCol) = GrandTotEnding(arrCol) + rsLedger!RemainingAmount
            
                GrandTotOpening(3) = GrandTotOpening(3) + rsLedger!OpeningPPN
                GrandtotFP = GrandtotFP + rsLedger!PPNIDR
                GrandTotcollection(3) = GrandTotcollection(3) + rsLedger!ARPPN
                GrandTotEnding(3) = GrandTotEnding(3) + rsLedger!RemainingPPN
                '***************************************************
                If stGeneral = 0 Then brs = brs + 1
                rsLedger.MoveNext
            Loop
            
            If stGeneral = 0 Then Call isiSubTotal
            brs = brs + 1
            Call isiGrandTotal
            Call bottom
        End With
    Else
        LblErrMsg = DisplayMsg(4006)
    End If
    Set rsLedger = Nothing
End Sub

Private Sub cmdReport_Click()
Dim rsLedger As New ADODB.Recordset
Dim SupplierCD As String
    
    If hakPrice(Me.Name) = 0 Then LblErrMsg = DisplayMsg("0006"): Exit Sub
    
    Me.MousePointer = vbHourglass
    LblErrMsg = ""
    If Trim(Cbo(0)) = "" Then
        LblErrMsg = DisplayMsg(1064) 'Please Input Supp
        Cbo(0).SetFocus
    ElseIf Cbo(0).MatchFound = False Then
        LblErrMsg = DisplayMsg(4050) 'Supp Not Found
        Cbo(0).SetFocus
    Else
        
        tglAkhir = Format(DateAdd("d", -1, Format(DateAdd("m", 1, dt), "yyyy-MM-01")), "yyyy-MM-dd")
        
        If Cbo(0).ListIndex <> 0 Then
            sqlcust = " And INV.Supplier_Code = '" & Cbo(0) & "' "
        Else
            sqlcust = ""
        End If
        
        If Option1(0).Value = True Then Call ExcelLedger(1) Else ExcelLedger (0)
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub cbo_Change(Index As Integer)
    lblNm(0) = ""
    LblErrMsg = ""
End Sub

Private Sub cbo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbo_Click(Index)
End Sub

Public Sub cbo_Click(Index As Integer)
    Cbo(0) = Cbo(0)
    If Cbo(0).MatchFound Then
        lblNm(0) = Cbo(0).Column(1)
        LblErrMsg = ""
    Else
        lblNm(0) = ""
        LblErrMsg = DisplayMsg(4011)
    End If
End Sub

Private Sub dt_change()
    Call dt_Click
    TampungDt = dt.Month
End Sub

Private Sub dt_Click()
    LblErrMsg = ""
    If dt.Month = 1 And Val(TampungDt) = 12 Then dt.Year = dt.Year + 1
    If dt.Month = 12 And Val(TampungDt) = 1 Then dt.Year = dt.Year - 1
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


Private Sub Option1_Click(Index As Integer)
    LblErrMsg = ""
End Sub
