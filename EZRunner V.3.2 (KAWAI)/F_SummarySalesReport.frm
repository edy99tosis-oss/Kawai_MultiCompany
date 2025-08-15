VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form F_SummarySalesReport 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Invoice Summary List"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   Icon            =   "F_SummarySalesReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   390
      TabIndex        =   13
      Top             =   2235
      Width           =   6870
      Begin VB.Label LblPesan 
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
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Width           =   6525
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xcel"
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
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2895
      Visible         =   0   'False
      Width           =   1245
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
      Height          =   285
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1395
      Width           =   4065
   End
   Begin VB.CommandButton Cmd_SubMenu 
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
      Left            =   405
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2895
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
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
      Left            =   6030
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2895
      Width           =   1245
   End
   Begin MSComCtl2.DTPicker MDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd MMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   345
      Left            =   1695
      TabIndex        =   1
      Top             =   1800
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
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
      CurrentDate     =   37798
   End
   Begin MSComCtl2.DTPicker MDate1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd MMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   345
      Left            =   3885
      TabIndex        =   2
      Top             =   1800
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
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
      CurrentDate     =   37798
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   5400
      TabIndex        =   6
      Top             =   240
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Line Line1 
      X1              =   3195
      X2              =   7215
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label LblCust 
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
      Height          =   255
      Left            =   3165
      TabIndex        =   11
      Top             =   1395
      Visible         =   0   'False
      Width           =   4050
   End
   Begin MSForms.ComboBox CboCust 
      Height          =   315
      Left            =   1695
      TabIndex        =   0
      Top             =   1365
      Width           =   1425
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "2514;556"
      ListRows        =   15
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
      Caption         =   "Customer"
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
      Left            =   450
      TabIndex        =   10
      Top             =   1425
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   3495
      TabIndex        =   9
      Top             =   1875
      Width           =   210
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
      Index           =   1
      Left            =   450
      TabIndex        =   8
      Top             =   1875
      Width           =   540
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Summary List"
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
      TabIndex        =   7
      Top             =   630
      Width           =   6870
   End
End
Attribute VB_Name = "F_SummarySalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tek As Integer, custCD As String
Dim Amounti As Double, PPni As Double, grandI As Double
Dim Curr As String * 2, BolCurr As Boolean, posawal As Double
Dim CustTotal(100000) As String, CurrTotal(10000) As String
Dim idx_Cust As Double, idx_curr As Double

Dim bteHakPrice As Byte

Private Sub CboCust_Change()
    If cboCust.MatchFound Then
        lblcust = cboCust.List(cboCust.ListIndex, 1)
        Text1 = cboCust.List(cboCust.ListIndex, 1)
        LblPesan = ""
    Else
        lblcust = ""
        Text1 = ""
        LblPesan = DisplayMsg(4072)
    End If
    If Trim(cboCust) = "" Then LblPesan = ""
End Sub

Private Sub Cmd_SubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim xlapp As New Excel.application
    Dim rsCek As New Recordset, Idx As Long, tempi As String, tempcust As String
    Dim bolcust As Boolean, bolinv As Boolean
    Dim rsCompany As New Recordset, rate As Double
    
    If Trim(cboCust) = "" Then LblPesan = DisplayMsg(1045): Exit Sub
    cboCust = Trim(cboCust)
    If Not cboCust.MatchFound Then LblPesan = DisplayMsg(4072): Exit Sub
    
    rate = tax("PPn")
    
    If lblcust = strAll Then
        sql = " select (invoice_master.cust_code)as sebango,trade_name,invoice_master.invoice_no,invoice_master.invoice_date, " & _
            " invoice_detail.po_no , invoice_detail.item_code , item_name, qty, price, invoice_detail.amount,  " & _
            " trade_master.country_cls,case trade_master.country_cls when '1' then  0 else  ppn end ppn, " & _
            " (select description from curr_cls where curr_cls= invoice_detail.currency_code) currency,  " & _
            " invoice_detail.currency_code, '10' PpnRate, invoice_master.Amount AmountTotal  "
            
            sql = sql + " From invoice_master, trade_master, invoice_detail, item_master  " & _
            " where invoice_master.cust_code=trade_master.trade_code  " & _
            " and invoice_master.invoice_no=invoice_detail.invoice_no  " & _
            " and invoice_detail.item_code=item_master.item_code  " & _
            " and fix_cls=1 " & _
            " and case trade_master.country_cls when '1' then (select ETD from packing_master where packing_no = invoice_master.list_do) else invoice_master.invoice_date end >= '" & Format(MDate, "YYYY-MM-DD") & "' " & _
            " and case trade_master.country_cls when '1' then (select ETD from packing_master where packing_no = invoice_master.list_do) else invoice_master.invoice_date end <= '" & Format(MDate1, "YYYY-MM-DD") & "' " & _
            " order by currency_code, invoice_master.cust_Code, invoice_master.invoice_no " & _
            "  "
    Else
        sql = " select (invoice_master.cust_code)as sebango,trade_name,invoice_master.invoice_no,invoice_master.invoice_date, " & _
            " invoice_detail.po_no , invoice_detail.item_code , item_name, qty, price, invoice_detail.amount,  " & _
            " trade_master.country_cls,case trade_master.country_cls when '1' then  0 else  ppn end ppn, " & _
            " (select description from curr_cls where curr_cls= invoice_detail.currency_code) currency,  " & _
            " invoice_detail.currency_code, '10' PpnRate, invoice_master.Amount AmountTotal  "
            
            sql = sql + " From invoice_master, trade_master, invoice_detail, item_master  " & _
            " where invoice_master.cust_code=trade_master.trade_code  " & _
            " and invoice_master.invoice_no=invoice_detail.invoice_no  " & _
            " and invoice_detail.item_code=item_master.item_code  " & _
            " and fix_cls=1 and invoice_master.cust_Code = '" & cboCust.Text & "' " & _
            " and case trade_master.country_cls when '1' then (select ETD from packing_master where packing_no = invoice_master.list_do) else invoice_master.invoice_date end >= '" & Format(MDate, "YYYY-MM-DD") & "' " & _
            " and case trade_master.country_cls when '1' then (select ETD from packing_master where packing_no = invoice_master.list_do) else invoice_master.invoice_date end <= '" & Format(MDate1, "YYYY-MM-DD") & "' " & _
            " order by currency_code, invoice_master.cust_Code, invoice_master.invoice_no " & _
            "  "
    End If
    
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
            
            .Range("a2", "f2").Merge
            .Range("a2") = rsCompany!company_name
            .Range("a3", "f3").Merge
            .Range("a3") = rsCompany!address1 & " " & rsCompany!address2 & " " & rsCompany!City & " " & rsCompany!Province & " " & rsCompany!postal_code
            .Range("a4", "f4").Merge
            .Range("a4") = "Phone: " & rsCompany!phone1 & " " & rsCompany!phone2 & " Fax: " & rsCompany!fax
            
            .Range("a6") = "Invoice Detail List"
            .Range("b6") = ""
            .Range("a6", "b6").Merge
            .Range("a6").horizontalAlignment = xlLeft
            .Range("a7") = "Date"
            .Range("b7") = ": " & Format(Now, "dd MMMM YYYY")
            .Range("a8") = "Period"
            .Range("b8") = ": " & Format(MDate, "dd MMMM YYYY") & " to " & Format(MDate1, "dd MMMM YYYY")
            
            Idx = 10
            tempcust = ""
            tempi = ""
            
            Amounti = 0
            PPni = 0
            grandI = 0
            
            posawal = 14
            idx_Cust = 0
            idx_curr = 0
            
            Do While Not rsCek.EOF
            
                bolcust = False
                bolinv = False
                BolCurr = False
                
                If Idx <> 10 And tempi <> Trim(rsCek!Invoice_No) Then
                    Call TotalInvoice(xlapp, posawal, Idx, "f", "e", rsCek!country_cls, rsCek!ppnrate)
                    bolinv = True
                    If lblcust = strAll Then
                        CustTotal(idx_Cust) = "f" & Idx
                        idx_Cust = idx_Cust + 1
                    Else
                        CurrTotal(idx_curr) = "f" & Idx
                        idx_curr = idx_curr + 1
                    End If
                End If
                
                If lblcust = strAll Then
                    If Idx <> 10 And tempcust <> Trim(rsCek!sebango) Then
                        Idx = Idx + 3
                        Call TotalCust(xlapp, Idx, "f", "e", "d", rsCek!country_cls, rsCek!ppnrate)
                        CurrTotal(idx_curr) = "f" & Idx
                        idx_Cust = 0
                        idx_curr = idx_curr + 1
                        bolcust = True
                    End If
                End If
                
                If Idx <> 10 And Curr <> Trim(rsCek!currency_code) Then
                    Idx = Idx + 3
                    Call TotalCurr(xlapp, Idx, "f", "e", "d", rsCek!country_cls, rsCek!ppnrate)
                    idx_curr = 0
                    BolCurr = True
                End If
                
                If Idx = 10 Or BolCurr Then
                    If Idx <> 10 Then
                        Idx = Idx + 4
                    End If
                    .Range("a" & Idx) = "Currency "
                    .Range("b" & Idx) = ": " & Trim(rsCek!Currency)
                    Idx = Idx + 1
                    .Range("a" & Idx) = "Customer "
                    .Range("b" & Idx) = ": " & Trim(rsCek!sebango) & " / " & Trim(rsCek!trade_name)
                    bolcust = False
                End If
                                
                If Idx = 11 Or bolcust Then
                    If Idx <> 11 Then
                        Idx = Idx + 4
                    End If
                    .Range("a" & Idx) = "Customer "
                    .Range("b" & Idx) = ": " & Trim(rsCek!sebango) & " / " & Trim(rsCek!trade_name)
                End If
                
                If Idx = 11 Or bolinv Then
                    If Idx = 11 Or bolcust Or BolCurr Then
                        Idx = Idx + 1
                    Else
                        Idx = Idx + 4
                    End If
                    .Range("a" & Idx) = "Invoice No."
                    .Range("b" & Idx) = ": " & Trim(rsCek!Invoice_No)
                    .Range("c" & Idx) = "Invoice Date "
                    .Range("d" & Idx) = ": " & Trim(Format(rsCek!Invoice_Date, "DD MMM YYYY"))
                    Idx = Idx + 1
                    .Range("a" & Idx) = "Product Code"
                    .Range("b" & Idx) = "Description"
                    .Range("c" & Idx) = "PO No"
                    .Range("d" & Idx) = "Qty"
                    .Range("e" & Idx) = "Price"
                    .Range("f" & Idx) = "Total"
                    .Range("a" & Idx, "f" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Range("a" & Idx, "f" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Idx = Idx + 1
                    posawal = Idx
                End If
                
                Idx = Idx
                'Content
                .Range("a" & Idx) = Trim(rsCek!Item_Code)
                .Range("b" & Idx) = Trim(rsCek!item_name)
                .Range("c" & Idx) = "'" & Trim(rsCek!po_no)
                .Range("d" & Idx) = Format(rsCek!Qty, gs_formatQty)
                If bteHakPrice = 1 Then
                    .Range("e" & Idx) = Format(rsCek!Price, gs_formatPrice)
                    .Range("f" & Idx) = Format(rsCek!Amount, gs_formatAmount)
                End If
                
                Idx = Idx + 1
                tempcust = Trim(rsCek!sebango)
                Curr = Trim(rsCek!currency_code)
                tempi = Trim(rsCek!Invoice_No)
                rsCek.MoveNext
            
            Loop
            rsCek.MoveLast
            
            If bteHakPrice = 1 Then
                Call TotalInvoice(xlapp, posawal, Idx, "f", "e", rsCek!country_cls, rsCek!ppnrate)
                If lblcust = strAll Then
                    CustTotal(idx_Cust) = "f" & Idx
                    idx_Cust = idx_Cust + 1
                Else
                    CurrTotal(idx_curr) = "f" & Idx
                    idx_curr = idx_curr + 1
                End If
                Idx = Idx + 3
                If lblcust = strAll Then
                    Call TotalCust(xlapp, Idx, "f", "e", "d", rsCek!country_cls, rsCek!ppnrate)
                    CurrTotal(idx_curr) = "f" & Idx
                    idx_curr = idx_curr + 1
                    Idx = Idx + 3
                End If
                Call TotalCurr(xlapp, Idx, "f", "e", "d", rsCek!country_cls, rsCek!ppnrate)
            Else
                .Range("a" & Idx, "f" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
            
            .Range("a1", "f" & Idx + 3).Columns.Font.Name = "Arial"
            .Range("a1", "f" & Idx + 3).Columns.Font.Size = 8
            .Range("a2", "f2").Columns.Font.Name = "Arial"
            .Range("a2", "f2").Columns.Font.Size = "10"
            .Range("a2", "f2").Columns.Font.Bold = True
            .Range("a2", "f4").horizontalAlignment = xlCenter
            .Range("a6", "f6").Columns.Font.Bold = True
            
            .Visible = True
            .Columns("D:D").Select
            .Selection.NumberFormat = gs_formatQty
            .Columns("E:E").Select
            .Selection.NumberFormat = gs_formatPrice
            .Columns("F:F").Select
            .Selection.NumberFormat = gs_formatAmount
            .ActiveSheet.PageSetup.PaperSize = xlPaperA4
            .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
            .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.4)
            .Range("A:F").Columns.AutoFit
            .WindowState = xlMaximized
        
        End With
    Else
        LblPesan = DisplayMsg(4006)
    End If
    Screen.MousePointer = vbDefault
End Sub

Sub TotalInvoice(xl As Excel.application, posawal As Double, Row As Long, Col As String, coltitle As String, Optional countrycls As String, Optional ppn As Double)
    With xl
        .Range("a" & Row, "f" & Row).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(coltitle & Row) = "Total"
        .Range(coltitle & Row + 1) = "PPN"
        .Range(coltitle & Row + 2) = "Grand Total"
        .Range(Col & Row).Formula = "=sum(" & Col & posawal & ":" & Col & Row - 1 & ")"  'Format(Amounti, "#,##0.#0")
        If countrycls = "1" Then
            .Range(Col & Row + 1) = 0
        Else
            .Range(Col & Row + 1).Formula = "=" & "(((" & Col & Row & ") * " & ppn & ") / 100)"
        End If
        .Range(Col & Row + 2).Formula = "=" & Col & Row & "+" & Col & Row + 1
    End With
End Sub

Sub TotalCust(xl As Excel.application, Row As Long, Col As String, coltitle As String, coltitle2 As String, countrycls As String, ppn As Double)
    Dim Y As Double, Formula As String
    With xl
        .Range("a" & Row, "f" & Row).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(coltitle2 & Row) = "Total Per Customer"
        .Range(coltitle & Row) = "Total"
        .Range(coltitle & Row + 1) = "PPN"
        .Range(coltitle & Row + 2) = "Grand Total"
        
        For Y = 0 To idx_Cust - 1
            If Y = 0 Then
                Formula = "= " & CustTotal(Y)
            Else
                Formula = Formula & "+" & CustTotal(Y)
            End If
        Next
        .Range(Col & Row).Formula = Formula
        If countrycls = "1" Then
            .Range(Col & Row + 1) = 0
        Else
            .Range(Col & Row + 1).Formula = "=" & "(((" & Col & Row & ") * " & ppn & ") / 100)"
        End If
        .Range(Col & Row + 2).Formula = "=" & Col & Row & "+" & Col & Row + 1
    End With
End Sub

Sub TotalCurr(xl As Excel.application, Row As Long, Col As String, coltitle As String, coltitle2 As String, countrycls As String, ppn As Double)
    Dim Y As Double, Formula As String
    With xl
        .Range("a" & Row, "f" & Row).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(coltitle2 & Row) = "Total Per Currency"
        .Range(coltitle & Row) = "Total"
        .Range(coltitle & Row + 1) = "PPN"
        .Range(coltitle & Row + 2) = "Grand Total"
        For Y = 0 To idx_curr - 1
            If Y = 0 Then
                Formula = "= " & CurrTotal(Y)
            Else
                Formula = Formula & "+" & CurrTotal(Y)
            End If
        Next
        .Range(Col & Row).Formula = Formula
        If countrycls = "1" Then
            .Range(Col & Row + 1) = 0
        Else
            .Range(Col & Row + 1).Formula = "=" & "(((" & Col & Row & ") * " & ppn & ") / 100)"
        End If
        .Range(Col & Row + 2).Formula = "=" & Col & Row & "+" & Col & Row + 1
    End With
End Sub

Sub Total(xl As Excel.application, Row As Long, Col As String, coltitle As String, coltitle2 As String)
    With xl
        .Range("a" & Row, "e" & Row).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(coltitle2 & Row) = "Grand Total"
        .Range(coltitle & Row) = "Total"
        .Range(coltitle & Row + 1) = "PPN"
        .Range(coltitle & Row + 2) = "Grand Total"
    End With
End Sub

Private Sub command2_Click()

    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsRpt As New ADODB.Recordset
    Dim rsRpt1 As New ADODB.Recordset
    
    LblPesan = ""
    If Trim(cboCust) = "" Then LblPesan = DisplayMsg(1045): Exit Sub
    cboCust = Trim(cboCust)
    If Not cboCust.MatchFound Then LblPesan = DisplayMsg(4072): Exit Sub
    
    Me.MousePointer = vbHourglass

    sql = "select a.country_cls, a.country_desc, a.cust_code, a.cust_name, a.invoice_no, a.invoice_date, " & _
          vbLf & "a.delivery_no, a.delivery_date, a.currency_code, a.curr_desc, sum(a.amount) as amount, " & _
          vbLf & "sum(a.nocommercial) as nocommercial, " & _
          vbLf & "sum(a.amount) - sum(a.nocommercial) as total, " & _
          vbLf & "(sum(a.amount) - sum(a.nocommercial)) * ppn_rate as total_rate, " & _
          vbLf & "(case when a.country_cls = 1 then 0 else " & _
          vbLf & "  case when a.ppn_value = 0 then 0 else " & _
          vbLf & "   case when currency_code = '03' then (sum(a.amount) - sum(a.nocommercial)) / a.ppn_value else " & _
          vbLf & "  ((sum(a.amount) - sum(a.nocommercial)) * ppn_rate) / a.ppn_value " & _
          vbLf & "  End " & _
          vbLf & " End " & _
          vbLf & "end) as ppn, " & _
          vbLf & "a.fakturpajak_no, a.company_name, a.address1, a.address2, a.province, a.city, a.postal_code, " & _
          vbLf & "a.phone1 , a.phone2, a.fax, a.pebno, a.pebdate " & _
          vbLf & "from ( "
    sql = sql & _
          vbLf & "select tm.country_cls, (case tm.country_cls when '0' then 'Local' else 'Export' end) as country_desc, " & _
          vbLf & "rtrim(invm.cust_code) cust_code, rtrim(tm.trade_name) cust_name, rtrim(invm.invoice_no) invoice_no, " & _
          vbLf & "invm.invoice_date, rtrim(invd.currency_code) currency_code, rtrim(cc.description) curr_desc, " & _
          vbLf & "invd.amount, " & _
          vbLf & "case when om.nocommercial_cls = 0 then 0 else invd.amount end as nocommercial, " & _
          vbLf & "ppn_value = isnull((select rate from tax_cls where tax_code = 'ppn' and start_date <= convert(varchar, invm.invoice_date, 112) and end_date >= convert(varchar, invm.invoice_date, 112)), 0), " & _
          vbLf & "ppn_rate = isnull((select tax_exchangerate from tax_exchangerate where currency_code = invd.currency_code and start_date <= convert(varchar, invm.invoice_date, 112) and end_date >= convert(varchar, invm.invoice_date, 112)), 0), " & _
          vbLf & "(case when tm.country_cls = 0 then invd.do_no else Invd.packing_no end) as Delivery_No, --Tambahan, peke pajak atau invoicedetail " & _
          vbLf & "(case when tm.country_cls = 0 then " & _
          vbLf & "(select distinct DO_Date from DO_Master where do_no = invd.do_no) " & _
          vbLf & "Else " & _
          vbLf & "(select distinct packing_date from packing_master where packing_no = invm.list_do) " & _
          vbLf & "end) as delivery_date, " & _
          vbLf & "rtrim(cp.company_name) company_name, rtrim(cp.address1) address1, rtrim(cp.address2) address2, " & _
          vbLf & "rtrim(cp.province) province, rtrim(cp.city) city, rtrim(cp.postal_code) postal_code, " & _
          vbLf & "rtrim(cp.phone1) phone1, rtrim(cp.phone2) phone2, rtrim(cp.fax) fax, " & _
          vbLf & "rtrim(invm.pebno) pebno, invm.pebdate, fpd.fakturpajak_no " & _
          vbLf & "from invoice_master invm " & _
          vbLf & "inner join invoice_detail invd on invm.invoice_no = invd.invoice_no " & _
          vbLf & "inner join trade_master tm on invm.cust_code = tm.trade_code " & _
          vbLf & "inner join orderentry_master om on invd.po_no = om.po_no " & _
          vbLf & "left outer join curr_cls cc on invd.currency_code = cc.curr_cls " & _
          vbLf & "left join FakturPajak_Detail fpd on invd.Invoice_No = fpd.Invoice_No and invd.DO_No = fpd.DO_No, " & _
          vbLf & "company_profile cp "
    sql = sql & _
          vbLf & "Where invm.fix_cls = 1 " & _
          vbLf & "and invm.invoice_date >= '" & Format(MDate, "yyyy-MM-dd") & "' " & _
          vbLf & "and invm.invoice_date <= '" & Format(MDate1, "yyyy-MM-dd") & "' " & _
          vbLf & ") a " & _
          vbLf & "group by a.country_cls, a.country_desc, a.cust_code, a.cust_name, " & _
          vbLf & "a.invoice_no, a.invoice_date,  a.delivery_no, a.delivery_date, a.currency_code, a.curr_desc, " & _
          vbLf & "a.ppn_value, a.ppn_rate, a.fakturpajak_no, " & _
          vbLf & "a.company_name , a.address1, a.address2, a.province, a.city, a.postal_code, a.phone1, " & _
          vbLf & "a.phone2 , a.fax, a.pebno, a.pebdate "

    If lblcust <> strAll Then
      sql = sql & _
            vbLf & "having a.cust_code = '" & cboCust & "' "
    End If
    
    sql = sql & _
          vbLf & "order by a.invoice_no "
    
    If rsRpt.State <> adStateClosed Then rsRpt.Close
    rsRpt.Open sql, Db, adOpenDynamic
    
    If Not rsRpt.EOF Then
    
        sqlprint = sql
        Fbulan = Format(MDate, "dd MMMM  yyyy")
        Ftahun = Format(MDate1, "dd MMMM  yyyy")
        reportcode = "InvoiceSummaryList"
        printorient = 1
        
        Set report = application.OpenReport(App.path & "\reports\SummaryOfSalesReport.rpt")
        report.Database.Tables(1).SetDataSource rsRpt, 3
        report.FormulaFields(1).Text = "'" & Format(MDate, "dd MMMM  yyyy") & " to " & Format(MDate1, "dd MMMM  yyyy") & "'"
                
         sqlprint2 = "select b.country_cls, b.currency_code, b.curr_desc, sum(b.amount) amount, sum(b.nocommercial) nocommercial, sum(total) total, sum(total_rate) total_rate, sum(ppn) ppn " & _
                     vbLf & "from ( " & _
                     vbLf & "select a.country_cls, a.cust_code, a.invoice_no, a.currency_code, a.curr_desc, sum(a.amount) amount, sum(a.nocommercial) nocommercial, " & _
                     vbLf & "sum(a.amount) - sum(a.nocommercial) total, (sum(a.amount) - sum(a.nocommercial)) * ppn_rate total_rate, " & _
                     vbLf & "case when a.country_cls = 1 then 0 else case when a.ppn_value = 0 then 0 else " & _
                     vbLf & "case when currency_code = '03' then (sum(a.amount) - sum(a.nocommercial)) / a.ppn_value else ((sum(a.amount) - sum(a.nocommercial)) * ppn_rate) / a.ppn_value end " & _
                     vbLf & "end end ppn " & _
                     vbLf & "from ( " & _
                     vbLf & "select tm.country_cls, rtrim(invm.cust_code) cust_code, rtrim(invm.invoice_no) invoice_no, rtrim(invd.currency_code) currency_code, rtrim(cc.description) curr_desc, " & _
                     vbLf & "invd.amount, case when om.nocommercial_cls = 0 then 0 else invd.amount end nocommercial, " & _
                     vbLf & "ppn_value = isnull((select rate from tax_cls where tax_code = 'ppn' and start_date <= convert(varchar, invm.invoice_date, 112) and end_date >= convert(varchar, invm.invoice_date, 112)), 0), " & _
                     vbLf & "ppn_rate = isnull((select tax_exchangerate from tax_exchangerate where currency_code = invd.currency_code and start_date <= convert(varchar, invm.invoice_date, 112) and end_date >= convert(varchar, invm.invoice_date, 112)), 0) " & _
                     vbLf & "from invoice_master invm " & _
                     vbLf & "inner join invoice_detail invd on invm.invoice_no = invd.invoice_no " & _
                     vbLf & "inner join trade_master tm on invm.cust_code = tm.trade_code " & _
                     vbLf & "inner join orderentry_master om on invd.po_no = om.po_no " & _
                     vbLf & "left outer join curr_cls cc on invd.currency_code = cc.curr_cls, company_profile cp " & _
                     vbLf & "Where invm.fix_cls = 1 " & _
                     vbLf & "and invm.invoice_date >= '" & Format(MDate, "yyyy-MM-dd") & "' " & _
                     vbLf & "and invm.invoice_date <= '" & Format(MDate1, "yyyy-MM-dd") & "' " & _
                     vbLf & ") a group by a.country_cls, a.cust_code, a.invoice_no, a.currency_code, a.curr_desc, a.ppn_value, a.ppn_rate " & _
                     vbLf & ") b group by b.country_cls, b.currency_code, b.curr_desc "
        
        If rsRpt1.State <> adStateClosed Then rsRpt1.Close
        rsRpt1.Open sqlprint2, Db, adOpenDynamic
        report.OpenSubreport("summary_country").Database.Tables(1).SetDataSource rsRpt1, 3
        
        Dim Rpt As New FrmRpt3
        With Rpt.CRViewer1
            .ReportSource = report
            .ViewReport
            .Zoom 1
        End With
        With Rpt
            .WindowState = 2
            .Show 1
        End With
    
    Else
    
        LblPesan.Caption = DisplayMsg("4006")
    
    End If
    
    Me.MousePointer = vbDefault

End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblPesan.Caption = ErrMsg
    End If
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(F_DetailSalesReport)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    bteHakPrice = hakPrice(Me.Name)
    Call Customer
    MDate = Now
    MDate1 = Now
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = True
End Sub

Private Sub Customer()
    Dim sql As String, RsCust As New ADODB.Recordset
    Dim i As Long
    
    If RsCust.State <> adStateClosed Then RsCust.Close
    RsCust.Open "trade_master order by trade_code asc", Db, adOpenDynamic, adLockOptimistic, adCmdTable
    
    cboCust.columnCount = 2
    cboCust.TextColumn = 1
    
    i = 0
        Do While Not RsCust.EOF
        cboCust.AddItem ""
        If i = 0 Then
            cboCust.List(i, 0) = strAll
            cboCust.List(i, 1) = strAll
            i = i + 1
            cboCust.AddItem ""
            cboCust.List(i, 0) = Trim(RsCust!Trade_Code)
            cboCust.List(i, 1) = Trim(RsCust!trade_name)
        Else
            cboCust.List(i, 0) = Trim(RsCust!Trade_Code)
            cboCust.List(i, 1) = Trim(RsCust!trade_name)
        End If
        i = i + 1
        RsCust.MoveNext
    Loop
    
    cboCust.ColumnWidths = "50 pt; 300 pt"
    cboCust.ListWidth = 350
    cboCust.ListRows = 15
    cboCust.ListIndex = 0
End Sub

Function tax(kode$) As String
    Dim rtax As Recordset, tempdate1 As String, tempdate2 As String
    
    tempdate1 = Format(MDate.Value, "yyyymmdd")
    tempdate2 = Format(MDate1.Value, "yyyymmdd")
    
    sql = "SELECT Tax_Code, Rate, start_Date, End_Date " & _
        "FROM tax_cls " & _
        "" & _
        "WHERE ((Start_date <= '" & tempdate1 & "' And End_Date >= '" & tempdate1 & "') " & _
        "Or (Start_date <= '" & tempdate2 & "' And End_Date >= '" & tempdate2 & "')) " & _
        "and tax_code= '" & kode & "'"
    
    Set rtax = New Recordset
    rtax.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not rtax.EOF Then
        tax = rtax!rate
    Else
        tax = 0
    End If
End Function





