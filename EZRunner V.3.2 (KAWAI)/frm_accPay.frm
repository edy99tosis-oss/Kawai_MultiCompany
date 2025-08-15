VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_accPay 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Part/Material Receipt List"
   ClientHeight    =   4455
   ClientLeft      =   1335
   ClientTop       =   2775
   ClientWidth     =   8220
   Icon            =   "frm_accPay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lblWHName 
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
      Left            =   3792
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "LblLocationName"
      Top             =   1845
      Width           =   4050
   End
   Begin VB.TextBox LblLocationName 
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
      Left            =   3792
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "LblLocationName"
      Top             =   1425
      Width           =   4050
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
      Left            =   5697
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3675
      Visible         =   0   'False
      Width           =   1035
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   5989
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   285
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Sub &Menu"
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
      Index           =   8
      Left            =   379
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3675
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   379
      TabIndex        =   8
      Top             =   2880
      Width           =   7470
      Begin VB.Label LblErrMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LblErrMsg"
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
         Left            =   105
         TabIndex        =   9
         Top             =   195
         Width           =   7260
      End
   End
   Begin VB.CommandButton Cmd_Save 
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
      Index           =   0
      Left            =   6814
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3675
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Left            =   1995
      TabIndex        =   2
      Top             =   2220
      Width           =   1500
      _ExtentX        =   2646
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
      Format          =   145227779
      CurrentDate     =   37798
   End
   Begin MSComCtl2.DTPicker Dmonth2 
      Height          =   315
      Left            =   4065
      TabIndex        =   3
      Top             =   2220
      Width           =   1500
      _ExtentX        =   2646
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
      Format          =   145227779
      CurrentDate     =   37798
   End
   Begin VB.Line Line2 
      X1              =   3795
      X2              =   7830
      Y1              =   2085
      Y2              =   2085
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warehouse Code"
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
      Left            =   390
      TabIndex        =   15
      Top             =   1845
      Width           =   1470
   End
   Begin MSForms.ComboBox CboWHCode 
      Height          =   315
      Left            =   1995
      TabIndex        =   1
      Top             =   1785
      Width           =   1725
      VariousPropertyBits=   612386843
      MaxLength       =   10
      DisplayStyle    =   3
      Size            =   "3043;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      X1              =   3795
      X2              =   7875
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date (Month)"
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
      Left            =   390
      TabIndex        =   14
      Top             =   2280
      Width           =   1125
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
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
      Left            =   390
      TabIndex        =   13
      Top             =   1425
      Width           =   705
   End
   Begin MSForms.ComboBox CboLocationCD 
      Height          =   315
      Left            =   1995
      TabIndex        =   0
      Top             =   1365
      Width           =   1725
      VariousPropertyBits=   612386843
      MaxLength       =   10
      DisplayStyle    =   3
      Size            =   "3043;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   3795
      TabIndex        =   12
      Top             =   2280
      Width           =   165
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Part/Material Receipt List"
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
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   7470
   End
End
Attribute VB_Name = "frm_accPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Amounti As Double
Dim PPni As Double
Dim grandI As Double

Dim bteHakPrice As Byte

Private Sub CboLocationCD_Change()
If CboLocationCD.MatchFound Then
   LblLocationName = CboLocationCD.List(CboLocationCD.ListIndex, 1)
   LblErrMsg = ""
Else
   LblLocationName = ""
   LblErrMsg = DisplayMsg("0032")
End If
End Sub

Private Sub CboLocationCD_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim j As Integer
If KeyCode = 13 Then
  j = 0
For i = 0 To CboLocationCD.ListCount - 1
    If UCase(Trim(CboLocationCD)) = UCase(Trim(CboLocationCD.List(i, 0))) Then
        CboLocationCD = Trim(CboLocationCD.List(i, 0))
        LblLocationName = Trim(CboLocationCD.List(i, 1))
        j = 1: Exit For
    End If
Next

If j = 0 Then
    LblErrMsg = DisplayMsg("0032"): Exit Sub
Else
    LblErrMsg = ""
End If
End If
End Sub

Private Sub CboLocationCD_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub CboWHCode_Change()
If cboWhCode.MatchFound Then
    lblWHName = cboWhCode.Column(1)
Else
    lblWHName = ""
End If
End Sub

Private Sub Cmd_Save_Click(Index As Integer)
Dim j As Integer
Dim sqlcd As String

Select Case Index
       Case 8:
                DoEvents
                frmMainMenu.Show
                DoEvents
                Unload Me

        Case 0:
                
              Call CboLocationCD_Change
              
              Dim application As New CRAXDDRT.application
              Dim report As New CRAXDDRT.report
              Dim rsRpt As New ADODB.Recordset
              Dim Rpt As New FrmRpt3
                 
                 j = 0
                For i = 0 To CboLocationCD.ListCount - 1
                    If UCase(Trim(CboLocationCD)) = UCase(Trim(CboLocationCD.List(i, 0))) Then
                        CboLocationCD = Trim(CboLocationCD.List(i, 0))
                        LblLocationName = Trim(CboLocationCD.List(i, 1))
                        j = 1: Exit For
                    End If
                Next
                
                If j = 0 Then LblErrMsg = DisplayMsg("0032"): Exit Sub
             
              LblErrMsg = ""
              Me.MousePointer = vbHourglass
                                
              If CboLocationCD.Text = strAll Then
                    sqlcd = ""
              Else
                    sqlcd = "and pr.supplier_code = '" & Trim(CboLocationCD) & "' "
              End If
              
              
              sql = "select '" & bteHakPrice & "' HakPrice, rtrim(pr.supplier_code) supplier_code, tm.trade_cls, rtrim(pr.po_no) po_no, rtrim(pr.warehouse_code) warehouse_code, pr.receipt_cls, pr.receipt_date, rtrim(pr.item_code) item_code, " & _
                "pr.qty, pr.currency_code, pr.price, pr.amount, rtrim(pr.suratjalan_no) suratjalan_no, rtrim(pr.remarks) remarks,rtrim(pr.bc_type) bc_type, rtrim(pr.bc40_no) bc40_no, Format(pr.BC40_Date,'yyyy-MM-dd') bc40_date , package_qty, " & _
                "rtrim(im.makeritem_code) makeritem_code, rtrim(im.item_name) item_name, im.sheetcoil_cls, im.length, im.width, im.thickness, " & _
                "rtrim(tm.trade_name) trade_name, rtrim(wh.wh_name) wh_name, rtrim(sh.description) sheetcoil_desc, rtrim(uc.description) unit_desc, rtrim(cc.description) curr_desc, " & _
                "rtrim(cp.company_name) company_name " & _
                "from part_receipt pr " & _
                "inner join item_master im on pr.item_code = im.item_code " & _
                "inner join trade_master tm on pr.supplier_code = tm.trade_code " & _
                "left outer join warehouse_master wh on pr.warehouse_code = wh.wh_code " & _
                "left outer join sheetcoil_cls sh on im.sheetcoil_cls = sh.sheetcoil_cls " & _
                "left outer join unit_cls uc on pr.unit_cls = uc.unit_cls " & _
                "left outer join curr_cls cc on pr.currency_code = cc.curr_cls, company_profile cp " & _
                "where (receipt_cls = 'R' or receipt_cls = 'R1') " & sqlcd & "" & _
                "and receipt_date >= '" & Format(DMonth, "yyyy-MM-dd") & "' " & _
                "and receipt_date <= '" & Format(Dmonth2, "yyyy-MM-dd") & "' "
                
                If cboWhCode <> strAll Then sql = sql & "and pr.warehouse_code = '" & Trim(cboWhCode) & "' "
                sql = sql & "order by pr.receipt_date, pr.item_code"
                                  
              If rsRpt.State <> adStateClosed Then rsRpt.Close
              rsRpt.Open sql, Db, adOpenForwardOnly, adLockReadOnly
              
              If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
              sqlprint = sql
              reportcode = "AccPay"
              printorient = 1
            
              Set report = application.OpenReport(App.path & "\Reports\rpt_accPay.rpt")
              report.Database.Tables(1).SetDataSource rsRpt
              
              
            report.FormulaFields(1).Text = "'" & Format(DMonth, "dd-MMM-yyyy") & " to " & Format(Dmonth2, "dd-MMM-yyyy") & "'"
            report.FormulaFields(6).Text = "" & gi_decimalDigitQty & ""
            report.FormulaFields(7).Text = "" & gi_decimalDigitPrice & ""
            report.FormulaFields(8).Text = "" & gi_decimalDigitPriceIDR & ""
            report.FormulaFields(9).Text = "" & gi_decimalDigitAmount & ""
            report.FormulaFields(10).Text = "" & gi_decimalDigitAmountIDR & ""
            
            Dim dates As String, dates2 As String, nPPn As String
            
            Rpt.CRViewer1.ReportSource = report
            Rpt.CRViewer1.ViewReport
            Rpt.CRViewer1.Zoom 1
            
            Rpt.WindowState = 2
           Rpt.Show 1
            
            Me.MousePointer = vbDefault
                
End Select
End Sub

Sub TotalCur(xl As Excel.application, Row As Long, Col As String, coltitle As String)
With xl
    .Range("a" & Row, "j" & Row).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(coltitle & Row) = "Total"
    .Range(coltitle & Row + 1) = "PPN"
    .Range(coltitle & Row + 2) = "Grand Total"
    
    .Range(Col & Row) = Format(Amounti, gs_formatAmount)
    .Range(Col & Row + 1) = Format(PPni, gs_formatAmount)
    .Range(Col & Row + 2) = Format(grandI, gs_formatAmount)
End With
End Sub

Private Sub Command1_Click() 'EXCEL
    Dim xlapp As New Excel.application
    Dim rsCek As New Recordset, Idx As Long, tempi As String, tempcls As String
    Dim bolcls As Boolean, bolcur As Boolean
    Dim rsCompany As New Recordset
    Dim AmountCls As Double, PPnCls As Double, GrandCls As Double
    
    MousePointer = vbHourglass
        sql = " select pr.Receipt_Cls, tm.country_cls,pr.supplier_code,tm.trade_name,receipt_date as deliveryDate,makeritem_code as partNumber,partName = case isnull(im.sheetcoil_cls,0) when 0 then  im.item_name else rtrim(im.item_name) + ' (' + rtrim(sh.description) + ', T' + cast(im.thickness as varchar(15)) + ' x W' + cast(im.width as varchar(15)) + ' x L' + cast(im.length as varchar(15)) + ')'  end, " & _
                    "suratjalan_no,po_no,qty,pr.unit_cls, " & _
                    " unitDesc = (select description from unit_cls a where a.unit_cls= pr.unit_cls ), " & _
                    " currDesc= (select description from curr_cls b where b.curr_cls= pr.currency_code ), " & _
                " price,amount,rtrim(company_name) company_name, company_code, rtrim(cp.address1) address1, rtrim(cp.address2) address2, rtrim(cp.phone1) phone1, rtrim(cp.phone2) phone2, rtrim(cp.fax) fax  " & _
                "from part_receipt pr join item_master im on pr.item_code=im.item_code join Trade_master tm on pr.supplier_code=tm.trade_code " & _
                " left join sheetcoil_cls sh on im.sheetcoil_cls = sh.sheetcoil_cls ,company_profile cp " & _
                " where pr.supplier_code='" & Trim(CboLocationCD) & "' and " & _
                " receipt_date>='" & Format(DMonth, "yyyy-MM-dd") & "' and " & _
                " receipt_date<='" & Format(Dmonth2, "yyyy-MM-dd") & "' and (receipt_cls ='R' or receipt_cls ='R1')" & _
                "order by pr.Receipt_Cls, currency_code "
    If rsCek.State <> adStateClosed Then rsCek.Close
    'rsCek.CursorLocation = adUseClient
    rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic
    
    If rsCek.EOF Then
        LblErrMsg.Caption = DisplayMsg(4006)
    Else
            
        With xlapp
            
            sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax  From company_profile "
            If rsCompany.State <> adStateClosed Then rsCompany.Close
            rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
            If rsCompany.EOF Then MousePointer = vbDefault: Exit Sub
            .Workbooks.Add
            
            .Range("a2", "j2").Merge
            .Range("a2") = "Part/Material Receipt List"
            .Range("e4", "j4").Merge
            .Range("e4") = rsCompany!company_name
            .Range("e5", "j5").Merge
            .Range("e5") = rsCompany!address1 & " " & rsCompany!address2 & " " & rsCompany!City
            .Range("e6", "j6").Merge
            .Range("e6") = "Phone : " & rsCompany!phone1 & " " & rsCompany!phone2
            .Range("e7", "j7").Merge
            .Range("e7") = "Fax : " & rsCompany!fax
            
            .Range("a8") = "Supplier Code"
            .Range("b8") = ": " & Trim(rsCek!Supplier_Code)
            .Range("c8", "d8").Merge
            .Range("c8") = "Supplier Name :  " & rsCek!trade_name
            .Range("b9", "d9").Merge
            .Range("a9") = "Date"
            .Range("b9") = ": " & Format(DMonth, "dd MMMM YYYY") & " to " & Format(Dmonth2, "dd MMMM YYYY")
            
            Idx = 11
            tempi = ""
            
            Amounti = 0
            PPni = 0
            grandI = 0
            
            AmountCls = 0
            PPnCls = 0
            GrandCls = 0
            
            Do While Not rsCek.EOF
                bolcls = False
                bolcur = False
                                
                If Idx <> 11 And (tempi <> Trim(rsCek!CurrDesc) Or tempcls <> Trim(rsCek!receipt_cls)) Then
                    Call TotalCur(xlapp, Idx, "j", "i")
                    bolcur = True
                    Amounti = 0
                    PPni = 0
                    grandI = 0
                End If
        
                If Idx <> 11 And tempcls <> Trim(rsCek!receipt_cls) Then
        '            Idx = Idx + 3
        '            Call TotalCls(xlapp, Idx, "j", "i", "e", "h", tempcls)
                    bolcls = True
                    AmountCls = 0
                    PPnCls = 0
                    GrandCls = 0
                End If
        
                If Idx = 11 Or bolcls Then
                    If Idx <> 11 Then
                        Idx = Idx + 3
                    End If
                    If Trim(rsCek!receipt_cls) = "R1" Then
                        .Range("a" & Idx) = "Return"
                    End If
                End If
        
                If Idx = 11 Or bolcur Then
                    If Idx = 11 Or bolcls Then
                        Idx = Idx + 1
                    Else
                        Idx = Idx + 4
                    End If
                    .Range("a" & Idx) = "Currency"
                    .Range("b" & Idx) = ": " & Trim(rsCek!CurrDesc)
        
                    Idx = Idx + 1
                    .Range("a" & Idx).HorizontalAlignment = xlCenter
                    .Range("a" & Idx) = "Delivery Date"
                    .Range("b" & Idx) = "Product Code"
                    .Range("c" & Idx) = "Description"
                    .Range("d" & Idx) = "Surat Jalan No"
                    .Range("e" & Idx) = "PO No"
                    .Range("f" & Idx) = "Qty"
                    .Range("g" & Idx) = "Unit"
                    .Range("h" & Idx) = "Curr"
                    .Range("i" & Idx) = "Price"
                    .Range("j" & Idx) = "Amount"
                    .Range("a" & Idx, "j" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Range("a" & Idx, "j" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Idx = Idx + 1
                End If
        
                Idx = Idx
                'Content
                .Range("a" & Idx).HorizontalAlignment = xlCenter
                .Range("a" & Idx) = Format(rsCek!DeliveryDate, "DD-MMM-YYYY")
                .Range("b" & Idx) = Trim(rsCek!partNumber)
                .Range("c" & Idx) = Trim(rsCek!PartName)
                .Range("d" & Idx) = "'" & Trim(rsCek!SuratJalan_No)
                .Range("e" & Idx) = "'" & Trim(rsCek!po_no)
                .Range("f" & Idx) = Format(rsCek!Qty, gs_formatQty)
                .Range("g" & Idx) = Trim(rsCek!unitdesc)
                If bteHakPrice = 1 Then
                    .Range("h" & Idx) = Trim(rsCek!CurrDesc)
                    .Range("i" & Idx) = Format(rsCek!Price, gs_formatPrice)
                    .Range("j" & Idx) = Format(rsCek!Amount, gs_formatAmount)
                End If
                
                Idx = Idx + 1
                tempcls = Trim(rsCek!receipt_cls)
                Amounti = Amounti + rsCek!Amount
                PPni = PPni + (rsCek!Amount * 0.1)
                grandI = grandI + (rsCek!Amount + (rsCek!Amount * 0.1))
        
                AmountCls = AmountCls + rsCek!Amount
                PPnCls = PPnCls + (rsCek!Amount * 0.1)
                GrandCls = GrandCls + (rsCek!Amount + (rsCek!Amount * 0.1))
                
                tempi = Trim(rsCek!CurrDesc)
                rsCek.MoveNext
            Loop
            
            If bteHakPrice = 1 Then
                Call TotalCur(xlapp, Idx, "j", "i")
            Else
                .Range("a" & Idx, "j" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
            Idx = Idx + 3
        
            .Range("F4:F" & Idx).Select
            .Selection.NumberFormat = gs_formatQty
            
            .Range("I4:I" & Idx).Select
            .Selection.NumberFormat = gs_formatPrice
            
            .Range("J4:J" & Idx).Select
            .Selection.NumberFormat = gs_formatAmount
            
            .Range("a1", "j" & Idx + 3).Columns.Font.Name = "Arial"
            .Range("a1", "j" & Idx + 3).Columns.Font.Size = 8
            
            .Range("a2", "j2").Columns.Font.Name = "Arial"
            .Range("a2", "j2").Columns.Font.Size = "10"
            .Range("a2", "j2").Columns.Font.Bold = True
            .Range("a2", "j2").HorizontalAlignment = xlCenter
            
            .Range("a1", "j" & Idx + 3).Columns.AutoFit
            .Range("A1").Select
            
            .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
            .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.04)
            .ActiveSheet.PageSetup.Orientation = 2
            .WindowState = xlMaximized
            .Visible = True
        
        End With
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub DMonth_Change()
If CDate(Dmonth2) < CDate(DMonth) Then
      LblErrMsg.Caption = DisplayMsg(4068)
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
End Sub

Private Sub DMonth_Click()
If CDate(Dmonth2) < CDate(DMonth) Then
      LblErrMsg.Caption = DisplayMsg(4068)
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
End Sub


Private Sub Dmonth2_Change()
If CDate(Dmonth2) < CDate(DMonth) Then
      LblErrMsg.Caption = DisplayMsg(4066)
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
End Sub

Private Sub Dmonth2_Click()
If CDate(Dmonth2) < CDate(DMonth) Then
      LblErrMsg.Caption = DisplayMsg(4066)
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
End Sub

Private Sub Form_Load()
Dim RsW As New Recordset
Dim ir As Integer
If gb_Simulation = True Then Call up_InitSimulation(Me)
LblLocationName = ""
LblErrMsg = ""
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
bteHakPrice = hakPrice(Me.Name)
Call StockLocation
DMonth = Format(Date, "dd MMM yyyy")
Dmonth2 = Format(Date, "dd MMM yyyy")

'##Tampilkan Warehouse code dari warehouse_master
sql = "Select rtrim(wh_code) as WC,wh_name as WN from warehouse_master order by wh_code"
RsW.Open sql, Db, adOpenKeyset, adLockOptimistic
cboWhCode.clear
cboWhCode.columnCount = 2
cboWhCode.TextColumn = 1
cboWhCode.AddItem ""
cboWhCode.List(0, 0) = strAll
cboWhCode.List(0, 1) = strAll
ir = 1
While Not RsW.EOF
    cboWhCode.AddItem ""
    cboWhCode.List(ir, 0) = RsW!wC
    cboWhCode.List(ir, 1) = Trim$(RsW!wn)
    ir = ir + 1
    RsW.MoveNext
Wend
cboWhCode.ColumnWidths = "80 pt; 180 pt"
cboWhCode.ListWidth = 260
cboWhCode.ListRows = 15
cboWhCode.ListIndex = 0
End Sub


Private Sub StockLocation()
Dim sql As String, RsStock As New ADODB.Recordset
Dim i As Integer

If RsStock.State <> adStateClosed Then RsStock.Close
RsStock.Open " trade_master where trade_cls='2' or trade_cls='3' order by trade_code asc", Db, adOpenDynamic, adLockOptimistic, adCmdTable

CboLocationCD.columnCount = 2
CboLocationCD.clear

CboLocationCD.AddItem
CboLocationCD.List(i, 0) = strAll
CboLocationCD.List(i, 1) = strAll


i = 1
Do While Not RsStock.EOF
   CboLocationCD.AddItem ""
   CboLocationCD.List(i, 0) = Trim(RsStock("trade_code"))
   CboLocationCD.List(i, 1) = Trim(RsStock("trade_name"))
   i = i + 1
   RsStock.MoveNext
Loop
CboLocationCD.ListIndex = 0
CboLocationCD.ColumnWidths = "80 pt; 250 pt"
CboLocationCD.ListWidth = 320
CboLocationCD.ListRows = 15
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub
Function tax(Tgl1$, Tgl2$) As String
Dim rtax As Recordset
    
    sql = "SELECT Tax_Code, Rate, start_Date, End_Date " & _
            "FROM tax_cls " & _
            "" & _
            "WHERE  " & _
            "start_date <= '" & Tgl1 & "'  and end_date >= '" & Tgl2 & "' " & _
            "and tax_code= 'PPN'"
    Set rtax = New Recordset
    rtax.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not rtax.EOF Then
        tax = rtax!rate
    Else
        tax = 0
    End If


End Function



