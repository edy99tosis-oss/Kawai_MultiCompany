VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_IncomingMaterialReport 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incoming Material Report"
   ClientHeight    =   6030
   ClientLeft      =   1320
   ClientTop       =   2760
   ClientWidth     =   8295
   Icon            =   "frm_IncomingMaterialReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkInclude 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Tools"
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
      Left            =   2160
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Upload 
      BackColor       =   &H0080FFFF&
      Caption         =   "Upload"
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
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4980
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   2190
      TabIndex        =   17
      Top             =   2130
      Width           =   3675
      Begin VB.OptionButton OptMaterial 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Material Cls"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2190
         TabIndex        =   19
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton OptGeneral 
         BackColor       =   &H00FDDFE3&
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   3
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptDetail 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Detail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   4
         Top             =   120
         Width           =   945
      End
   End
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
      Top             =   3345
      Visible         =   0   'False
      Width           =   3870
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
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "LblLocationName"
      Top             =   1785
      Width           =   3870
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
      Left            =   6780
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4965
      Width           =   1035
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   5989
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   285
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
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
      TabIndex        =   7
      Top             =   4965
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   379
      TabIndex        =   9
      Top             =   4170
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
         TabIndex        =   10
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
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4950
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker DTo 
      Height          =   315
      Left            =   4290
      TabIndex        =   0
      Top             =   1260
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
      Format          =   141230083
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin MSComCtl2.DTPicker DFrom 
      Height          =   315
      Left            =   2190
      TabIndex        =   23
      Top             =   1260
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
      Format          =   141230083
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin VB.Label Label5 
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
      Left            =   3870
      TabIndex        =   24
      Top             =   1350
      Width           =   210
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Tools"
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
      Left            =   570
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Material Cls"
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
      Left            =   570
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   3795
      X2              =   7650
      Y1              =   3585
      Y2              =   3585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview In"
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
      Left            =   570
      TabIndex        =   15
      Top             =   2325
      Width           =   915
   End
   Begin MSForms.ComboBox CboMaterialCls 
      Height          =   315
      Left            =   2175
      TabIndex        =   1
      Top             =   3285
      Visible         =   0   'False
      Width           =   1500
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "2646;556"
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
      X2              =   7635
      Y1              =   2025
      Y2              =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date From"
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
      Left            =   570
      TabIndex        =   14
      Top             =   1320
      Width           =   900
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
      Left            =   570
      TabIndex        =   13
      Top             =   1785
      Width           =   705
   End
   Begin MSForms.ComboBox CboLocationCD 
      Height          =   315
      Left            =   2175
      TabIndex        =   2
      Top             =   1725
      Width           =   1500
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "2646;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Incoming Material Report"
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
      Left            =   375
      TabIndex        =   11
      Top             =   570
      Width           =   7470
   End
End
Attribute VB_Name = "frm_IncomingMaterialReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Amounti As Double
Dim PPni As Double
Dim grandI As Double
Dim TotalBC_Ori, TotalBC_USD As Double
Dim TotalSupplier_Ori, TotalSupplier_USD As Double
Dim TotalMaterial_Ori, TotalMaterial_USD As Double
Dim GrandTotal_Ori, GrandTotal_USD As Double
Const Include = "07"
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

Private Sub CboMaterialCls_Change()
If CboMaterialCls.MatchFound Then
    lblWHName = CboMaterialCls.Column(1)
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
                "pr.qty, pr.currency_code, pr.price, pr.amount, rtrim(pr.suratjalan_no) suratjalan_no, rtrim(pr.remarks) remarks, rtrim(pr.bc40_no) bc40_no, package_qty, " & _
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
                " pr.receipt_date>='" & Format(DFrom, "yyyy-mm-dd") & "' and pr.receipt_date<='" & Format(DTo, "yyyy-mm-dd") & "'  "
                
                
                If CboMaterialCls <> strAll Then sql = sql & "and pr.warehouse_code = '" & Trim(CboMaterialCls) & "' "
                sql = sql & "order by pr.receipt_date, pr.item_code"
                                  
              If rsRpt.State <> adStateClosed Then rsRpt.Close
              rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
              
            
              If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
              sqlprint = sql
              reportcode = "AccPay"
              printorient = 1
            
              Set report = application.OpenReport(App.path & "\Reports\rpt_accPay.rpt")
              report.Database.Tables(1).SetDataSource rsRpt
              
              
            report.FormulaFields(1).Text = "'" & Format(DFrom, "dd-MMM-yyyy") & " To " & Format(DTo, "dd-MMM-yyyy") & ""
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
    Dim rsCek As New Recordset, Idx As Long, tempi As String, tempcls, tempsupp, tempbc, tempmatcls, tempsuppname As String
    Dim tempmatname As String, TempCountry As String
    Dim bolcls As Boolean, bolcur As Boolean
    Dim rsCompany As New Recordset
    Dim i As Long
    Dim totalQty As Double, TotalAmount As Double, subtotalQty As Double, SubTotalAmount As Double
    
    
    'If MsgBox("Choose 'Yes' To Preview In General Or Choose 'No' To Preview In Detail.", vbYesNo + vbInformation) = vbYes Then
    If OptGeneral.Value = True Then
        DoEvents
        MousePointer = vbHourglass
        sql = " select a.supplier_code,a.trade_name,sum(a.qty)qty,a.currdesc,sum(a.amount)Original_amount, sum(a.AmountConvertion)Amount_Convertion from ---tambah qty dl " & vbCrLf & _
                    " ( " & vbCrLf & _
                    "  select pr.Receipt_Cls, tm.country_cls,pr.suratjalan_no surat_jalan,  " & vbCrLf & _
                    "    pr.supplier_code,tm.trade_name,pr.receipt_date as deliveryDate,im.item_code as Item_Code,  " & vbCrLf & _
                    "    partName =  case isnull(im.sheetcoil_cls,0) when 0 then  im.item_name   " & vbCrLf & _
                    "                    else rtrim(im.item_name) + ' (' +   " & vbCrLf & _
                    "                        rtrim(sh.description) + ', T' +   " & vbCrLf & _
                    "                        cast(im.thickness as varchar(15)) + ' x W' +   " & vbCrLf & _
                    "                        cast(im.width as varchar(15)) + ' x L' +   " & vbCrLf & _
                    "                        cast(im.length as varchar(15)) + ')'  end,   " & vbCrLf & _
                    "    isnull(im.Material_Cls,'')material_cls,isnull((Select isnull(Description,'')description From Material_Cls m Where m.Material_Cls=im.Material_Cls),'') Material_Description,  " & vbCrLf
        
        sql = sql + "    pr.suratjalan_no,pr.po_no,pr.qty,pr.unit_cls, --pr.BC40_No, pr.BC40_Date,  " & vbCrLf & _
                    "    unitDesc = (select description from unit_cls a where a.unit_cls= pr.unit_cls ),      currDesc= (select description from curr_cls b where b.curr_cls= pr.currency_code ),   " & vbCrLf & _
                    "    TaxExchangeRate=(Select Tax_ExchangeRate From Tax_ExchangeRate Where Currency_Code='02' and Start_Date<=convert(varchar, pr.receipt_date, 112) and End_Date>=convert(varchar, pr.receipt_date, 112) ),   " & vbCrLf & _
                    "    AmountConvertion=   case pr.Currency_Code  " & vbCrLf & _
                    "                            when '02' then pr.amount  " & vbCrLf & _
                    "                            when '03' then pr.amount/(Select Tax_ExchangeRate From Tax_ExchangeRate Where Currency_Code='02' and Start_Date<=convert(varchar, pr.receipt_date, 112) and End_Date>=convert(varchar, pr.receipt_date, 112)  )  " & vbCrLf & _
                    "                            else (pr.amount*(Select (Tax_ExchangeRate)TaxExchangeRateRupiah From Tax_ExchangeRate Where Currency_Code='03' and Start_Date<=convert(varchar, pr.receipt_date, 112)  and End_Date>=convert(varchar, pr.receipt_date, 112)  ))/(Select Tax_ExchangeRate From Tax_ExchangeRate Where Currency_Code='02' and Start_Date<=convert(varchar, pr.receipt_date, 112)  and End_Date>=convert(varchar, pr.receipt_date, 112) )  " & vbCrLf & _
                    "                            end ,  " & vbCrLf & _
                    "    pr.price,pr.amount,ism.invoice_no,ism.invoice_date,  " & vbCrLf & _
                    "    rtrim(company_name) company_name, cp.company_code, rtrim(cp.address1) address1, rtrim(cp.address2) address2,   " & vbCrLf & _
                    "    rtrim(cp.phone1) phone1, rtrim(cp.phone2) phone2, rtrim(cp.fax) fax, POM.PO_Date, im.Group_Cls   " & vbCrLf
        
        sql = sql + "  from part_receipt pr     join item_master im on pr.item_code=im.item_code   " & vbCrLf & _
                    "    join Trade_master tm on pr.supplier_code=tm.trade_code    " & vbCrLf & _
                    "    Left Join PurchaseOrder_Master POM on PR.PO_NO = POM.PO_No  " & vbCrLf & _
                    "    left join InvoiceSupplier_Detail isd on pr.seq_no=isd.receiptseq_no  " & vbCrLf & _
                    "    left join InvoiceSupplier_Master ism on isd.invoice_no=ism.invoice_no  " & vbCrLf & _
                    "    left join sheetcoil_cls sh on im.sheetcoil_cls = sh.sheetcoil_cls ,company_profile cp    " & vbCrLf & _
                    "  where pr.receipt_date>='" & Format(DFrom, "yyyy-mm-dd") & "' and pr.receipt_date<='" & Format(DTo, "yyyy-mm-dd") & "'  " & vbCrLf & _
                    "    and (pr.receipt_cls ='R' or pr.receipt_cls ='R1') " & vbCrLf & _
                    "    " & IIf(Trim(CboLocationCD.Text) <> strAll, "   and pr.supplier_code='" & Trim(CboLocationCD.Text) & "'", "") & vbCrLf & _
                    " ) a " & vbCrLf
                    
'        If ChkInclude.Value = 0 Then
'            sql = sql + " Where a.Group_Cls<>'" & Include & "' " & vbCrLf
'        End If
        
        sql = sql + " group by a.supplier_code,a.trade_name,currdesc --group by a.Receipt_Cls, a.country_cls,a.suratjalan_no,a.unitdesc,  " & vbCrLf & _
                    " --   a.supplier_code,a.trade_name,a.item_code,a.material_description, " & vbCrLf & _
                    " --    a.Material_Cls,a.po_no,a.unit_cls, a.BC40_No, a.BC40_Date, " & vbCrLf & _
                    " --   a.amount,a.surat_jalan,a.deliverydate,a.partname,a.currdesc,  " & vbCrLf & _
                    " --   a.price,a.invoice_no,a.invoice_date,a.taxexchangerate,a.amountconvertion,  " & vbCrLf & _
                    " --   a.company_name, a.company_code, a.address1, a.address2,   " & vbCrLf & _
                    " --   a.phone1, a.phone2, a.fax, a.PO_Date    " & vbCrLf & _
                    " order by a.Supplier_Code--,a.BC40_No  "
                    
  If rsCek.State <> adStateClosed Then rsCek.Close
        
        rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic
        
        If rsCek.EOF Then
            LblErrMsg.Caption = DisplayMsg(4006)
        Else
                
            With xlapp
            
                LblErrMsg.Caption = "[1719] Please Wait While Export To Excel....!"
                
                sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax  From company_profile "
                If rsCompany.State <> adStateClosed Then rsCompany.Close
                rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
                If rsCompany.EOF Then MousePointer = vbDefault: Exit Sub
                .Workbooks.Add
                
                .Range("a1") = rsCompany!company_name
                .Range("a2") = "INCOMING MATERIAL FOR"
                .Range("a3") = "PERIOD " & Format(DFrom, "dd MMM yyyy") & " To " & Format(DTo, "dd MMM yyyy")
                .Range("a4") = " "
                
                .Range("a5") = "Supplier Name"
                
                .Range("b5") = "Qty"
                .Range("b5").horizontalAlignment = xlRight
                
                .Range("c5") = "Curr"
                .Range("c5").horizontalAlignment = xlCenter
                
                .Range("d5") = "Original Amount"
                .Range("d5").horizontalAlignment = xlRight
                
                .Range("e5") = "Amount (USD)"
                .Range("e5").horizontalAlignment = xlRight
                
                .Range("a" & 5, "e" & 5).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Range("a" & 5, "e" & 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range("a" & 5, "e" & 5).Font.Bold = True
                .Range("a" & 5, "e" & 5).horizontalAlignment = xlCenter
                .Range("a" & 5, "e" & 5).Interior.ColorIndex = 38
                
                
                Idx = 6
                GrandTotal_Ori = 0
                totalQty = 0
                
                
                
                Do While Not rsCek.EOF
                    
                    'Content
                    .Range("a" & Idx).horizontalAlignment = xlLeft
                    .Range("a" & Idx) = "" & Trim(rsCek!Supplier_Code) & " - " & Trim(rsCek!trade_name) & ""
                    
                    .Range("b" & Idx) = IIf(IsNull(rsCek!Qty), 0, rsCek!Qty)
                    .Range("b" & Idx).horizontalAlignment = xlRight
                    .Range("b" & Idx).Select
                    .Selection.NumberFormat = gs_formatQty
                    
                    .Range("c" & Idx) = IIf(IsNull(rsCek!CurrDesc), "", Trim(rsCek!CurrDesc))
                    .Range("c" & Idx).horizontalAlignment = xlCenter
                    
                    .Range("d" & Idx) = IIf(IsNull(rsCek!original_amount), "", rsCek!original_amount)
                    .Range("d" & Idx).horizontalAlignment = xlRight
                    
                    .Range("e" & Idx) = IIf(IsNull(rsCek!amount_convertion), "", rsCek!amount_convertion)
                    .Range("e" & Idx).horizontalAlignment = xlRight
                    
                    GrandTotal_Ori = GrandTotal_Ori + IIf(IsNull(rsCek!amount_convertion), 0, rsCek!amount_convertion)
                    totalQty = totalQty + IIf(IsNull(rsCek!Qty), "", rsCek!Qty)
                    
                    If Trim(rsCek!CurrDesc) = "IDR" Then
                        .Range("d" & Idx).Select
                        .Selection.NumberFormat = gs_formatPriceIDR
                    Else
                        .Range("d" & Idx).Select
                        .Selection.NumberFormat = gs_formatPrice
                    End If
                    
                    .Range("e" & Idx).Select
                    .Selection.NumberFormat = gs_formatPrice
                    .Range("a" & Idx, "e" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    
                    rsCek.MoveNext
                    Idx = Idx + 1
                    
                Loop
                
                .Range("a" & 5, "d" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                
                'Grand Total
                '.Range("a" & Idx, "c" & Idx).Merge
                
                .Range("a" & Idx) = "Grand Total"
                .Range("a" & Idx).Font.Bold = True
                .Range("a" & Idx).horizontalAlignment = xlLeft
                .Range("e" & Idx) = GrandTotal_Ori
                .Range("e" & Idx).Select
                .Selection.NumberFormat = gs_formatPrice
                
                'Grandtotal Qty
                
                .Range("b" & Idx) = totalQty
                .Range("b" & Idx).Select
                .Selection.NumberFormat = gs_formatQty
                
                .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
                .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.04)
                .ActiveSheet.PageSetup.Orientation = xlPortrait
                .ActiveSheet.PageSetup.PaperSize = xlPaperA4
                .ActiveSheet.PageSetup.PrintArea = "a1:e" & Idx
                
                .Range("a:f").Columns.AutoFit
                .WindowState = xlMaximized
                
                .Visible = True
                LblErrMsg.Caption = ""
            End With
        End If
        
    ElseIf OptMaterial.Value = True Then
    
        DoEvents
        
        MousePointer = vbHourglass
                
        sql = "  select a.item_code,a.partName,sum(a.qty)qty, sum(a.amount)amount, sum(a.amountconvertion)amount_convertion, Country_Cls=case a.country_cls when '0' then 'Domestic' when '1' then 'Overseas' end from  " & vbCrLf & _
                    "  (  " & vbCrLf & _
                    "   select pr.Receipt_Cls, tm.country_cls,pr.suratjalan_no surat_jalan,   " & vbCrLf & _
                    "     pr.supplier_code,tm.trade_name,pr.receipt_date as deliveryDate,im.item_code as Item_Code,   " & vbCrLf & _
                    "     partName =  case isnull(im.sheetcoil_cls,0) when 0 then  im.item_name    " & vbCrLf & _
                    "                     else rtrim(im.item_name) + ' (' +    " & vbCrLf & _
                    "                         rtrim(sh.description) + ', T' +    " & vbCrLf & _
                    "                         cast(im.thickness as varchar(15)) + ' x W' +    " & vbCrLf & _
                    "                         cast(im.width as varchar(15)) + ' x L' +    " & vbCrLf & _
                    "                         cast(im.length as varchar(15)) + ')'  end,    " & vbCrLf & _
                    "     isnull(im.Material_Cls,'')material_cls,isnull((Select isnull(Description,'')description From Material_Cls m Where m.Material_Cls=im.Material_Cls),'') Material_Description,      pr.suratjalan_no,pr.po_no,pr.qty,pr.unit_cls, pr.BC40_No, pr.BC40_Date,   "
        
        sql = sql + "     unitDesc = (select description from unit_cls a where a.unit_cls= pr.unit_cls ),      currDesc= (select description from curr_cls b where b.curr_cls= pr.currency_code ),    " & vbCrLf & _
                    "     TaxExchangeRate=(Select Tax_ExchangeRate From Tax_ExchangeRate Where Currency_Code='02' and Start_Date<=convert(varchar, pr.receipt_date, 112)  and End_Date>=convert(varchar, pr.receipt_date, 112) ),    " & vbCrLf & _
                    "     AmountConvertion=   case pr.Currency_Code   " & vbCrLf & _
                    "                             when '02' then pr.amount   " & vbCrLf & _
                    "                             when '03' then pr.amount/(Select Tax_ExchangeRate From Tax_ExchangeRate Where Currency_Code='02' and Start_Date<=convert(varchar, pr.receipt_date, 112)  and End_Date>=convert(varchar, pr.receipt_date, 112) )   " & vbCrLf & _
                    "                             else (pr.amount*(Select (Tax_ExchangeRate)TaxExchangeRateRupiah From Tax_ExchangeRate Where Currency_Code='03' and Start_Date<=convert(varchar, pr.receipt_date, 112)  and End_Date>=convert(varchar, pr.receipt_date, 112) ))/(Select Tax_ExchangeRate From Tax_ExchangeRate Where Currency_Code='02' and Start_Date<=convert(varchar, pr.receipt_date, 112) and End_Date>=convert(varchar, pr.receipt_date, 112) )   " & vbCrLf & _
                    "                             end ,   " & vbCrLf & _
                    "     pr.price,pr.amount,ism.invoice_no,ism.invoice_date,   " & vbCrLf & _
                    "     rtrim(company_name) company_name, cp.company_code, rtrim(cp.address1) address1, rtrim(cp.address2) address2,    " & vbCrLf & _
                    "     rtrim(cp.phone1) phone1, rtrim(cp.phone2) phone2, rtrim(cp.fax) fax, POM.PO_Date       " & vbCrLf & _
                    "   from part_receipt pr      "
        
        sql = sql + "   join item_master im on pr.item_code=im.item_code    " & vbCrLf & _
                    "     join Trade_master tm on pr.supplier_code=tm.trade_code     " & vbCrLf & _
                    "     Left Join PurchaseOrder_Master POM on PR.PO_NO = POM.PO_No   " & vbCrLf & _
                    "     left join InvoiceSupplier_Detail isd on pr.seq_no=isd.receiptseq_no   " & vbCrLf & _
                    "     left join InvoiceSupplier_Master ism on isd.invoice_no=ism.invoice_no   " & vbCrLf & _
                    "     left join sheetcoil_cls sh on im.sheetcoil_cls = sh.sheetcoil_cls ,company_profile cp     " & vbCrLf & _
                    "   where pr.receipt_date>='" & Format(DFrom, "yyyy-mm-dd") & "' and pr.receipt_date<='" & Format(DTo, "yyyy-mm-dd") & "'  " & vbCrLf & _
                    "     and (pr.receipt_cls ='R' or pr.receipt_cls ='R1')  " & vbCrLf & _
                    IIf(Trim(CboLocationCD.Text) <> strAll, "  and pr.supplier_code='" & Trim(CboLocationCD.Text) & "' ", "") & " " & IIf(Trim(CboMaterialCls.Text) <> strAll, "  and im.material_cls='" & Trim(CboMaterialCls.Text) & "' ", "") & " " & IIf(ChkInclude.Value = 0, "  and im.Group_Cls<>'" & Include & "' ", "") & vbCrLf & _
                    "  ) a  " & vbCrLf & _
                    "  group by a.item_code,a.partName,a.country_cls   "
        
        sql = sql + "  order by a.country_cls, a.item_code "

  If rsCek.State <> adStateClosed Then rsCek.Close
        
        rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic
        
        If rsCek.EOF Then
            LblErrMsg.Caption = DisplayMsg(4006)
        Else
                
            With xlapp
            
                LblErrMsg.Caption = "[1719] Please Wait While Export To Excel....!"
                
                sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax  From company_profile "
                If rsCompany.State <> adStateClosed Then rsCompany.Close
                rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
                If rsCompany.EOF Then MousePointer = vbDefault: Exit Sub
                .Workbooks.Add
                
                .Range("a1") = rsCompany!company_name
                .Range("a2") = "INCOMING MATERIAL FOR"
                .Range("a3") = "PERIOD " & Format(DFrom, "dd MMM yyyy") & " To " & Format(DTo, "dd MMM yyyy")
                .Range("a4") = " "
                
                .Range("a5") = "Item Code"
                .Range("a5").horizontalAlignment = xlLeft
                
                .Range("b5") = "Item Name"
                .Range("b5").horizontalAlignment = xlLeft
                
                .Range("c5") = "Qty"
                .Range("c5").horizontalAlignment = xlRight
                
                .Range("d5") = "Amount Original"
                .Range("d5").horizontalAlignment = xlRight
                
                .Range("e5") = "Amount (USD)"
                .Range("e5").horizontalAlignment = xlRight
                
                .Range("a" & 5, "e" & 5).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Range("a" & 5, "e" & 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range("a" & 5, "e" & 5).Font.Bold = True
                .Range("a" & 5, "e" & 5).horizontalAlignment = xlCenter
                .Range("a" & 5, "e" & 5).Interior.ColorIndex = 38
                
                
                Idx = 6
                totalQty = 0
                TotalAmount = 0
                SubTotalAmount = 0
                subtotalQty = 0
                TempCountry = ""
                
                Do While Not rsCek.EOF
                    
                    'Content
                    If TempCountry = "" Or TempCountry <> rsCek!country_cls Then
                        .Range("a" & Idx).horizontalAlignment = xlLeft
                        .Range("a" & Idx) = "'" & Trim(rsCek!country_cls)
                        .Range("a" & Idx, "e" & Idx).Font.Bold = True
                        .Range("a" & Idx, "e" & Idx).Interior.ColorIndex = 19
                        .Range("a" & Idx, "e" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                        Idx = Idx + 1
                    End If
                    .Range("a" & Idx).horizontalAlignment = xlLeft
                    .Range("a" & Idx) = "'" & Trim(rsCek!Item_Code)
                    
                    .Range("b" & Idx) = IIf(IsNull(rsCek!PartName), "", Trim(rsCek!PartName))
                    .Range("b" & Idx).horizontalAlignment = xlLeft
                    
                    .Range("c" & Idx) = IIf(IsNull(rsCek!Qty), 0, rsCek!Qty)
                    .Range("c" & Idx).horizontalAlignment = xlRight
                    .Range("c" & Idx).Select
                    .Selection.NumberFormat = gs_formatQty
                    
                    .Range("d" & Idx) = IIf(IsNull(rsCek!Amount), 0, rsCek!Amount)
                    .Range("d" & Idx).horizontalAlignment = xlRight
                    .Range("d" & Idx).Select
                    .Selection.NumberFormat = gs_formatAmount
                    
                    .Range("e" & Idx) = IIf(IsNull(rsCek!amount_convertion), 0, rsCek!amount_convertion)
                    .Range("e" & Idx).horizontalAlignment = xlRight
                    .Range("e" & Idx).Select
                    .Selection.NumberFormat = gs_formatAmount
                    
                    totalQty = totalQty + IIf(IsNull(rsCek!Qty), "", rsCek!Qty)
                    TotalAmount = TotalAmount + IIf(IsNull(rsCek!amount_convertion), 0, rsCek!amount_convertion)
                    SubTotalAmount = SubTotalAmount + IIf(IsNull(rsCek!amount_convertion), 0, rsCek!amount_convertion)
                    subtotalQty = subtotalQty + IIf(IsNull(rsCek!Qty), "", rsCek!Qty)
                    
                    .Range("a" & Idx, "e" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    TempCountry = rsCek!country_cls
                    rsCek.MoveNext
                    Idx = Idx + 1
                    If rsCek.EOF Then
                        .Range("a" & Idx).horizontalAlignment = xlLeft
                        .Range("a" & Idx) = "Sub Total " & TempCountry
                        'SubTotal Qty
                        .Range("c" & Idx) = subtotalQty
                        .Range("c" & Idx).Select
                        .Selection.NumberFormat = gs_formatQty
                        
                        'SubTotal Amount
                        .Range("e" & Idx) = SubTotalAmount
                        .Range("e" & Idx).Select
                        .Selection.NumberFormat = gs_formatAmount
                        
                        .Range("a" & Idx, "e" & Idx).Font.Bold = True
                        .Range("a" & Idx, "e" & Idx).Interior.ColorIndex = 20
                        .Range("a" & Idx, "e" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                        SubTotalAmount = 0
                        subtotalQty = 0
                        Idx = Idx + 1
                    ElseIf Not rsCek.EOF And TempCountry <> rsCek!country_cls Then
                        .Range("a" & Idx).horizontalAlignment = xlLeft
                        .Range("a" & Idx) = "Sub Total " & TempCountry
                        'SubTotal Qty
                        .Range("c" & Idx) = subtotalQty
                        .Range("c" & Idx).Select
                        .Selection.NumberFormat = gs_formatQty
                        
                        'SubTotal Amount
                        .Range("e" & Idx) = SubTotalAmount
                        .Range("e" & Idx).Select
                        .Selection.NumberFormat = gs_formatAmount
                        
                        .Range("a" & Idx, "e" & Idx).Font.Bold = True
                        .Range("a" & Idx, "e" & Idx).Interior.ColorIndex = 20
                        .Range("a" & Idx, "e" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                        SubTotalAmount = 0
                        subtotalQty = 0
                        Idx = Idx + 1
                    End If
                    
                Loop
                
                .Range("a" & 5, "e" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range("a" & Idx) = "Grand Total"
                
                'Grandtotal Qty
                .Range("c" & Idx) = totalQty
                .Range("c" & Idx).Select
                .Selection.NumberFormat = gs_formatQty
                
                'Grandtotal Amount
                .Range("e" & Idx) = TotalAmount
                .Range("e" & Idx).Select
                .Selection.NumberFormat = gs_formatAmount
                
                .Range("a" & Idx, "e" & Idx).Font.Bold = True
                .Range("a" & Idx, "e" & Idx).Interior.ColorIndex = 15
                .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
                .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.04)
                .ActiveSheet.PageSetup.Orientation = xlPortrait
                .ActiveSheet.PageSetup.PaperSize = xlPaperA4
                .ActiveSheet.PageSetup.PrintArea = "a1:e" & Idx
                
                .Range("a:e").Columns.AutoFit
                .WindowState = xlMaximized
'                For Idx = 1 To 70
'                    .Range("e" & Idx).Interior.ColorIndex = Idx
'                Next Idx
                .Visible = True
                LblErrMsg.Caption = ""
            End With
        End If
        
    Else
        DoEvents
        MousePointer = vbHourglass
    
    '        sql = " select pr.Receipt_Cls, tm.country_cls,pr.supplier_code,tm.trade_name,receipt_date as deliveryDate,makeritem_code as partNumber,partName = case isnull(im.sheetcoil_cls,0) when 0 then  im.item_name else rtrim(im.item_name) + ' (' + rtrim(sh.description) + ', T' + cast(im.thickness as varchar(15)) + ' x W' + cast(im.width as varchar(15)) + ' x L' + cast(im.length as varchar(15)) + ')'  end, " & _
    '                    "suratjalan_no,po_no,qty,pr.unit_cls, " & _
    '                    " unitDesc = (select description from unit_cls a where a.unit_cls= pr.unit_cls ), " & _
    '                    " currDesc= (select description from curr_cls b where b.curr_cls= pr.currency_code ), " & _
    '                " price,amount,rtrim(company_name) company_name, company_code, rtrim(cp.address1) address1, rtrim(cp.address2) address2, rtrim(cp.phone1) phone1, rtrim(cp.phone2) phone2, rtrim(cp.fax) fax  " & _
    '                "from part_receipt pr join item_master im on pr.item_code=im.item_code join Trade_master tm on pr.supplier_code=tm.trade_code " & _
    '                " left join sheetcoil_cls sh on im.sheetcoil_cls = sh.sheetcoil_cls ,company_profile cp " & _
    '                " where pr.supplier_code='" & Trim(CboLocationCD) & "' and " & _
    '                " receipt_date='" & Format(DMonth, "yyyy-MM-dd") & "' and " & _
    '                " (receipt_cls ='R' or receipt_cls ='R1')" & _
    '                "order by pr.Receipt_Cls, currency_code "
    
            sql = " select pr.Receipt_Cls, tm.country_cls,rtrim(pr.suratjalan_no)suratjalan_no,pr.price,pr.BC_type, " & vbCrLf & _
                        "   pr.supplier_code,tm.trade_name,pr.receipt_date as deliveryDate,im.item_code as Item_Code, " & vbCrLf & _
                        "   partName =  case isnull(im.sheetcoil_cls,0) when 0 then  im.item_name  " & vbCrLf & _
                        "                   else rtrim(im.item_name) + ' (' +  " & vbCrLf & _
                        "                       rtrim(sh.description) + ', T' +  " & vbCrLf & _
                        "                       cast(im.thickness as varchar(15)) + ' x W' +  " & vbCrLf & _
                        "                       cast(im.width as varchar(15)) + ' x L' +  " & vbCrLf & _
                        "                       cast(im.length as varchar(15)) + ')'  end,  " & vbCrLf & _
                        "   isnull(im.Material_Cls,'')material_cls,isnull((Select isnull(Description,'')description From Material_Cls m Where m.Material_Cls=im.Material_Cls),'') Material_Description, " & vbCrLf & _
                        "   pr.suratjalan_no,pr.po_no,pr.qty,pr.unit_cls, pr.BC40_No, pr.BC40_Date, " & vbCrLf & _
                        "   unitDesc = (select description from unit_cls a where a.unit_cls= pr.unit_cls ),   " & vbCrLf
            
            sql = sql + "   currDesc= (select description from curr_cls b where b.curr_cls= pr.currency_code ),  " & vbCrLf & _
                        "   TaxExchangeRate=(Select Tax_ExchangeRate From Tax_ExchangeRate Where Currency_Code='02' and Start_Date<=convert(varchar, pr.receipt_date, 112)  and End_Date>=convert(varchar, pr.receipt_date, 112) ),  " & vbCrLf & _
                        "   AmountConvertion=   case pr.Currency_Code " & vbCrLf & _
                        "                           when '02' then pr.amount " & vbCrLf & _
                        "                           when '03' then pr.amount/(Select Tax_ExchangeRate From Tax_ExchangeRate Where Currency_Code='02' and Start_Date<=convert(varchar, pr.receipt_date, 112)  and End_Date>=convert(varchar, pr.receipt_date, 112) ) " & vbCrLf & _
                        "                           else (pr.amount*(Select (Tax_ExchangeRate)TaxExchangeRateRupiah From Tax_ExchangeRate Where Currency_Code='03' and Start_Date<=convert(varchar, pr.receipt_date, 112)  and End_Date>=convert(varchar, pr.receipt_date, 112) ))/(Select Tax_ExchangeRate From Tax_ExchangeRate Where Currency_Code='02' and Start_Date<=convert(varchar, pr.receipt_date, 112) and End_Date>=convert(varchar, pr.receipt_date, 112) )  " & vbCrLf & _
                        "                           end , " & vbCrLf & _
                        "   pr.price,pr.amount,ism.invoice_no,ism.invoice_date, " & vbCrLf & _
                        "   rtrim(company_name) company_name, cp.company_code, rtrim(cp.address1) address1, rtrim(cp.address2) address2,  " & vbCrLf & _
                        "   rtrim(cp.phone1) phone1, rtrim(cp.phone2) phone2, rtrim(cp.fax) fax, POM.PO_Date, im.Group_Cls   " & vbCrLf & _
                        " from part_receipt pr  " & vbCrLf
            
            sql = sql + "   join item_master im on pr.item_code=im.item_code  " & vbCrLf & _
                        "   join Trade_master tm on pr.supplier_code=tm.trade_code   " & vbCrLf & _
                        "   Left Join PurchaseOrder_Master POM on PR.PO_NO = POM.PO_No " & vbCrLf & _
                        "   left join InvoiceSupplier_Detail isd on pr.seq_no=isd.receiptseq_no " & vbCrLf & _
                        "   left join InvoiceSupplier_Master ism on isd.invoice_no=ism.invoice_no " & vbCrLf & _
                        "   left join sheetcoil_cls sh on im.sheetcoil_cls = sh.sheetcoil_cls ,company_profile cp   " & vbCrLf & _
                        " where pr.receipt_date>='" & Format(DFrom, "yyyy-mm-dd") & "' and pr.receipt_date<='" & Format(DTo, "yyyy-mm-dd") & "'  " & vbCrLf & _
                        "   and (pr.receipt_cls ='R' or pr.receipt_cls ='R1')  " & vbCrLf
                        
            If Trim(CboLocationCD.Text) <> strAll Then
                sql = sql + "   and pr.supplier_code='" & Trim(CboLocationCD.Text) & "'" & vbCrLf
            End If
            

            
            sql = sql + " order by pr.Supplier_Code,pr.BC40_No,pr.suratjalan_no "
    
        If rsCek.State <> adStateClosed Then rsCek.Close
        'rsCek.CursorLocation = adUseClient
        rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic
        
        If rsCek.EOF Then
            LblErrMsg.Caption = DisplayMsg(4006)
        Else
                
            With xlapp
            
                LblErrMsg.Caption = "[1719] Please Wait While Export To Excel....!"
                
                sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax  From company_profile "
                If rsCompany.State <> adStateClosed Then rsCompany.Close
                rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
                If rsCompany.EOF Then MousePointer = vbDefault: Exit Sub
                .Workbooks.Add
                
                .Range("a1") = rsCompany!company_name
                .Range("a2") = "INCOMING MATERIAL FOR"
                .Range("a3") = "PERIOD " & Format(DFrom, "dd MMM yyyy") & " To " & Format(DTo, "dd MMM yyyy")
                .Range("a4") = " "
                
                .Range("a5") = "Product Code"
                .Range("b5").columnWidth = 12
                
                .Range("b5") = "Description"
                .Range("b5").columnWidth = 40
                
                .Range("c5") = "Surat Jalan No."
                .Range("c5").columnWidth = 12
                .Range("c5").horizontalAlignment = xlCenter
                
                .Range("d5") = "Delivery Date"
                .Range("d5").columnWidth = 12
                .Range("d5").horizontalAlignment = xlCenter
                
                .Range("e5") = "Invoice Date"
                .Range("e5").columnWidth = 12
                .Range("e5").horizontalAlignment = xlCenter
                
                .Range("f5") = "Invoice No"
                .Range("f5").columnWidth = 14
                
                .Range("g5") = "Qty"
                .Range("g5").columnWidth = 7
                .Range("g5").horizontalAlignment = xlCenter
                
                .Range("h5") = "Unit"
                .Range("h5").columnWidth = 5
                
                .Range("i5") = "Original"
                .Range("i5").horizontalAlignment = xlCenter
                .Range("i5").columnWidth = 7
                .Range("i6") = "Curr"
                .Range("i6").horizontalAlignment = xlCenter
                
'Calips Upgrade 2013-01-2013 ===========================================

                .Range("j5") = "Unit"
                .Range("j5").horizontalAlignment = xlCenter
                .Range("j5").columnWidth = 15
                .Range("j6") = "Price"
                .Range("j6").horizontalAlignment = xlCenter
                
'=======================================================================
                
                .Range("k5") = "Original"
                .Range("k5").horizontalAlignment = xlCenter
                .Range("k5").columnWidth = 15
                .Range("k6") = "Amount"
                .Range("k6").horizontalAlignment = xlCenter
                
                .Range("l5", "m5").Merge
                .Range("l5") = "Amount"
                .Range("l5").horizontalAlignment = xlCenter
                .Range("l6", "m6").Merge
                .Range("l6") = "USD"
                .Range("l6").horizontalAlignment = xlCenter
                .Range("l5").columnWidth = 3
                .Range("m5").columnWidth = 16
                
'Calip Upgrade 31-01-2013==============================================

                .Range("n5") = "Type BC"
                .Range("n5").columnWidth = 18
                
'======================================================================
                
                .Range("o5") = "NO BC"
                .Range("o5").columnWidth = 18
                
                .Range("p5") = "TGL BC"
                .Range("p5").columnWidth = 18
                
                .Range("q5") = "PO NO"
                .Range("q5").columnWidth = 25
                
                .Range("a5", "a6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("b5", "b6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("c5", "c6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("d5", "d6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("e5", "e6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("f5", "f6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("g5", "g6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("h5", "h6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("i5", "i6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("j5", "j6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("l5", "l6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("m5", "m6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("n5", "n6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("o5", "o6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("p5", "p6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("q5", "q6").Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("q5", "q6").Borders(xlEdgeRight).LineStyle = xlContinuous
                
                .Range("a" & 6, "q" & 6).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range("a" & 5, "q" & 5).Borders(xlEdgeTop).LineStyle = xlContinuous
                            
                .Range("a5", "q5").Interior.ColorIndex = 15
                .Range("a6", "q6").Interior.ColorIndex = 15
                                   
                Idx = 7
                TotalBC_Ori = 0
                TotalSupplier_Ori = 0
                TotalMaterial_Ori = 0
                GrandTotal_Ori = 0
                
                TotalBC_USD = 0
                TotalSupplier_USD = 0
                TotalMaterial_USD = 0
                GrandTotal_USD = 0
                
                tempbc = "0"
                tempmatcls = "0"
                tempmatname = ""
                tempsupp = "0"
                tempsuppname = ""
                            
                Do While Not rsCek.EOF
                    'View Total Per BC40 No

                    
                    If tempbc <> "0" And Trim(tempbc) <> Trim(rsCek!BC40_No) Then
                        .Range("a" & Idx).horizontalAlignment = xlLeft
                        .Range("a" & Idx) = ""
                        '.Range("a" & Idx, "f" & Idx).Merge
                        .Range("k" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Interior.ColorIndex = 37


                        .Range("l" & Idx) = "$ "
                        .Range("l" & Idx).horizontalAlignment = xlLeft
                        .Range("l" & Idx).Font.Bold = True
                        .Range("m" & Idx) = Format(TotalBC_USD, gs_formatPrice)

                        .Range("m" & Idx).Select
                        .Selection.NumberFormat = gs_formatPrice

                        .Range("m" & Idx).horizontalAlignment = xlRight
                        .Range("m" & Idx).Font.Bold = True
                        TotalBC_Ori = 0
                        TotalBC_USD = 0
                        tempbc = "0"
                        Idx = Idx + 1
                    End If
                    
   
                    
                    'Header & Total Per Supplier Code
                    If (tempsupp = "0" Or Trim(tempsupp) <> Trim(rsCek!Supplier_Code)) Then  'And (tempmatcls = "0" Or Trim(tempmatcls) <> Trim(rsCek!material_cls))
                        '.Visible = True
                        If tempsupp <> "0" Then
                            .Range("a" & Idx).horizontalAlignment = xlLeft
                            .Range("a" & Idx) = "Total " + tempsuppname
                            .Range("a" & Idx).Font.Bold = True
                            '.Range("a" & Idx, "f" & Idx).Merge
                            .Range("a" & Idx).Interior.ColorIndex = 45
                            .Range("j" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                            .Range("a" & Idx, "q" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                            .Range("a" & Idx, "q" & Idx).Borders(xlEdgeRight).LineStyle = xlContinuous
                            .Range("a" & Idx, "q" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
                            .Range("a" & Idx, "q" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                            .Range("a" & Idx, "q" & Idx).Interior.ColorIndex = 38

                            .Range("l" & Idx) = "$ "
                            .Range("l" & Idx).horizontalAlignment = xlLeft
                            .Range("l" & Idx).Font.Bold = True
    
                            .Range("m" & Idx) = TotalSupplier_USD
                            
                            .Range("m" & Idx).Select
                            .Selection.NumberFormat = gs_formatPrice
                            
                            .Range("m" & Idx).horizontalAlignment = xlRight
                            .Range("m" & Idx).Font.Bold = True
                            TotalSupplier_Ori = 0
                            TotalSupplier_USD = 0
                            tempsupp = "0"
                            Idx = Idx + 1
                        End If
                        
                        .Range("a" & Idx).horizontalAlignment = xlLeft
                        .Range("a" & Idx) = "Supplier : " & Trim(rsCek!Supplier_Code) & " - " & Trim(rsCek!trade_name)
                        '.Range("a" & Idx).Font.Bold = True
                        '.Range("a" & Idx, "n" & Idx).Merge
                        .Range("a" & Idx, "q" & Idx).Interior.ColorIndex = 45
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                        tempsuppname = "" & Trim(rsCek!Supplier_Code) & " - " & Trim(rsCek!trade_name)
                        Idx = Idx + 1
                    End If
                   
     
                    'Content
    
                    .Range("a" & Idx).horizontalAlignment = xlLeft
                    .Range("A:A").NumberFormat = "@"
                    .Range("a" & Idx) = Trim(rsCek!Item_Code)
                    .Range("a" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    
                    
                    .Range("b" & Idx) = Trim(rsCek!PartName)
                    .Range("b" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Range("b" & Idx).horizontalAlignment = xlLeft
                    
                    .Range("c" & Idx) = "'" & IIf(IsNull(rsCek!SuratJalan_No), "", Trim(rsCek!SuratJalan_No))
                    .Range("c" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Range("c" & Idx).horizontalAlignment = xlCenter
                    
                    .Range("d" & Idx) = Format(rsCek!DeliveryDate, "DD/MM/YYYY")
                    .Range("d" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Range("d" & Idx).horizontalAlignment = xlCenter
                    
                    .Range("e" & Idx) = Format(rsCek!Invoice_Date, "DD/MM/YYYY")
                    .Range("e" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Range("e" & Idx).horizontalAlignment = xlCenter
                    
                    .Range("f" & Idx) = Trim(rsCek!Invoice_No)
                    .Range("f" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    
                    .Range("g" & Idx) = Trim(rsCek!Qty)
                    .Range("g" & Idx).horizontalAlignment = xlCenter
                    .Range("g" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    
                    .Range("h" & Idx) = Trim(rsCek!unitdesc)
                    .Range("h" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    
                    .Range("i" & Idx) = Trim(rsCek!CurrDesc)
                    .Range("i" & Idx).horizontalAlignment = xlCenter
                    .Range("i" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    
                    .Range("j" & Idx) = Trim(rsCek!Price)
                    .Range("j" & Idx).horizontalAlignment = xlCenter
                    .Range("j" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    
                    .Range("K" & Idx) = rsCek!Amount
                    
                    If Trim(rsCek!CurrDesc) = "IDR" Then
                        .Range("K" & Idx).Select
                        .Selection.NumberFormat = gs_formatPriceIDR
                    Else
                        .Range("K" & Idx).Select
                        .Selection.NumberFormat = gs_formatPrice
                    End If
                    
                    .Range("K" & Idx).horizontalAlignment = xlRight
                    .Range("K" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    
                    .Range("l" & Idx) = "$"
                    .Range("l" & Idx).horizontalAlignment = xlLeft
                    .Range("l" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Range("l" & Idx, "l" & Idx).Interior.ColorIndex = 2
                             
                    .Range("m" & Idx) = rsCek!AmountConvertion
                    
                    .Range("m" & Idx).Select
                    .Selection.NumberFormat = gs_formatPrice
                    
                    .Range("m" & Idx).horizontalAlignment = xlRight
                    '.Range("j" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    
                    '.Range("m" & Idx, "m" & Idx).NumberFormat = "@"
                    .Range("n" & Idx) = IIf(IsNull(rsCek!BC_Type), "", Trim(rsCek!BC_Type))
                    .Range("n" & Idx).horizontalAlignment = xlLeft
                    .Range("n" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Range("n" & Idx).Borders(xlEdgeRight).LineStyle = xlContinuous
                    
                    'If Not IsNull(rsCek!po_date) Then
                        .Range("o" & Idx, "o" & Idx).NumberFormat = "@"
                        .Range("o" & Idx) = IIf(IsNull(rsCek!BC40_No), "", Trim(rsCek!BC40_No))
                    'End If
                    .Range("o" & Idx).horizontalAlignment = xlLeft
                    .Range("o" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Range("o" & Idx).Borders(xlEdgeRight).LineStyle = xlContinuous
                    
                                       
                    'If Not IsNull(rsCek!po_date) Then
                        .Range("p" & Idx, "p" & Idx).NumberFormat = "@"
                        .Range("p" & Idx) = IIf(IsNull(rsCek!BC40_No), "", Trim(rsCek!BC40_Date))
                    'End If
                    .Range("p" & Idx).horizontalAlignment = xlLeft
                    .Range("p" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Range("p" & Idx).Borders(xlEdgeRight).LineStyle = xlContinuous
                    
                    If Not IsNull(rsCek!po_date) Then
                        .Range("q" & Idx) = "'" & Trim(rsCek!po_no) & "/" & Format(rsCek!po_date, "DD.MM.YYYY")
                    End If
                    .Range("q" & Idx).horizontalAlignment = xlLeft
                    .Range("q" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Range("q" & Idx).Borders(xlEdgeRight).LineStyle = xlContinuous
                    
                    '.Range("a" & Idx, "j" & Idx).Columns.AutoFit
                    
                    .Range("a" & Idx, "q" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Range("a" & Idx, "q" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    
                    'temporary total
                    'tempbc = Trim(rsCek!bc40_no)
                    tempmatcls = Trim(IIf(IsNull(rsCek!Material_Cls), "0", rsCek!Material_Cls))
                    tempsupp = Trim(rsCek!Supplier_Code)
                                 
                    'total per BC No
                    TotalBC_Ori = TotalBC_Ori + rsCek!Amount
                    TotalBC_USD = TotalBC_USD + IIf(IsNull(rsCek!AmountConvertion), 0, rsCek!AmountConvertion)
                                    
                    'Total Per Supplier
                    TotalSupplier_Ori = TotalSupplier_Ori + rsCek!Amount
                    TotalSupplier_USD = TotalSupplier_USD + IIf(IsNull(rsCek!AmountConvertion), 0, rsCek!AmountConvertion)
                    
                    'Total per Material Cls
                    TotalMaterial_Ori = TotalMaterial_Ori + rsCek!Amount
                    TotalMaterial_USD = TotalMaterial_USD + IIf(IsNull(rsCek!AmountConvertion), 0, rsCek!AmountConvertion)
                    
                    'Grand TotaL
                    GrandTotal_Ori = GrandTotal_Ori + rsCek!Amount
                    GrandTotal_USD = GrandTotal_USD + IIf(IsNull(rsCek!AmountConvertion), 0, rsCek!AmountConvertion)
                                   
                    rsCek.MoveNext
                    Idx = Idx + 1
                    If rsCek.EOF Then
                        'total BC40 Last
                        .Range("a" & Idx).horizontalAlignment = xlLeft
                        .Range("a" & Idx) = ""
                        '.Range("a" & Idx, "f" & Idx).Merge
                        .Range("k" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                        '.Range("a" & Idx, "m" & Idx).Interior.ColorIndex = 37
    
    '                    .Range("h" & Idx) = Format(TotalBC_Ori, gs_formatPrice)
    '                    .Range("h" & Idx).HorizontalAlignment = xlRight
                        
                        '.Range("i" & Idx) = "$ "
                        '.Range("i" & Idx).HorizontalAlignment = xlLeft
                        '.Range("i" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        '.Range("i" & Idx).Font.Bold = True
                      
                        '.Range("j" & Idx) = TotalBC_USD
                        
                        '.Range("J" & Idx).Select
                        '.Selection.NumberFormat = gs_formatPrice
                        
                        '.Range("j" & Idx).HorizontalAlignment = xlRight
                        '.Range("j" & Idx).Font.Bold = True
                        TotalBC_Ori = 0
                        TotalBC_USD = 0
                        tempbc = "0"
                        'Total Supplier Last
                        'Idx = Idx + 1
                        .Range("a" & Idx).horizontalAlignment = xlLeft
                        .Range("a" & Idx) = "Total " + tempsuppname
                        .Range("a" & Idx).Font.Bold = True
                        '.Range("a" & Idx, "f" & Idx).Merge
                        .Range("k" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Range("a" & Idx, "q" & Idx).Interior.ColorIndex = 38
    

                        
                        .Range("l" & Idx) = "$ "
                        .Range("l" & Idx).horizontalAlignment = xlLeft
                        .Range("l" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Range("l" & Idx).Font.Bold = True
      
                        .Range("m" & Idx) = TotalSupplier_USD
                        
                        .Range("m" & Idx).Select
                        .Selection.NumberFormat = gs_formatPrice
                        
                        .Range("m" & Idx).horizontalAlignment = xlRight
                        .Range("m" & Idx).Font.Bold = True
                        TotalSupplier_Ori = 0
                        TotalSupplier_USD = 0
                        tempsupp = "0"
                        'Total Material Cls Last
                        Idx = Idx + 1
                       
                        TotalMaterial_USD = 0
                        tempmatcls = "0"
                        'Idx = Idx + 1
                    End If
                    
                Loop
                'Grand Total
                '.Range("a" & Idx, "f" & Idx).Merge
                .Range("a" & Idx) = "Grand Total"
                .Range("a" & Idx).Font.Bold = True
                .Range("a" & Idx, "q" & Idx).Interior.ColorIndex = 33
                .Range("a" & Idx).horizontalAlignment = xlLeft
                    

                    
                '.Range("i" & Idx).CurrentRegion.FormatConditions
                .Range("l" & Idx) = "$ "
                .Range("l" & Idx).horizontalAlignment = xlLeft
                .Range("l" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                '.Range("i" & Idx).Font.Bold = True
                
                .Range("m" & Idx) = GrandTotal_USD
                
                .Range("m" & Idx).Select
                .Selection.NumberFormat = gs_formatPrice
                
                .Range("m" & Idx).horizontalAlignment = xlRight
                '.Range("j" & Idx).Font.Bold = True
                
                '.Range("i" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("a" & Idx, "q" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("a" & Idx, "q" & Idx).Borders(xlEdgeRight).LineStyle = xlContinuous
                .Range("a" & Idx, "q" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Range("a" & Idx, "q" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                
                .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
                .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.04)
                .ActiveSheet.PageSetup.Orientation = xlLandscape
                .ActiveSheet.PageSetup.PaperSize = xlPaperA4
                .ActiveSheet.PageSetup.PrintArea = "a1:o" & Idx
                .ActiveSheet.PageSetup.Orientation = 2
                .WindowState = xlMaximized
                
                .Visible = True
                LblErrMsg.Caption = ""
            End With
        End If
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
'If CDate(Dmonth2) < CDate(DMonth) Then
'      LblErrMsg.Caption = DisplayMsg(4068)
'      Exit Sub
'   Else
'      LblErrMsg.Caption = ""
'   End If
End Sub

Private Sub DMonth_Click()
'If CDate(Dmonth2) < CDate(DMonth) Then
'      LblErrMsg.Caption = DisplayMsg(4068)
'      Exit Sub
'   Else
'      LblErrMsg.Caption = ""
'   End If
End Sub

Private Sub Form_Load()
Dim RsW As New Recordset
Dim ir As Integer
If gb_Simulation = True Then Call up_InitSimulation(Me)
LblLocationName = ""
LblErrMsg = ""
ChkInclude.Value = 0
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
bteHakPrice = hakPrice(Me.Name)
Call StockLocation
DFrom = Format(Date, "dd MMM yyyy")
DTo = Format(Date, "dd MMM yyyy")

'##Tampilkan Material Cls
sql = " Select rtrim(Material_Cls) Material_Cls, Description From Material_Cls Order By Material_Cls "
'sql = "Select rtrim(wh_code) as WC,wh_name as WN from warehouse_master order by wh_code"
RsW.Open sql, Db, adOpenKeyset, adLockOptimistic
CboMaterialCls.clear
CboMaterialCls.columnCount = 2
CboMaterialCls.TextColumn = 1
CboMaterialCls.AddItem ""
CboMaterialCls.List(0, 0) = strAll
CboMaterialCls.List(0, 1) = strAll
ir = 1
While Not RsW.EOF
    CboMaterialCls.AddItem ""
    CboMaterialCls.List(ir, 0) = RsW!Material_Cls
    CboMaterialCls.List(ir, 1) = Trim$(RsW!Description)
    ir = ir + 1
    RsW.MoveNext
Wend
CboMaterialCls.ColumnWidths = "60 pt; 180 pt"
CboMaterialCls.ListWidth = 240
CboMaterialCls.ListRows = 15
CboMaterialCls.ListIndex = 0
End Sub

Private Sub StockLocation()
Dim sql As String, RsStock As New ADODB.Recordset
Dim i As Long

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
CboLocationCD.ColumnWidths = "50 pt; 250 pt"
CboLocationCD.ListWidth = 300
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


Private Sub OptDetail_Click()
If OptMaterial.Value = True Then
    Label2.Visible = True
    CboMaterialCls.Visible = True
    lblWHName.Visible = True
    Line2.Visible = True
Else
    Label2.Visible = False
    CboMaterialCls.Visible = False
    lblWHName.Visible = False
    Line2.Visible = False
End If
End Sub

Private Sub OptGeneral_Click()
If OptMaterial.Value = True Then
    Label2.Visible = True
    CboMaterialCls.Visible = True
    lblWHName.Visible = True
    Line2.Visible = True
Else
    Label2.Visible = False
    CboMaterialCls.Visible = False
    lblWHName.Visible = False
    Line2.Visible = False
End If
End Sub

Private Sub OptMaterial_Click()
If OptMaterial.Value = True Then
    Label2.Visible = True
    CboMaterialCls.Visible = True
    lblWHName.Visible = True
    Line2.Visible = True
Else
    Label2.Visible = False
    CboMaterialCls.Visible = False
    lblWHName.Visible = False
    Line2.Visible = False
End If
End Sub

Private Sub Upload_Click()
Dim rsupload As New ADODB.Recordset
Dim rsCek As New ADODB.Recordset
Dim a As Double
Dim b As Double

sql = "select * from upload"
If rsupload.State <> adStateClosed Then rsupload.Close
rsupload.Open sql, Db, adOpenKeyset, adLockOptimistic
a = 0
b = 0

Do While Not rsupload.EOF
    sql = "select * from stock_master where warehouse_code='" & Trim(rsupload("warehouse_code")) & "' and item_code='" & Trim(rsupload("item_code")) & "'  "
    
    If rsCek.State <> adStateClosed Then rsCek.Close
    rsCek.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If rsCek.EOF Then
        rsCek.AddNew
        rsCek(0) = Trim(rsupload("warehouse_Code"))
        rsCek(1) = Trim(rsupload("Item_Code"))
        rsCek(2) = 0
        rsCek(3) = 0
        rsCek(4) = 0
        rsCek(5) = 0
        rsCek(6) = 0
        rsCek(7) = 0
        rsCek(8) = 0
        rsCek(9) = 0
        rsCek(10) = 0
        rsCek(11) = 0
        rsCek(12) = 0
        rsCek(13) = rsupload("qty")
        rsCek(14) = 0
        rsCek(15) = 0
        rsCek(16) = 0
        rsCek(17) = 0
        rsCek(18) = 0
        'rscek(19) = Now
        'rscek(20) = Trim(rsupload("warehouse_Code"))
        'rscek(21) = Trim(rsupload("warehouse_Code"))
        'rscek(22) = Trim(rsupload("warehouse_Code"))
        'rscek(23) = Trim(rsupload("warehouse_Code"))
        rsCek(24) = Now
        rsCek(25) = "Initial"
        rsCek(26) = Now
        rsCek.update
        a = a + 1
    Else
    sql = "update stock_master set TM_Inventory='" & rsupload("qty") & "' where warehouse_code='" & Trim(rsupload("warehouse_code")) & "' and item_code='" & Trim(rsupload("item_code")) & "'  "
    Db.Execute sql
    b = b + 1
    
    
    End If
    
    rsupload.MoveNext
Loop

LblErrMsg.Caption = "data Inserted " & a & " Record, Updated " & b




End Sub
