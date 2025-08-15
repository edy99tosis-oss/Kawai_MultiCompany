VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmForecast 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forecast Part/Material"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "FrmForecast.frx":0000
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   315
      Left            =   300
      TabIndex        =   16
      Top             =   2700
      Visible         =   0   'False
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1335
      Left            =   300
      TabIndex        =   11
      Top             =   1065
      Width           =   8145
      Begin VB.TextBox lblcust 
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
         Height          =   210
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   4860
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   780
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   293470211
         UpDown          =   -1  'True
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   315
         Left            =   2910
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   780
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   293535747
         UpDown          =   -1  'True
         CurrentDate     =   37798
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   2745
         X2              =   7740
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   540
      End
      Begin VB.Label LblCode 
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
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   330
         Width           =   900
      End
      Begin MSForms.ComboBox cbocust 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   300
         Width           =   1710
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3016;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
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
         Index           =   1
         Left            =   2580
         TabIndex        =   12
         Top             =   840
         Width           =   165
      End
   End
   Begin VB.OptionButton OptForecast 
      BackColor       =   &H00FDDFE3&
      Caption         =   "&Material"
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
      Index           =   1
      Left            =   1290
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1245
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.OptionButton OptForecast 
      BackColor       =   &H00FDDFE3&
      Caption         =   "&Part"
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
      Index           =   0
      Left            =   300
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1245
      Visible         =   0   'False
      Width           =   1005
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
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3705
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   300
      TabIndex        =   9
      Top             =   3045
      Width           =   8100
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
         Left            =   60
         TabIndex        =   10
         Top             =   210
         Width           =   7875
      End
   End
   Begin VB.CommandButton CmdPreview 
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
      Left            =   7245
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3705
      Width           =   1140
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   6660
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   270
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   17
      Top             =   2460
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Forecast Part/Material"
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
      Height          =   375
      Left            =   345
      TabIndex        =   8
      Top             =   270
      Width           =   8055
   End
End
Attribute VB_Name = "frmForecast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HakU As Integer
Dim StrPeriode As String
Dim tgl_sb As String * 2
Private Sub CboCust_Change()
cboCust_Click
End Sub

Private Sub cboCust_Click()
lblcust = ""
If cboCust.ListIndex <= 0 Then cboCust.ListIndex = 0
If cboCust.MatchFound Then lblcust.Text = cboCust.List(cboCust.ListIndex, 1)
End Sub

Private Sub CmdPreview_Click()
Dim SqlRpt As String
Dim rsRpt As New ADODB.Recordset

Dim X As Integer

Dim bln As Integer, thn As Integer, CName As String
Dim VClosing As String, strSQL As String
Dim bulan(6) As Integer, tahun(6) As String

Dim rsCP As New ADODB.Recordset

Set rsCP = Db.Execute("Select Company_name from company_profile")
CName = ""

If Not rsCP.EOF Then CName = Trim$(rsCP!company_name)
If Trim$(cboCust) = "" Then LblErrMsg = DisplayMsg(1054): Me.MousePointer = vbDefault: Exit Sub
cboCust = cboCust
LblErrMsg = ""

' Cek Begin Periode
LblErrMsg = up_ValidateDateRange(Tgl1.Value, False)
If Trim(LblErrMsg) <> "" Then Exit Sub
LblErrMsg = ""
'---

If cboCust.MatchFound = False Then LblErrMsg = DisplayMsg(1054): Me.MousePointer = vbDefault: Exit Sub
    
If OptForecast(0).Value Then  'Part

        'SqlRpt = "select master.itm, master.des, b.hasil Bln1, c.hasil bln2, d.hasil bln3, e.hasil bln4, f.hasil bln5, g.hasil bln6 from " & vbCrLf & _
        '       "(select itm,rtrim(item_name)des from(Select item_code itm from price_master where trade_code='" & Trim$(CboCust) & "' Union Select item_code itm from item_master where supplier_code='" & Trim$(CboCust) & "') c,item_master IM where c.itm=IM.Item_code and IM.sheetcoil_cls is null ) master, " & vbCrLf & _
        '       " (Select R.childitem_code itm ,Childrequirement_qty Qty,ChildrequirementResult_qty Result,Childrequirement_qty-ChildrequirementResult_qty  Hasil from Requirement_master R" & vbCrLf & _
        '       "   where R.childitem_code  in " & vbCrLf & _
        '       "   (Select item_code from price_master where trade_code='" & Trim$(CboCust) & "' Union Select item_code from item_master where supplier_code='" & Trim$(CboCust) & "') and ChildRequirement_year='" & year(DateAdd("m", 0, Format(Tgl1, "YYYY-mm-DD"))) & "' and ChildRequirement_month ='" & month(DateAdd("m", 0, Format(Tgl1, "YYYY-mm-DD"))) & "') b, " & vbCrLf & _
        '       " (Select R.childitem_code itm ,Childrequirement_qty Qty,ChildrequirementResult_qty Result,Childrequirement_qty-ChildrequirementResult_qty  Hasil from Requirement_master R " & vbCrLf & _
        '       "  where R.childitem_code  in " & vbCrLf & _
        '       "   (Select item_code from price_master where trade_code='" & Trim$(CboCust) & "' Union Select item_code from item_master where supplier_code='" & Trim$(CboCust) & "') and ChildRequirement_year='" & year(DateAdd("m", 1, Format(Tgl1, "YYYY-mm-DD"))) & "' and ChildRequirement_month ='" & month(DateAdd("m", 1, Format(Tgl1, "YYYY-mm-DD"))) & "') c , " & vbCrLf & _
        '       " (Select R.childitem_code itm ,Childrequirement_qty Qty,ChildrequirementResult_qty Result,Childrequirement_qty-ChildrequirementResult_qty  Hasil  from Requirement_master R " & vbCrLf & _
        '       "  where R.childitem_code  in " & vbCrLf & _
        '       "  (Select item_code from price_master where trade_code='" & Trim$(CboCust) & "' Union Select item_code from item_master where supplier_code='" & Trim$(CboCust) & "') and ChildRequirement_year='" & year(DateAdd("m", 2, Format(Tgl1, "YYYY-mm-DD"))) & "' and ChildRequirement_month ='" & month(DateAdd("m", 2, Format(Tgl1, "YYYY-mm-DD"))) & "') d, " & vbCrLf & _
        '       " (Select R.childitem_code itm ,Childrequirement_qty Qty,ChildrequirementResult_qty Result,Childrequirement_qty-ChildrequirementResult_qty  Hasil from Requirement_master R " & vbCrLf & _
        '       "  Where R.childitem_code  in " & vbCrLf & _
        '       "  (Select item_code from price_master where trade_code='" & Trim$(CboCust) & "' Union Select item_code from item_master where supplier_code='" & Trim$(CboCust) & "') and ChildRequirement_year='" & year(DateAdd("m", 3, Format(Tgl1, "YYYY-mm-DD"))) & "' and ChildRequirement_month ='" & month(DateAdd("m", 3, Format(Tgl1, "YYYY-mm-DD"))) & "') e, " & vbCrLf & _
        '       " (Select R.childitem_code itm ,Childrequirement_qty Qty,ChildrequirementResult_qty Result,Childrequirement_qty-ChildrequirementResult_qty  Hasil from Requirement_master R " & vbCrLf & _
        '       "  where R.childitem_code  in " & vbCrLf & _
        '       "  (Select item_code from price_master where trade_code='" & Trim$(CboCust) & "' Union Select item_code from item_master where supplier_code='" & Trim$(CboCust) & "') and ChildRequirement_year='" & year(DateAdd("m", 4, Format(Tgl1, "YYYY-mm-DD"))) & "' and ChildRequirement_month ='" & month(DateAdd("m", 4, Format(Tgl1, "YYYY-mm-DD"))) & "') f, " & vbCrLf & _
        '       " (Select R.childitem_code itm ,Childrequirement_qty Qty,ChildrequirementResult_qty Result,Childrequirement_qty-ChildrequirementResult_qty  Hasil from Requirement_master R " & vbCrLf & _
        '       "  where R.childitem_code  in " & vbCrLf & _
        '       " (Select item_code from price_master where trade_code='" & Trim$(CboCust) & "' Union  Select item_code from item_master where supplier_code='" & Trim$(CboCust) & "') and ChildRequirement_year='" & year(DateAdd("m", 5, Format(Tgl1, "YYYY-mm-DD"))) & "' and ChildRequirement_month ='" & month(DateAdd("m", 5, Format(Tgl1, "YYYY-mm-DD"))) & "') g " & vbCrLf & _
        '       "  where master.itm *= b.itm " & vbCrLf & _
        '       "  and master.itm *= c.itm and master.itm *= d.itm " & vbCrLf & _
        '       "  and master.itm *= e.itm " & vbCrLf & _
        '       "  and master.itm *= f.itm and master.itm *= g.itm "
    
'---- Forecast for KAWAI with Stock Calculation

        Dim rscount As New ADODB.Recordset
        Dim SqlCount As Long
        Dim StrCount As String
        
        If cboCust = strAll Then
            StrCount = "Select Count(Item_Code) From Item_Master A Where Supplier_Code<>'999' "
        Else
            StrCount = "Select Count(Item_Code) From Item_Master A Where Supplier_Code='" & Trim(cboCust) & "'"
        End If

        Set rscount = Db.Execute(StrCount)
        SqlCount = rscount(0)
        
        ' --------------------
        Select Case up_GetDateRange(Tgl1.Value)
            Case Is = 0
                VClosing = "LM"
            Case Is = 1
                VClosing = "TM"
            Case Is = 2
                VClosing = "NM"
        End Select

        Me.MousePointer = vbHourglass
        
        strSQL = " Select A.Item_Code,A.Item_Name,  " & vbCrLf & _
                          "     isnull((Select top 1 RTrim(Description) From Price_Master Inner Join Curr_Cls On Price_Master.Currency_Code=Curr_Cls.Curr_Cls " & vbCrLf & _
                          "     Where Trade_Code=A.Supplier_Code And Item_Code=A.Item_Code) ,'') Currency, " & vbCrLf & _
                          "     isnull((Select top 1 Price From Price_Master Where Trade_Code=A.Supplier_Code And Item_Code=A.Item_Code) ,0) Price, " & vbCrLf & _
                          "     isnull((Select LM_PreMonth From Stock_Master Where Warehouse_Code=A.WH_Code And Item_Code=a.Item_Code),0) BegStock, " & vbCrLf & _
                          "     isnull((Select LM_Inventory From Stock_Master Where Warehouse_Code=A.WH_Code And Item_Code=a.Item_Code),0) InventStock, " & vbCrLf & _
                          "     isnull((Select " & VClosing & "_Current From Stock_Master Where Warehouse_Code=A.WH_Code And Item_Code=a.Item_Code),0) EndStock, " & vbCrLf & _
                          "     isnull((Select " & VClosing & "_Current From Stock_Master Where Warehouse_Code='WH-003' And Item_Code=a.Item_Code),0) NGProcess, " & vbCrLf & _
                          "     isnull((Select " & VClosing & "_Current From Stock_Master Where Warehouse_Code='WH-004' And Item_Code=a.Item_Code),0) NGVendor, " & vbCrLf & _
                          "     isnull((Select " & VClosing & "_Current From Stock_Master Where Warehouse_Code='WH-007' And Item_Code=a.Item_Code),0) NGInSupplier, " & vbCrLf
                
        strSQL = strSQL + " (Select isnull(sum(TotalQty-TotalReceipt),0) BO  " & vbCrLf & _
                          " From ( " & vbCrLf & _
                          "     Select Month(Delivery_Date) Bln, Year(Delivery_Date) Thn, Item_Code,Sum(Qty) TotalQty,  " & vbCrLf & _
                          "         isnull((Select Sum(Qty) From Part_Receipt Where Item_Code=P.Item_Code And Po_No=P.Po_no  " & vbCrLf & _
                          "             --And Year(receipt_Date)=Year(Delivery_Date) And Month(Receipt_Date)=Month(Delivery_Date) " & vbCrLf & _
                          "             Group By PO_No,Item_Code),0) TotalReceipt " & vbCrLf & _
                          "     From PurchaseOrder_Detail P " & vbCrLf & _
                          "         Group By month(Delivery_Date), Year(Delivery_Date), PO_No,Item_Code " & vbCrLf & _
                          " ) PO Where Item_Code=A.Item_Code And Bln='" & Month(Tgl1) & "' and thn='" & Year(Tgl1) & "') BO, " & vbCrLf
                
        strSQL = strSQL + " (Select isnull(sum(TotalReceipt),0) QtyReceipt  " & vbCrLf & _
                          " From ( " & vbCrLf & _
                          "     Select Month(Delivery_Date) Bln, Year(Delivery_Date) Thn, Item_Code,Sum(Qty) TotalQty,  " & vbCrLf & _
                          "         isnull((Select Sum(Qty) From Part_Receipt Where Item_Code=P.Item_Code And Po_No=P.Po_no  " & vbCrLf & _
                          "             --And Year(receipt_Date)=Year(Delivery_Date) And Month(Receipt_Date)=Month(Delivery_Date) " & vbCrLf & _
                          "             Group By PO_No,Item_Code),0) TotalReceipt " & vbCrLf & _
                          "     From PurchaseOrder_Detail P " & vbCrLf & _
                          "         Group By month(Delivery_Date), Year(Delivery_Date), PO_No,Item_Code " & vbCrLf & _
                          " ) PO Where Item_Code=A.Item_Code And Bln='" & Month(Tgl1) & "' and thn='" & Year(Tgl1) & "') QtyReceipt, " & vbCrLf
        
        
        ' Po dan Requirement sebelumnya berdasarkan periode Closing
        '-------------------------------------------------------------------------
        
        If VClosing = "LM" Then
            strSQL = strSQL & "   0 Reqm2, 0 PoM2, 0 Reqm1, 0 PoM1,  " & vbCrLf
            
        ElseIf VClosing = "TM" Then
            strSQL = strSQL & "      0 Reqm2, 0 PoM2," & vbCrLf & _
              "     isnull((Select sum(Childrequirement_qty-ChildrequirementResult_qty) from Requirement_master where childitem_code=a.Item_Code  " & vbCrLf & _
              "         and ChildRequirement_year='" & Year(DateAdd("M", -1, Tgl1.Value)) & "' and ChildRequirement_month ='" & Month(DateAdd("M", -1, Tgl1.Value)) & "'),0) Reqm1, " & vbCrLf & _
              "     isnull((Select sum(Qty) from PurchaseOrder_Detail PD where item_code=a.Item_Code   " & vbCrLf & _
              "         and Year(PD.delivery_Date)='" & Year(DateAdd("M", -1, Tgl1.Value)) & "' and Month(PD.delivery_Date) ='" & Month(DateAdd("M", -1, Tgl1.Value)) & "'),0) Pom1, " & vbCrLf
              
        Else
            strSQL = strSQL & "     isnull((Select sum(Childrequirement_qty-ChildrequirementResult_qty) from Requirement_master where childitem_code=a.Item_Code  " & vbCrLf & _
              "         and ChildRequirement_year='" & Year(DateAdd("M", -2, Tgl1.Value)) & "' and ChildRequirement_month ='" & Month(DateAdd("M", -2, Tgl1.Value)) & "'),0) Reqm2, " & vbCrLf & _
              "     isnull((Select sum(Qty) from PurchaseOrder_Detail PD where item_code=a.Item_Code   " & vbCrLf & _
              "         and Year(PD.delivery_Date)='" & Year(DateAdd("M", -2, Tgl1.Value)) & "' and Month(PD.delivery_Date) ='" & Month(DateAdd("M", -2, Tgl1.Value)) & "'),0) Pom2, " & vbCrLf & _
              "     isnull((Select sum(Childrequirement_qty-ChildrequirementResult_qty) from Requirement_master where childitem_code=a.Item_Code  " & vbCrLf & _
              "         and ChildRequirement_year='" & Year(DateAdd("M", -1, Tgl1.Value)) & "' and ChildRequirement_month ='" & Month(DateAdd("M", -1, Tgl1.Value)) & "'),0) Reqm1, " & vbCrLf & _
              "     isnull((Select sum(Qty) from PurchaseOrder_Detail PD where item_code=a.Item_Code   " & vbCrLf & _
              "         and Year(PD.delivery_Date)='" & Year(DateAdd("M", -1, Tgl1.Value)) & "' and Month(PD.delivery_Date) ='" & Month(DateAdd("M", -1, Tgl1.Value)) & "'),0) Pom1, " & vbCrLf
        End If
                           
        '-------------------------------------------------------------------------
        
        bln = Month(Tgl1.Value) - 1
        thn = Year(Tgl1.Value)
                
        For X = 0 To 5
            bln = bln + 1
            
            If bln > 12 Then
                bln = 1: thn = thn + 1
            End If
            bulan(X) = bln
            tahun(X) = thn
            
            strSQL = strSQL & "     isnull((Select sum(Childrequirement_qty-ChildrequirementResult_qty) from Requirement_master where childitem_code=a.Item_Code  " & vbCrLf & _
                  "         and ChildRequirement_year='" & thn & "' and ChildRequirement_month ='" & bln & "'),0) Req" & X + 1 & ", " & vbCrLf & _
                  "     isnull((Select sum(Qty) from PurchaseOrder_Detail PD where item_code=a.Item_Code   " & vbCrLf & _
                  "         and Year(PD.delivery_Date)='" & thn & "' and Month(PD.delivery_Date) ='" & bln & "'),0) Po" & X + 1 & ", " & vbCrLf
        Next
        
                'StrSql = StrSql + "     isnull((Select sum(Childrequirement_qty-ChildrequirementResult_qty) from Requirement_master where childitem_code=a.Item_Code  " & vbCrLf & _
                '                  "         and ChildRequirement_year='2009' and ChildRequirement_month ='8'),0) Req2, " & vbCrLf & _
                '                  "     isnull((Select sum(Qty) from PurchaseOrder_Detail PD inner Join PurchaseOrder_Master PM On PD.PO_No=PD.PO_No where supplier_Code=A.Supplier_Code And item_code=a.Item_Code   " & vbCrLf & _
                '                  "         and Year(PD.delivery_Date)='2009' and Month(PD.delivery_Date) ='8'),0) Po2, " & vbCrLf & _
                '                  "     isnull((Select sum(Childrequirement_qty-ChildrequirementResult_qty) from Requirement_master where childitem_code=a.Item_Code  " & vbCrLf & _
                '                  "         and ChildRequirement_year='2009' and ChildRequirement_month ='9'),0) Req3, " & vbCrLf & _
                '                  "     isnull((Select sum(Qty) from PurchaseOrder_Detail PD inner Join PurchaseOrder_Master PM On PD.PO_No=PD.PO_No where supplier_Code=A.Supplier_Code And item_code=a.Item_Code   " & vbCrLf & _
                '                  "         and Year(PD.delivery_Date)='2009' and Month(PD.delivery_Date) ='9'),0) Po3, " & vbCrLf & _
                '                  "     isnull((Select sum(Childrequirement_qty-ChildrequirementResult_qty) from Requirement_master where childitem_code=a.Item_Code  " & vbCrLf & _
                '                  "         and ChildRequirement_year='2009' and ChildRequirement_month ='10'),0) Req4, " & vbCrLf & _
                '                  "     isnull((Select sum(Qty) from PurchaseOrder_Detail PD inner Join PurchaseOrder_Master PM On PD.PO_No=PD.PO_No where supplier_Code=A.Supplier_Code And item_code=a.Item_Code   "
                '
                'StrSql = StrSql + "         and Year(PD.delivery_Date)='2009' and Month(PD.delivery_Date) ='10'),0) Po4, " & vbCrLf & _
                '                  "     isnull((Select sum(Childrequirement_qty-ChildrequirementResult_qty) from Requirement_master where childitem_code=a.Item_Code  " & vbCrLf & _
                '                  "         and ChildRequirement_year='2009' and ChildRequirement_month ='11'),0) Req5, " & vbCrLf & _
                '                  "     isnull((Select sum(Qty) from PurchaseOrder_Detail PD inner Join PurchaseOrder_Master PM On PD.PO_No=PD.PO_No where supplier_Code=A.Supplier_Code And item_code=a.Item_Code   " & vbCrLf & _
                '                  "         and Year(PD.delivery_Date)='2009' and Month(PD.delivery_Date) ='11'),0) Po5, " & vbCrLf & _
                '                  "     isnull((Select sum(Childrequirement_qty-ChildrequirementResult_qty) from Requirement_master where childitem_code=a.Item_Code  " & vbCrLf & _
                '                  "         and ChildRequirement_year='2009' and ChildRequirement_month ='12'),0) Req6, " & vbCrLf & _
                '                  "     isnull((Select sum(Qty) from PurchaseOrder_Detail PD inner Join PurchaseOrder_Master PM On PD.PO_No=PD.PO_No where supplier_Code=A.Supplier_Code And item_code=a.Item_Code   " & vbCrLf & _
                '                  "         and Year(PD.delivery_Date)='2009' and Month(PD.delivery_Date) ='12'),0) Po6 " & vbCrLf & _

    If cboCust = strAll Then
        strSQL = strSQL & " Supplier_Code From Item_Master A Where Supplier_Code<>'999' Order By A.Supplier_Code, A.Item_Code "
    Else
        strSQL = strSQL & " Supplier_Code From Item_Master A Where Supplier_Code='" & Trim(cboCust) & "' Order By A.Item_Code"
    End If
    
    If rsRpt.State <> adStateClosed Then rsRpt.Close
    Set rsRpt = Db.Execute(strSQL)
    
    If rsRpt.EOF Then
        LblErrMsg = DisplayMsg(13)
        Me.MousePointer = vbDefault
        Exit Sub
    End If

    '--------------
    'Print To Excel
    
    Pb1.Visible = True
    Label1.Visible = True
    
    Pb1.Max = SqlCount
    
    Dim xlapp As New Excel.application
    Dim Baris As Integer
    Dim Idx As Long
    
    With xlapp
        .Workbooks.Add
        '.Visible = True
       
        .Range("A2") = "Material Forecast Plan Report"
        .Range("A2").VerticalAlignment = xlCenter
        .Range("A2").Columns.Font.Name = "Arial"
        .Range("A2").Columns.Font.Size = 10
        .Range("A2").Columns.Font.Bold = True
    
        .Range("A3") = "Supplier : " & Trim(cboCust) & " / " & Trim(lblcust)
        .Range("A3").VerticalAlignment = xlCenter
        .Range("A3").Columns.Font.Name = "Arial"
        .Range("A3").Columns.Font.Size = 8
        .Range("A3").Columns.Font.Bold = True
    
        .Range("A4") = "Period : " & Format(Tgl1, "MMM YYYY") & " to " & Format(Tgl2, "MMM YYYY")
        .Range("A4").VerticalAlignment = xlCenter
        .Range("A4").Columns.Font.Name = "Arial"
        .Range("A4").Columns.Font.Size = 8
        .Range("A4").Columns.Font.Bold = True
        
        Dim barisawal As Byte, NoKolom As Integer, StrKolom As String, Depan As String
        Dim Y As Integer, ColSebelum As String, NO As Integer
        
        barisawal = 6
        
        .Range("A" & barisawal) = "No"
        .Range("B" & barisawal) = "Item Code"
        .Range("C" & barisawal) = "Description"
        .Range("D" & barisawal) = "Currency"
        .Range("E" & barisawal) = "Unit Price"
        
        .Range("F" & barisawal) = MonthName(bulan(0)) & " " & tahun(0)
        .Range("F" & barisawal + 1) = "Beg Stock"
        .Range("G" & barisawal + 1) = "Req"
        .Range("H" & barisawal + 1) = "Fix Order"
        .Range("I" & barisawal + 1) = "Received"
        .Range("J" & barisawal + 1) = "B O"
        .Range("K" & barisawal + 1) = "End Stock"
        .Range("L" & barisawal + 1) = "NG Process"
        .Range("M" & barisawal + 1) = "NG Vendor"
        .Range("N" & barisawal + 1) = "NG Supplier"
        .Range("O" & barisawal + 1) = "Adj Stock"
        
        .Range("P" & barisawal) = MonthName(bulan(1)) & " " & tahun(1)
        .Range("P" & barisawal + 1) = "Beg Stock"
        .Range("Q" & barisawal + 1) = "Req "
        .Range("R" & barisawal + 1) = "Fix Order"
        .Range("S" & barisawal + 1) = "End Stock"
        
        .Range("T" & barisawal) = MonthName(bulan(2)) & " " & tahun(2)
        .Range("T" & barisawal + 1) = "Beg Stock"
        .Range("U" & barisawal + 1) = "Req "
        .Range("V" & barisawal + 1) = "Fix Order"
        .Range("W" & barisawal + 1) = "End Stock"
        
        .Range("X" & barisawal) = MonthName(bulan(3)) & " " & tahun(3)
        .Range("X" & barisawal + 1) = "Beg Stock"
        .Range("Y" & barisawal + 1) = "Req "
        .Range("Z" & barisawal + 1) = "Fix Order"
        .Range("AA" & barisawal + 1) = "End Stock"
        
        .Range("AB" & barisawal) = MonthName(bulan(4)) & " " & tahun(4)
        .Range("AB" & barisawal + 1) = "Beg Stock"
        .Range("AC" & barisawal + 1) = "Req "
        .Range("AD" & barisawal + 1) = "Fix Order"
        .Range("AE" & barisawal + 1) = "End Stock"
        
        .Range("AF" & barisawal) = MonthName(bulan(5)) & " " & tahun(5)
        .Range("AF" & barisawal + 1) = "Beg Stock"
        .Range("AG" & barisawal + 1) = "Req "
        .Range("AH" & barisawal + 1) = "Fix Order"
        .Range("AI" & barisawal + 1) = "End Stock"
        
        Baris = barisawal + 2
        Idx = 1
        
        Do While Not rsRpt.EOF
            .Range("A" & Baris) = Idx
            .Range("B" & Baris) = rsRpt("Item_Code")
            .Range("C" & Baris) = rsRpt("Item_Name")
            .Range("D" & Baris) = rsRpt("Currency")
            .Range("E" & Baris) = rsRpt("Price")
            
            Pb1.Value = (Idx)
            Label1.Caption = "Transfering data " & Idx & " of " & SqlCount
            
            If .Range("D" & Baris) <> "IDR" Then
                .Range("E" & barisawal + 2, "E" & Baris).NumberFormat = gs_formatAmount
            Else
                .Range("E" & barisawal + 2, "E" & Baris).NumberFormat = gs_formatAmountIDR
            End If
                
            ' Sebelum Perubahan BegStock diambil dari Posisi Closing (LM, TM atau NM)
            '.Range("F" & Baris) = rsRpt("BegStock")
            
            ' Setelah Perubahan BegStock diambil dari LM_PreMonth untuk periode LM  atau LM_Inventory untuk Periode TM & NM
            If VClosing = "LM" Then
                .Range("F" & Baris) = rsRpt("BegStock")                            '+ rsRpt("Pom1") + rsRpt("Pom2") - rsRpt("ReqM1") - rsRpt("ReqM2")
            ElseIf VClosing = "TM" Then
                .Range("F" & Baris) = rsRpt("InventStock")                              '+ rsRpt("Pom1") + rsRpt("Pom2") - rsRpt("ReqM1") - rsRpt("ReqM2")
            Else
                .Range("F" & Baris) = rsRpt("InventStock") + rsRpt("Pom1") - rsRpt("ReqM1")                           '+ rsRpt("Pom1") + rsRpt("Pom2") - rsRpt("ReqM1") - rsRpt("ReqM2")
            End If
            '---------------------------------------------------
            
            .Range("G" & Baris) = rsRpt("Req1")
            .Range("H" & Baris) = rsRpt("PO1")
            .Range("I" & Baris) = rsRpt("QtyReceipt")
            .Range("J" & Baris) = rsRpt("BO")
            
            ' Sebelum Perubahan (End Stock mempertimbangkan actual stock di WH NG)
            '.Range("K" & Baris) = rsRpt("BegStock") - rsRpt("Req1") + rsRpt("QtyReceipt") - rsRpt("NGProcess") - rsRpt("NGVendor") - rsRpt("NGInSupplier")
            
            ' Sesudah Perubahan berdasarkan request 20 April 2011 (Yang diperhitungkan hanya beginstock, Po, dan Requirement) - Yudi
            ' Formula End Stock menjadi BeginStock + Po - Req (tidak memperhitungkan stock actual di warehouse)
            .Range("K" & Baris) = "=F" & Baris & "+H" & Baris & "-G" & Baris                                '.Range("F" & Baris) + rsRpt("PO1") + rsRpt("BO") - rsRpt("Req1")
            '------------------------------------------------------
            
            .Range("L" & Baris) = rsRpt("NGProcess")
            .Range("M" & Baris) = rsRpt("NGVendor")
            .Range("N" & Baris) = rsRpt("NGInSupplier")
            .Range("O" & Baris) = ""
            
            .Range("P" & Baris) = "=K" & Baris
            .Range("Q" & Baris) = rsRpt("Req2")
            .Range("R" & Baris) = rsRpt("PO2")
            '.Range("S" & Baris) = .Range("P" & Baris) - .Range("Q" & Baris) + .Range("R" & Baris)
            .Range("S" & Baris) = "=P" & Baris & "+R" & Baris & "-Q" & Baris
            
            .Range("T" & Baris) = "=S" & Baris
            .Range("U" & Baris) = rsRpt("Req3")
            .Range("V" & Baris) = rsRpt("PO3")
            '.Range("W" & Baris) = .Range("T" & Baris) - .Range("U" & Baris) + .Range("V" & Baris)
            .Range("W" & Baris) = "=T" & Baris & "+V" & Baris & "-U" & Baris
            
            .Range("X" & Baris) = "=W" & Baris
            .Range("Y" & Baris) = rsRpt("Req4")
            .Range("Z" & Baris) = rsRpt("PO4")
            '.Range("AA" & Baris) = .Range("X" & Baris) - .Range("Y" & Baris) + .Range("Z" & Baris)
            .Range("AA" & Baris) = "=X" & Baris & "+Z" & Baris & "-Y" & Baris
            
            .Range("AB" & Baris) = "=AA" & Baris
            .Range("AC" & Baris) = rsRpt("Req5")
            .Range("AD" & Baris) = rsRpt("PO5")
            .Range("AE" & Baris) = .Range("AB" & Baris) - .Range("AC" & Baris) + .Range("AD" & Baris)
            .Range("AE" & Baris) = "=AB" & Baris & "+AD" & Baris & "-AC" & Baris
            
            .Range("AF" & Baris) = "=AE" & Baris
            .Range("AG" & Baris) = rsRpt("Req6")
            .Range("AH" & Baris) = rsRpt("PO6")
            .Range("AI" & Baris) = .Range("AF" & Baris) - .Range("AG" & Baris) + .Range("AH" & Baris)
            .Range("AI" & Baris) = "=AF" & Baris & "+AH" & Baris & "-AG" & Baris
            
            rsRpt.MoveNext
            Baris = Baris + 1
            Idx = Idx + 1
        Loop
    
        .WindowState = xlMaximized
        .Range("a" & barisawal, "AI" & barisawal + Idx).Columns.Font.Name = "Arial"
        .Range("a" & barisawal, "AI" & barisawal + Idx).Columns.Font.Size = 8
        
        .Range("A" & barisawal, "A" & barisawal + 1).Merge
        .Range("B" & barisawal, "B" & barisawal + 1).Merge
        .Range("C" & barisawal, "C" & barisawal + 1).Merge
        .Range("D" & barisawal, "D" & barisawal + 1).Merge
        .Range("E" & barisawal, "E" & barisawal + 1).Merge
    
        
        .Range("F" & barisawal, "O" & barisawal).Merge
        .Range("P" & barisawal, "S" & barisawal).Merge
        .Range("T" & barisawal, "W" & barisawal).Merge
        .Range("X" & barisawal, "AA" & barisawal).Merge
        .Range("AB" & barisawal, "AE" & barisawal).Merge
        .Range("AF" & barisawal, "AI" & barisawal).Merge
    
        .Range("a" & barisawal, "AI" & barisawal + 1).Columns.Font.Bold = True
        .Range("a" & barisawal, "AI" & barisawal + 1).VerticalAlignment = xlCenter
        .Range("a" & barisawal, "AI" & barisawal + 1).HorizontalAlignment = xlCenter
    
        .Range("A" & barisawal, "AI" & barisawal + Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("A" & barisawal, "AI" & barisawal + Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A" & barisawal, "AI" & barisawal + Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("A" & barisawal, "AI" & barisawal + Idx).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("A" & barisawal, "AI" & barisawal + Idx).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range("A" & barisawal, "AI" & barisawal + Idx).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        .Range("E" & barisawal, "AI" & barisawal).NumberFormat = "MMM YYYY"
        .Range("F" & barisawal + 2, "AI" & Baris).NumberFormat = gs_formatQty
        
        .Range("B" & barisawal, "E" & barisawal + Idx).Columns.AutoFit
        .Range("F" & barisawal, "AI" & barisawal + Idx).ColumnWidth = 9
        
        .Visible = True
        
    End With

            ' -----------------------
            '    Set rsRpt = New Recordset
            '
            '    If rsRpt.State <> adStateClosed Then rsRpt.Close
            '      rsRpt.Open SqlRpt, Db, adOpenForwardOnly, adLockReadOnly
            '
            '      If rsRpt.EOF Then lblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
            '
            '      Set Repot = application.OpenReport(App.path & "\Reports\Rpt_forecastparts.rpt")
            '      Repot.Database.Tables(1).SetDataSource rsRpt
            '      Repot.ReportTitle = "Forecast Part"
            '
            '      Fbulan = Tgl1
            '      Ftahun = Tgl2
            '      F_Factory = Trim$(lblcust)
            '      F_Cust_Name = CName
            '
            '      Repot.FormulaFields(6).Text = "'" & Format(DateAdd("m", 0, Format(Tgl1, "YYYY-MM-DD")), "MMM  YYYY") & "'"
            '      Repot.FormulaFields(5).Text = "'" & Format(DateAdd("m", 1, Format(Tgl1, "YYYY-MM-DD")), "MMM  YYYY") & "'"
            '      Repot.FormulaFields(4).Text = "'" & Format(DateAdd("m", 2, Format(Tgl1, "YYYY-MM-DD")), "MMM  YYYY") & "'"
            '      Repot.FormulaFields(3).Text = "'" & Format(DateAdd("m", 3, Format(Tgl1, "YYYY-MM-DD")), "MMM  YYYY") & "'"
            '      Repot.FormulaFields(2).Text = "'" & Format(DateAdd("m", 4, Format(Tgl1, "YYYY-MM-DD")), "MMM  YYYY") & "'"
            '      Repot.FormulaFields(1).Text = "'" & Format(DateAdd("m", 5, Format(Tgl1, "YYYY-MM-DD")), "MMM  YYYY") & "'"
            '
            '      Repot.FormulaFields(13).Text = "'" & Format(Tgl1, "MMMM YYYY") & " to " & Format(Tgl2, "MMMM YYYY") & "'"  'Periode
            '      Repot.FormulaFields(14).Text = "'" & Trim$(lblcust) & "'"  'Supplier Code
            '      Repot.FormulaFields(15).Text = "'" & CName & "'" 'Nama Perusahaan
            '
            '        '#####################################################################
            '        '# Qty Digit and decimal
            '        Repot.FormulaFields(16).Text = "" & gi_decimalDigitQty & ""
            '        Repot.FormulaFields(17).Text = "" & gi_decimalDigitQty & ""
            '        '#####################################################################
            '      reportcode = "ForecastPart"
            '      sqlprint = SqlRpt
            '      printorient = 1
            '      With FrmRpt3
            '        .CRViewer1.ReportSource = Repot
            '        .CRViewer1.ViewReport
            '        .CRViewer1.Zoom 1
            '        .WindowState = 2
            '        .Show 1
            '      End With
            '      Me.MousePointer = vbDefault
'

ElseIf OptForecast(1).Value Then 'receipt

            'SqlRpt = "select master.itm,master.Mdes, master.des,master.thick,master.width,master.length, b.hasil Bln1, c.hasil bln2, d.hasil bln3, e.hasil bln4, f.hasil bln5, g.hasil bln6 from " & _
            '       "(select itm,rtrim(MC.Description) MDes,rtrim(CS.Description)des,rtrim(IM.Thickness) Thick,rtrim(IM.width) width,rtrim(IM.length) length from(Select item_code itm from price_master where trade_code='" & Trim$(CboCust) & "' Union Select item_code itm from item_master where supplier_code='" & Trim$(CboCust) & "') c,item_master IM,sheetcoil_cls CS,material_cls MC  where c.itm=IM.Item_code and IM.sheetcoil_cls=CS.sheetcoil_cls and IM.material_cls=MC.Material_cls and (IM.sheetcoil_cls is not null) ) master, " & _
            '       " (Select R.childitem_code itm ,Childrequirement_qty Qty,ChildrequirementResult_qty Result,Childrequirement_qty-ChildrequirementResult_qty  Hasil from Requirement_master R" & _
            '       "   where R.childitem_code  in " & _
            '       "   (Select item_code from price_master where trade_code='" & Trim$(CboCust) & "' Union Select item_code from item_master where supplier_code='" & Trim$(CboCust) & "') and ChildRequirement_year='" & year(DateAdd("m", 0, Format(Tgl1, "YYYY-mm-DD"))) & "' and ChildRequirement_month ='" & month(DateAdd("m", 0, Format(Tgl1, "YYYY-mm-DD"))) & "') b, " & _
            '       " (Select R.childitem_code itm ,Childrequirement_qty Qty,ChildrequirementResult_qty Result,Childrequirement_qty-ChildrequirementResult_qty  Hasil from Requirement_master R " & _
            '       "  where R.childitem_code  in " & _
            '       "   (Select item_code from price_master where trade_code='" & Trim$(CboCust) & "' Union Select item_code from item_master where supplier_code='" & Trim$(CboCust) & "') and ChildRequirement_year='" & year(DateAdd("m", 1, Format(Tgl1, "YYYY-mm-DD"))) & "' and ChildRequirement_month ='" & month(DateAdd("m", 1, Format(Tgl1, "YYYY-mm-DD"))) & "') c , " & _
            '       " (Select R.childitem_code itm ,Childrequirement_qty Qty,ChildrequirementResult_qty Result,Childrequirement_qty-ChildrequirementResult_qty  Hasil  from Requirement_master R " & _
            '       "  where R.childitem_code  in " & _
            '       "  (Select item_code from price_master where trade_code='" & Trim$(CboCust) & "' Union Select item_code from item_master where supplier_code='" & Trim$(CboCust) & "') and ChildRequirement_year='" & year(DateAdd("m", 2, Format(Tgl1, "YYYY-mm-DD"))) & "' and ChildRequirement_month ='" & month(DateAdd("m", 2, Format(Tgl1, "YYYY-mm-DD"))) & "') d, " & _
            '       " (Select R.childitem_code itm ,Childrequirement_qty Qty,ChildrequirementResult_qty Result,Childrequirement_qty-ChildrequirementResult_qty  Hasil from Requirement_master R " & _
            '       "  Where R.childitem_code  in " & _
            '       "  (Select item_code from price_master where trade_code='" & Trim$(CboCust) & "' Union Select item_code from item_master where supplier_code='" & Trim$(CboCust) & "') and ChildRequirement_year='" & year(DateAdd("m", 3, Format(Tgl1, "YYYY-mm-DD"))) & "' and ChildRequirement_month ='" & month(DateAdd("m", 3, Format(Tgl1, "YYYY-mm-DD"))) & "') e, " & _
            '       " (Select R.childitem_code itm ,Childrequirement_qty Qty,ChildrequirementResult_qty Result,Childrequirement_qty-ChildrequirementResult_qty  Hasil from Requirement_master R " & _
            '       "  where R.childitem_code  in " & _
            '       "  (Select item_code from price_master where trade_code='" & Trim$(CboCust) & "' Union Select item_code from item_master where supplier_code='" & Trim$(CboCust) & "') and ChildRequirement_year='" & year(DateAdd("m", 4, Format(Tgl1, "YYYY-mm-DD"))) & "' and ChildRequirement_month ='" & month(DateAdd("m", 4, Format(Tgl1, "YYYY-mm-DD"))) & "') f, " & _
            '       " (Select R.childitem_code itm ,Childrequirement_qty Qty,ChildrequirementResult_qty Result,Childrequirement_qty-ChildrequirementResult_qty  Hasil from Requirement_master R " & _
            '       "  where R.childitem_code  in " & _
            '       " (Select item_code from price_master where trade_code='" & Trim$(CboCust) & "' Union  Select item_code from item_master where supplier_code='" & Trim$(CboCust) & "') and ChildRequirement_year='" & year(DateAdd("m", 5, Format(Tgl1, "YYYY-mm-DD"))) & "' and ChildRequirement_month ='" & month(DateAdd("m", 5, Format(Tgl1, "YYYY-mm-DD"))) & "') g " & _
            '       "  where master.itm *= b.itm " & _
            '       "  and master.itm *= c.itm and master.itm *= d.itm " & _
            '       "  and master.itm *= e.itm " & _
            '       "  and master.itm *= f.itm and master.itm *= g.itm " & _
            '       " order by master.des asc, master.Mdes asc,  master.thick desc, master.itm asc "
            '
            '
            '    Set rsRpt = New Recordset
            '
            '
            '    If rsRpt.State <> adStateClosed Then rsRpt.Close
            '      rsRpt.Open SqlRpt, Db, adOpenForwardOnly, adLockReadOnly
            '
            '      If rsRpt.EOF Then lblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
            '
            '      Set Repot = application.OpenReport(App.path & "\Reports\Rpt_forecastmaterial.rpt")
            '      Repot.Database.Tables(1).SetDataSource rsRpt
            '      Repot.ReportTitle = "Forecast Part"
            '
            '        '#####################################################################
            '        '# Qty Digit and decimal
            '        Repot.FormulaFields(16).Text = "" & gi_decimalDigitQty & ""
            '        Repot.FormulaFields(17).Text = "" & gi_decimalDigitQty & ""
            '        Repot.FormulaFields(18).Text = "" & gi_decimalDigitThickness & ""
            '        Repot.FormulaFields(19).Text = "" & gi_decimalDigitThickness & ""
            '        Repot.FormulaFields(20).Text = "" & gi_decimalDigitWidth & ""
            '        Repot.FormulaFields(21).Text = "" & gi_decimalDigitWidth & ""
            '        Repot.FormulaFields(22).Text = "" & gi_decimalDigitLength & ""
            '        Repot.FormulaFields(23).Text = "" & gi_decimalDigitLength & ""
            '        '#####################################################################
            '
            '      Fbulan = Tgl1
            '      Ftahun = Tgl2
            '      F_Factory = Trim$(lblcust)
            '      F_Cust_Name = CName
            '
            '      Repot.FormulaFields(6).Text = "'" & Format(DateAdd("m", 0, Format(Tgl1, "YYYY-MM-DD")), "MMM  YYYY") & "'"
            '      Repot.FormulaFields(5).Text = "'" & Format(DateAdd("m", 1, Format(Tgl1, "YYYY-MM-DD")), "MMM  YYYY") & "'"
            '      Repot.FormulaFields(4).Text = "'" & Format(DateAdd("m", 2, Format(Tgl1, "YYYY-MM-DD")), "MMM  YYYY") & "'"
            '      Repot.FormulaFields(3).Text = "'" & Format(DateAdd("m", 3, Format(Tgl1, "YYYY-MM-DD")), "MMM  YYYY") & "'"
            '      Repot.FormulaFields(2).Text = "'" & Format(DateAdd("m", 4, Format(Tgl1, "YYYY-MM-DD")), "MMM  YYYY") & "'"
            '      Repot.FormulaFields(1).Text = "'" & Format(DateAdd("m", 5, Format(Tgl1, "YYYY-MM-DD")), "MMM  YYYY") & "'"
            ''
            '      Repot.FormulaFields(13).Text = "'" & Format(Tgl1, "MMMM YYYY") & " to " & Format(Tgl2, "MMMM YYYY") & "'"  'Periode
            '      Repot.FormulaFields(14).Text = "'" & Trim$(lblcust) & "'"   'Supplier Code
            '      Repot.FormulaFields(15).Text = "'" & CName & "'"  'Nama Perusahaan
            '
            '      reportcode = "ForecastMaterial"
            '      sqlprint = SqlRpt
            '      printorient = 2
            '      With FrmRpt3
            '        .CRViewer1.ReportSource = Repot
            '        .CRViewer1.ViewReport
            '        .CRViewer1.Zoom 1
            '        .WindowState = 2
            '        .Show 1
            '      End With

End If

Me.MousePointer = vbDefault
Pb1.Visible = False
Label1.Visible = False

End Sub

'************ Unload **********
Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
HakU = hakUpdate(Me.Name)
adtocboCust
cboCust.ListIndex = 0
Tgl1 = Format(Now, "MMM YYYY")
tgl1_Click
OptForecast(0).Value = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub
'**************
Sub adtocboCust()
Dim sqlcust As String
Dim RsCust As New Recordset
Dim i As Long
    sqlcust = "select trade_code, trade_name, address1 from trade_master where trade_cls='2' or trade_cls='3'"
    Set RsCust = Db.Execute(sqlcust)

    With cboCust
        .clear
        .columnCount = 3
        .ColumnWidths = "80pt;300pt;0pt"
        .ListWidth = 380
        .ListRows = 15
        
        .AddItem
        .List(0, 0) = strAll
        .List(0, 1) = strAll
        .List(0, 2) = strAll
        
        i = 1
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("Trade_code"))
            .List(i, 1) = IIf(IsNull(RsCust("trade_name")), " ", Trim(RsCust("Trade_Name")))
            .List(i, 2) = IIf(IsNull(RsCust("address1")), " ", Trim(RsCust("Address1")))
            RsCust.MoveNext
            i = i + 1

        Loop
    End With
End Sub

Private Sub Tgl1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
tgl1_Click
End Sub

Private Sub tgl1_Change()
tgl1_Click
tgl_sb = Tgl1.Month
End Sub

Private Sub tgl1_Click()
If Tgl1.Month = 1 And Val(tgl_sb) = 12 Then Tgl1.Year = Tgl1.Year + 1
If Tgl1.Month = 12 And Val(tgl_sb) = 1 Then Tgl1.Year = Tgl1.Year - 1
Tgl2 = Format(DateAdd("m", 5, Format(Tgl1, "yyyy-mm-dd")), "MMM YYYY")
End Sub
