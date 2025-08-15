VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmTradeMasterInquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Trade Master Inquiry"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmTradeMasterInquiry.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      Index           =   0
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8325
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Filter"
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
      Left            =   9435
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8325
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Refresh"
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
      Left            =   10710
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8325
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Copy"
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
      Left            =   11265
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9720
      Width           =   1155
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
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      TabIndex        =   7
      Top             =   9720
      Width           =   1155
   End
   Begin VB.CommandButton cmdreport 
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
      Left            =   10050
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9720
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
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
      Index           =   0
      Left            =   13695
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9720
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   390
      TabIndex        =   13
      Top             =   8880
      Width           =   14445
      Begin VB.Label LblErrMsg 
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
         Top             =   210
         Width           =   14175
      End
   End
   Begin VB.ComboBox cbocari 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FrmTradeMasterInquiry.frx":0E42
      Left            =   5400
      List            =   "FrmTradeMasterInquiry.frx":0E44
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   8355
      Width           =   2415
   End
   Begin VB.TextBox txtcari 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   8355
      Width           =   3255
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12960
      TabIndex        =   15
      Top             =   240
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   7125
      Left            =   330
      TabIndex        =   16
      Top             =   900
      Width           =   14490
      _cx             =   25559
      _cy             =   12568
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
      GridColor       =   12582912
      GridColorFixed  =   12582912
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
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
      BackStyle       =   0  'Transparent
      Caption         =   "By :"
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
      Left            =   4920
      TabIndex        =   12
      Top             =   8400
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      Left            =   480
      TabIndex        =   11
      Top             =   8400
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Master Inquiry"
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
      Left            =   330
      TabIndex        =   10
      Top             =   240
      Width           =   14505
   End
End
Attribute VB_Name = "FrmTradeMasterInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS As New ADODB.Recordset
Dim i As Integer
Dim cari As String
Dim sql As String

Dim bteColSelect As Byte
Dim bteColTradeCode As Byte
Dim bteColTradeCls As Byte
Dim bteColTradeName As Byte
Dim bteColTradeAbbr As Byte
Dim bteColContact As Byte
Dim bteColAddress1 As Byte
Dim bteColAddress2 As Byte
Dim bteColCity As Byte
Dim bteColPostCode As Byte
Dim bteColPhone As Byte
Dim bteColFax As Byte
Dim bteColClosing As Byte
Dim bteColPay As Byte
Dim bteColNPWPNo As Byte
Dim bteColNPWPName As Byte
Dim bteColNPWPAddress As Byte
Dim bteColNPWPCity As Byte
Dim bteColTerms As Byte
Dim bteColCountry As Byte
Dim bteColCountryCls As Byte
Dim bteColEpteCls As Byte
Dim bteColInvoiceTo As Byte
Dim bteColPOCls As Byte
Dim bteColRegion As Byte
Dim bteColAffiliateCompany As Byte

Sub Header()
    
    bteColSelect = 0
    bteColTradeCode = 1
    bteColTradeCls = 2
    bteColTradeName = 3
    bteColTradeAbbr = 4
    bteColContact = 5
    bteColAddress1 = 6
    bteColAddress2 = 7
    bteColCity = 8
    bteColPostCode = 9
    bteColPhone = 10
    bteColFax = 11
    bteColClosing = 12
    bteColPay = 13
    bteColNPWPNo = 14
    bteColNPWPName = 15
    bteColNPWPAddress = 16
    bteColNPWPCity = 17
    bteColTerms = 18
    bteColCountry = 19
    bteColCountryCls = 20
    bteColEpteCls = 21
    bteColInvoiceTo = 22
    bteColPOCls = 23
    bteColRegion = 24
    bteColAffiliateCompany = 25
    
    With grid
        .clear
        .Rows = 1
        .ColS = 26
                
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColTradeCode) = "Trade Code"
        .TextMatrix(0, bteColTradeCls) = "Trade Cls"
        .TextMatrix(0, bteColTradeName) = "Trade Name"
        .TextMatrix(0, bteColTradeAbbr) = "Trade ABBR"
        .TextMatrix(0, bteColContact) = "Contact Person"
        .TextMatrix(0, bteColAddress1) = "Address 1"
        .TextMatrix(0, bteColAddress2) = "Address 2"
        .TextMatrix(0, bteColCity) = "City"
        .TextMatrix(0, bteColPostCode) = "Postal Code"
        .TextMatrix(0, bteColPhone) = "Telephone"
        .TextMatrix(0, bteColFax) = "Fax"
        .TextMatrix(0, bteColClosing) = "Closing Day"
        .TextMatrix(0, bteColPay) = "Pay Day"
        .TextMatrix(0, bteColNPWPNo) = "NPWP No"
        .TextMatrix(0, bteColNPWPName) = "NPWP Name"
        .TextMatrix(0, bteColNPWPAddress) = "NPWP Address"
        .TextMatrix(0, bteColNPWPCity) = "NPWP City"
        .TextMatrix(0, bteColTerms) = "Payment Terms"
        .TextMatrix(0, bteColCountry) = "Country"
        .TextMatrix(0, bteColCountryCls) = "Country Cls"
        .TextMatrix(0, bteColEpteCls) = "Epte Cls"
        .TextMatrix(0, bteColInvoiceTo) = "Invoice To"
        .TextMatrix(0, bteColPOCls) = "PO Cls"
        .TextMatrix(0, bteColRegion) = "Region"
        .TextMatrix(0, bteColAffiliateCompany) = "Affiliate Company"
             
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColTradeCode) = 1200
        .ColWidth(bteColTradeCls) = 1000
        .ColWidth(bteColTradeName) = 3000
        .ColWidth(bteColTradeAbbr) = 2500
        .ColWidth(bteColContact) = 2500
        .ColWidth(bteColAddress1) = 3000
        .ColWidth(bteColAddress2) = 3000
        .ColWidth(bteColCity) = 2000
        .ColWidth(bteColPostCode) = 1200
        .ColWidth(bteColPhone) = 1800
        .ColWidth(bteColFax) = 1800
        .ColWidth(bteColClosing) = 1200
        .ColWidth(bteColPay) = 1000
        .ColWidth(bteColNPWPNo) = 2000
        .ColWidth(bteColNPWPName) = 3000
        .ColWidth(bteColNPWPAddress) = 4500
        .ColWidth(bteColNPWPCity) = 2000
        .ColWidth(bteColTerms) = 3000
        .ColWidth(bteColCountry) = 3000
        .ColWidth(bteColCountryCls) = 1500
        .ColWidth(bteColEpteCls) = 1000
        .ColWidth(bteColInvoiceTo) = 1200
        .ColWidth(bteColPOCls) = 700
        .ColWidth(bteColRegion) = 1500
        .ColWidth(bteColAffiliateCompany) = 2000
        
        .Cell(flexcpAlignment, 0, 0, 0, bteColAffiliateCompany) = flexAlignCenterCenter
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColTradeCode) = flexAlignLeftCenter
        .ColAlignment(bteColTradeCls) = flexAlignCenterCenter
        For i = bteColTradeName To bteColFax
        .ColAlignment(i) = flexAlignLeftCenter
        Next i
        .ColAlignment(bteColClosing) = flexAlignCenterCenter
        .ColAlignment(bteColPay) = flexAlignCenterCenter
        For i = bteColNPWPNo To bteColAffiliateCompany
        .ColAlignment(i) = flexAlignLeftCenter
        Next i
        .EditMaxLength = 1
    End With
End Sub

Sub Kosong()
  txtCari.Text = ""
  LblErrMsg.Caption = ""
'  cbocari.Text = ""
  cboCari.ListIndex = -1
End Sub

Sub Browse()
  Header
  i = 1
  
  RS.Requery
  If RS.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Exit Sub

  RS.MoveFirst
  Do While Not RS.EOF
     
      grid.Rows = grid.Rows + 1
      grid.Cell(flexcpBackColor, i, bteColSelect) = &HFFFFFF
      With grid
        .TextMatrix(i, bteColTradeCode) = IIf(IsNull(RS!Trade_Code), "", Trim(RS!Trade_Code))
        .TextMatrix(i, bteColTradeCls) = IIf(IsNull(RS!trade_cls), "", Trim(RS!trade_cls))
        .TextMatrix(i, bteColTradeName) = IIf(IsNull(RS!trade_name), "", Trim(RS!trade_name))
        .TextMatrix(i, bteColTradeAbbr) = IIf(IsNull(RS!trade_abbr), "", Trim(RS!trade_abbr))
        .TextMatrix(i, bteColContact) = IIf(IsNull(RS!contact_person), "", Trim(RS!contact_person))
        .TextMatrix(i, bteColAddress1) = IIf(IsNull(RS!address1), "", Trim(RS!address1))
        .TextMatrix(i, bteColAddress2) = IIf(IsNull(RS!address2), "", Trim(RS!address2))
        .TextMatrix(i, bteColCity) = IIf(IsNull(RS!City), "", Trim(RS!City))
        .TextMatrix(i, bteColPostCode) = IIf(IsNull(RS!postal_code), "", Trim(RS!postal_code))
        .TextMatrix(i, bteColPhone) = IIf(IsNull(RS!Telephone), "", Trim(RS!Telephone))
        .TextMatrix(i, bteColFax) = IIf(IsNull(RS!fax), "", Trim(RS!fax))
        .TextMatrix(i, bteColClosing) = IIf(IsNull(RS!Closing_Day), "", Trim(RS!Closing_Day))
        .TextMatrix(i, bteColPay) = IIf(IsNull(RS!Pay_Day), "", Trim(RS!Pay_Day))
        .TextMatrix(i, bteColNPWPNo) = IIf(IsNull(RS!NPWP_No), "", Trim(RS!NPWP_No))
        .TextMatrix(i, bteColNPWPName) = IIf(IsNull(RS!NPWP_Name), "", Trim(RS!NPWP_Name))
        .TextMatrix(i, bteColNPWPAddress) = IIf(IsNull(RS!NPWP_Address), "", Trim(RS!NPWP_Address))
        .TextMatrix(i, bteColNPWPCity) = IIf(IsNull(RS!NPWP_City), "", Trim(RS!NPWP_City))
        .TextMatrix(i, bteColTerms) = uf_Description(IIf(IsNull(RS!POPayment_Terms), "", Trim(RS!POPayment_Terms)), "PaymentTerm_Cls", "PaymentTerm_Cls")
        .TextMatrix(i, bteColNPWPCity) = IIf(IsNull(RS!NPWP_City), "", Trim(RS!NPWP_City))
        .TextMatrix(i, bteColCountry) = IIf(IsNull(RS!Country), "", Trim(RS!Country))
        
        If RS("country_cls") = 1 Then
         .TextMatrix(i, bteColCountryCls) = "Overseas"
        Else
         .TextMatrix(i, bteColCountryCls) = "Domestic"
        End If
        
        If RS("Epte_cls") = 1 Then
        .TextMatrix(i, bteColEpteCls) = "Yes"
        Else
        .TextMatrix(i, bteColEpteCls) = "No"
        End If
        
        .TextMatrix(i, bteColInvoiceTo) = IIf(IsNull(RS("Invoice_to")), "", Trim(RS("Invoice_to")))
        
        If RS("po_cls") = 1 Then
        .TextMatrix(i, bteColPOCls) = "Yes"
        Else
        .TextMatrix(i, bteColPOCls) = "No"
        End If
               
        .TextMatrix(i, bteColRegion) = uf_Description(IIf(IsNull(RS!Region_Cls), "", Trim(RS!Region_Cls)), "Region_Cls", "Region_Cls")
        
        If RS("affiliate_cls") = 1 Then
        .TextMatrix(i, bteColAffiliateCompany) = "Yes"
        Else
        .TextMatrix(i, bteColAffiliateCompany) = "No"
        End If
        
     End With
     
     RS.MoveNext
     i = i + 1
  Loop

End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
  sql = "select * from Trade_Master"
  Set RS = Db.Execute(sql)
  CtrlMenu1.FormName = Me.Name
  Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
  cmdsubmenu.Caption = "Sub &Menu"
  
  With cboCari
    .AddItem "Trade Code"
    .AddItem "Trade Cls"
    .AddItem "Trade Name"
    .AddItem "Trade ABBR"
    .AddItem "Contact Person"
    .AddItem "Address 1"
    .AddItem "Address 2"
    .AddItem "City"
    .AddItem "Postal Code"
    .AddItem "Telephone"
    .AddItem "Fax"
    .AddItem "Closing Day"
    .AddItem "Pay Day"
    .AddItem "NPWP No"
    .AddItem "NPWP Name"
    .AddItem "NPWP Address"
    .AddItem "NPWP City"
    .AddItem "Payment Terms"
    .AddItem "Country"
    .AddItem "Country Cls"
    .AddItem "Epte Cls"
    .AddItem "Invoice To"
    .AddItem "PO Cls"
    .AddItem "Region"
    .AddItem "Affiliate Company"
   End With
  
  Kosong
  'Header
  Browse
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim TextGrid As String

With grid
    TextGrid = grid.Text

    If TextGrid = "S" Then
       Call kosongColGrid
    ElseIf TextGrid = "D" Then
       Call kosongColGrid("S")
    End If

    .TextMatrix(Row, Col) = TextGrid
End With

End Sub

Private Sub kosongColGrid(Optional Kolom As String)
    Dim i As Integer
    
    With grid
        .Col = bteColSelect
    
        If Kolom <> "" Then
           For i = 1 To .Rows - 1

              If .Text = Kolom Then .Text = ""
              If .TextMatrix(i, bteColSelect) <> "D" Then .TextMatrix(i, bteColSelect) = ""
           Next i
        Else
           For i = 1 To .Rows - 1

              If .TextMatrix(i, bteColSelect) <> "" Then .TextMatrix(i, bteColSelect) = ""

           Next i
        End If
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  If grid.Col <> bteColSelect Then Cancel = True
End Sub

Private Sub grid_Click()
  With grid
    If .Row = 1 And .Col <> bteColSelect Then
      If .ColSort(.Col) = flexSortStringAscending Then
        .ColSort(.Col) = flexSortStringDescending
      Else
        .ColSort(.Col) = flexSortStringAscending
      End If
      .Sort = .ColSort(.Col)
    End If
  End With
End Sub

'Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
'  With grid
'    If KeyCode = vbKeyReturn Then
'      If .Row = .Rows - 1 Then
'        .TopRow = 1
'        .Row = 1
'      Else
'        .Row = .Row + 1
'        .TopRow = .TopRow + 1
'      End If
'      .SetFocus
'    End If
'  End With
'
'End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
  If grid.Col = bteColSelect Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> Asc("S") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
  End If
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbocari_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cbocari_Click()
  Select Case cboCari
    Case "Trade Code": cari = "Trade_Code"
    Case "Trade Cls": cari = "Trade_Cls"
    Case "Trade Name": cari = "Trade_Name"
    Case "Trade ABBR": cari = "Trade_Abbr"
    Case "Contact Person": cari = "Contact_Person"
    Case "Address 1": cari = "Address1"
    Case "Address 2": cari = "Address2"
    Case "City": cari = "City"
    Case "Postal Code": cari = "Postal_Code"
    Case "Telephone": cari = "Telephone"
    Case "Fax": cari = "Fax"
    Case "Closing Day": cari = "Closing_Day"
    Case "Pay Day": cari = "Pay_Day"
    Case "NPWP No": cari = "NPWP_No"
    Case "NPWP Name": cari = "NPWP_Name"
    Case "NPWP Address": cari = "NPWP_Address"
    Case "NPWP City": cari = "NPWP_City"
    Case "Payment Terms": cari = "POPayment_Terms"
    Case "Country": cari = "Country"
    Case "Country Cls": cari = "Country_Cls"
    Case "Epte Cls": cari = "Epte_Cls"
    Case "Invoice To": cari = "Invoice_To"
    Case "PO Cls": cari = "Po_Cls"
    Case "Region": cari = "Region_Cls"
    Case "Affiliate Company": cari = "Affiliate_Cls"
  End Select
  
End Sub

Sub carigrid()
Dim NO As Integer

LblErrMsg.Caption = ""

  With grid
    For i = 1 To .Rows - 1
      If InStr(LCase(.TextMatrix(i, cboCari.ListIndex + 1)), LCase(txtCari.Text)) Then
        .Row = i
        .SetFocus
        NO = 0
        If i <> 1 Then .TopRow = i - 1
        Exit Sub
      Else
        NO = 1
      End If
    Next i
  End With
  
  If NO = 1 Then LblErrMsg.Caption = DisplayMsg(4006)
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0:
        If txtCari.Text = "" Then
            LblErrMsg.Caption = DisplayMsg(4007)
            txtCari.SetFocus
        ElseIf cboCari.Text = "" Then
            LblErrMsg.Caption = DisplayMsg(4008)
            cboCari.SetFocus
        Else
            carigrid
        End If
Case 1:
        If txtCari.Text = "" Then
            LblErrMsg.Caption = DisplayMsg(4007)
        ElseIf cboCari.Text = "" Then
            LblErrMsg.Caption = DisplayMsg(4008)
        Else
            LblErrMsg.Caption = ""
            
           If cari = "POPayment_Terms" Then
            sql = "select * from Trade_Master tm where exists " & _
                 "(select * from PaymentTerm_Cls pc where tm.POPayment_Terms=pc.PaymentTerm_Cls and description like '%" & txtCari.Text & "%')"
            Set RS = Db.Execute(sql)
           ElseIf cari = "Country_Cls" Then
                If InStr(1, "oes", LCase(Trim(txtCari.Text))) > 0 Then
                    sql = "select * from trade_master"
                    Set RS = Db.Execute(sql)
                ElseIf InStr(1, "overseas", LCase(Trim(txtCari.Text))) > 0 Then
                    sql = "select * from trade_master where country_cls=1"
                    Set RS = Db.Execute(sql)
'                    Rs.Filter = "country_cls = 1"
                ElseIf InStr(1, "domestic", LCase(Trim(txtCari.Text))) > 0 Then
                    sql = "select * from trade_master where country_cls is null or country_cls =0"
                    Set RS = Db.Execute(sql)
'                    Rs.Filter = "(country_cls is null) or (country_cls =0)"
                End If
           ElseIf cari = "Epte_Cls" Then
                If InStr(1, "yes", LCase(Trim(txtCari.Text))) > 0 Then
                    sql = "select * from trade_master where Epte_Cls=1"
                    Set RS = Db.Execute(sql)
                ElseIf InStr(1, "no", LCase(Trim(txtCari.Text))) > 0 Then
                    sql = "select * from trade_master where Epte_Cls is null or Epte_Cls=0"
                    Set RS = Db.Execute(sql)
                End If
           ElseIf cari = "Po_Cls" Then
                If InStr(1, "yes", LCase(Trim(txtCari.Text))) > 0 Then
                    sql = "select * from trade_master where po_cls=1"
                    Set RS = Db.Execute(sql)
                ElseIf InStr(1, "no", LCase(Trim(txtCari.Text))) > 0 Then
                    sql = "select * from trade_master where po_cls is null or po_cls=0"
                    Set RS = Db.Execute(sql)
                End If
           ElseIf cari = "Region_Cls" Then
             sql = "select * from Trade_Master tm where exists " & _
                 "(select * from Region_Cls rc where tm.region_cls=rc.region_cls and description like '%" & txtCari.Text & "%')"
             Set RS = Db.Execute(sql)
           ElseIf cari = "Affiliate_Cls" Then
                If InStr(1, "yes", LCase(Trim(txtCari.Text))) > 0 Then
                    sql = "select * from trade_master where Affiliate_cls=1"
                    Set RS = Db.Execute(sql)
                ElseIf InStr(1, "no", LCase(Trim(txtCari.Text))) > 0 Then
                    sql = "select * from trade_master where Affiliate_cls is null or Affiliate_cls=0"
                    Set RS = Db.Execute(sql)
                End If
           Else
                sql = "select * from trade_master where " & cari & " like '%" & txtCari.Text & "%'"
                Set RS = Db.Execute(sql)
'                Rs.Filter = cari & " like '%" & txtcari.Text & "%'"
           End If
            
            Browse
        End If
 Case 2:
        'rs.Requery
        sql = "select * from trade_master"
        Set RS = Db.Execute(sql)
'        Rs.Filter = ""
        'header
        Browse
        Kosong
 End Select

End Sub

Private Sub command2_Click(Index As Integer)
Dim tanya
Dim hapus As Boolean
Dim sql1 As String
Dim rs1 As New Recordset

hapus = False
Select Case Index
    Case 0:
        If hakAkses("FrmTradeMaster") = 0 Then LblErrMsg = DisplayMsg(3007):  Exit Sub
        If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Exit Sub
    
        With grid
            For i = 1 To .Rows - 1
                If .TextMatrix(i, bteColSelect) = "S" Then
                  FrmTradeMaster.Text1(0).Text = .TextMatrix(i, bteColTradeCode)
                  FrmTradeMaster.enter
                  FrmTradeMaster.Show
                  Unload Me
                  Exit Sub
                ElseIf .TextMatrix(i, bteColSelect) = "D" Then
                  If IsEmpty(tanya) Then tanya = MsgBox("Do You Really Want To Delete This Data?", vbQuestion & vbYesNo, "Confirmation")
                  
                  If tanya = vbYes Then
                    sql1 = "select * from Item_Master where Supplier_code='" & .TextMatrix(i, bteColTradeCode) & "'"
                    Set rs1 = Db.Execute(sql1)
                    If Not (rs1.BOF And rs1.EOF) Then
                        LblErrMsg.Caption = DisplayMsg(1204)
                        .Row = i
                        .SetFocus
                        Exit Sub
                    End If
                    
                    sql1 = "select * from Warehouse_Master where adm_group='" & .TextMatrix(i, bteColTradeCode) & "'"
                    Set rs1 = Db.Execute(sql1)
                    If Not (rs1.BOF And rs1.EOF) Then
                        LblErrMsg.Caption = DisplayMsg(1204)
                        .Row = i
                        .SetFocus
                        Exit Sub
                    End If
                        
                    sql = "delete from Trade_Master where Trade_Code='" & .TextMatrix(i, bteColTradeCode) & "'"
                    Db.Execute (sql)
                            
                    hapus = True
                  Else
                    Exit For
                  End If
                End If
            Next i
        End With
        If (hapus) Then LblErrMsg = DisplayMsg(1201)
        hapus = False
        'rs.Requery
        'header
        Browse
    
    Case 1:
        'rs.Requery
        Kosong
        'header
        Browse
    Case 2:
        With grid
            For i = 1 To .Rows - 1
                If .TextMatrix(i, bteColSelect) = "S" Then
                  FrmTradeMaster.Text1(0).Text = .TextMatrix(i, bteColTradeCode)
                  FrmTradeMaster.enter
                  FrmTradeMaster.ubahtrade = False
                  FrmTradeMaster.Text1(0).Text = ""
                  FrmTradeMaster.Text1(0).DataChanged = False
                  FrmTradeMaster.Text1(0).Enabled = True
                  FrmTradeMaster.headerGrid
                  FrmTradeMaster.Show
                  Unload Me
                  Exit Sub
                End If
            Next i
        End With
   
End Select

End Sub

Private Sub cmdReport_Click()
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset
  Dim Rpt As New FrmRpt3
  
  Me.MousePointer = vbHourglass
  
  sql = "select Trade_Code, Trade_Cls, Trade_Name, Trade_abbr, Contact_Person, " & _
        "Address1, Address2, City, Country, Country_Cls, Epte_Cls, IsNull(tm.region_cls,'') " & _
        "Region_Cls, IsNull(rc.description,'') Region_Desc, Postal_Code, " & _
        "Telephone, Fax, Closing_Day, Pay_Day, IsNull(Affiliate_Cls,'0') Affiliate_Cls, " & _
        "NPWP_No, NPWP_Name, NPWP_Address, NPWP_City, IsNull(Invoice_To,'') Invoice_To, PO_Cls, " & _
        "IsNull(POPayment_Terms,'') POPayment_Terms, IsNull(pc.description,'') Payment_Desc " & _
        "From Trade_Master tm Left join PaymentTerm_Cls pc on tm.POPayment_Terms = pc.PaymentTerm_Cls " & _
        " Left join Region_Cls rc  on tm.region_cls = rc.region_cls " & _
        " "
  
 ' Sql = "Select * from Trade_Master "
  
  
  If txtCari.Text <> "" And cboCari.Text <> "" Then
    If cari = "POPayment_Terms" Then
        sql = sql & "and exists " & _
              "(select * from PaymentTerm_cls pc where description like " & _
              "'%" & txtCari.Text & "%' and tm.POPayment_Terms = pc.PaymentTerm_Cls)"
    ElseIf cari = "Country_cls" Then
        If InStr(1, "oes", LCase(Trim(txtCari.Text))) > 0 Then
           sql = sql
        ElseIf InStr(1, "overseas", LCase(Trim(txtCari.Text))) > 0 Then
           sql = sql & "and country_cls=1"
        ElseIf InStr(1, "domestic", LCase(Trim(txtCari.Text))) > 0 Then
           sql = sql & "and (country_cls is null or country_cls=0)"
        End If
    ElseIf cari = "Epte_Cls" Then
       If InStr(1, "yes", LCase(Trim(txtCari.Text))) > 0 Then
            sql = sql & "and Epte_Cls=1"
       ElseIf InStr(1, "no", LCase(Trim(txtCari.Text))) > 0 Then
            sql = sql & "and (Epte_Cls is null or Epte_Cls=0)"
       End If
    ElseIf cari = "po_cls" Then
        If InStr(1, "yes", LCase(Trim(txtCari.Text))) > 0 Then
            sql = sql & "and po_cls=1"
        ElseIf InStr(1, "no", LCase(Trim(txtCari.Text))) > 0 Then
            sql = sql & "and (po_cls is null or po_cls=0)"
        End If
    ElseIf cari = "Region_Cls" Then
        sql = sql & "and exists " & _
              "(select * from region_cls rc where description like " & _
              "'%" & txtCari.Text & "%' and tm.Region_Cls = rc.Region_Cls)"
    ElseIf cari = "Affiliate_cls" Then
       If InStr(1, "yes", LCase(Trim(txtCari.Text))) > 0 Then
            sql = sql & "and Affiliate_cls=1"
       ElseIf InStr(1, "no", LCase(Trim(txtCari.Text))) > 0 Then
            sql = sql & "and (Affiliate_cls is null or Affiliate_cls=0)"
       End If
    Else
        sql = sql & " and " & cari & " like '%" & txtCari.Text & "%' "
    End If

  End If
  
  sql = sql & " order by Trade_Code"
  
  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
  
  If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
  sqlprint = sql
  reportcode = "trademaster"
  printorient = 2
  Set report = application.OpenReport(App.path & "\Reports\rptTradeMaster.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  report.ReportTitle = "Trade Master"
  
  Rpt.CRViewer1.ReportSource = report
  Rpt.CRViewer1.ViewReport
  Rpt.CRViewer1.Zoom 1
  
  Rpt.WindowState = 2
  Rpt.Show 1
  
  Me.MousePointer = vbDefault
  
End Sub

Private Sub CmdSubMenu_Click()
  If cmdsubmenu.Caption = "&Back" Then
    FrmTradeMaster.Show
  Else
    frmMainMenu.Show
  End If
  
  Unload Me
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If RS.State <> adStateClosed Then RS.Close
  cmdsubmenu.Caption = "Sub &Menu"
End Sub

