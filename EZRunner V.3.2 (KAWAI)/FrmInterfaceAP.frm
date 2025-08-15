VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmInterfaceAP 
   BackColor       =   &H00FDDFE3&
   Caption         =   "AP Interface"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmInterfaceAP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10680
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet inetFTP 
      Left            =   1440
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1800
      Left            =   285
      TabIndex        =   7
      Tag             =   "TTTF*/"
      Top             =   1080
      Width           =   14640
      Begin VB.CommandButton CmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Search"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "TTFF*/"
         Top             =   1305
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker InvFrom 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1530
         _ExtentX        =   2699
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
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker InvTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   13
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1530
         _ExtentX        =   2699
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
         CurrentDate     =   37798
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interface Cls"
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
         Left            =   240
         TabIndex        =   19
         Tag             =   "TTFF*/"
         Top             =   1395
         Width           =   1110
      End
      Begin MSForms.ComboBox cbointeface 
         Height          =   345
         Left            =   1680
         TabIndex        =   18
         Tag             =   "TTFF*/"
         Top             =   1320
         Width           =   1530
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2699;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         X1              =   3360
         X2              =   7620
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label Label2 
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
         Left            =   3360
         TabIndex        =   14
         Tag             =   "TTFF*/"
         Top             =   960
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice From"
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
         Left            =   240
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   900
         Width           =   1125
      End
      Begin VB.Label LblCustomer 
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
         Left            =   3420
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   60
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
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   1215
      End
      Begin MSForms.ComboBox cboSupplier 
         Height          =   345
         Left            =   1680
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   300
         Width           =   1530
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2699;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   270
      TabIndex        =   5
      Tag             =   "TTTF*/"
      Top             =   9300
      Width           =   14640
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   375
         Left            =   75
         TabIndex        =   20
         Tag             =   "TTTF*/"
         Top             =   180
         Visible         =   0   'False
         Width           =   14490
         _ExtentX        =   25559
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
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
         TabIndex        =   6
         Tag             =   "TTTF*/"
         Top             =   180
         Width           =   14325
      End
   End
   Begin VB.CommandButton Cmd_Clear 
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
      Left            =   12570
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "FFTT*/"
      Top             =   10020
      Width           =   1125
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
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "TTFF*/"
      Top             =   10020
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_save 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Export"
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
      Left            =   13770
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "FFTT*/"
      Top             =   10020
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13080
      TabIndex        =   4
      Tag             =   "FTTF*/"
      Top             =   360
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5970
      Left            =   285
      TabIndex        =   17
      Tag             =   "TTTT*/"
      Top             =   2940
      Width           =   14640
      _cx             =   25823
      _cy             =   10530
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
   Begin MSForms.Label LblRecord 
      Height          =   255
      Left            =   11580
      TabIndex        =   16
      Tag             =   "FFTT*/"
      Top             =   9060
      Width           =   3345
      BackColor       =   16637923
      VariousPropertyBits=   8388627
      Caption         =   "0 Record(s)"
      Size            =   "5900;450"
      FontName        =   "Verdana"
      FontEffects     =   1073741827
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AP Interface"
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
      Index           =   0
      Left            =   300
      TabIndex        =   3
      Tag             =   "TTTF*/"
      Top             =   390
      Width           =   14610
   End
End
Attribute VB_Name = "FrmInterfaceAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_GettingDir As Boolean
Dim LocalPath As String

Dim sql As String
Dim RS As New Recordset
Dim ubah As Boolean, hapus As Boolean, gavalid As Boolean, ubahedate As Boolean
Dim SDate, EDate, sdateawal, edateakhir
Dim i As Long

Dim bteColSelect As Byte
Dim bteColPostingCode As Byte
Dim bteColLedger As Byte
Dim bteColCompCode As Byte
Dim bteColCalTaxAuto As Byte
Dim bteColDocDate As Byte
Dim bteColPosDate As Byte
Dim bteColDocType As Byte
Dim bteColCurr As Byte
Dim bteColExRate As Byte
Dim bteColInvNo As Byte
Dim bteColDocHeader As Byte
Dim bteColTrade As Byte
Dim bteColPostKey As Byte
Dim bteColAccNo As Byte
Dim bteColSpGL As Byte
Dim bteColAssets As Byte
Dim bteColReconAcc As Byte
Dim bteColAmount As Byte
Dim bteColAmountLocal As Byte
Dim bteColGroupCurr As Byte
Dim bteColCompCurr As Byte
Dim bteColTaxCode As Byte
Dim bteColTaxType As Byte
Dim bteColCostCenter As Byte
Dim bteColProfCenter As Byte
Dim bteColBussArea As Byte
Dim bteColTradePartner As Byte
Dim bteColPartnerProfit As Byte
Dim bteColPaymentTerm As Byte
Dim bteColSite As Byte
Dim bteColPayBlock As Byte
Dim bteColPayRef As Byte
Dim bteColPayMeth As Byte
Dim bteColBaseLine As Byte
Dim bteColNoteIssue As Byte
Dim bteColWBSelement As Byte
Dim bteColOrderNo As Byte
Dim bteColProdCode As Byte
Dim bteColFinPlan As Byte
Dim bteColStartDate As Byte
Dim bteColRef1 As Byte
Dim bteColRef2 As Byte
Dim bteColRef3 As Byte
Dim bteColCentBank As Byte
Dim bteColsupplCountry As Byte
Dim bteColHouseBank As Byte
Dim bteColAssetVal As Byte
Dim bteColQty As Byte
Dim bteColUnitQty As Byte
Dim bteColSegment As Byte
Dim bteColSortKey As Byte
Dim bteColItemText As Byte
Dim bteColTransType As Byte
Dim bteColBussPart As Byte

Sub Header()

    Dim X As Integer
    
    LblErrMsg = ""
    LblRecord = "0 Record(s)"
    
    bteColSelect = 0
    bteColPostingCode = 1
    bteColLedger = 2
    bteColCompCode = 3
    bteColCalTaxAuto = 4
    bteColDocDate = 5
    bteColPosDate = 6
    bteColDocType = 7
    bteColCurr = 8
    bteColExRate = 9
    bteColInvNo = 11
    bteColDocHeader = 10
    bteColTrade = 12
    bteColPostKey = 13
    bteColAccNo = 14
    bteColSpGL = 15
    bteColAssets = 16
    bteColReconAcc = 17
    bteColAmount = 18
    bteColAmountLocal = 19
    bteColGroupCurr = 20
    bteColCompCurr = 21
    bteColTaxCode = 22
    bteColTaxType = 23
    bteColCostCenter = 24
    bteColProfCenter = 25
    bteColBussArea = 26
    bteColTradePartner = 27
    bteColPartnerProfit = 28
    bteColPaymentTerm = 29
    bteColSite = 30
    bteColPayBlock = 31
    bteColPayRef = 32
    bteColPayMeth = 33
    bteColBaseLine = 34
    bteColNoteIssue = 35
    bteColWBSelement = 36
    bteColOrderNo = 37
    bteColProdCode = 38
    bteColFinPlan = 39
    bteColStartDate = 40
    bteColRef1 = 41
    bteColRef2 = 42
    bteColRef3 = 43
    bteColCentBank = 44
    bteColsupplCountry = 45
    bteColHouseBank = 46
    bteColAssetVal = 47
    bteColQty = 48
    bteColUnitQty = 49
    bteColSegment = 50
    bteColSortKey = 51
    bteColItemText = 52
    bteColTransType = 53
    bteColBussPart = 54
  
    With grid
        .clear
        
        .Rows = 1
        .ColS = 55
        
        '.TextMatrix(0, bteColSelect) = ""
        .Cell(flexcpChecked, 0, bteColSelect) = flexUnchecked
        .TextMatrix(0, bteColPosDate) = "Posting Date"
        .TextMatrix(0, bteColDocType) = "Doc. Type"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColInvNo) = "Invoice No"
        .TextMatrix(0, bteColPostKey) = "Posting Key"
        .TextMatrix(0, bteColAccNo) = "Account No"
        .TextMatrix(0, bteColAmount) = "Document Amount"
        .TextMatrix(0, bteColTaxCode) = "Tax Code"
        .TextMatrix(0, bteColCostCenter) = "Cost Center"
        .TextMatrix(0, bteColDocHeader) = "Refence No."
        .TextMatrix(0, bteColItemText) = "Item Line Text"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColPosDate) = 1500
        .ColWidth(bteColDocType) = 1000
        .ColWidth(bteColCurr) = 850
        .ColWidth(bteColInvNo) = 2500
        .ColWidth(bteColPostKey) = 1250
        .ColWidth(bteColAccNo) = 2000
        .ColWidth(bteColAmount) = 2000
        .ColWidth(bteColTaxCode) = 1000
        .ColWidth(bteColCostCenter) = 1800
        .ColWidth(bteColDocHeader) = 1800
        .ColWidth(bteColItemText) = 3000
        
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColPosDate) = flexAlignCenterCenter
        .ColAlignment(bteColDocType) = flexAlignCenterCenter
        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        .ColAlignment(bteColInvNo) = flexAlignLeftCenter
        .ColAlignment(bteColPostKey) = flexAlignCenterCenter
        .ColAlignment(bteColAccNo) = flexAlignCenterCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
        .ColAlignment(bteColTaxCode) = flexAlignCenterCenter
        .ColAlignment(bteColCostCenter) = flexAlignCenterCenter

        For X = 0 To 54
            .ColHidden(X) = True
        Next
        
        .ColHidden(bteColSelect) = False
        .ColHidden(bteColPosDate) = False
        .ColHidden(bteColDocType) = False
        .ColHidden(bteColCurr) = False
        .ColHidden(bteColInvNo) = False
        .ColHidden(bteColPostKey) = False
        .ColHidden(bteColAccNo) = False
        .ColHidden(bteColAmount) = False
        .ColHidden(bteColTaxCode) = False
        .ColHidden(bteColCostCenter) = False
        .ColHidden(bteColDocHeader) = False
        .ColHidden(bteColItemText) = False
        
        .EditMaxLength = 1
    End With

End Sub

Function fc_WriteIniFile(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    fc_WriteIniFile = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

Private Sub CboSupplier_Change()
    Call cbosupplier_Click
End Sub

Private Sub cbosupplier_Click()
    If cboSupplier.ListIndex < 0 Then
        LblCustomer = ""
    Else
        LblCustomer = cboSupplier.Column(1)
    End If
    Call Header
End Sub

Private Sub Cmd_Save_Click()
    Dim adoStream As ADODB.Stream
    Dim adoStreamOut As ADODB.Stream
    Dim fs
    Dim a
    Dim InvoiceNo As String
    Dim Q As Double
    Dim Y As Double
    Dim XData As Integer
    Dim IFPart As String
    Dim IFPart1 As String
    Dim ListOfData As String
    Dim PbMax As Integer
    Dim strSQL As String
    On Error GoTo ErrExport
    
    LblErrMsg = ""
    
    IFPart = App.path & "\IFData" & "\IF_AP_" & Trim(cboSupplier) & "_" & Format(InvTo, "yyyyMMdd")
    IFPart1 = App.path & "\IFData" & "\"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(IFPart & ".tsv", True)
    
    XData = 1
    
    If grid.Rows <= 1 Then
        LblErrMsg = DisplayMsg("0013")
        Exit Sub
    End If
    
    PbMax = grid.Rows - 1
    PBar.Visible = True
    PBar.Max = PbMax
    InvoiceNo = ""
    
    Do While XData <= grid.Rows - 1
        
        If grid.TextMatrix(XData, 1) = "X" And grid.Cell(flexcpChecked, XData, 0) = flexChecked Then
        
            '**Update Invoice No yang sudah di interface**
                
                InvoiceNo = grid.TextMatrix(XData, bteColInvNo)
                strSQL = "Update InvoiceSupplier_Master" & vbCrLf & _
                        "Set Interface_Cls='1',Interface_Date=Getdate(),Interface_User='" & userLogin & "'" & vbCrLf & _
                        "Where Invoice_No='" & Trim(InvoiceNo) & "'"
                Db.Execute strSQL
            
            '**End Update Invoice No yang sudah di interface**
            
        Q = grid.FindRow(InvoiceNo, 0, bteColInvNo, False)
            Y = 0
            For Y = Q To grid.Rows - 1
                
                If grid.TextMatrix(Y, bteColInvNo) <> InvoiceNo Then
                    XData = Y - 2
                    Exit For
                End If
                
                ListOfData = Trim(grid.TextMatrix(Y, 1)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 2)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 3)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 4)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 5)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 6)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 7)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 8)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 9)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 10)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 11)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 12)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 13)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 14)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 15)) & vbTab
                                    
                ListOfData = ListOfData & _
                                    Trim(grid.TextMatrix(Y, 16)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 17)) & vbTab & _
                                    Trim(CDbl(grid.TextMatrix(Y, 18))) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 19)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 20)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 21)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 22)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 23)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 24)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 25)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 26)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 27)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 28)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 29)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 30)) & vbTab
                                    
                ListOfData = ListOfData & _
                                    Trim(grid.TextMatrix(Y, 31)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 32)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 33)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 34)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 35)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 36)) & vbTab & _
                                    Trim(grid.TextMatrix(Y, 37)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 38)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 39)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 40)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 41)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 42)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 43)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 44)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 45)) & vbTab
                                    
                ListOfData = ListOfData & _
                                    RTrim(grid.TextMatrix(Y, 46)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 47)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 48)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 49)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 50)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 51)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 52)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 53)) & vbTab & _
                                    RTrim(grid.TextMatrix(Y, 54))
                                                                        
                a.WriteLine (ListOfData)

            Next Y
            
        End If
        
        PBar.Value = XData
        
        XData = XData + 1
         

    Loop
    
    a.Close
    
    Set adoStream = New ADODB.Stream

    adoStream.Charset = "ASCII"
    adoStream.Open
    adoStream.LoadFromFile IFPart & ".tsv"
    
    adoStream.Position = 0
    Set adoStreamOut = New ADODB.Stream
    adoStreamOut.Charset = "UTF-8"
    adoStreamOut.Open
    adoStreamOut.WriteText adoStream.ReadText
    adoStreamOut.SaveToFile IFPart1 & "520FID03.TXT", adSaveCreateOverWrite
 
    Kill (IFPart & ".tsv")
    
    PBar.Visible = False
    
    
    
    LocalPath = IFPart1 & "520FID03.TXT"
    
    Shell IFPart1 & "FTPuploadAP.bat", vbHide
    Sleep 350
    LblErrMsg = "Export A/P Data Success !"
    
    'UploadFTP
    'LblErrMsg = "Export AP Data Success !"
    Exit Sub

ErrExport:
    LblErrMsg = "[" & err.number & "]-" & err.Description
    err.clear
    
End Sub

Private Sub cmdSearch_Click()

    Dim RsSearch As New ADODB.Recordset
    Dim StrSearch As String
    Dim CData As Integer
    Dim XData As Integer
    Dim JmlTran As Integer

    LblErrMsg = ""
    
    'On Error GoTo ErrSearch
    
    
    If cboSupplier.MatchFound = False Then
        LblErrMsg.Caption = DisplayMsg("4050")
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Call Header
        
'
'StrSearch = "    DECLARE @StartDate DATETIME    " & vbCrLf & _
'                        "    DECLARE @EndDate DATETIME    " & vbCrLf & _
'                        "    DECLARE  @Supplier VARCHAR(25)   " & vbCrLf & _
'                        "        " & vbCrLf & _
'                        "  SET @StartDate = '" & Format(InvFrom, "yyyy-MM-dd") & "'  " & vbCrLf & _
'                        "  SET @endDate = '" & Format(InvTo, "yyyy-MM-dd") & "'  " & vbCrLf & _
'                        "  SET @Supplier = '" & Trim(cboSupplier) & "'  " & vbCrLf & _
'                        "        " & vbCrLf & _
'                        "    -- ############################    " & vbCrLf & _
'                        "    -- AP    " & vbCrLf & _
'                        "    -- ############################         " & vbCrLf & _
'                        " SELECT   " & vbCrLf & _
'                        " JournalType ='AP',     " & vbCrLf & _
'                        "   TransCode ,     "
'
'StrSearch = StrSearch + "   CASE TransCode     " & vbCrLf & _
'                        "   WHEN 'AP' THEN 'X'     " & vbCrLf & _
'                        "   ELSE ''     " & vbCrLf & _
'                        "   END PostingControl ,     " & vbCrLf & _
'                        "   '' Ledger ,     " & vbCrLf & _
'                        "   'C520' Company_Code ,     " & vbCrLf & _
'                        "   '' Calculate_Tax_Auto ,     " & vbCrLf & _
'                        "   Invoice.PostingDate Document_Date ,                " & vbCrLf & _
'                        "   Invoice.PostingDate Posting_Date , 'I2' Document_Tpye ,     " & vbCrLf & _
'                        "   COALESCE(CM.SAP_Curr, '') Currency ,     " & vbCrLf & _
'                        "   '' Exchenge_Rate ,     "
'
'StrSearch = StrSearch + "   coalesce(Voucher_no,'') /*'AP Doc.'*/ Reference ,     " & vbCrLf & _
'                        "   Invoice.Invoice_No Doc_Header ,     " & vbCrLf & _
'                        "   '' Tradding_PartNer ,     " & vbCrLf & _
'                        "   Posting_Key, Account_No, " & vbCrLf & _
'                        "   '' GL_Indicator ,     " & vbCrLf & _
'                        "   '' AssetTran , '' RecAccount ,     " & vbCrLf & _
'                        "   CASE WHEN TransCode = 'AP' THEN SalesAmount + TaxAmount     " & vbCrLf & _
'                        "        WHEN TransCode = 'PC' THEN SalesAmount     " & vbCrLf & _
'                        "        WHEN TransCode = 'PT' THEN TaxAmount     " & vbCrLf & _
'                        "   END Amount, '' Local_Amount, '' Group_Currency, '' Global_Curr_Comp,     " & vbCrLf & _
'                        "   Tax_Code, '' TaxType,  "
'
'StrSearch = StrSearch + "   case when transCode <>'AP' then  " & vbCrLf & _
'                        "       case when rtrim(adm_group) ='998' then  " & vbCrLf & _
'                        "               invoice.cost_center  " & vbCrLf & _
'                        "           WHEN rtrim(adm_group) ='997'  " & vbCrLf & _
'                        "               then invoice.cost_Center  " & vbCrLf & _
'                        "       else Cost_Center  " & vbCrLf & _
'                        "       END  " & vbCrLf & _
'                        "   end Cost_Center,  " & vbCrLf & _
'                        "   Profit_Center, '' Buss_Area, '' Trad_Partner,     " & vbCrLf & _
'                        "   '' Partner_Profit, '' PayTerm, '' FSITE, '' PayBlock, '' PayRef, '' PayMeth, '' Baseline,     " & vbCrLf & _
'                        "   '' Note_IssueDate, '' WBS_elemen, '' OrderNo, '' ProductCode, '' Fin_Plan, '' StartDate,                '' Ref1, '' Ref2, '' Ref3, '' CentralBank, '' Suppl_Country, '' HouseBank, '' Assets_Val,     "
'
'StrSearch = StrSearch + "   '' Qty, '' UnitOfQty, '' Segment, '' SortKey, CASE WHEN TransCode ='AP' then RTRIM(TM.Trade_Name) ELSE REPLACE(coalesce(voucher_description,''),char(13),'') END /*'AP Doc.'*/ Item_Text, '' Trans_Type, '' Buss_Partner,Interface_Cls     " & vbCrLf & _
'                        "  " & vbCrLf & _
'                        " FROM " & vbCrLf & _
'                        "       ( " & vbCrLf & _
'                        "       SELECT Transcode,Supplier_Code, Invoice_no, postingDate, Currency, SalesAmount = Sum(salesAmount), TaxAmount, PphAmount, " & vbCrLf & _
'                        "           Interface_Cls, Voucher_No, Voucher_description, Adm_Group, " & vbCrLf & _
'                        "           Account_No ,posting_key, tax_Code, Cost_Center ='', Profit_Center " & vbCrLf & _
'                        "       FROM ( " & vbCrLf & _
'                        "           SELECT  TransCode ='AP',InvM.Supplier_Code ,     " & vbCrLf & _
'                        "                   InvD.Invoice_No ,                                    " & vbCrLf & _
'                        "                   CONVERT(VARCHAR(8), InvM.Invoice_Date, 112) PostingDate ,     "
'
'StrSearch = StrSearch + "                   COALESCE(InvD.Currency_Code, '') Currency ,             " & vbCrLf & _
'                        "                   invd.Amount SalesAmount,          " & vbCrLf & _
'                        "                   TaxAmount=0, " & vbCrLf & _
'                        "                   PPHAmount=0, " & vbCrLf & _
'                        "                   Interface_Cls, " & vbCrLf & _
'                        "                   PR.Warehouse_Code, " & vbCrLf & _
'                        "                   TRM.Country_Cls,voucher_no, voucher_description, adm_group,  " & vbCrLf & _
'                        "                   cost_center = coalesce(wm.cost_center,''), " & vbCrLf & _
'                        "                   OtherCls = coalesce(PM.Others_Cls,'0'), " & vbCrLf & _
'                        "                   CASE WHEN TransCode = 'AP' And invm.Supplier_Code='C001' THEN  'D10001'  " & vbCrLf & _
'                        "                       When TransCode = 'AP' And invm.Supplier_Code='C066' THEN 'D76100'  "
'
'StrSearch = StrSearch + "                       When TransCode = 'AP' And invm.Supplier_Code='C163' THEN coalesce(Account_No,'') + 'P110'  " & vbCrLf & _
'                        "                       When TransCode = 'AP' And invm.Supplier_Code='C169' THEN coalesce(Account_No,'')  + 'P097'  " & vbCrLf & _
'                        "                       When TransCode = 'AP' And invm.Supplier_Code='C165' THEN coalesce(Account_No,'')  + 'P029'  " & vbCrLf & _
'                        "                       When TransCode = 'AP' And invm.Supplier_Code='C160' THEN coalesce(Account_No,'')  + 'P090'  " & vbCrLf & _
'                        "                       When TransCode = 'AP' And invm.Supplier_Code='C172' THEN coalesce(Account_No,'')  + 'P042'  " & vbCrLf & _
'                        "                       When TransCode = 'AP' And invm.Supplier_Code='C098' THEN coalesce(Account_No,'')  + 'P020'  " & vbCrLf & _
'                        "                       When TransCode = 'AP' And invm.Supplier_Code='C164' THEN coalesce(Account_No,'')  + 'P207'  " & vbCrLf & _
'                        "                       When TransCode = 'AP' THEN  " & vbCrLf & _
'                        "                       coalesce(Li.Account_No,'') + invm.Supplier_Code  " & vbCrLf & _
'                        "                   END Account_No, " & vbCrLf & _
'                        "                   li.posting_key, li.tax_Code, Profit_Center "
'
'StrSearch = StrSearch + "           FROM    InvoiceSupplier_detail InvD     " & vbCrLf & _
'                        "                   LEFT JOIN Invoicesupplier_Master InvM ON InvD.Invoice_No = InvM.Invoice_No                                     " & vbCrLf & _
'                        "                   Left Join Trade_Master TRM On InvM.Supplier_Code=TRM.Trade_Code          " & vbCrLf & _
'                        "                   LEFT JOIN (SELECT PO_No, Others_Cls from PurchaseOrder_Master) PM on PM.PO_no =InvD.PO_No and Others_Cls ='1'  " & vbCrLf & _
'                        "                   LEFT JOIN part_Receipt PR on INvD.receiptSeq_no = pR.Seq_no " & vbCrLf & _
'                        "                   LEFT JOIN WareHouse_Master wm on Pr.Warehouse_Code = wm.WH_Code  " & vbCrLf & _
'                        "                   LEFT JOIN  " & vbCrLf & _
'                        "                   ( " & vbCrLf & _
'                        "                   Select B.Supplier_Code,C.*,1 orderBy  " & vbCrLf & _
'                        "                   From (  " & vbCrLf & _
'                        "                       Select Coalesce(A.Trade_Code,'0000')Link,Supplier_Code  "
'
'StrSearch = StrSearch + "                       From  " & vbCrLf & _
'                        "                           ( " & vbCrLf & _
'                        "                           Select Distinct Supplier_Code From InvoiceSupplier_Master Where Invoice_Date between  @StartDate and @EndDate  " & vbCrLf & _
'                        "                           )AP Left Join   " & vbCrLf & _
'                        "                           (Select * From LinkInterface Where JournalType='AP' And TransCode='AP')A On ap.Supplier_Code=a.Trade_Code  " & vbCrLf & _
'                        "                       )B  " & vbCrLf & _
'                        "                       Left JOin (Select * From LinkInterface Where JournalType='AP' And TransCode='AP')C On B.Link=C.Trade_Code  " & vbCrLf & _
'                        "                   ) LI on li.supplier_code  = Invm.Supplier_Code  " & vbCrLf & _
'                        "                                " & vbCrLf & _
'                        "           WHERE   Invoice_Date BETWEEN @StartDate     " & vbCrLf & _
'                        "                                AND     @EndDate           "
'
'If cboSupplier.Text <> "ALL" Then
'    StrSearch = StrSearch + " and InvM.supplier_code ='" & cboSupplier.Text & "' "
'End If
'StrSearch = StrSearch + "               " & vbCrLf & _
'                        "       ) PT  " & vbCrLf & _
'                        "       GROUP BY Transcode,Supplier_Code, Invoice_no, postingDate, Currency, TaxAmount, PphAmount, " & vbCrLf & _
'                        "           Interface_Cls, COuntry_Cls,  " & vbCrLf & _
'                        "           Voucher_No, Voucher_description, Adm_Group,posting_key,Account_No,tax_Code,Profit_Center " & vbCrLf & _
'                        "            " & vbCrLf & _
'                        "       UNION " & vbCrLf & _
'                        "  " & vbCrLf & _
'                        "       SELECT Transcode,Supplier_Code, Invoice_no, postingDate, Currency, SalesAmount = Sum(salesAmount), TaxAmount, PphAmount, " & vbCrLf & _
'                        "           Interface_Cls, Voucher_No, Voucher_description, Adm_Group, " & vbCrLf & _
'                        "            Account_No,posting_key, tax_Code, cost_center,Profit_Center "
'
'StrSearch = StrSearch + "       FROM ( " & vbCrLf & _
'                        "           SELECT  TransCode ='PC',InvM.Supplier_Code ,     " & vbCrLf & _
'                        "                   InvD.Invoice_No ,                                    " & vbCrLf & _
'                        "                   CONVERT(VARCHAR(8), InvM.Invoice_Date, 112) PostingDate ,     " & vbCrLf & _
'                        "                   COALESCE(InvD.Currency_Code, '') Currency ,             " & vbCrLf & _
'                        "                   invd.Amount SalesAmount,          " & vbCrLf & _
'                        "                   TaxAmount=0, " & vbCrLf & _
'                        "                   PPHAmount=0, " & vbCrLf & _
'                        "                   Interface_Cls, " & vbCrLf & _
'                        "                   PR.Warehouse_Code, " & vbCrLf & _
'                        "                   TRM.Country_Cls,voucher_no, voucher_description, adm_group,  "
'
'StrSearch = StrSearch + "                   cost_center =  case when wm.Adm_Group='998' or  wm.Adm_Group='997' then wm.cost_center else coalesce(li.cost_center,'') end, " & vbCrLf & _
'                        "                   OtherCls = coalesce(PM.Others_Cls,'0'), " & vbCrLf & _
'                        "                   Case  " & vbCrLf & _
'                        "                       When Country_Cls='0'  and Others_Cls ='1' Then '82350000'  " & vbCrLf & _
'                        "                       When Country_Cls='1' and Others_Cls ='1' Then '82350080'  " & vbCrLf & _
'                        "                       When Warehouse_Code='W011' Then '84220000'  " & vbCrLf & _
'                        "                       When Warehouse_Code='W0112' Then '84220080'  " & vbCrLf & _
'                        "                       When Warehouse_Code='W038' Then '84220080'  " & vbCrLf & _
'                        "                       When Warehouse_Code='W0382' Then '84220080'  " & vbCrLf & _
'                        "                       When Warehouse_Code='WH015' Then '84220000' /*'84220080'*/  " & vbCrLf & _
'                        "                       When INVm.Supplier_Code='C001' Then '82311180'  "
'
'StrSearch = StrSearch + "                       When Country_Cls='0' Then  " & vbCrLf & _
'                        "                           case when coalesce(li.account_no,'') ='' then '82311100'  " & vbCrLf & _
'                        "                           Else  " & vbCrLf & _
'                        "                            '82311100' /*li.account_no*/  " & vbCrLf & _
'                        "                           End  " & vbCrLf & _
'                        "                       when Country_Cls='1' Then  " & vbCrLf & _
'                        "                           case when coalesce(li.account_no,'') ='' then '82311172'  " & vbCrLf & _
'                        "                           else  " & vbCrLf & _
'                        "                               li.account_no  " & vbCrLf & _
'                        "                           End  " & vbCrLf & _
'                        "                       WHEN TransCode = 'PC' THEN Li.Account_No  "
'
'StrSearch = StrSearch + "                   END Account_No, " & vbCrLf & _
'                        "                   li.posting_key, li.tax_Code,Profit_Center " & vbCrLf & _
'                        "           FROM    InvoiceSupplier_detail InvD     " & vbCrLf & _
'                        "                   LEFT JOIN Invoicesupplier_Master InvM ON InvD.Invoice_No = InvM.Invoice_No                                     " & vbCrLf & _
'                        "                   Left Join Trade_Master TRM On InvM.Supplier_Code=TRM.Trade_Code          " & vbCrLf & _
'                        "                   LEFT JOIN (SELECT PO_No, Others_Cls from PurchaseOrder_Master) PM on PM.PO_no =InvD.PO_No and Others_Cls ='1'  " & vbCrLf & _
'                        "                   LEFT JOIN part_Receipt PR on INvD.receiptSeq_no = pR.Seq_no " & vbCrLf & _
'                        "                   LEFT JOIN WareHouse_Master wm on Pr.Warehouse_Code = wm.WH_Code  " & vbCrLf & _
'                        "                   LEFT JOIN  " & vbCrLf & _
'                        "                   ( " & vbCrLf & _
'                        "                   Select B.Supplier_Code,C.*,1 orderBy  "
'
'StrSearch = StrSearch + "                   From (  " & vbCrLf & _
'                        "                       Select Coalesce(A.Trade_Code,'0000')Link,Supplier_Code  " & vbCrLf & _
'                        "                       From  " & vbCrLf & _
'                        "                           ( " & vbCrLf & _
'                        "                           Select Distinct Supplier_Code From InvoiceSupplier_Master Where Invoice_Date between  @StartDate and @EndDate  " & vbCrLf & _
'                        "                           )AP Left Join   " & vbCrLf & _
'                        "                           (Select * From LinkInterface Where JournalType='AP' And TransCode='PC')A On ap.Supplier_Code=a.Trade_Code  " & vbCrLf & _
'                        "                       )B  " & vbCrLf & _
'                        "                       Left JOin (Select * From LinkInterface Where JournalType='AP' And TransCode='PC')C On B.Link=C.Trade_Code  " & vbCrLf & _
'                        "                   ) LI on li.supplier_code  = Invm.Supplier_Code  " & vbCrLf & _
'                        "                   --Select * From LinkInterface Where JournalType='AP' And TransCode='PC') LI on "
'
'StrSearch = StrSearch + "                   --  case when LI.Trade_code  ='0000' then Invm.Supplier_Code else li.trade_code end = Invm.Supplier_Code " & vbCrLf & _
'                        "           WHERE   Invoice_Date BETWEEN @StartDate     " & vbCrLf & _
'                        "                                AND     @EndDate           " & vbCrLf
'
'If cboSupplier.Text <> "ALL" Then
'    StrSearch = StrSearch + " and InvM.supplier_code ='" & cboSupplier.Text & "' "
'End If
'StrSearch = StrSearch + "               " & vbCrLf & _
'                        "       ) PT  " & vbCrLf & _
'                        "       GROUP BY Transcode,Supplier_Code, Invoice_no, postingDate, Currency, TaxAmount, PphAmount, " & vbCrLf & _
'                        "           Interface_Cls,  " & vbCrLf & _
'                        "           Warehouse_Code, COuntry_Cls,  " & vbCrLf & _
'                        "           Voucher_No, Voucher_description, Adm_Group,posting_key,Account_No,tax_Code,cost_center,Profit_Center " & vbCrLf & _
'                        "       ) Invoice " & vbCrLf & _
'                        "         LEFT JOIN Curr_Mapping CM ON CM.EZRCurr = Invoice.Currency     "
'
'StrSearch = StrSearch + "         LEFT JOIN Trade_master TM on invoice.supplier_Code = TM.trade_code  " & vbCrLf & _
'                        " WHERE     SalesAmount > 0 " & vbCrLf & _
'                        IIf(cbointeface.ListIndex = 0, " and Coalesce(Interface_Cls,'0')='1' ", "and Coalesce(Interface_Cls,'0')='0'") & vbCrLf & _
'                        IIf(cboSupplier.ListIndex = 0, "And Supplier_Code NOT IN " & strExcludeCustSup & " ", " And Supplier_Code='" & Trim(cboSupplier.Text) & "' ") & vbCrLf & _
'                        " AND  Invoice.Supplier_Code<>'C020' " & vbCrLf & _
'                        " ORDER BY Invoice.Invoice_No, transCode "
                        
                        
                        
'===================================================NEW Query=================================================
                        
                        
StrSearch = "     DECLARE @StartDate DATETIME     " & vbCrLf & _
                        "     DECLARE @EndDate DATETIME     " & vbCrLf & _
                        "     DECLARE  @Supplier VARCHAR(25)    " & vbCrLf & _
                        "          " & vbCrLf & _
                        "  SET @StartDate = '" & Format(InvFrom, "yyyy-MM-dd") & "'  " & vbCrLf & _
                        "  SET @endDate = '" & Format(InvTo, "yyyy-MM-dd") & "'  " & vbCrLf & _
                        "  SET @Supplier = '" & Trim(cboSupplier) & "'  " & vbCrLf & _
                        "          " & vbCrLf & _
                        "     -- ############################     " & vbCrLf & _
                        "     -- AP     " & vbCrLf & _
                        "     -- ############################          " & vbCrLf

StrSearch = StrSearch + "  SELECT    " & vbCrLf & _
                        "  JournalType ='AP',      " & vbCrLf & _
                        "    TransCode ,        CASE TransCode      " & vbCrLf & _
                        "    WHEN 'AP' THEN 'X'      " & vbCrLf & _
                        "    ELSE ''      " & vbCrLf & _
                        "    END PostingControl ,      " & vbCrLf & _
                        "    '' Ledger ,      " & vbCrLf & _
                        "    'C520' Company_Code ,      " & vbCrLf & _
                        "    '' Calculate_Tax_Auto ,      " & vbCrLf & _
                        "    Invoice.PostingDate Document_Date ,                 " & vbCrLf & _
                        "    Invoice.PostingDate Posting_Date , 'I5' Document_Tpye ,      " & vbCrLf

StrSearch = StrSearch + "    (select [Description] from Curr_Cls where Curr_Cls= Currency) Currency ,      " & vbCrLf & _
                        "    '' Exchenge_Rate ,         " & vbCrLf & _
                        "    coalesce(PaymentVoucher_No,'') /*'AP Doc.'*/ Reference ,      " & vbCrLf & _
                        "    Invoice.Invoice_No Doc_Header ,      " & vbCrLf & _
                        "    '' Tradding_PartNer ,      " & vbCrLf & _
                        "    Posting_Key, Account_No,  " & vbCrLf & _
                        "    '' GL_Indicator ,      " & vbCrLf & _
                        "    '' AssetTran , '' RecAccount ,      " & vbCrLf & _
                        "    CASE WHEN TransCode = 'AP' THEN SalesAmount + TaxAmount      " & vbCrLf & _
                        "         WHEN TransCode = 'PC' THEN SalesAmount      " & vbCrLf & _
                        "         WHEN TransCode = 'PT' THEN TaxAmount      " & vbCrLf

StrSearch = StrSearch + "    END Amount, '' Local_Amount, '' Group_Currency, '' Global_Curr_Comp,      " & vbCrLf & _
                        "    Tax_Code, '' TaxType,     case when transCode <>'AP' then   " & vbCrLf & _
                        "        case when rtrim(adm_group) ='998' then   " & vbCrLf & _
                        "                invoice.cost_center   " & vbCrLf & _
                        "            WHEN rtrim(adm_group) ='997'   " & vbCrLf & _
                        "                then invoice.cost_Center   " & vbCrLf & _
                        "        else Cost_Center   " & vbCrLf & _
                        "        END   " & vbCrLf & _
                        "    end Cost_Center,   " & vbCrLf & _
                        "    Profit_Center, '' Buss_Area, '' Trad_Partner,      " & vbCrLf & _
                        "    '' Partner_Profit, '' PayTerm, '' FSITE, '' PayBlock, '' PayRef, '' PayMeth, '' Baseline,      " & vbCrLf

StrSearch = StrSearch + "    '' Note_IssueDate, '' WBS_elemen, '' OrderNo, '' ProductCode, '' Fin_Plan, '' StartDate,                '' Ref1, '' Ref2, '' Ref3, '' CentralBank, '' Suppl_Country, '' HouseBank, '' Assets_Val,        '' Qty, '' UnitOfQty, '' Segment, '' SortKey, CASE WHEN TransCode ='AP' then RTRIM(TM.Trade_Name) ELSE REPLACE(coalesce(VoucherDesc,''),char(13),'') END /*'AP Doc.'*/ Item_Text, '' Trans_Type, '' Buss_Partner,Interface_Cls      " & vbCrLf & _
                        "    " & vbCrLf & _
                        "  FROM  " & vbCrLf & _
                        "        (  " & vbCrLf & _
                        "        SELECT Transcode,Supplier_Code, Invoice_no, postingDate, Currency, SalesAmount = Sum(salesAmount), TaxAmount, PphAmount,  " & vbCrLf & _
                        "            Interface_Cls, PaymentVoucher_No, VoucherDesc, Adm_Group,  " & vbCrLf & _
                        "            Account_No ,posting_key, tax_Code, Cost_Center ='', Profit_Center  " & vbCrLf & _
                        "        FROM (  " & vbCrLf & _
                        "            SELECT  TransCode ='AP',Supplier_Code=Coalesce(SAP_CODE,InvM.Supplier_Code) ,      " & vbCrLf & _
                        "                    InvD.Invoice_No ,                                     " & vbCrLf & _
                        "                    CONVERT(VARCHAR(8), InvM.Invoice_Date, 112) PostingDate ,                        COALESCE(InvD.Currency_Code, '') Currency ,              " & vbCrLf

StrSearch = StrSearch + "                    invd.Amount SalesAmount,           " & vbCrLf & _
                        "                    TaxAmount=0,  " & vbCrLf & _
                        "                    PPHAmount=0,  " & vbCrLf & _
                        "                    Interface_Cls,  " & vbCrLf & _
                        "                    PR.Warehouse_Code,  " & vbCrLf & _
                        "                    TRM.Country_Cls, BL_no AS PaymentVoucher_No, VoucherDesc, adm_group,   " & vbCrLf & _
                        "                    cost_center = coalesce(li.cost_center,''),  " & vbCrLf & _
                        "                    OtherCls = coalesce(PM.Others_Cls,'0'),  " & vbCrLf & _
                        "                    CASE WHEN TransCode = 'AP' And TRM.SAP_CODE='C001' THEN  'D10001'   " & vbCrLf & _
                        "                        When TransCode = 'AP' And TRM.SAP_CODE='C066' THEN 'D76100'                         When TransCode = 'AP' And invm.Supplier_Code='C163' THEN coalesce(Account_No,'') + 'P110'   " & vbCrLf & _
                        "                        When TransCode = 'AP' And TRM.SAP_CODE='C169' THEN coalesce(Account_No,'')  + 'P097'   "

StrSearch = StrSearch + "                        When TransCode = 'AP' And TRM.SAP_CODE='C165' THEN coalesce(Account_No,'')  + 'P029'   " & vbCrLf & _
                        "                        When TransCode = 'AP' And TRM.SAP_CODE='C160' THEN coalesce(Account_No,'')  + 'P090'   " & vbCrLf & _
                        "                        When TransCode = 'AP' And TRM.SAP_CODE='C172' THEN coalesce(Account_No,'')  + 'P042'   " & vbCrLf & _
                        "                        When TransCode = 'AP' And TRM.SAP_CODE='C098' THEN coalesce(Account_No,'')  + 'P020'   " & vbCrLf & _
                        "                        When TransCode = 'AP' And TRM.SAP_CODE='C164' THEN coalesce(Account_No,'')  + 'P207'   " & vbCrLf & _
                        "                        When TransCode = 'AP' THEN   " & vbCrLf & _
                        "                        SAP_Code   " & vbCrLf & _
                        "                    END Account_No,  " & vbCrLf & _
                        "                    li.posting_key, case when li.posting_key='31' then '' else li.tax_Code end tax_Code, Profit_Center            FROM    InvoiceSupplier_detail InvD      " & vbCrLf & _
                        "                    LEFT JOIN Invoicesupplier_Master InvM ON InvD.Invoice_No = InvM.Invoice_No                                      " & vbCrLf & _
                        "                    Left Join Trade_Master TRM On InvM.Supplier_Code=TRM.Trade_Code           " & vbCrLf

StrSearch = StrSearch + "                    LEFT JOIN (SELECT PO_No, Others_Cls from PurchaseOrder_Master) PM on PM.PO_no =InvD.PO_No and Others_Cls ='1'   " & vbCrLf & _
                        "                    INNER JOIN part_Receipt PR on INvD.receiptSeq_no = pR.Seq_no  " & vbCrLf & _
                        "                    LEFT JOIN WareHouse_Master wm on Pr.Warehouse_Code = wm.WH_Code   " & vbCrLf & _
                        "                    LEFT JOIN   " & vbCrLf & _
                        "                    (  " & vbCrLf & _
                        "                    Select B.Supplier_Code,C.*,1 orderBy   " & vbCrLf & _
                        "                    From (   " & vbCrLf & _
                        "                        Select Coalesce(A.Trade_Code,'0000')Link,Supplier_Code                          " & vbCrLf & _
                        "                      From   " & vbCrLf & _
                        "                            (  " & vbCrLf & _
                        "                            Select Distinct Supplier_Code From InvoiceSupplier_Master Where Invoice_Date between  @StartDate and @EndDate   " & vbCrLf

StrSearch = StrSearch + "                            )AP Left Join    " & vbCrLf & _
                        "                            (Select * From LinkInterface Where JournalType='AP' And TransCode='AP')A On ap.Supplier_Code=a.Trade_Code   " & vbCrLf & _
                        "                        )B   " & vbCrLf & _
                        "                        Left JOin (Select * From LinkInterface Where JournalType='AP' And TransCode='AP')C On B.Link=C.Trade_Code   " & vbCrLf & _
                        "                    ) LI on li.supplier_code  = Invm.Supplier_Code   " & vbCrLf & _
                        "                                  " & vbCrLf & _
                        "            WHERE   Invoice_Date BETWEEN @StartDate      " & vbCrLf & _
                        "                                 AND     @EndDate                           " & vbCrLf
                        
If cboSupplier.Text <> "ALL" Then
    StrSearch = StrSearch + " and InvM.supplier_code ='" & cboSupplier.Text & "' "
End If

StrSearch = StrSearch + "               " & vbCrLf & _
                        "        ) PT   " & vbCrLf & _
                        "        GROUP BY Transcode,Supplier_Code, Invoice_no, postingDate, Currency, TaxAmount, PphAmount,  " & vbCrLf & _
                        "            Interface_Cls, COuntry_Cls,   " & vbCrLf

StrSearch = StrSearch + "            PaymentVoucher_No, VoucherDesc, Adm_Group,posting_key,Account_No,tax_Code,Profit_Center  " & vbCrLf & _
                        "              " & vbCrLf & _
                        "        UNION  " & vbCrLf & _
                        "    " & vbCrLf & _
                        "        SELECT Transcode,Supplier_Code, Invoice_no, postingDate, Currency, SalesAmount = Sum(salesAmount), TaxAmount, PphAmount,  " & vbCrLf & _
                        "            Interface_Cls, PaymentVoucher_No, VoucherDesc, Adm_Group,  " & vbCrLf & _
                        "             Account_No,posting_key,Case When  posting_key='31' then ' ' else tax_Code end tax_Code , cost_center,Profit_Center        FROM (  " & vbCrLf & _
                        "            SELECT  TransCode ='PC',Supplier_Code=Coalesce(SAP_CODE,InvM.Supplier_Code) ,      " & vbCrLf & _
                        "                    InvD.Invoice_No ,                                     " & vbCrLf & _
                        "                    CONVERT(VARCHAR(8), InvM.Invoice_Date, 112) PostingDate ,      " & vbCrLf & _
                        "                    COALESCE(InvD.Currency_Code, '') Currency ,              " & vbCrLf

StrSearch = StrSearch + "                    invd.Amount SalesAmount,           " & vbCrLf & _
                        "                    TaxAmount=0,  " & vbCrLf & _
                        "                    PPHAmount=0,  " & vbCrLf & _
                        "                    Interface_Cls,  " & vbCrLf & _
                        "                    PR.Warehouse_Code,  " & vbCrLf & _
                        "                    TRM.Country_Cls,PaymentVoucher_No, VoucherDesc, adm_group,                      " & vbCrLf & _
                        "                    cost_center = coalesce(li.cost_center,''), " & vbCrLf & _
                        "                    OtherCls = coalesce(PM.Others_Cls,'0'),  " & vbCrLf & _
                        "                    Li.Account_No  Account_No,  " & vbCrLf & _
                        "                    li.posting_key, li.tax_Code,Profit_Center  " & vbCrLf & _
                        "            FROM    InvoiceSupplier_detail InvD      " & vbCrLf

StrSearch = StrSearch + "                    LEFT JOIN Invoicesupplier_Master InvM ON InvD.Invoice_No = InvM.Invoice_No                                      " & vbCrLf & _
                        "                    Left Join Trade_Master TRM On InvM.Supplier_Code=TRM.Trade_Code           " & vbCrLf & _
                        "                    LEFT JOIN (SELECT PO_No, Others_Cls from PurchaseOrder_Master) PM on PM.PO_no =InvD.PO_No and Others_Cls ='1'   " & vbCrLf & _
                        "                    INNER JOIN part_Receipt PR on INvD.receiptSeq_no = pR.Seq_no  " & vbCrLf & _
                        "                    LEFT JOIN WareHouse_Master wm on Pr.Warehouse_Code = wm.WH_Code   " & vbCrLf & _
                        "                    LEFT JOIN   " & vbCrLf & _
                        "                    (  " & vbCrLf & _
                        "                    Select B.Supplier_Code,C.*,1 orderBy                      " & vbCrLf & _
                        "                  From (   " & vbCrLf & _
                        "                        Select Coalesce(A.Trade_Code,'0000')Link,Supplier_Code   " & vbCrLf & _
                        "                        From   " & vbCrLf

StrSearch = StrSearch + "                            (  " & vbCrLf & _
                        "                            Select Distinct Supplier_Code From InvoiceSupplier_Master Where Invoice_Date between  @StartDate and @EndDate   " & vbCrLf & _
                        "                            )AP Left Join    " & vbCrLf & _
                        "                            (Select * From LinkInterface Where JournalType='AP' And TransCode='PC')A On ap.Supplier_Code=a.Trade_Code   " & vbCrLf & _
                        "                        )B   " & vbCrLf & _
                        "                        Left JOin (Select * From LinkInterface Where JournalType='AP' And TransCode='PC')C On B.Link=C.Trade_Code   " & vbCrLf & _
                        "                    ) LI on li.supplier_code  = Invm.Supplier_Code   " & vbCrLf & _
                        "                     " & vbCrLf & _
                        "            WHERE   Invoice_Date BETWEEN @StartDate      " & vbCrLf & _
                        "                                 AND     @EndDate            " & vbCrLf & _
                        "                 " & vbCrLf

StrSearch = StrSearch + "        ) PT   " & vbCrLf & _
                        "        GROUP BY Transcode,Supplier_Code, Invoice_no, postingDate, Currency, TaxAmount, PphAmount,  " & vbCrLf & _
                        "            Interface_Cls,   " & vbCrLf & _
                        "            Warehouse_Code, COuntry_Cls,   " & vbCrLf & _
                        "            PaymentVoucher_No, VoucherDesc, Adm_Group,posting_key,Account_No,tax_Code,cost_center,Profit_Center  " & vbCrLf & _
                        "        ) Invoice  " & vbCrLf & _
                        "          --LEFT JOIN Curr_Mapping CM ON CM.EZRCurr = Invoice.Currency               " & vbCrLf & _
                        "        LEFT JOIN Trade_master TM on invoice.supplier_Code = TM.trade_code   " & vbCrLf & _
                        "  WHERE     SalesAmount > 0  " & vbCrLf & _
                        IIf(cbointeface.ListIndex = 0, " and Coalesce(Interface_Cls,'0')='1' ", "and Coalesce(Interface_Cls,'0')='0'") & vbCrLf & _
                        IIf(cboSupplier.ListIndex = 0, "And Supplier_Code NOT IN " & strExcludeCustSup & " ", " And Supplier_Code='" & Trim(cboSupplier.Text) & "' ") & vbCrLf & _
                        "  AND  Invoice.Supplier_Code<>'C020'  "

StrSearch = StrSearch + "  ORDER BY Invoice.Invoice_No, transCode  " & vbCrLf & _
                        "  "
                        


    If RsSearch.State <> adStateClosed Then RsSearch.Close
    
    Set RsSearch = Db.Execute(StrSearch)
    
    If RsSearch.EOF Then
        LblErrMsg = DisplayMsg("0013")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    XData = 0
    JmlTran = 0
    
    Do While Not RsSearch.EOF
        grid.AddItem ""
        
        If RsSearch.Fields(2) = "X" Then
        
            If RsSearch("Interface_Cls") = "1" Then
                grid.Cell(flexcpChecked, grid.Rows - 1, 0) = Checked
            Else
                grid.Cell(flexcpChecked, grid.Rows - 1, 0) = flexUnchecked
            End If
            
            JmlTran = JmlTran + 1
        End If
        
        For XData = 2 To RsSearch.Fields.Count - 4
            If XData - 1 = bteColAmount Then
                grid.TextMatrix(grid.Rows - 1, XData - 1) = Format(RsSearch.Fields(XData), "#,##0.00")
            Else
                grid.TextMatrix(grid.Rows - 1, XData - 1) = Trim(RsSearch.Fields(XData)) & ""
        End If
        
        Next XData
        RsSearch.MoveNext
    Loop
    
    LblRecord = Format(JmlTran, "#,##0") & " Record(s)"
    
    Me.MousePointer = vbDefault
    Exit Sub
    
ErrSearch:

    LblErrMsg = "[" & err.number & "]-" & err.Description
    err.clear
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    Call Kosong
    Call Header
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
'    With Anchor1
'      .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
'      .DoInit
'    End With
End Sub

Sub Kosong()
    Dim RsSuppl As New ADODB.Recordset
    Dim strSQL As String
    Dim X As Integer
    
    strSQL = "Select Trade_Code Supplier_Code, Trade_Name supplier_Name" & vbCrLf & _
                  " From Trade_Master " & vbCrLf & _
                  "        WHERE Trade_Cls=2 AND left(trade_Code,1)='S'--AND TradeExternal_Cls IN ('1','2') AND Trade_Code NOT IN " & strExcludeCustSup & "  " & vbCrLf & _
                  "            ORDER BY Trade_Code " & vbCrLf
                
    If RsSuppl.State <> adStateClosed Then RsSuppl.Close
    
    Set RsSuppl = Db.Execute(strSQL)
    
    
    cboSupplier.clear
    cboSupplier.ListWidth = 350
    cboSupplier.columnCount = 2
    cboSupplier.ColumnWidths = "100 pt;250 pt"
    
    cboSupplier.AddItem ""
    cboSupplier.List(0, 0) = "ALL"
    cboSupplier.List(0, 1) = "ALL"
    
    X = 1
    Do While Not RsSuppl.EOF
        cboSupplier.AddItem ""
        cboSupplier.List(X, 0) = Trim(RsSuppl("Supplier_Code") & "")
        cboSupplier.List(X, 1) = Trim(RsSuppl("supplier_Name") & "")
        RsSuppl.MoveNext
        X = X + 1
    Loop
    
    cboSupplier.ListIndex = 0
    
    InvFrom = Format(Now(), "yyyy-MMM-") & "01"
    InvTo = DateAdd("m", 1, InvFrom) - 1


    With cbointeface
        .clear
        .AddItem "Yes"
        .AddItem "No"
        
        .ListIndex = 1
    End With
    
End Sub


Private Sub UploadFTP()
Dim host_name As String
Dim RsFtp As New ADODB.Recordset
Dim strSQL As String
Dim Host As String
Dim user As String
Dim Pwd As String
Dim Folder As String

    Enabled = False
    MousePointer = vbHourglass
    
    LblErrMsg.Caption = "Working"
    'txtResults.SelStart = Len(txtResults.Text)
    
    strSQL = "Select * From ftp_Setting"
    If RsFtp.State <> adStateClosed Then RS.Close
    
    Set RsFtp = Db.Execute(strSQL)
    
    If Not RsFtp.EOF Then
    
        Host = IIf(IsNull(RsFtp("Host1")), "", Trim(RsFtp("host1")))
        user = IIf(IsNull(RsFtp("user1")), "", Trim(RsFtp("user1")))
        Pwd = IIf(IsNull(RsFtp("pwd1")), "", Trim(RsFtp("pwd1")))
        Folder = IIf(IsNull(RsFtp("Folder1")), "", Trim(RsFtp("Folder1")))

    End If

    
        
    DoEvents

    ' You must set the URL before the user name and
    ' password. Otherwise the control cannot verify
    ' the user name and password and you get the error:
    '
    '       Unable to connect to remote host
    host_name = Host
    If LCase$(Left$(host_name, 6)) <> "ftp://" Then _
        host_name = "ftp://" & host_name
    Inetftp.URL = host_name

    Inetftp.userName = user
    Inetftp.Password = Pwd & "ss"
    Folder = Folder & "520FID01.TXT"
    ' Do not include the host name here. That will make
    ' the control try to use its default user name and
    ' password and you'll get the error again.
'    inetFTP.Execute , "Put " & _
'        txtLocalFile.Text & " " & txtRemoteFile.Text
        
'    Inetftp.Execute , "Put " & _
        LocalPath & " " & Folder


'    m_GettingDir = True
    Inetftp.Execute , "open"

End Sub

Private Sub Cmd_SubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Sub Browse()

End Sub

Private Sub cmd_clear_Click()
    Kosong
    Header
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim i As Long
Dim Q As Double
If Col = 0 And Row = 0 Then

    For Q = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, 0, 0) = flexChecked And grid.TextMatrix(Q, 1) = "X" Then
           grid.Cell(flexcpChecked, Q, 0) = flexChecked
        ElseIf grid.Cell(flexcpChecked, 0, 0) = flexUnchecked And grid.TextMatrix(Q, 1) = "X" Then
            grid.Cell(flexcpChecked, Q, 0) = flexUnchecked
        End If
    Next Q
    
End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 If grid.Col <> bteColSelect Then Cancel = True
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
  If grid.Col = bteColSelect Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
  End If
End Sub

Private Sub Inetftp_StateChanged(ByVal State As Integer)
Select Case State
        Case icError
            LblErrMsg.Caption = "Error: " & _
                "    " & Inetftp.ResponseCode & vbCrLf & _
                "    " & Inetftp.ResponseInfo
        Case icNone
            LblErrMsg.Caption = "None"
        Case icConnecting
            LblErrMsg.Caption = "Connecting"
        Case icConnected
            LblErrMsg.Caption = "Connected"
        Case icDisconnecting
            LblErrMsg.Caption = "Disconnecting"
        Case icDisconnected
            LblErrMsg.Caption = "Disconnected"
        Case icRequestSent
            LblErrMsg.Caption = "Request Sent"
        Case icRequesting
            LblErrMsg.Caption = "Requesting"
        Case icReceivingResponse
            LblErrMsg.Caption = "Receiving Response"
        Case icRequestSent
            LblErrMsg.Caption = "Request Sent"
        Case icResponseReceived
            LblErrMsg.Caption = "Response Received"
        Case icResolvingHost
            LblErrMsg.Caption = "Resolving Host"
        Case icHostResolved
            LblErrMsg.Caption = "Host Resolved"

        Case icResponseCompleted
            LblErrMsg.Caption = Inetftp.ResponseInfo

            If m_GettingDir Then
                Dim txt As String
                Dim chunk As Variant

                m_GettingDir = False

                ' Get the first chunk.
                chunk = Inetftp.GetChunk(1024, icString)
                DoEvents
                Do While Len(chunk) > 0
                    txt = txt & chunk
                    chunk = Inetftp.GetChunk(1024, icString)
                    DoEvents
                Loop

                LblErrMsg.Caption = "----------"
                LblErrMsg.Caption = txt
            End If

       Case Else
            LblErrMsg.Caption = "State = " & Format$(State)
    End Select

    Enabled = True
    MousePointer = vbDefault
End Sub

Private Sub InvFrom_Change()
    Call Header
End Sub


Private Sub InvTo_Change()
    Call Header
End Sub
