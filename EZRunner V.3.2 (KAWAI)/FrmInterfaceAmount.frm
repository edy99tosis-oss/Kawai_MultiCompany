VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmInterfaceAmount 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Interface Inventory Valuation"
   ClientHeight    =   10140
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15120
   Icon            =   "FrmInterfaceAmount.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10140
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   915
      Left            =   135
      TabIndex        =   10
      Tag             =   "TTTF*/"
      Top             =   705
      Width           =   14760
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   285
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   360
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
         CustomFormat    =   "MMM yyyy"
         Format          =   294191107
         UpDown          =   -1  'True
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DtEnd 
         Height          =   315
         Left            =   5640
         TabIndex        =   14
         Tag             =   "TTFF*/"
         Top             =   360
         Visible         =   0   'False
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
         CustomFormat    =   "MMM yyyy"
         Format          =   294191107
         UpDown          =   -1  'True
         CurrentDate     =   37798
      End
      Begin VB.Label LblStatus 
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
         Left            =   13545
         TabIndex        =   16
         Top             =   555
         Visible         =   0   'False
         Width           =   75
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
         Index           =   0
         Left            =   5280
         TabIndex        =   15
         Tag             =   "TTFF*/"
         Top             =   435
         Visible         =   0   'False
         Width           =   165
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
         Index           =   2
         Left            =   600
         TabIndex        =   13
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   540
      End
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "FFTT*/"
      Top             =   9660
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "TFFT*/"
      Top             =   9660
      Width           =   1125
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "FFTT*/"
      Top             =   9660
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Tag             =   "TFTT*/"
      Top             =   8880
      Width           =   14760
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   375
         Left            =   90
         TabIndex        =   1
         Tag             =   "TTTF*/"
         Top             =   180
         Visible         =   0   'False
         Width           =   14565
         _ExtentX        =   25691
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
         Left            =   210
         TabIndex        =   2
         Tag             =   "TTTF*/"
         Top             =   180
         Width           =   14205
      End
   End
   Begin InetCtlsObjects.Inet Inetftp 
      Left            =   3210
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6870
      Left            =   135
      TabIndex        =   6
      Tag             =   "TTTT*/"
      Top             =   1695
      Width           =   14760
      _cx             =   26035
      _cy             =   12118
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13080
      TabIndex        =   7
      Tag             =   "FTTF*/"
      Top             =   0
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Interface Inventory Valuation"
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
      Left            =   30
      TabIndex        =   9
      Tag             =   "TTTF*/"
      Top             =   120
      Width           =   14610
   End
   Begin MSForms.Label LblRecord 
      Height          =   255
      Left            =   11310
      TabIndex        =   8
      Tag             =   "FFTT*/"
      Top             =   8640
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
End
Attribute VB_Name = "FrmInterfaceAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_GettingDir As Boolean
Dim LocalPath As String

Dim sql As String, StrSearch
Dim RS As New Recordset
Dim ubah As Boolean, hapus As Boolean, gavalid As Boolean, ubahedate As Boolean
Dim SDate, EDate, sdateawal, edateakhir
Dim i As Integer, CData As Integer, XData As Integer, JmlTran As Integer
Dim RsSearch As New ADODB.Recordset

Dim bteColSelect As Byte
Dim bteColCustCode As Byte
Dim bteColJournalType As Byte
Dim bteColTransCode As Byte
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
Dim bteColPostControl As Byte

Sub Header()

    Dim X As Integer
    
    'LblErrMsg = ""
    LblRecord = "0 Record(s)"
    
    bteColPostingCode = 0
    bteColPostControl = 1
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
        
        .TextMatrix(0, bteColPostControl) = "Posting Control"
        .TextMatrix(0, bteColCompCode) = "Company Code"
        .TextMatrix(0, bteColPosDate) = "Posting Date"
        .TextMatrix(0, bteColDocType) = "Doc. Type"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColInvNo) = "Invoice No"
        .TextMatrix(0, bteColPostKey) = "Posting Key"
        .TextMatrix(0, bteColAccNo) = "Account No"
        .TextMatrix(0, bteColAmount) = "Document Amount"
        .TextMatrix(0, bteColTaxCode) = "Tax Code"
        .TextMatrix(0, bteColCostCenter) = "Cost Center"
        .TextMatrix(0, bteColDocHeader) = "Reference No"
        .TextMatrix(0, bteColItemText) = "Item Line Text"
        
        .ColWidth(bteColPostControl) = 300
        .ColWidth(bteColCompCode) = 1500
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
        
        
        .ColAlignment(bteColPostControl) = flexAlignCenterCenter
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

        .ColHidden(bteColPostControl) = True
        .ColHidden(bteColCompCode) = True
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
        .ColHidden(bteColInvNo) = True
        
        .EditMaxLength = 1
    End With

End Sub
'Ini editan
Function fc_WriteIniFile(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    fc_WriteIniFile = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

Private Sub cmd_clear_Click()
    Header
    clear
End Sub

Private Sub Cmd_Save_Click()
    Dim adoStream As ADODB.Stream
    Dim adoStreamOut As ADODB.Stream
    Dim fs
    Dim a
    Dim InvoiceNo As String
    Dim Y As Double
    Dim Q As Double
    
    Dim XData As Integer
    Dim IFPart As String
    Dim IFPart1 As String
    Dim ListOfData As String
    Dim PbMax As Integer
    Dim CLoop As Long
    Dim strSQL As String
    Dim icek As Integer
    
    On Error GoTo ErrExport
    
    LblErrMsg = ""
    
    
    For i = 1 To grid.Rows - 1
    
            strSQL = "SELECT * FROM InterfaceInv_Valuation WHERE Period= '" & grid.TextMatrix(i, bteColPosDate) & "' AND Account_No='" & grid.TextMatrix(i, bteColAccNo) & "'  "
                    Db.Execute strSQL
                    
            If RsSearch.State <> adStateClosed Then RsSearch.Close
            RsSearch.Open strSQL, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            If RsSearch.EOF Then
            
            strSQL = " Insert into InterfaceInv_Valuation (Period, Account_No, Amount, Last_User, Register_Date, Posting_Key)" & vbCrLf & _
                             " values('" & grid.TextMatrix(i, bteColPosDate) & "','" & grid.TextMatrix(i, bteColAccNo) & "', '" & CDbl(grid.TextMatrix(i, bteColAmount)) & "', " & vbCrLf & _
                             " '" & userLogin & "', getdate(), '" & grid.TextMatrix(i, bteColPostKey) & "')"
                Db.Execute strSQL
            Else
            strSQL = " Update InterfaceInv_Valuation  set Period = '" & grid.TextMatrix(i, bteColPosDate) & "'," & vbCrLf & _
                             " Amount = '" & CDbl(grid.TextMatrix(i, bteColAmount)) & "', Last_Update= GetDate(), Last_User='" & userLogin & "' " & vbCrLf & _
                             " WHERE Account_No= '" & grid.TextMatrix(i, bteColAccNo) & "' AND Posting_Key = '" & grid.TextMatrix(i, bteColPostKey) & "' "
                             
                    Db.Execute strSQL
        End If
    Next i
    
    
    IFPart = App.path & "\IFData" & "\"
    IFPart1 = App.path & "\IFData" & "\IF_INV_" & Format(DtFrom, "yyyyMMdd") & "_" & Format(DtEnd, "yyyyMMdd")
    
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(IFPart & ".txt", True)
    
    XData = 1
    
    If grid.Rows <= 1 Then
        LblErrMsg = DisplayMsg("0013")
        Exit Sub
    End If
    
    PbMax = grid.Rows - 1
    PBar.Visible = True
    PBar.Max = PbMax
    
    'Do While XData <= Grid.Rows - 1
'
        For Y = 1 To grid.Rows - 1
        
        ListOfData = grid.TextMatrix(Y, 1) & vbTab & _
                            grid.TextMatrix(Y, 2) & vbTab & _
                            grid.TextMatrix(Y, 3) & vbTab & _
                            grid.TextMatrix(Y, 4) & vbTab & _
                            grid.TextMatrix(Y, 5) & vbTab & _
                            grid.TextMatrix(Y, 6) & vbTab & _
                            grid.TextMatrix(Y, 7) & vbTab & _
                            grid.TextMatrix(Y, 8) & vbTab & _
                            grid.TextMatrix(Y, 9) & vbTab & _
                            grid.TextMatrix(Y, 10) & vbTab & _
                            grid.TextMatrix(Y, 11) & vbTab & _
                            grid.TextMatrix(Y, 12) & vbTab & _
                            grid.TextMatrix(Y, 13) & vbTab & _
                            grid.TextMatrix(Y, 14) & vbTab & _
                            grid.TextMatrix(Y, 15) & vbTab
                            
        ListOfData = ListOfData & _
                            grid.TextMatrix(Y, 16) & vbTab & _
                            grid.TextMatrix(Y, 17) & vbTab & _
                            grid.TextMatrix(Y, 18) & vbTab & _
                            grid.TextMatrix(Y, 19) & vbTab & _
                            grid.TextMatrix(Y, 20) & vbTab & _
                            grid.TextMatrix(Y, 21) & vbTab & _
                            grid.TextMatrix(Y, 22) & vbTab & _
                            grid.TextMatrix(Y, 23) & vbTab & _
                            grid.TextMatrix(Y, 24) & vbTab & _
                            grid.TextMatrix(Y, 25) & vbTab & _
                            grid.TextMatrix(Y, 26) & vbTab & _
                            grid.TextMatrix(Y, 27) & vbTab & _
                            grid.TextMatrix(Y, 28) & vbTab & _
                            grid.TextMatrix(Y, 29) & vbTab & _
                            grid.TextMatrix(Y, 30) & vbTab
                            
        ListOfData = ListOfData & _
                            grid.TextMatrix(Y, 31) & vbTab & _
                            grid.TextMatrix(Y, 32) & vbTab & _
                            grid.TextMatrix(Y, 33) & vbTab & _
                            grid.TextMatrix(Y, 34) & vbTab & _
                            grid.TextMatrix(Y, 35) & vbTab & _
                            grid.TextMatrix(Y, 36) & vbTab & _
                            grid.TextMatrix(Y, 37) & vbTab & _
                            grid.TextMatrix(Y, 38) & vbTab & _
                            grid.TextMatrix(Y, 39) & vbTab & _
                            grid.TextMatrix(Y, 40) & vbTab & _
                            grid.TextMatrix(Y, 41) & vbTab & _
                            grid.TextMatrix(Y, 42) & vbTab & _
                            grid.TextMatrix(Y, 43) & vbTab & _
                            grid.TextMatrix(Y, 44) & vbTab & _
                            grid.TextMatrix(Y, 45) & vbTab
                            
        ListOfData = ListOfData & _
                            grid.TextMatrix(Y, 46) & vbTab & _
                            grid.TextMatrix(Y, 47) & vbTab & _
                            grid.TextMatrix(Y, 48) & vbTab & _
                            grid.TextMatrix(Y, 49) & vbTab & _
                            grid.TextMatrix(Y, 50) & vbTab & _
                            grid.TextMatrix(Y, 51) & vbTab & _
                            grid.TextMatrix(Y, 52) & vbTab & _
                            grid.TextMatrix(Y, 53) & vbTab & _
                            grid.TextMatrix(Y, 54)
                            
                            
        a.WriteLine (ListOfData)
        
        Next Y
'        End If
        PBar.Value = XData
    
    a.Close
    
    Set adoStream = New ADODB.Stream

    adoStream.Charset = "ASCII"
    adoStream.Open
    adoStream.LoadFromFile IFPart & ".txt"
    
    adoStream.Position = 0
    Set adoStreamOut = New ADODB.Stream
    adoStreamOut.Charset = "UTF-8"
    adoStreamOut.Open
    adoStreamOut.WriteText adoStream.ReadText
    adoStreamOut.SaveToFile IFPart & "520FID08.txt", adSaveCreateOverWrite
    Kill (IFPart & ".txt")

    PBar.Visible = False
    
    LocalPath = IFPart & "520FID08.txt"
    
    Shell IFPart & "FTPuploadInventoryValuation.bat", vbHide
    Sleep 350

    LblErrMsg = "Export Inventory  Data Success !"
    LblStatus = "Status Approved"
    Cmd_save.Enabled = False
    
    Exit Sub

ErrExport:
    LblErrMsg = "[" & err.number & "]-" & err.Description
    err.clear
End Sub

Private Sub Cmd_SubMenu_Click()
Unload Me
    frmMainMenu.Show
End Sub

Private Sub cmdSearch_Click()
Dim ls_ClosingMonth As String
Dim ls_ClosingYear As String
Dim strSQL As String
Dim rsSQL As New ADODB.Recordset
Dim rsSp As New ADODB.Recordset
Dim cmd As ADODB.Command

    'Call Header

    LblErrMsg = ""
    
    On Error GoTo ErrSearch
    
    Me.MousePointer = vbHourglass
    Call Header
    
    strSQL = " DECLARE @StartDate char(6)    " & vbCrLf & _
          " SET @StartDate = '" & Format(DtFrom, "yyyyMM") & "'  " & vbCrLf & _
          " SELECT * FROM InterfaceInv_Valuation WHERE Period = @StartDate"
    'Db.Execute strSQL
                    
    If rsSQL.State <> adStateClosed Then rsSQL.Close
    rsSQL.Open strSQL, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    
   
    strSQL = "exec sp_InterfaceInv_Valuation '" & Format(DtFrom, "yyyy-MM-dd") & "'"

    If RsSearch.State <> adStateClosed Then RsSearch.Close
    RsSearch.Open strSQL, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsSQL.EOF Then
    LblStatus.Visible = True
    LblStatus = "Status Unapproved"
    Cmd_save.Enabled = True
    Else
        LblStatus.Visible = True
        LblStatus = "Status Approved"
        Cmd_save.Enabled = False
    End If
        
    If RsSearch.EOF Then
        LblErrMsg = DisplayMsg("0013")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    XData = 0
    JmlTran = 0
    i = 0
    
    Do While Not RsSearch.EOF

        i = i + 1
        
        grid.AddItem ""

        grid.TextMatrix(i, bteColPostControl) = Trim(RsSearch!PostingControl)
        grid.TextMatrix(i, bteColCompCode) = Trim(RsSearch!Company_Code)
        grid.TextMatrix(i, bteColPosDate) = Format(RsSearch!Posting_Date, "yyyyMM")
        grid.TextMatrix(i, bteColDocType) = Trim(RsSearch!Document_Type)
        grid.TextMatrix(i, bteColCurr) = Trim(RsSearch!SAP_Curr)
        grid.TextMatrix(i, bteColPostKey) = Trim(RsSearch!Posting_Key)
        grid.TextMatrix(i, bteColAccNo) = Trim(RsSearch!Account_No)
        grid.TextMatrix(i, bteColAmount) = Format(RsSearch!SalesAmount, "#,##0.00")
        grid.TextMatrix(i, bteColTaxCode) = Trim(RsSearch!tax_code)
        grid.TextMatrix(i, bteColCostCenter) = Trim(RsSearch!Cost_Center)
        
        RsSearch.MoveNext
    Loop
    
    If i >= 0 Then
            LblRecord = Format(i, "#,##0 Record")
        Else
            LblRecord = Format("0", "#,##0")
        End If
        
    
    Me.MousePointer = vbDefault
    Exit Sub
    
ErrSearch:

    LblErrMsg = "[" & err.number & "]-" & err.Description
    err.clear
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(Me) 'Editan
Call Header
CtrlMenu1.FormName = Me.Name    'Editan
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"  'Editan
DtFrom.Value = Now
DtEnd.Value = Now
End Sub

'Editan
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
    
    strSQL = "Select * From ftp_Setting"
    If RsFtp.State <> adStateClosed Then RS.Close
    
    Set RsFtp = Db.Execute(strSQL)
    
    If Not RsFtp.EOF Then
    
        Host = IIf(IsNull(RsFtp("Host1")), "", Trim(RsFtp("host2")))
        user = IIf(IsNull(RsFtp("user1")), "", Trim(RsFtp("user2")))
        Pwd = IIf(IsNull(RsFtp("pwd1")), "", Trim(RsFtp("pwd2")))
        Folder = IIf(IsNull(RsFtp("Folder1")), "", Trim(RsFtp("Folder2")))

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
    Inetftp.Password = Pwd
    Folder = Folder & "520FID08.TXT"
    ' Do not include the host name here. That will make
    ' the control try to use its default user name and
    ' password and you'll get the error again.
'    inetFTP.Execute , "Put " & _
'        txtLocalFile.Text & " " & txtRemoteFile.Text
        
    Inetftp.Execute , "Put " & _
        LocalPath & " " & Folder

End Sub


'Editan
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

                'Get the first chunk.
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

Private Function uf_GetLastClosing(Request As String) As String

    '###############################################################
    '#                                                             #
    '#  Notes : To Get Last Closing Month,Year, or full date       #
    '#                                                             #
    '###############################################################

    Dim sqlControl As String, RsInvControl As New ADODB.Recordset
    Dim InvYear As String
    Dim InvMonth As String
    Dim lotno As String

    sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year desc ,inventory_month desc"

    If Request = "fulldate" Then
        sqlControl = "   select    " & _
        "        cast (   " & _
        "        cast(year as varchar(4) ) +case when month <10 then '0' else'' end +cast (month as varchar(2) )+'01'    " & _
        "            as dateTime)ClosingDate     " & _
        "        from    " & _
        "        (   " & _
        "        select top 1 max(inventory_month)month,inventory_year year   " & _
        "         from inventory_control    " & _
        "        where fix_cls='1'   " & _
        "        group by inventory_year   " & _
        "        order by inventory_year desc   " & _
        "        )tbA  "
    End If

    If RsInvControl.State <> adStateClosed Then RsInvControl.Close
    RsInvControl.Open sqlControl, Db, adOpenForwardOnly, adLockReadOnly

    If RsInvControl.EOF = False Then '#Inventory CLosing Data exist
        If Request <> "fulldate" Then
            RsInvControl.MoveFirst
            InvYear = Trim(RsInvControl!Inventory_Year)
            InvMonth = Trim(RsInvControl!Inventory_Month)
        End If
    End If

    If Request = "month" Then '#Request for month
        uf_GetLastClosing = InvMonth
    ElseIf Request = "year" Then '#Request for year
        uf_GetLastClosing = InvYear
    ElseIf Request = "fulldate" Then '#Request for fulldate
        uf_GetLastClosing = IIf(IsNull(RsInvControl!closingdate), 0, Format(RsInvControl!closingdate, "yyyy-MM-dd"))
    End If

    RsInvControl.Close

End Function

Public Function up_ValidateDateRange(n As Date, booUpdate As Boolean) As String

    'Dim li_diff As Integer
    li_diff = DateDiff("M", uf_GetLastClosing("fulldate"), n)
    If booUpdate Then
'        Disable untuk standar

            If li_diff > 0 Then
                up_ValidateDateRange = DisplayMsg(8022) 'Please input valid daterange
            Else
                up_ValidateDateRange = ""
            End If
    Else
        If li_diff > 2 Or li_diff < 0 Then
            up_ValidateDateRange = DisplayMsg(8022) 'Please input valid daterange
            
        Else
            up_ValidateDateRange = ""
        End If
    End If

End Function

Private Sub clear()
    DtFrom = Now
    LblStatus = ""
    LblStatus.Visible = True
    Cmd_save.Enabled = True
    LblErrMsg = ""
End Sub

