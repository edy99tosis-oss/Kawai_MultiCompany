VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmBOMReceiptInquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BOM Receipt Inquiry"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBOMReceiptInquiry.frx":0000
   ScaleHeight     =   10200
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin EZRunnerv3.Anchor Anchor1 
      Left            =   480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6060
      Left            =   120
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   2280
      Width           =   14805
      _cx             =   26114
      _cy             =   10689
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   275
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
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000010&
      Height          =   315
      Left            =   12840
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "FFTT*/"
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton Cmd_Excel 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Excel"
      Height          =   375
      Left            =   13770
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "FFTT*/"
      Top             =   9660
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_SubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "TFFT*/"
      Top             =   9660
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Clear 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "FFTT*/"
      Top             =   9660
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   90
      TabIndex        =   0
      Tag             =   "TFTT*/"
      Top             =   8880
      Width           =   14760
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
         TabIndex        =   8
         Tag             =   "TFTF*/"
         Top             =   240
         Width           =   14205
      End
   End
   Begin InetCtlsObjects.Inet Inetftp 
      Left            =   3180
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13080
      TabIndex        =   9
      Tag             =   "FTTF*/"
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin MSComCtl2.DTPicker DTEffective 
      Height          =   330
      Left            =   1680
      TabIndex        =   1
      Tag             =   "TTFF*/"
      Top             =   1080
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
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
      Format          =   138936323
      UpDown          =   -1  'True
      CurrentDate     =   37860
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1395
      Left            =   120
      TabIndex        =   10
      Tag             =   "TTTF*/"
      Top             =   720
      Width           =   14760
      Begin VB.TextBox txtItemCode 
         Height          =   315
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton CmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Search"
         Height          =   380
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   800
         Width           =   1125
      End
      Begin VB.CommandButton cmdBrowser 
         BackColor       =   &H0080FFFF&
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   3120
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   405
         Width           =   540
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   3480
         X2              =   8400
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label txtItemName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1980
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Receipt Inquiry"
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
      Left            =   0
      TabIndex        =   19
      Tag             =   "TFTF*/"
      Top             =   240
      Width           =   14970
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Left            =   11520
      TabIndex        =   17
      Tag             =   "FFTT*/"
      Top             =   8580
      Width           =   1140
   End
   Begin MSForms.Label LblRecord 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Tag             =   "TFFT*/"
      Top             =   8520
      Width           =   1185
      ForeColor       =   16711935
      BackColor       =   16637923
      VariousPropertyBits=   8388627
      Size            =   "2090;450"
      FontName        =   "Verdana"
      FontEffects     =   1073741827
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label lblNm 
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   0
      Left            =   4080
      TabIndex        =   14
      Tag             =   "TTFF*/"
      Top             =   840
      Width           =   2565
   End
End
Attribute VB_Name = "FrmBOMReceiptInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dateUp As Date
Dim dbTransfer As New ADODB.Connection
Dim DblJml As Double

Dim ColA, colNo, ColDesc, ColHSNumber, ColItemCode, ColDescription, ColRegion, ColOriginCountry, ColInvoiceNumber, ColDate As Byte
Dim ColValue, ColPercent As Byte, ColBCNO As Byte, colBCDate As Byte, colBCType As Byte, ColUSD As Byte, colVendor As Byte, colClasification As Byte, ColQty As Byte
Dim ColUnit As Byte, ColCurrency As Byte, ColPrice As Byte


Dim sqlfg As String, sqlmb As String, FGCls As String
Dim SqlData As String

Private Sub up_Header()
    Dim X As Integer
    
    ' Definisi indeks kolom
    ColItemCode = 0
    ColDescription = 1
    ColRegion = 2
    ColOriginCountry = 3
    ColQty = 4
    ColUnit = 5
    ColCurrency = 6
    ColPrice = 7
    ColValue = 8
    ColInvoiceNumber = 9
    colBCType = 10
    ColDate = 11
    ColPercent = 12
    colVendor = 13
    colClasification = 14
    
    With Grid
        .clear
        .ColS = 15  ' Sesuaikan jumlah kolom setelah menghapus colNo
        .Rows = 1   ' Pastikan ada satu baris untuk header
        
        ' Set header teks
        Dim headers As Variant
        headers = Array("Item Code", "Description", "Region", "Origin Country", _
                        "Qty", "Unit", "Curr", "Unit Price", "Amount (USD)", "BC Number", _
                        "BC Type", "Date", "Percent (%)", "Vendor", "Clasification Part")
        
        For X = LBound(headers) To UBound(headers)
            .TextMatrix(0, X) = headers(X)
        Next X
        
        ' Alignment isi data
        Dim alignments As Variant
        alignments = Array(flexAlignLeftCenter, flexAlignLeftCenter, flexAlignCenterCenter, _
                           flexAlignCenterCenter, flexAlignRightCenter, flexAlignLeftCenter, flexAlignLeftCenter, _
                           flexAlignRightCenter, flexAlignRightCenter, flexAlignLeftCenter, flexAlignLeftCenter, _
                           flexAlignCenterCenter, flexAlignCenterCenter, flexAlignLeftCenter, flexAlignLeftCenter)
        
        For X = LBound(alignments) To UBound(alignments)
            .ColAlignment(X) = alignments(X)
        Next X

        ' Sembunyikan kolom tertentu
        Dim hiddenCols As Variant
        hiddenCols = Array(ColRegion, ColOriginCountry, ColPercent)
        
        For X = LBound(hiddenCols) To UBound(hiddenCols)
            .ColHidden(hiddenCols(X)) = True
        Next X

        ' Set lebar kolom
        Dim colWidths As Variant
        colWidths = Array(1800, 5500, 1500, 1000, 1000, 900, 900, _
                          1000, 1500, 1250, 1000, 2000, 2000, 4500, 2000)

        For X = LBound(colWidths) To UBound(colWidths)
            .ColWidth(X) = colWidths(X)
        Next X
        
        ' === Paksa header rata tengah ===
        Grid.Cell(flexcpAlignment, 0, 0, 0, Grid.ColS - 1) = flexAlignCenterCenter
        
        .EditMaxLength = 1
        .OutlineCol = ColItemCode
        .OutlineBar = flexOutlineBarSimpleLeaf
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False

        ' Refresh tampilan grid
        .Refresh
    End With
End Sub

Private Sub cmdBrowser_Click(Index As Integer)
    Me.MousePointer = vbHourglass
 Select Case Index
  Case 0:
   frm_BrowseItemCode.getItemCode = txtItemCode.Text
   frm_BrowseItemCode.Show 1
   txtItemCode.Text = frm_BrowseItemCode.getItemCode
   txtItemName(1) = frm_BrowseItemCode.getItemName
  Case 1:
   If txtItemCode.Enabled = True Then
    frm_BrowseItemCode.getItemCode = txtItemCode.Text
    frm_BrowseItemCode.Show 1
    txtItemCode.Text = frm_BrowseItemCode.getItemCode
    txtItemName(1) = frm_BrowseItemCode.getItemName
   End If
  End Select
 Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'LblRecord = "0"
    If gb_Simulation = True Then Call up_InitSimulation(Me) 'Editan

    CtrlMenu1.FormName = Me.Name    'Editan
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"  'Editan
    DTEffective.Value = Now
        
    Call up_Header
'    Call up_Header2
    
     With Anchor1
          .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
          .DoInit
    End With
    
    
End Sub

Private Sub Cmd_Excel_Click()
    Dim lb_OFFICE_UNDER2010 As Boolean
    Dim xlapp As New Excel.application
    Dim a As Integer, X As Integer
    Dim TglEnd1 As String, strSQL As String
    Dim Region As String
    Dim Country As String
    Dim RsSearch As New ADODB.Recordset

    Dim Idx As Integer
    Dim IdxTotal As Integer
    Dim IdxB As Integer
    Dim IdxC As Integer
    Dim TotalAsean As Integer
    Dim TotalAseanV As Integer
    Dim xlColA As String
    Dim xlColNo As String
    Dim xlColItemCode As String
    Dim xlColHSNumber As String
    Dim xlColDescription As String
    Dim xlColRegion As String
    Dim xlColOriginCountry As String
    Dim xlColCurrency As String
    Dim xlColPrice As String
    Dim xlColQty As String
    Dim xlColUnit As String
    Dim xlColInvoiceNumber As String
    Dim xlColDate As String
    Dim xlColVendor As String
    Dim xlColValue As String
    Dim xlColPercent As String
    Dim xlcolUSD As String
    Dim xlColBCNo As String
    Dim xlColBCDate As String
    Dim xlColAmount As String
    Dim xlColBCType As String
    Dim xlColClasification As String


    If Grid.Rows > 1 Then

    FGCls = "01"

    'For x = 1 To 3

            LblErrMsg = ""
            Me.MousePointer = vbHourglass


             xlColItemCode = "a"
             xlColDescription = "b"
             xlColQty = "c"
             xlColUnit = "d"
             xlColCurrency = "e"
             xlColPrice = "f"
             xlcolUSD = "g"
             xlColInvoiceNumber = "h"
             xlColBCType = "i"
             xlColBCDate = "j"
             xlColVendor = "k"
             xlColClasification = "l"
             
            With xlapp

                    .Workbooks.Add

                    .Range(xlColItemCode & "1:" & xlColClasification & "3").Merge
                    .Range(xlColItemCode & "1") = "PT KAWAI INDONESIA"
                    .Range(xlColItemCode & "1").Font.Size = 18
                    .Range(xlColItemCode & "1").horizontalAlignment = xlLeft
                    .Range(xlColItemCode & "1").verticalAlignment = xlCenter
                    .Range(xlColItemCode & "1").Font.Bold = True

                    .Range(xlColItemCode & "4:" & xlColClasification & "4").Merge
                    .Range(xlColItemCode & "4") = "BOM"
                    .Range(xlColItemCode & "4").Font.Size = 15
                    .Range(xlColItemCode & "4").horizontalAlignment = xlCenter
                    .Range(xlColItemCode & "4").verticalAlignment = xlCenter
                    .Range(xlColItemCode & "4").Font.Bold = True

                    .Range(xlColItemCode & "5:" & xlColClasification & "5").Merge
                    .Range(xlColItemCode & "5") = "Receipt Inquiry"
                    .Range(xlColItemCode & "5").Font.Size = 15
                    .Range(xlColItemCode & "5").horizontalAlignment = xlCenter
                    .Range(xlColItemCode & "5").verticalAlignment = xlCenter
                    .Range(xlColItemCode & "5").Font.Bold = True

                    '.Range(xlColA & "7:" & xlColNo & "7").Merge
                    .Range(xlColItemCode & "7") = "Part Number"
                    .Range(xlColItemCode & "7").horizontalAlignment = xlLeft
                    .Range(xlColItemCode & "7").horizontalAlignment = xlLeft

                    .Range(xlColDescription & "7") = txtItemCode
                    .Range(xlColDescription & "7").horizontalAlignment = xlLeft
                    .Range(xlColDescription & "7").verticalAlignment = xlCenter
                    .Range(xlColDescription & "7").Font.Bold = True

                    .Range(xlColItemCode & "8") = "Description"
                    .Range(xlColItemCode & "8").horizontalAlignment = xlLeft
                    .Range(xlColItemCode & "8").verticalAlignment = xlCenter

                    .Range(xlColDescription & "8") = txtItemName(1)
                    .Range(xlColDescription & "8").horizontalAlignment = xlLeft
                    .Range(xlColDescription & "8").verticalAlignment = xlCenter
                    .Range(xlColDescription & "8").Font.Bold = True

                    '------------Components Imported or Unknown Origin

                    .Range(xlColItemCode & "10") = "Item Code"
                    .Range(xlColDescription & "10") = "Description"
                    .Range(xlColQty & "10") = "Qty"
                    .Range(xlColUnit & "10") = "Unit"
                    .Range(xlColCurrency & "10") = "Currency"
                    .Range(xlColPrice & "10") = "Unit Price"
                    .Range(xlcolUSD & "10") = "Amount (USD)"
                    .Range(xlColInvoiceNumber & "10") = "BC Number"
                    .Range(xlColBCType & "10") = "BC Type"
                    .Range(xlColBCDate & "10") = "Date"
                    .Range(xlColVendor & "10") = "Vendor"
                    .Range(xlColClasification & "10") = "Clasification"

                    Idx = 10
                    'IdxA = Idx

                    TglEnd1 = DTEffective
                    TglEnd1 = Format(TglEnd1, "mm/dd/yyyy")

                     strSQL = "EXEC sp_BOMReceiptInquiry_Excel '" & txtItemCode & "', '" & TglEnd1 & "'"

                    If RsSearch.State <> adStateClosed Then RsSearch.Close
                    RsSearch.CursorLocation = adUseClient
                    RsSearch.Open strSQL, Db, adOpenForwardOnly, adLockOptimistic

                    i = 10
                    '***
                    Do While Not RsSearch.EOF
                        i = i + 1
                        
                        .Cells(i, xlColItemCode).Value = Trim(RsSearch!Item_Code)
                        .Range(xlColDescription & i) = Trim(RsSearch!item_name)
                        .Range(xlColQty & i) = RsSearch!Qty
                        .Range(xlColUnit & i) = Trim(RsSearch!unit)
                        .Range(xlColCurrency & i) = Trim(RsSearch!Currency)
                        .Range(xlColPrice & i) = RsSearch!Price
                        .Range(xlcolUSD & i) = RsSearch!Value
                        .Range(xlColInvoiceNumber & i) = RsSearch!BC40_No
                        .Range(xlColBCType & i) = RsSearch!BC_Type
                        .Range(xlColBCDate & i) = Format(RsSearch!BC40_Date, "dd MMM yyyy")
                        .Range(xlColVendor & i) = Trim(RsSearch!Vendor)
                        .Range(xlColClasification & i) = Trim(RsSearch!ClasificationPart_Cls)
                        
                        
                        RsSearch.MoveNext
                        
                        
                    Loop

                    Idx = i
                     
                    .Columns(xlColItemCode & ":" & xlColItemCode).columnWidth = 13
                    .Columns(xlColDescription & ":" & xlColDescription).columnWidth = 62
                    .Columns(xlColQty & ":" & xlColQty).columnWidth = 11
                    .Columns(xlColUnit & ":" & xlColUnit).columnWidth = 11
                    .Columns(xlColCurrency & ":" & xlColCurrency).columnWidth = 11
                    .Columns(xlColPrice & ":" & xlColPrice).columnWidth = 16
                    .Columns(xlcolUSD & ":" & xlcolUSD).columnWidth = 16
                    .Columns(xlColInvoiceNumber & ":" & xlColInvoiceNumber).columnWidth = 16
                    .Columns(xlColBCType & ":" & xlColBCType).columnWidth = 16
                    .Columns(xlColBCDate & ":" & xlColBCDate).columnWidth = 16
                    .Columns(xlColVendor & ":" & xlColVendor).columnWidth = 45
                    .Columns(xlColClasification & ":" & xlColClasification).columnWidth = 20
                    .Columns(xlColItemCode).horizontalAlignment = xlLeft
                    .Columns(xlColBCDate).horizontalAlignment = xlCenter

                    .Range(xlColItemCode & "10:" & xlColClasification & Idx).Select

                    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone

                    With .Selection.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With

                    With .Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With

                    .Range(xlColItemCode & "10:" & xlColClasification & "10").Select
                    .Range(xlColItemCode & "10:" & xlColClasification & "10").Font.Bold = True

                    With .Selection
                        .horizontalAlignment = xlCenter
                        .verticalAlignment = xlCenter
                    End With


                     IdxTotal = Idx + 2

                    .Range(xlColQty & IdxTotal) = "Total"
                    .Range(xlcolUSD & IdxTotal).Formula = "=Sum(" & xlcolUSD & "11:" & xlcolUSD & Idx & ")"

                    .Range(xlColItemCode & (IdxTotal & ":") & xlColClasification & IdxTotal).Select

                    .Range(xlColItemCode & (IdxTotal & ":") & xlColClasification & IdxTotal).Font.Bold = True


                    With .Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With

                .Sheets.Select
                .Visible = True
                .WindowState = xlMaximized
                .ActiveWindow.Zoom = 80
                End With
        'Next x


    Else
        LblErrMsg.Caption = DisplayMsg("8012")
    End If

    LblErrMsg.Caption = DisplayMsg("9008")
    Me.MousePointer = vbDefault

End Sub

Private Sub Cmd_SubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Dim TglEnd1 As String, strSQL As String
    Dim RsSearch As New ADODB.Recordset
    Dim i As Integer

    ' Kosongkan pesan error
    LblErrMsg.Caption = ""
    DblJml = 0
    txtamount = Format("0", gs_formatAmount)
    LblRecord.Caption = 0

    ' Panggil ulang header agar tetap konsisten
    Call up_Header

    ' Ubah cursor menjadi jam pasir saat loading
    Me.MousePointer = vbHourglass

    ' Format tanggal
    TglEnd1 = Format(DTEffective, "mm/dd/yyyy")

    ' Query untuk memanggil stored procedure
    strSQL = "EXEC sp_BOMReceiptInquiry_Sel '" & txtItemCode & "', '" & TglEnd1 & "'"

    ' Buka recordset
    If RsSearch.State <> adStateClosed Then RsSearch.Close
    RsSearch.CursorLocation = adUseClient
    RsSearch.Open strSQL, Db, adOpenForwardOnly, adLockOptimistic

    ' Cek jika ada data
    If Not RsSearch.EOF Then
        With Grid
            Do While Not RsSearch.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, ColItemCode) = Trim(RsSearch!Item_Code)
                .TextMatrix(.Rows - 1, ColDescription) = Trim(RsSearch!item_name)
                .TextMatrix(.Rows - 1, ColRegion) = IIf(IsNull(RsSearch("Region")), "", Trim(RsSearch("Region")))
                .TextMatrix(.Rows - 1, ColOriginCountry) = IIf(IsNull(RsSearch("Origin_Country")), "", Trim(RsSearch("Origin_Country")))
                .TextMatrix(.Rows - 1, ColQty) = Trim(RsSearch!Qty)
                .TextMatrix(.Rows - 1, ColUnit) = UCase(Trim(RsSearch!unit))
                .TextMatrix(.Rows - 1, ColCurrency) = Trim(RsSearch!Currency)
                .TextMatrix(.Rows - 1, ColPrice) = Trim(RsSearch!Price)
                .TextMatrix(.Rows - 1, ColInvoiceNumber) = Trim(RsSearch!BC40_No)
                .TextMatrix(.Rows - 1, colBCType) = Trim(RsSearch!BC_Type)
                .TextMatrix(.Rows - 1, ColDate) = Format(RsSearch!BC40_Date, "dd MMM yyyy")
                .TextMatrix(.Rows - 1, ColValue) = Trim(RsSearch!Value)
                .TextMatrix(.Rows - 1, ColPercent) = Format(RsSearch!Percentage, gs_formatPercentage)
                .TextMatrix(.Rows - 1, colVendor) = Trim(RsSearch!Vendor)
                .TextMatrix(.Rows - 1, colClasification) = Trim(RsSearch!ClasificationPart_Cls)
                
                ' Update jumlah total
                DblJml = DblJml + RsSearch("Value")

                ' Outline row untuk tampilan hirarki
                .Col = 1
                .IsSubtotal(.Rows - 1) = True
                .RowData(.Rows - 1) = .Rows - 1
                .RowOutlineLevel(.Rows - 1) = Trim(RsSearch!level)

                RsSearch.MoveNext
            Loop
        End With

        ' Update jumlah record
        LblRecord.Caption = Format(RsSearch.RecordCount, "#,##0 Record")

    Else
        LblErrMsg.Caption = DisplayMsg("0013")
    End If

    ' Tutup recordset jika masih terbuka
    If RsSearch.State <> adStateClosed Then RsSearch.Close

    ' Kembalikan cursor ke default
    Me.MousePointer = vbDefault

    ' Update total amount
    txtamount = Format(DblJml, gs_formatAmount)

    ' === Refresh Grid agar tampilan tetap benar ===
    Grid.Refresh
End Sub

Private Sub cmd_clear_Click()
    Call up_Header
'    Call up_Header2
    DTEffective.Value = Now
    LblRecord = "0 Record"
    LblErrMsg.Caption = ""
    txtamount = ""
    txtItemCode = ""
End Sub

Private Sub txtItemCode_Change()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    sql = "Select * From item_master where FinishGoodPart_Cls='01' and Item_Code='" & txtItemCode.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        txtItemName(1).Caption = Trim(RS("Item_Name"))
    Else
        txtItemName(1).Caption = ""
        Exit Sub
    End If
    
    If txtItemCode.Text = "" Then txtItemName(1).Caption = ""
    
End Sub
