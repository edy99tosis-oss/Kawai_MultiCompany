VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmPart_RecUpload 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Part (Material) Receipt Scheduled Upload"
   ClientHeight    =   10425
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15090
   Icon            =   "FrmPart_RecUpload.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   34766.63
   ScaleMode       =   0  'User
   ScaleWidth      =   15119.68
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1275
      Left            =   240
      TabIndex        =   17
      Top             =   7560
      Width           =   14595
      Begin VB.Label lblSupplierCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( 0 )"
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
         Left            =   8760
         TabIndex        =   35
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Supplier Code"
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
         Left            =   6720
         TabIndex        =   34
         Top             =   600
         Width           =   1860
      End
      Begin VB.Label lblWHCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( 0 )"
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
         Left            =   2500
         TabIndex        =   32
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Warehouse Code"
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
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   2115
      End
      Begin MSForms.Label Label5 
         Height          =   255
         Left            =   649
         TabIndex        =   30
         Top             =   1440
         Width           =   2655
         BackColor       =   16637923
         Caption         =   "Invalid Format End Date"
         Size            =   "4683;450"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Item Code"
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
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid PO No"
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
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Surat Jalan No"
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
         Left            =   3480
         TabIndex        =   27
         Top             =   240
         Width           =   1890
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Delivery Date"
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
         Left            =   3480
         TabIndex        =   26
         Top             =   600
         Width           =   1830
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Qty"
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
         Left            =   3480
         TabIndex        =   25
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid BC Date"
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
         Left            =   6720
         TabIndex        =   24
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lblItemCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( 0 )"
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
         Left            =   2500
         TabIndex        =   23
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblPONo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( 0 )"
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
         Left            =   2500
         TabIndex        =   22
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblSJNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( 0 )"
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
         Left            =   5640
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblDelDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( 0 )"
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
         Left            =   5640
         TabIndex        =   20
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( 0 )"
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
         Left            =   5640
         TabIndex        =   19
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblBCDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( 0 )"
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
         Left            =   8760
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FDDFE3&
      Height          =   1260
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   14625
      Begin VB.TextBox lblSupp 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   3255
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   270
         Width           =   3510
      End
      Begin VB.CommandButton cmd_Browser 
         Caption         =   "..."
         Height          =   300
         Left            =   6840
         TabIndex        =   13
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox txtUpload 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   7935
      End
      Begin VB.Line Line4 
         X1              =   3255
         X2              =   6755
         Y1              =   555
         Y2              =   555
      End
      Begin MSForms.ComboBox cboSupplier 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2672;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code "
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
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   1275
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   1215
         BackColor       =   16637923
         Caption         =   "File Name"
         Size            =   "2143;450"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton btnUpload 
         Height          =   375
         Left            =   9720
         TabIndex        =   11
         Top             =   720
         Width           =   495
         BackColor       =   14737632
         Caption         =   "..."
         Size            =   "873;661"
         FontName        =   "Verdana"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton btnTemplate 
         Height          =   375
         Left            =   10320
         TabIndex        =   10
         Top             =   720
         Width           =   1215
         BackColor       =   8454143
         Caption         =   "Template"
         Size            =   "2143;661"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.CommandButton btnCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Cancel"
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9720
      Width           =   1185
   End
   Begin VB.CommandButton CmdMenu 
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9720
      Width           =   1185
   End
   Begin VB.CommandButton btnSubmit 
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
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9720
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   240
      TabIndex        =   3
      Top             =   9000
      Width           =   14595
      Begin VB.Label LblErr 
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
         Height          =   255
         Left            =   105
         TabIndex        =   4
         Top             =   195
         Width           =   14370
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12960
      TabIndex        =   0
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4890
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2280
      Width           =   14595
      _cx             =   25744
      _cy             =   8625
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
      BackColorFixed  =   12640511
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483624
      BackColorAlternate=   -2147483624
      GridColor       =   8421504
      GridColorFixed  =   12582912
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
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
   Begin MSComDlg.CommonDialog CDExcel 
      Left            =   11520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.Label LblRecord 
      Height          =   255
      Left            =   11520
      TabIndex        =   33
      Tag             =   "FFTT*/"
      Top             =   7227
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Part (Material) Receipt Scheduled Upload"
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
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   14595
   End
End
Attribute VB_Name = "FrmPart_RecUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim simpan As Boolean, ubah As Boolean, hapus As Boolean 'Status Ubah/Hapus
Dim parentItem As String 'Parent Item
Dim gridTglAwal As String, gridTglAkhir As String 'Klik Grid
Dim sama As Boolean 'cek Parent
Dim nilKosong As Boolean
Dim kondisi As String
Dim provisionCls As String
Dim KeyProd As String
Dim seqNo As New clsMRP
Dim blnFix As Integer, thnFix As Integer
Dim TglReceipt, DateActual, receiptDate, bcDate As Date
Dim JmlTran As Integer

Dim tglAwal, tglAkhir As String
Dim tglSesdh, tglSeblm As String
Dim l_prod_code As Double, l_qty_format As Double, l_Start_Date As Double, l_End_Date As Double
Dim IInvalidWHCode, IInvalidProdCode, IInvalildPONo, IInvalidSJNo, IInvalidDelDate, IInvalidQty, IInvalidTypeBC, IInvalidSupplier As Integer
Dim IInvalidNoBC, IInvalidBCDate As Integer
Dim validate As Boolean
Dim sql As String
Dim SupplierCode, PONO, SJNo, WHCode, ItemCode, BC40_No, bcType  As String
Dim Qty As Double
Dim newDb As New ADODB.Connection

Dim bteColWarehouseCode As Byte
Dim bteColProdCode As Byte
Dim bteColDesc As Byte
Dim bteColPONo As Byte
Dim bteColSuratJalanNo As Byte
Dim bteColDeliveryDate As Byte
Dim bteColQty As Byte
Dim bteColTypeBC As Byte
Dim bteColNoBC As Byte
Dim bteColBCDate As Byte
Dim bteColRemark As Byte

Dim ls_PathExcel As String

Private Sub up_GridHeader()

    Dim i As Integer
    
    LblRecord = "0 Record(s)"
    
    bteColWarehouseCode = 0
    bteColProdCode = 1
    bteColDesc = 2
    bteColPONo = 3
    bteColSuratJalanNo = 4
    bteColDeliveryDate = 5
    bteColQty = 6
    bteColTypeBC = 7
    bteColNoBC = 8
    bteColBCDate = 9
    bteColRemark = 10
    
    With grid
        .clear
        .ColS = 11
        .Rows = 1
        
        .TextMatrix(0, bteColWarehouseCode) = "WH Code"
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColPONo) = "PO No"
        .TextMatrix(0, bteColSuratJalanNo) = "Surat Jalan No"
        .TextMatrix(0, bteColDeliveryDate) = "Delivery Date"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColTypeBC) = "BC Type"
        .TextMatrix(0, bteColNoBC) = "BC No"
        .TextMatrix(0, bteColBCDate) = "BC Date"
        .TextMatrix(0, bteColRemark) = "Remark"
        
        .ColWidth(bteColWarehouseCode) = 1500
        .ColWidth(bteColProdCode) = 1500
        .ColWidth(bteColDesc) = 3000
        .ColWidth(bteColPONo) = 3000
        .ColWidth(bteColSuratJalanNo) = 2500
        .ColWidth(bteColDeliveryDate) = 1500
        .ColWidth(bteColQty) = 1000
        .ColWidth(bteColTypeBC) = 1000
        .ColWidth(bteColNoBC) = 1000
        .ColWidth(bteColBCDate) = 1500
        .ColWidth(bteColRemark) = 4500
        
                
        .ColAlignment(bteColWarehouseCode) = flexAlignLeftCenter
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColPONo) = flexAlignLeftCenter
        .ColAlignment(bteColSuratJalanNo) = flexAlignLeftCenter
        .ColAlignment(bteColDeliveryDate) = flexAlignLeftCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColTypeBC) = flexAlignLeftCenter
        .ColAlignment(bteColNoBC) = flexAlignLeftCenter
        .ColAlignment(bteColBCDate) = flexAlignLeftCenter
        .ColAlignment(bteColRemark) = flexAlignLeftCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, bteColRemark) = flexAlignCenterCenter
        
        .EditMaxLength = 1
    End With

End Sub

Private Sub up_Clear()
    TglReceipt = Now()
    lblWHCode = "( 0 )"
    lblItemCode = "( 0 )"
    lblPONo = "( 0 )"
    lblSJNo = "( 0 )"
    lblWHCode = "( 0 )"
    lblDelDate = "( 0 )"
    lblQty = "( 0 )"
    lblBCDate = "( 0 )"
    
    IInvalidWHCode = 0
    IInvalidProdCode = 0
    IInvalildPONo = 0
    IInvalidSJNo = 0
    IInvalidDelDate = 0
    IInvalidQty = 0
    IInvalidTypeBC = 0
    IInvalidNoBC = 0
    IInvalidBCDate = 0
    btnSubmit.Enabled = False
    
    CDExcel.filename = ""
    LblErr.Caption = ""
    txtUpload = ""
    cboSupplier(0).Text = ""
    lblSupp.Text = ""
    up_GridHeader
End Sub

Private Sub btnCancel_Click()
    up_Clear
End Sub

Private Sub btnSubmit_Click()
    Dim strS As Integer, Jawab As Integer, CekS As Boolean, CekD As Boolean
    Dim strD As Integer, ie As Integer
    Dim CC As Boolean, ubah As Integer
    Dim totalQty As Double
    Dim strSQL As String
    Dim X As Double
    
    Dim FixMonth As String * 2
    Dim FixYear As String * 4
    Dim rsProv As New ADODB.Recordset
    
    Dim icg As Long
    Dim tampungBln As String
            
    ' ---------------
    Me.MousePointer = vbHourglass
    LblErr = ""

    If HakU = 0 Then LblErr = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
    tampungBln = seqNo.blnAkhir()
    blnFix = Split(tampungBln, ",")(0)
    thnFix = Split(tampungBln, ",")(1)
           
    If cek Then
    
     Db.BeginTrans
     
     For X = 1 To grid.Rows - 1
        SupplierCode = cboSupplier(0).Text
        PONO = grid.TextMatrix(X, bteColPONo)
        WHCode = grid.TextMatrix(X, bteColWarehouseCode)
        receiptDate = grid.TextMatrix(X, bteColDeliveryDate)
        ItemCode = grid.TextMatrix(X, bteColProdCode)
        Qty = CDbl(grid.TextMatrix(X, bteColQty))
        SJNo = grid.TextMatrix(X, bteColSuratJalanNo)
        BC40_No = grid.TextMatrix(X, bteColNoBC)
        bcDate = grid.TextMatrix(X, bteColBCDate)
        bcType = grid.TextMatrix(X, bteColTypeBC)
        
        up_UploadPartReceipt
        
        
        '*********** Proses Insert ke Supply **************
        sql = "Select Provision_Cls from Item_MaSter where ITem_Code = '" & ItemCode & "'"
        Set rsProv = Db.Execute(sql)
        If Not rsProv.EOF Then provisionCls = Trim(rsProv(0)) Else provisionCls = ""

        '*************************************************
        
        '#Process Konsumsi Subcon
        '===============================================
        If KeyProd = "R" Then
           KeyProd = ""
        End If
        If uf_GetSubConStatus(Trim(SupplierCode)) = "3" Then
            KeyProd = "R" & KeyProd
            Call up_SubConInputConsumption(Trim$(ItemCode), CDbl(Qty), Trim$(WHCode))
        End If
        '===============================================
        
        If grid.TextMatrix(X, bteColTypeBC) = "4.0" Then
        ElseIf grid.TextMatrix(X, bteColTypeBC) = "2.6.2" Then
        ElseIf grid.TextMatrix(X, bteColTypeBC) = "2.3" Then
        Else
            ProsesStock 1, Trim$(ItemCode), Trim$(WHCode), Trim$(WHCode), Format(receiptDate, "YYYYMM"), Trim$(Qty), ""
        End If
        
        Db.Execute "update purchaseorder_detail with (updlock) set complete_cls = " & _
            "case when " & _
                "(select isnull(sum(qty), 0) from part_receipt where receipt_cls = 'R' and po_no = purchaseorder_detail.po_no and item_code = purchaseorder_detail.item_code) - " & _
                "(select isnull(sum(qty), 0) from part_receipt where receipt_cls = 'R1' and po_no = purchaseorder_detail.po_no and item_code = purchaseorder_detail.item_code) >= qty " & _
            "then 1 else 0 end " & _
            "where po_no='" & Trim$(PONO) & "' and item_code='" & Trim$(ItemCode) & "' "
          
    Next X
    
    If err.number = 0 Then
        Db.CommitTrans
    Else
        Db.RollbackTrans
        LblErr = err.Description
        err.clear
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    LblErr = DisplayMsg(1000)
    btnSubmit.Enabled = False
       
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub btnTemplate_Click()
Dim objExcel As New Excel.application
    Dim RS As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Double

'On Error GoTo errHandler

    If G_CekExcelApp = False Then LblErr = DisplayMsg(5000): Exit Sub

    LblErr.Caption = ""
    CDExcel.filter = "Excel Files (*.xls)|*.xls"
    CDExcel.filename = "Part (Material) Receipt (Upload) "
    CDExcel.CancelError = True

    On Error GoTo errCancel
    CDExcel.ShowSave

   On Error GoTo errHandler
    If Len(CDExcel.filename) = 0 Then Exit Sub
    If Dir(CDExcel.filename) <> "" Then
        If MsgBox("Overwrite existing file?", vbExclamation + vbYesNo, "Overwrite") = vbNo Then Exit Sub
    End If
    ls_PathExcel = Mid(CDExcel.filename, 1, Len(CDExcel.filename) - Len(CDExcel.FileTitle))

    MousePointer = MousePointerConstants.vbHourglass

    Set objExcel = New Excel.application
    With objExcel
        .Workbooks.Add
        .Visible = True
        .Cells.Select
        .Cells.EntireColumn.delete

        .Range("A1").EntireColumn.delete xlDown
        .Range("A1:I1").Borders.Weight = xlThin
        .Range("A2:I2").Borders.Weight = xlThin
        .Rows("1:" & grid.Rows).Select
        .Selection.Interior.Pattern = xlNone
        
        .Range("A1:I1").Select
        .Selection.Font.Bold = True
        
        .Range("A2:I2").Select
        .Selection.Font.color = &H80FF80
        
        .Range("A1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("A1").Value = "Warehouse Code"
        
        .Range("B1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("B1").Value = "Product Code"

        .Range("C1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("C1").Value = "PO No"

        .Range("D1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("D1").Value = "Surat Jalan No"

        .Range("E1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("E1").Value = "Delivery Date"
         
        .Range("F1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("F1").Value = "Qty"
        
        .Range("G1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("G1").Value = "Type BC"
        
        .Range("H1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("H1").Value = "No BC"
        
        .Range("I1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("I1").Value = "BC Date"
        
        .Range("A2").Value = "Char (25)"
        .Range("B2").Value = "Char (25)"
        .Range("C2").Value = "Char (50)"
        .Range("D2").Value = "Char (100)"
        .Range("E2").Value = "Date (MM/dd/yyyy)"
        .Range("F2").Value = "Numeric (18, 2)"
        .Range("G2").Value = "Char (3)"
        .Range("H2").Value = "Char (10)"
        .Range("I2").Value = "Date (MM/dd/yyyy)"
        
        .Cells.Select
        .Cells.EntireColumn.AutoFit

        .ActiveWorkbook.SaveAs filename:= _
        CDExcel.filename, FileFormat:= _
        xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
        , CreateBackup:=False

    End With
    
    MousePointer = MousePointerConstants.vbDefault
    Exit Sub

errHandler:
    If err.number <> 0 Then
        MousePointer = MousePointerConstants.vbDefault
        LblErr = err.Description
        grid.FixedRows = 1
    End If
    If RS.State = adStateOpen Then
        RS.Close
        Set RS = Nothing
    End If
errCancel:
End Sub

Private Function uf_ValidateInput() As Boolean
    If cboSupplier(0).Text = "" Then
        cboSupplier(0).SetFocus
        LblErr = "Please Select Supplier Code !"
        uf_ValidateInput = False
        Exit Function
    End If
    uf_ValidateInput = True
    
End Function

Private Sub btnUpload_Click()
Dim objExcel As New Excel.application
Dim objWorkSheet As New Worksheet
Dim objWorkBook As Workbook
Dim i As Long
Dim iCol As Integer
Dim colcount As Integer
Dim RS As New ADODB.Recordset
Dim strSQL As String
Dim filename As String
Dim rsTglAwal As String
Dim rs_DB As New ADODB.Recordset
Dim rs_Unit As New ADODB.Recordset
Dim rs_PO As New ADODB.Recordset
Dim rs_Qty As New ADODB.Recordset
Dim rs_Supp As New ADODB.Recordset
Dim ls_invalidMsg As String
Dim iGrdRow As Double
Dim rsUpdate As New ADODB.Recordset 'utk Update Data
Dim WHDesc As String
Dim StatusRow As Boolean
    
If uf_ValidateInput = False Then Exit Sub
    
Me.MousePointer = vbHourglass
    
    lblWHCode = "( 0 )"
    lblItemCode = "( 0 )"
    lblPONo = "( 0 )"
    lblSJNo = "( 0 )"
    lblWHCode = "( 0 )"
    lblDelDate = "( 0 )"
    lblQty = "( 0 )"
    lblBCDate = "( 0 )"
    
    IInvalidWHCode = 0
    IInvalidProdCode = 0
    IInvalildPONo = 0
    IInvalidSJNo = 0
    IInvalidDelDate = 0
    IInvalidQty = 0
    IInvalidTypeBC = 0
    IInvalidNoBC = 0
    IInvalidBCDate = 0
    
    CDExcel.filename = ""
    LblErr.Caption = ""
    txtUpload = ""
    up_GridHeader
    
    JmlTran = 0
    
    If G_CekExcelApp = False Then LblErr.Caption = "Excel Application is not found": Exit Sub
        
    LblErr.Caption = ""
    CDExcel.filter = "Excel Worksheets (*.xls)|*.xls|"
    
    On Error GoTo errCancel
    CDExcel.CancelError = True
    
    On Error GoTo err
    
    CDExcel.ShowOpen
    filename = CDExcel.filename

    txtUpload = filename
    txtUpload.SetFocus
    
    If CDExcel.filename <> "" Then
            
        Set objExcel = New Excel.application
        Set objWorkBook = objExcel.Workbooks.Open(CDExcel.filename)
        Set objWorkSheet = objWorkBook.Sheets("Sheet1")
        objExcel.Visible = False
        
        i = 2
        iGrdRow = 1
        colcount = 22
        With objWorkSheet
            Do While .Cells(i, 1).Value <> ""
                    ls_invalidMsg = ""
                     StatusRow = True

                        grid.AddItem ""
                        
                       '1. Cek WH Code
                       rs_DB.Open "SELECT WH_Code FROM WareHouse_Master WHERE WH_Code='" & Trim(.Cells(i, 1)) & "'", Db, adOpenKeyset, adLockOptimistic
                       If rs_DB.EOF = True Then
                            IInvalidWHCode = IInvalidWHCode + 1
                            StatusRow = False
                            ls_invalidMsg = "Warehouse Code not register in warehouse master"
                        Else
                            grid.TextMatrix(iGrdRow, bteColWarehouseCode) = Trim(rs_DB!wh_code)
                        End If
                        
                       '2. Cek Product Code
                       rs_Unit.Open "SELECT Item_Code, Item_Name FROM Item_Master WHERE Item_Code='" & Trim(.Cells(i, 2)) & "'", Db, adOpenKeyset, adLockOptimistic
                       If rs_Unit.EOF = True Then
                            IInvalidProdCode = IInvalidProdCode + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "Product Code not register in Item Master"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", Product Code not register in Item Master"
                            End If
                        Else
                            grid.TextMatrix(iGrdRow, bteColProdCode) = Trim(rs_Unit!Item_Code)
                            grid.TextMatrix(iGrdRow, bteColDesc) = Trim(rs_Unit!item_name)
                        End If
                        
                       '3. Cek PO No
                       rs_PO.Open "SELECT PO_No FROM PurchaseOrder_Master WHERE PO_No='" & Trim(.Cells(i, 3)) & "'", Db, adOpenKeyset, adLockOptimistic
                       If rs_PO.EOF = True Then
                            IInvalildPONo = IInvalildPONo + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "PO No Not Register In Purchase Order Master"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", PO No Not Register In Purchase Order Master"
                            End If
                        Else
                            grid.TextMatrix(iGrdRow, bteColPONo) = Trim(rs_PO!po_no)
                        End If
                        
                        '4. Surat Jalan
                        grid.TextMatrix(iGrdRow, bteColSuratJalanNo) = .Cells(i, 4)
                        
                        '5. Cek Delivery Date
                        If IsDate(Format(.Cells(i, 5), "dd MMM yyyy")) = False Then
                            IInvalidDelDate = IInvalidDelDate + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "Invalid Delivery Date Format"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", Invalid Delivery Date Format"
                            End If
                        Else
                            grid.TextMatrix(iGrdRow, bteColDeliveryDate) = Format(.Cells(i, 5), "dd MMM yyyy")
                        End If
                        
                        '6. Cek Qty
                        rs_Qty.Open " SELECT (Qty)POQty, ISNULL((select sum(Qty) from Part_Receipt where PO_No='" & Trim(.Cells(i, 3)) & "'" & vbCrLf & _
                                    " and Item_Code='" & Trim(.Cells(i, 2)) & "'),0)PRQty FROM PurchaseOrder_Detail" & vbCrLf & _
                                    " WHERE PO_No='" & Trim(.Cells(i, 3)) & "' and Item_Code='" & Trim(.Cells(i, 2)) & "'", Db, adOpenKeyset, adLockOptimistic
                         
                         If rs_Qty.EOF = False Then
                             If ((rs_Qty!PRQty) + .Cells(i, 6)) > (rs_Qty!poqty) Then
                                IInvalidQty = IInvalidQty + 1
                                StatusRow = False
                                If ls_invalidMsg = "" Then
                                    ls_invalidMsg = "Qty receipt must be lower then purchase qty"
                                    grid.TextMatrix(iGrdRow, bteColQty) = Format(.Cells(i, 6), gs_formatQtyBOM)
                                Else
                                    ls_invalidMsg = ls_invalidMsg & ", Qty receipt must be lower then purchase qty"
                                    grid.TextMatrix(iGrdRow, bteColQty) = Format(.Cells(i, 6), gs_formatQtyBOM)
                                End If
                            Else
                                grid.TextMatrix(iGrdRow, bteColQty) = Format(.Cells(i, 6), gs_formatQtyBOM)
                            End If
                        End If
                        
                        grid.TextMatrix(iGrdRow, bteColTypeBC) = .Cells(i, 7)
                        grid.TextMatrix(iGrdRow, bteColNoBC) = .Cells(i, 8)
                        
                       '9. Cek BC Date
                        If IsDate(Format(.Cells(i, 9), "dd MMM yyyy")) = False Then
                            IInvalidBCDate = IInvalidBCDate + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "Invalid BC Date Format"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", Invalid BC Date Format"
                            End If
                        Else
                            grid.TextMatrix(iGrdRow, bteColBCDate) = Format(.Cells(i, 9), "dd MMM yyyy")
                        End If
                        
                       '10. Cek Supplier
                       rs_Supp.Open "SELECT Supplier_Code FROM PurchaseOrder_Master WHERE PO_No='" & Trim(.Cells(i, 3)) & "'", Db, adOpenKeyset, adLockOptimistic
                       If rs_Supp.EOF = False Then
                        If Trim(rs_Supp!Supplier_Code) <> Trim(cboSupplier(0).Text) Then
                            IInvalidSupplier = IInvalidSupplier + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "Supplier Code Not Register In Purchase Order Master"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", Supplier Code Not Register In Purchase Order Master"
                            End If
                        End If
                        End If
                           
                       rs_DB.Close
                       rs_Unit.Close
                       rs_PO.Close
                       rs_Qty.Close
                       rs_Supp.Close
                       
                        If StatusRow = False Then
                            grid.Cell(flexcpBackColor, iGrdRow, bteColWarehouseCode, iGrdRow, bteColRemark) = vbRed
                            grid.TextMatrix(iGrdRow, bteColRemark) = ls_invalidMsg
                            btnSubmit.Enabled = False
                        Else
                            btnSubmit.Enabled = True
                        End If
                        iGrdRow = iGrdRow + 1

                    DoEvents
                    i = i + 1
                    
                    JmlTran = JmlTran + 1
            Loop
            
            lblWHCode.Caption = "( " & IInvalidWHCode & " )"
            lblItemCode.Caption = "( " & IInvalidProdCode & " )"
            lblPONo.Caption = "( " & IInvalildPONo & " )"
            lblSJNo.Caption = "( " & IInvalidSJNo & " )"
            lblDelDate.Caption = "( " & IInvalidDelDate & " )"
            lblQty.Caption = "( " & IInvalidQty & " )"
            lblBCDate.Caption = "( " & IInvalidBCDate & " )"
            
            LblRecord = Format(JmlTran, "#,##0") & " Record(s)"
            
        End With
        
        objWorkBook.Close
        Set objWorkSheet = Nothing
        Set objWorkBook = Nothing
        Set objExcel = Nothing
        
    LblErr.Caption = "Reading Excel Finish"
        
    Me.MousePointer = vbDefault
        
    End If
    Exit Sub
    
errCancel:
err:
    LblErr.Caption = err.Description
    objExcel.Workbooks.Close
    Set objWorkSheet = Nothing
    Set objWorkBook = Nothing
    Set objExcel = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub CboSupplier_Change(Index As Integer)
Dim RS As New ADODB.Recordset
Dim rsUnit As New ADODB.Recordset

    If cboSupplier(0).MatchFound Then
        lblSupp = cboSupplier(0).Column(1)
        LblErr = ""
    Else
        LblErr = ""
    End If
    
    If LblErr = "" Then
        rsUnit.Open "select * from Item_master where item_code='" & cboSupplier(0).Text & "'", Db, adOpenKeyset, adLockOptimistic
        If rsUnit.EOF = False Then
            LblErr = IIf(IsNull(Trim(rsUnit("item_name"))), "", Trim(rsUnit("item_name")))
        End If
        If rsUnit.State <> adStateClosed Then rsUnit.Close
    End If
    
    'penambahan 20220604
'    RS.Open "SELECT TOP 1 WH_Code FROM dbo.WareHouse_Master WHERE Adm_Group = '" & Trim(cboSupplier(0)) & "' ", Db, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If Not RS.EOF Then
'        WHCode = Trim(RS.Fields("WH_Code"))
'    Else
'        LblErr = DisplayMsg(1042)
'    End If
'    RS.Close
    
End Sub

Private Sub cmd_Browser_Click()
 Me.MousePointer = vbHourglass
  frm_BrowseSupp.getItemCode = cboSupplier(0).Text
  frm_BrowseSupp.Show 1
  cboSupplier(0).Text = frm_BrowseSupp.getPartNumber
  cboSupplier(0).SetFocus
  
  Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    HakU = hakUpdate(Me.Name)
    
    If newDb.State <> adStateClosed Then newDb.Close
    newDb.Open Db.ConnectionString
    
    up_Clear
        
    up_FillCombo
    up_GridHeader
End Sub

Private Sub up_FillCombo()

    Dim rscbo As New ADODB.Recordset 'Isi Combo

   '##Tampilkan Combo Customer code dari trade_master
    sql = "Select rtrim(trade_code) as TC,Trade_name as TN, Address1 as A from trade_master where trade_cls in ('2', '3') order by trade_code"
    
    Set rscbo = New Recordset
    rscbo.Open sql, Db, adOpenKeyset, adLockOptimistic
    cboSupplier(0).clear
    cboSupplier(0).columnCount = 2
    cboSupplier(0).TextColumn = 1
    i = 0
    While Not rscbo.EOF
        cboSupplier(0).AddItem ""
        cboSupplier(0).List(i, 0) = rscbo!TC
        cboSupplier(0).List(i, 1) = Trim$(rscbo!TN)
        i = i + 1
        rscbo.MoveNext
    Wend
    cboSupplier(0).ColumnWidths = "60 pt; 300 pt"
    cboSupplier(0).ListWidth = 360
    cboSupplier(0).ListRows = 15

    Set rscbo = Nothing
    
End Sub

Private Sub CmdMenu_Click()
'    rst.Close
'    Set rst = Nothing
'    RsW.Close
'    Set RsW = Nothing
    frmMainMenu.Show
    Unload Me
    'DoEvents
End Sub

Private Function uf_GetSubConStatus(ls_TradeCode As String) As String
    Dim RS As New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    RS.Open "select trade_cls from trade_master where trade_code='" & Trim(ls_TradeCode) & "'", Db, adOpenKeyset, adLockOptimistic
    If RS.EOF = False Then
        uf_GetSubConStatus = Trim(RS!trade_cls & "")
    Else
        uf_GetSubConStatus = ""
    End If
    If RS.State = 1 Then RS.Close
End Function

Function cek() As Boolean
Dim X As Integer
cek = True

 For X = 1 To grid.Rows - 1
    PONO = grid.TextMatrix(X, bteColProdCode)
    SJNo = grid.TextMatrix(X, bteColSuratJalanNo)
    
    up_ValidateSuratJalan
    If validate <> True Then
        LblErr = DisplayMsg(9013)
        cek = False
        Exit Function
    End If
 Next X

'Validasi closing
'LblErr = up_ValidateDateRange(Format(TglReceipt, "yyyy-MM-dd"), True)
'If LblErr <> "" Then
'    cek = False
'    Exit Function
'End If

End Function

Public Function up_ValidateSuratJalan() As Boolean
Dim sqlcek As String
Dim rsCek As New ADODB.Recordset
    
    DateActual = Format(TglReceipt, "MM/DD/YYYY")
    
    sqlcek = " SELECT DISTINCT * FROM Part_Receipt WHERE PO_No='" & PONO & "' " & vbCrLf & _
            " AND SuratJalan_No='" & SJNo & "' "
    Set rsCek = Db.Execute(sqlcek)
    
    If Not rsCek.EOF Then
    
        sqlcek = " SELECT DISTINCT Receipt_Date FROM Part_Receipt WHERE PO_No='" & PONO & "' " & vbCrLf & _
                 " AND SuratJalan_No='" & SJNo & "' "
        Set rsCek = Db.Execute(sqlcek)
        
        If Not rsCek.EOF Then
            receiptDate = Trim(rsCek!Receipt_Date & "")
        End If
        
        If receiptDate = DateActual Then
            validate = True
        Else
            validate = False
        End If
    Else
        validate = True
    End If
    
End Function

Private Sub up_UploadPartReceipt()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim prm1 As ADODB.Parameter
    Dim prm2 As ADODB.Parameter
    Dim prm3 As ADODB.Parameter
    Dim prm4 As ADODB.Parameter
    Dim prm5 As ADODB.Parameter
    Dim prm6 As ADODB.Parameter
    Dim prm7 As ADODB.Parameter
    Dim prm8 As ADODB.Parameter
    Dim prm9 As ADODB.Parameter
    Dim prm10 As ADODB.Parameter
    Dim prm11 As ADODB.Parameter
        
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_UploadPartReceipt_Ins"
     
    Set prm1 = cmd.CreateParameter("Supplier", adVarChar, adParamInput, 15, SupplierCode)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("PONo", adVarChar, adParamInput, 35, PONO)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("WarehouseCode", adVarChar, adParamInput, 15, WHCode)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("ReceiptDate", adDate, adParamInput, , receiptDate)
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 25, ItemCode)
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("Qty", adDouble, adParamInput, , Qty)
    cmd.Parameters.append prm6
    Set prm7 = cmd.CreateParameter("SuratJalanNo", adVarChar, adParamInput, 25, SJNo)
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("BC40_No", adVarChar, adParamInput, 30, BC40_No)
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("BCDate", adDate, adParamInput, , bcDate)
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("BCType", adVarChar, adParamInput, 15, bcType)
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("UserID", adVarChar, adParamInput, 15, userLogin)
    cmd.Parameters.append prm11
    
    cmd.Execute

End Sub

'End Sub

'Proses Stock
Sub ProsesStock(nTipe As Byte, ItemCode As String, OldWHCode As String, NewWHCode As String, RecDate As String, QtyX As String, OldQty As String)
Dim Sqlc As String, RsInvControl As New ADODB.Recordset
Dim RsWHS As Recordset, RSIS As Recordset, rswh As Recordset
Dim WHStock_ctrl As Boolean
Dim ItemStock_ctrl As Boolean


ItemStock_ctrl = False
WHStock_ctrl = False

QtyX = Format(QtyX, gs_formatQty)
OldQty = Format(OldQty, gs_formatQty)

Dim FlagU As Byte
        
'Proses Cek nya adalah dari warehouse dulu baru cek item nya
Set RsWHS = Db.Execute("Select stockcontrol_cls,WH_name from warehouse_master where WH_code='" & Trim$(OldWHCode) & "'")
    If Not RsWHS.EOF Then
        If RsWHS!stockcontrol_cls <> "01" Then
            'Stock tidak boleh di update
            Exit Sub
        Else
            'Stock boleh di update
            'Cek apakah item boleh  update stock atau tidak
            Set RSIS = Db.Execute("Select StockControl_cls from item_master where item_code='" & Trim$(ItemCode) & "'") ' and (stockcontrol_cls='01')")
            If Not RSIS.EOF Then
                If Trim$(RSIS!stockcontrol_cls) <> "01" Then
                    'Stock tidak boleh di update
                    Exit Sub
                End If
                '## Update Stock Master
                    updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(QtyX), 0
                
                ItemStock_ctrl = True
            Else
                ItemStock_ctrl = False
                Exit Sub
            End If
            RSIS.Close
            Set RSIS = Nothing
            
        End If
        WHStock_ctrl = True
    Else
        WHStock_ctrl = False
    End If
    RsWHS.Close
    Set RsWHS = Nothing

End Sub

Sub updateStock(WHCode As String, ItemCode As String, RecMonth As String, RecYear As String, QtyX As Double, OldQty As Double)
Dim RsI As Recordset
Dim DBi As New Connection
Dim FixMonth As String
Dim FixYear As String

Dim LM_P As Double, LM_S As Double, LM_RJ As Double, LM_I As Double
Dim TM_P As Double, TM_S As Double, TM_RJ As Double, TM_I As Double
Dim NM_P As Double, NM_S As Double, NM_RJ As Double, NM_I As Double

'Db.BeginTrans
If GetLastMonthStock = "" Then LblErr = DisplayMsg(4019): Exit Sub

FixMonth = Right(GetLastMonthStock, 2)
FixYear = Mid(GetLastMonthStock, 1, 4)

Set RsI = New Recordset
RsI.Open "select * from stock_master (updlock) where " & _
      " warehouse_code='" & Trim(WHCode) & "' and " & _
      " item_code='" & Trim(ItemCode) & "'", Db, adOpenDynamic, adLockOptimistic
If RsI.EOF Then
  RsI.AddNew
  'Receipt
  RsI!Item_Code = Trim$(ItemCode)
  RsI!Warehouse_Code = Trim$(WHCode)
  RsI!lm_premonth = "0"
  RsI!tm_premonth = "0"
  RsI!nm_premonth = "0"

  RsI!lm_supply = "0"
  RsI!tm_supply = "0"
  RsI!nm_supply = "0"
  RsI!lm_lossreject = "0"
  RsI!tm_lossreject = "0"
  RsI!nm_lossreject = "0"

  RsI!lm_current = "0"
  RsI!tm_current = "0"
  RsI!nm_current = "0"

  RsI!lm_inventory = Null
  RsI!tm_inventory = Null
  RsI!nm_inventory = Null

  If Val(FixYear) = Val(RecYear) Then
    RsI!lm_receipt = IIf(Val(RecMonth) = Val(FixMonth), QtyX, 0)
    RsI!tm_receipt = IIf(Val(RecMonth) = Val(FixMonth) + 1, QtyX, 0)
    RsI!nm_receipt = IIf(Val(RecMonth) = Val(FixMonth) + 2, QtyX, 0)
    'Current
    RsI!lm_current = Val(RsI!lm_premonth) + Val(RsI!lm_receipt) - Val(RsI!lm_supply) - Val(RsI!lm_lossreject)
    RsI!tm_current = Val(RsI!tm_premonth) + Val(RsI!tm_receipt) - Val(RsI!tm_supply) - Val(RsI!tm_lossreject)
    'Next Proses
    RsI!nm_premonth = RsI!tm_current
    RsI!nm_current = Val(RsI!nm_premonth) + Val(RsI!nm_receipt) - Val(RsI!nm_supply) - Val(RsI!nm_lossreject)
  ElseIf Val(FixYear) < Val(RecYear) Then
    RsI!tm_receipt = RsI!tm_receipt + IIf((Val(RecMonth) + 12) - Val(FixMonth) = 1, QtyX, 0)
    RsI!nm_receipt = RsI!nm_receipt + IIf((Val(RecMonth) + 12) - Val(FixMonth) = 2, QtyX, 0)
    'Current
    RsI!lm_current = Val(RsI!lm_premonth) + Val(RsI!lm_receipt) - Val(RsI!lm_supply) - Val(RsI!lm_lossreject)
    RsI!tm_current = Val(RsI!tm_premonth) + Val(RsI!tm_receipt) - Val(RsI!tm_supply) - Val(RsI!tm_lossreject)
    RsI!nm_current = Val(RsI!nm_premonth) + Val(RsI!nm_receipt) - Val(RsI!nm_supply) - Val(RsI!nm_lossreject)
  End If
  RsI!Last_Update = Now
  RsI!last_user = userLogin
  RsI.update
Else
  'LM Null offset
    If IsNull(RsI!lm_premonth) Then
        LM_P = 0
    Else
        LM_P = CDbl(RsI!lm_premonth)
    End If
    If IsNull(RsI!lm_supply) Then
        LM_S = 0
    Else
        LM_S = CDbl(RsI!lm_supply)
    End If
    If IsNull(RsI!lm_lossreject) Then
        LM_RJ = 0
    Else
        LM_RJ = CDbl(RsI!lm_lossreject)
    End If
    'TM Null Offset
    If IsNull(RsI!tm_premonth) Then
        TM_P = 0
    Else
        TM_P = CDbl(RsI!tm_premonth)
    End If
    If IsNull(RsI!tm_supply) Then
        TM_S = 0
    Else
        TM_S = CDbl(RsI!tm_supply)
    End If
    If IsNull(RsI!tm_lossreject) Then
        TM_RJ = 0
    Else
        TM_RJ = CDbl(RsI!tm_lossreject)
    End If
    'NM Null Offset
    If IsNull(RsI!nm_premonth) Then
        NM_P = 0
    Else
        NM_P = CDbl(RsI!nm_premonth)
    End If
    If IsNull(RsI!nm_supply) Then
        NM_S = 0
    Else
        NM_S = CDbl(RsI!nm_supply)
    End If
    If IsNull(RsI!nm_lossreject) Then
        NM_RJ = 0
    Else
        NM_RJ = CDbl(RsI!nm_lossreject)
    End If

  If Val(FixYear) = Val(RecYear) Then
    RsI!lm_receipt = RsI!lm_receipt - IIf(Val(RecMonth) = Val(FixMonth), (OldQty - QtyX), 0)
    RsI!tm_receipt = RsI!tm_receipt - IIf(Val(RecMonth) = Val(FixMonth) + 1, (OldQty - QtyX), 0)
    RsI!nm_receipt = RsI!nm_receipt - IIf(Val(RecMonth) = Val(FixMonth) + 2, (OldQty - QtyX), 0)
    'Current
    RsI!lm_current = (LM_P + Val(RsI!lm_receipt)) - LM_S - LM_RJ
    RsI!tm_current = (TM_P + Val(RsI!tm_receipt)) - TM_S - TM_RJ
    'Next Proses
    RsI!nm_premonth = RsI!tm_current
    NM_P = RsI!nm_premonth
    RsI!nm_current = (NM_P + Val(RsI!nm_receipt)) - NM_S - NM_RJ
  ElseIf Val(FixYear) < Val(RecYear) Then
    RsI!tm_receipt = RsI!tm_receipt - IIf((Val(RecMonth) + 12) - Val(FixMonth) = 1, (OldQty - QtyX), 0)
    RsI!nm_receipt = RsI!nm_receipt - IIf((Val(RecMonth) + 12) - Val(FixMonth) = 2, (OldQty - QtyX), 0)
    'Current
    RsI!lm_current = (LM_P + Val(RsI!lm_receipt)) - LM_S - LM_RJ
    RsI!tm_current = (TM_P + Val(RsI!tm_receipt)) - TM_S - TM_RJ
    RsI!nm_premonth = RsI!tm_current
    NM_P = RsI!nm_premonth
    RsI!nm_current = (NM_P + Val(RsI!nm_receipt)) - NM_S - NM_RJ
  End If
  RsI!Last_Update = Now
  RsI!last_user = userLogin
  RsI.update
End If

On Error GoTo erri

    RsI.Close
    Set RsI = Nothing
Exit Sub
erri:

    RsI.Close
    Set RsI = Nothing
    DBi.Close
    Set DBi = Nothing
End Sub

Sub up_SubConInputConsumption(ibu As String, Qty As Double, ls_WhCode As String)
Dim rsAnak As New ADODB.Recordset, rsc As New ADODB.Recordset
Dim fromWHCode As String, fromAddress As String, toWHCode As String, toAddress As String
Dim UnitCls As String, currCD As String, Price As Double, Amount As Double
Dim itemAnak As String, nilPrice As String, currAnak As String, qtyAnak As Double
Dim stockWH As String, stockItem As String, Sn As Double

    '*********Update Supply Anak2nya diambil dr BOM MaSter ***********
    sql = "Select c.Manufacture_Code as Factory_Code,a.Item_Code,a.Qty as qtyAnak,a.Unit_Cls," & _
        "b.WH_Code,b.Address,b.Stockcontrol_Cls as stockItem,(select Stockcontrol_Cls from warehouse_master where wh_code='" & ls_WhCode & "') as stockWH, b.Provision_Cls " & _
        "from BOM_Master a,Item_Master b, Item_Master c " & _
        "where a.Item_Code = b.Item_Code " & _
        "And a.Parent_ItemCode = c.Item_Code " & _
        "  " & _
        "And a.Parent_ItemCode = '" & ibu & _
        "' And Start_Date <='" & Format(TglReceipt, "yyyyMMdd") & _
        "' And End_Date >= '" & Format(TglReceipt, "yyyyMMdd") & "' order by a.Item_Code"
    Set rsAnak = newDb.Execute(sql)

    If Not rsAnak.EOF Then
        Do While Not rsAnak.EOF
            If KeyProd <> "R" Then
                KeyProd = "R"
            End If
        
            fromWHCode = rsAnak("Wh_Code")
            fromAddress = IIf(IsNull(rsAnak("Address")), "", rsAnak("Address"))
            toWHCode = IIf(IsNull(rsAnak("Factory_Code")), "", rsAnak("Factory_Code"))
            itemAnak = rsAnak("Item_Code")
            qtyAnak = rsAnak("QtyAnak") * CDbl(Qty)
            UnitCls = rsAnak("Unit_Cls")
            nilPrice = isiPrice(itemAnak, Format(TglReceipt, "yyyy-MM-dd"), currCD)
            currAnak = Split(nilPrice, ",")(0)
            Price = Split(nilPrice, ",")(1)
            Amount = CDbl(qtyAnak) * CDbl(Price)
            stockWH = rsAnak("StockWH")
            stockItem = rsAnak("StockItem")
            
            Set rsc = Db.Execute("Select isnull(Max(seq_No),0)+1 From Part_Supply")
            Sn = rsc(0)
            rsc.Close
            
            KeyProd = KeyProd & seqNo.keyReceipt
            sql = "insert into Part_Supply(FromWarehouse_Code,From_Address,ToWarehouse_Code,ChildSupply_date,ChildItem_Code,Supply_Cls," & _
                "ChildRequirement_Qty,ChildUnit_Cls,Currency_Code,Price,Amount,ParentItem_Code,Lot_No,Remarks,SubConPartReceipt_SeqNo,Do_NO," & _
                "Last_Update,Last_User) " & _
                "values ('" & ls_WhCode & "','" & "" & "','" & ls_WhCode & "','" & Format(TglReceipt, "yyyy-MM-dd") & "','" & itemAnak & "','S'," & _
                CDbl(qtyAnak) & ",'" & UnitCls & "','" & currAnak & "'," & Price & "," & Amount & ",'" & ItemCode & "','','','" & KeyProd & "', ''," & _
                "getdate(),'" & userLogin & "')"
            Db.Execute sql
            
            If stockWH = "01" And stockItem = "01" Then _
                Call seqNo.updateStock(ls_WhCode, itemAnak, qtyAnak, "", Format(TglReceipt, "yyyy-MM-dd"), blnFix, thnFix, Db, "Supply", 0, 1)
            If Not (rsAnak.EOF) Then rsAnak.MoveNext
        '******************n
        Loop
        '******************
End If
End Sub

Function isiPrice(ItemCode As String, tglDO As String, currCode As String) As String
    Dim rsPrice As New ADODB.Recordset
    
    sql = "select top 1 currency_code,isnull(price,0) Price from price_master where " & _
        "item_code='" & ItemCode & _
        "' and price_cls='01' " & _
        "and start_date<='" & Format(tglDO, "yyyymmdd") & _
        "' and end_date>='" & Format(tglDO, "yyyymmdd") & _
        "' order by trade_code desc, priority_cls desc"
        
    Set rsPrice = newDb.Execute(sql)
    If rsPrice.EOF Then
        isiPrice = currCode & ",0"
    Else
        isiPrice = Trim(rsPrice(0)) & "," & Trim(rsPrice(1))
    End If
    Set rsPrice = Nothing
End Function

Private Sub Label11_Click()

End Sub
