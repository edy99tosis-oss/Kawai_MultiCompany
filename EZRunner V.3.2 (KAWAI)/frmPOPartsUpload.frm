VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPOPartsUpload 
   BackColor       =   &H00FDDFE3&
   Caption         =   "PO Parts Upload"
   ClientHeight    =   10995
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15210
   Icon            =   "frmPOPartsUpload.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   15210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   29
      Tag             =   "FTTF*/"
      Top             =   10080
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   28
      Tag             =   "TTFF*/"
      Top             =   10080
      Width           =   1185
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
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   27
      Tag             =   "FTTF*/"
      Top             =   10080
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   360
      TabIndex        =   23
      Tag             =   "TTTF*/"
      Top             =   9240
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
         TabIndex        =   24
         Tag             =   "TTTF*/"
         Top             =   195
         Width           =   14370
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1275
      Left            =   360
      TabIndex        =   7
      Tag             =   "TTTF*/"
      Top             =   7920
      Width           =   14595
      Begin VB.Label lblQtyContract 
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
         TabIndex        =   33
         Tag             =   "TTFF*/"
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Qty Contract"
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
         TabIndex        =   32
         Tag             =   "TTFF*/"
         Top             =   960
         Width           =   1740
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Price"
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
         TabIndex        =   31
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label lblPrice 
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
         TabIndex        =   30
         Tag             =   "TTFF*/"
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
         Left            =   8760
         TabIndex        =   22
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   375
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
         Left            =   5640
         TabIndex        =   21
         Tag             =   "TTFF*/"
         Top             =   960
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
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblPODate 
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
         Tag             =   "TTFF*/"
         Top             =   240
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
         TabIndex        =   18
         Tag             =   "TTFF*/"
         Top             =   960
         Width           =   375
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
         TabIndex        =   17
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label8 
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
         Left            =   6720
         TabIndex        =   16
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label7 
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
         Left            =   3480
         TabIndex        =   15
         Tag             =   "TTFF*/"
         Top             =   960
         Width           =   1560
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
         TabIndex        =   14
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   1830
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid PO Date"
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
         TabIndex        =   13
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   1350
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
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label LblCode 
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
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   2115
      End
      Begin MSForms.Label Label5 
         Height          =   255
         Left            =   649
         TabIndex        =   10
         Tag             =   "TTFF*/"
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
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   1860
      End
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
         Left            =   2500
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FDDFE3&
      Height          =   780
      Left            =   360
      TabIndex        =   2
      Tag             =   "TTTF*/"
      Top             =   720
      Width           =   14625
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
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   280
         Width           =   7935
      End
      Begin MSForms.CommandButton btnTemplate 
         Height          =   375
         Left            =   10320
         TabIndex        =   6
         Tag             =   "TTFF*/"
         Top             =   240
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
      Begin MSForms.CommandButton btnUpload 
         Height          =   375
         Left            =   9720
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   240
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
      Begin MSForms.Label Label3 
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   310
         Width           =   1215
         BackColor       =   16637923
         Caption         =   "File Name"
         Size            =   "2143;450"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13080
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "FTTF*/"
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5805
      Left            =   360
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   "TTTF*/"
      Top             =   1680
      Width           =   14595
      _cx             =   25744
      _cy             =   10239
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
      Left            =   1920
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin EZRunnerv3.Anchor Anchor1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin MSForms.Label LblRecord 
      Height          =   255
      Left            =   11640
      TabIndex        =   26
      Tag             =   "FFTT*/"
      Top             =   7560
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
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order (Parts) Upload"
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
      TabIndex        =   1
      Tag             =   "TTTF*/"
      Top             =   240
      Width           =   15105
   End
End
Attribute VB_Name = "frmPOPartsUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TglReceipt, PODate, DelDate As Date
Dim JmlTran As Integer

Dim IInvalidSupplier, IInvalidWHCode, IInvalildPONo, IInvalildPODate, IInvalidDelDate, IInvalidItemCode, IInvalidQty, IInvalidPrice, IInvalidQtyContract As Integer
Dim IInvalidNoBC, IInvalidBCDate As Integer
Dim sql As String
Dim rsSp As New ADODB.Recordset
Dim Parm, Parm2, Parm3, Parm4, Parm5, Char As String
Dim SupplierCode, WHCode, PONO, ItemCode, Period, UnitCls, Curr, PriceContract  As String
Dim Amount, Price, TotalAmount As Double
Dim Qty As Long
Dim newDb As New ADODB.Connection

Dim bteColSupplierCode As Byte
Dim bteColWarehouseCode As Byte
Dim bteColPONo As Byte
Dim bteColPODate As Byte
Dim bteColDeliveryDate As Byte
Dim bteColItemCode As Byte
Dim bteColDesc As Byte
Dim bteColQtyUnit As Byte
Dim bteColQty As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColRemark As Byte
Dim bteColPriceContractCls As Byte

Dim ls_PathExcel As String

Private Sub up_GridHeader()

    Dim i As Long
    
    LblRecord = "0 Record(s)"
    
    bteColSupplierCode = 0
    bteColWarehouseCode = 1
    bteColPONo = 2
    bteColPODate = 3
    bteColDeliveryDate = 4
    bteColItemCode = 5
    bteColDesc = 6
    bteColQtyUnit = 7
    bteColQty = 8
    bteColCurr = 9
    bteColPrice = 10
    bteColAmount = 11
    bteColRemark = 12
    bteColPriceContractCls = 13
    
    With grid
        .clear
        .ColS = 14
        .Rows = 1
        
        .TextMatrix(0, bteColSupplierCode) = "Supplier Code"
        .TextMatrix(0, bteColWarehouseCode) = "WH Code"
        .TextMatrix(0, bteColPONo) = "PO No"
        .TextMatrix(0, bteColPODate) = "PO Date"
        .TextMatrix(0, bteColDeliveryDate) = "Delivery Date"
        .TextMatrix(0, bteColItemCode) = "Item Code"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColQtyUnit) = "Qty Unit"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColCurr) = "Currency"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColRemark) = "Remark"
        .TextMatrix(0, bteColPriceContractCls) = "PO Price Contract Cls"
        
        .ColWidth(bteColSupplierCode) = 1500
        .ColWidth(bteColWarehouseCode) = 1500
        .ColWidth(bteColPONo) = 2000
        .ColWidth(bteColPODate) = 2000
        .ColWidth(bteColDeliveryDate) = 2000
        .ColWidth(bteColItemCode) = 2000
        .ColWidth(bteColDesc) = 2500
        .ColWidth(bteColQtyUnit) = 850
        .ColWidth(bteColQty) = 1000
        .ColWidth(bteColCurr) = 1000
        .ColWidth(bteColPrice) = 1500
        .ColWidth(bteColAmount) = 1500
        .ColWidth(bteColRemark) = 4500
        .ColWidth(bteColPriceContractCls) = 500
        
        .ColHidden(bteColPriceContractCls) = True
                
        .ColAlignment(bteColSupplierCode) = flexAlignLeftCenter
        .ColAlignment(bteColWarehouseCode) = flexAlignLeftCenter
        .ColAlignment(bteColPONo) = flexAlignLeftCenter
        .ColAlignment(bteColPODate) = flexAlignLeftCenter
        .ColAlignment(bteColDeliveryDate) = flexAlignLeftCenter
        .ColAlignment(bteColItemCode) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColQtyUnit) = flexAlignCenterCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        .ColAlignment(bteColPrice) = flexAlignRightCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
        .ColAlignment(bteColRemark) = flexAlignLeftCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, bteColRemark) = flexAlignCenterCenter
        
        .EditMaxLength = 1
        
    End With

End Sub

Private Sub up_Clear()
    TglReceipt = Now()
    lblSupplierCode = "( 0 )"
    lblWHCode = "( 0 )"
    lblPONo = "( 0 )"
    lblPODate = "( 0 )"
    lblDelDate = "( 0 )"
    lblItemCode = "( 0 )"
    lblQty = "( 0 )"
    lblPrice = "( 0 )"
    
    IInvalidSupplier = 0
    IInvalidWHCode = 0
    IInvalildPONo = 0
    IInvalildPODate = 0
    IInvalidDelDate = 0
    IInvalidItemCode = 0
    IInvalidQty = 0
    IInvalidPrice = 0
    btnSubmit.Enabled = True
    
    CDExcel.filename = ""
    LblErr.Caption = ""
    txtUpload = ""
    up_GridHeader
End Sub

Private Sub btnSubmit_Click()
Dim X As Double
    
    Me.MousePointer = vbHourglass
    LblErr = ""
    SupplierCode = ""
    TotalAmount = 0
    
    If HakU = 0 Then LblErr = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
     For X = 1 To grid.Rows - 1
       Db.BeginTrans
        If SupplierCode = "" Then
            SupplierCode = grid.TextMatrix(X, bteColSupplierCode)
            PONO = grid.TextMatrix(X, bteColPONo)
            Period = Format(grid.TextMatrix(X, bteColPODate), "yyyyMM")
            PODate = grid.TextMatrix(X, bteColPODate)
            DelDate = grid.TextMatrix(X, bteColDeliveryDate)
            WHCode = grid.TextMatrix(X, bteColWarehouseCode)
            Amount = 0
            PriceContract = grid.TextMatrix(X, bteColPriceContractCls)
            
            up_UploadPOMaster
        End If
                
        PONO = grid.TextMatrix(X, bteColPONo)
        ItemCode = grid.TextMatrix(X, bteColItemCode)
        DelDate = grid.TextMatrix(X, bteColDeliveryDate)
        Curr = grid.TextMatrix(X, bteColCurr)
        Price = CDbl(grid.TextMatrix(X, bteColPrice))
        UnitCls = grid.TextMatrix(X, bteColQtyUnit)
        Qty = CDbl(grid.TextMatrix(X, bteColQty))
        Amount = grid.TextMatrix(X, bteColAmount)
        PriceContract = grid.TextMatrix(X, bteColPriceContractCls)
        
        up_UploadPODetail
        
        If err.number = 0 Then
            Db.CommitTrans
            up_UpdatePriceContract
        Else
            Db.RollbackTrans
            LblErr = err.Description
            err.clear
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        TotalAmount = TotalAmount + Amount
        
        up_UploadPOMaster
        
    Next X
    
    LblErr = DisplayMsg(1000)
    btnSubmit.Enabled = False
       
    
    Me.MousePointer = vbDefault
End Sub

Private Sub btnTemplate_Click()
Dim objExcel As New Excel.application
    Dim RS As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Double

    If G_CekExcelApp = False Then LblErr = DisplayMsg(5000): Exit Sub

    LblErr.Caption = ""
    CDExcel.filter = "Excel Files (*.xls)|*.xls"
    CDExcel.filename = "Purchase Order (Parts) Upload"
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
        .Range("A1:G1").Borders.Weight = xlThin
        .Range("A2:G2").Borders.Weight = xlThin
        .Rows("1:" & grid.Rows).Select
        .Selection.Interior.Pattern = xlNone
        
        .Range("A1:G1").Select
        .Selection.Font.Bold = True
        
        .Range("A2:G2").Select
        .Selection.Font.color = &H80FF80
        
        .Range("A1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("A1").Value = "Supplier Code"
        
        .Range("B1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("B1").Value = "Warehouse Code"

        .Range("C1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("C1").Value = "PO No"

        .Range("D1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("D1").Value = "PO Date"

        .Range("E1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("E1").Value = "Delivery Date"
         
        .Range("F1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("F1").Value = "Item Code"
        
        .Range("G1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("G1").Value = "Order Qty"
        
        .Range("A2").Value = "Char (25)"
        .Range("B2").Value = "Char (25)"
        .Range("C2").Value = "Char (25)"
        .Range("D2").Value = "Date (MM/dd/yyyy)"
        .Range("E2").Value = "Date (MM/dd/yyyy)"
        .Range("F2").Value = "Char (25)"
        .Range("G2").Value = "Numeric"
        
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
Dim ls_invalidMsg As String
Dim iGrdRow As Double
Dim rsUpdate As New ADODB.Recordset
Dim WHDesc As String
Dim StatusRow As Boolean
Dim pQty As Integer

Me.MousePointer = vbHourglass
    
    TglReceipt = Now()
    lblSupplierCode = "( 0 )"
    lblWHCode = "( 0 )"
    lblPONo = "( 0 )"
    lblPODate = "( 0 )"
    lblDelDate = "( 0 )"
    lblItemCode = "( 0 )"
    lblQty = "( 0 )"
    lblPrice = "( 0 )"
    
    IInvalidSupplier = 0
    IInvalidWHCode = 0
    IInvalildPONo = 0
    IInvalildPODate = 0
    IInvalidDelDate = 0
    IInvalidItemCode = 0
    IInvalidQty = 0
    IInvalidPrice = 0
    IInvalidQtyContract = 0
    
    btnSubmit.Enabled = True
    
    CDExcel.filename = ""
    LblErr.Caption = ""
    txtUpload.Text = ""
    ItemCode = ""
    
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

                       '1. Cek Supplier Code
                       Parm = Trim(.Cells(i, 1))
                       Parm2 = ""
                       Parm3 = ""
                       Parm4 = ""
                       Parm5 = "0"
                       Char = "Supplier"
                       up_ValidateUpload
                       
                       If rsSp.EOF = True Then
                            IInvalidSupplier = IInvalidSupplier + 1
                            StatusRow = False
                            ls_invalidMsg = "Supplier Code not register in Trade Master"
                            grid.TextMatrix(iGrdRow, bteColSupplierCode) = Trim(.Cells(i, 1))
                        Else
                            grid.TextMatrix(iGrdRow, bteColSupplierCode) = Trim(.Cells(i, 1))
                        End If
                        rsSp.Close

                       '2. Cek Warehouse Code
                       Parm = Trim(.Cells(i, 2))
                       Parm2 = ""
                       Parm3 = ""
                       Parm4 = ""
                       Parm5 = "0"
                       Char = "Warehouse"
                                              
                       up_ValidateUpload
                       If rsSp.EOF = True Then
                            IInvalidWHCode = IInvalidWHCode + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "Warehouse Code not register in Warehouse Master"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", Warehouse Code not register in Warehouse Master"
                            End If
                            grid.TextMatrix(iGrdRow, bteColWarehouseCode) = Trim(.Cells(i, 2))
                        Else
                            grid.TextMatrix(iGrdRow, bteColWarehouseCode) = Trim(.Cells(i, 2))
                        End If
                        rsSp.Close

                       '3. Cek PO No
'                       rs_PO.Open "SELECT PO_No FROM PurchaseOrder_Master WHERE PO_No='" & Trim(.Cells(i, 3)) & "'", Db, adOpenKeyset, adLockOptimistic
'                       If rs_PO.EOF = True Then
'                            IInvalildPONo = IInvalildPONo + 1
'                            StatusRow = False
'                            If ls_invalidMsg = "" Then
'                                ls_invalidMsg = "PO No Not Register In Purchase Order Master"
'                            Else
'                                ls_invalidMsg = ls_invalidMsg & ", PO No Not Register In Purchase Order Master"
'                            End If
'                        Else
                            grid.TextMatrix(iGrdRow, bteColPONo) = Trim(.Cells(i, 3))
'                        End If

                        '4. PO Date
                        If IsDate(Format(.Cells(i, 4), "dd MMM yyyy")) = False Then
                            IInvalildPODate = IInvalildPODate + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "Invalid PO Date Format"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", Invalid PO Date Format"
                            End If
                            grid.TextMatrix(iGrdRow, bteColPODate) = Format(.Cells(i, 4), "dd MMM yyyy")
                        Else
                            grid.TextMatrix(iGrdRow, bteColPODate) = Format(.Cells(i, 4), "dd MMM yyyy")
                        End If

                        '5. Delivery Date
                        If IsDate(Format(.Cells(i, 5), "dd MMM yyyy")) = False Then
                            IInvalidDelDate = IInvalidDelDate + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "Invalid Delivery Date Format"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", Invalid Delivery Date Format"
                            End If
                            grid.TextMatrix(iGrdRow, bteColDeliveryDate) = Format(.Cells(i, 5), "dd MMM yyyy")
                        Else
                            grid.TextMatrix(iGrdRow, bteColDeliveryDate) = Format(.Cells(i, 5), "dd MMM yyyy")
                        End If

                        '6. Cek Item Code
                         Parm = Trim(.Cells(i, 6))
                         Parm2 = ""
                         Parm3 = ""
                         Parm4 = ""
                         Parm5 = "0"
                         Char = "Item"
                         up_ValidateUpload
                        
                         If rsSp.EOF = True Then
                            IInvalidItemCode = IInvalidItemCode + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "Item Code not register in Item Master"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", Item Code not register in Item Master"
                            End If
                            grid.TextMatrix(iGrdRow, bteColItemCode) = Trim(.Cells(i, 6))
                         Else
                             grid.TextMatrix(iGrdRow, bteColItemCode) = Trim(rsSp!code)
                             grid.TextMatrix(iGrdRow, bteColDesc) = Trim(rsSp!Description)
                             grid.TextMatrix(iGrdRow, bteColQtyUnit) = Trim(rsSp!unit)
                         End If
                         rsSp.Close
'

                       '7. Cek Qty
                        Parm = Trim(.Cells(i, 6))
                        Parm2 = ""
                        Parm3 = ""
                        Parm4 = ""
                        Parm5 = "0"
                        Char = "Qty"
                        up_ValidateUpload
                        
                        If Format(CDbl(IsNull(rsSp!MinOrder)), 1) > Format(CDbl(.Cells(i, 7)), gs_formatQtyBOM) Then
                            IInvalidQty = IInvalidQty + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "Invalid Qty Format"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", Invalid Qty Format"
                            End If
                            grid.TextMatrix(iGrdRow, bteColQty) = Format(CDbl(.Cells(i, 7)), gs_formatQtyBOM)
                        Else
                            grid.TextMatrix(iGrdRow, bteColQty) = Format(CDbl(.Cells(i, 7)), gs_formatQtyBOM)
                        End If
                        rsSp.Close
                        
                        '8. Cek Price
                         Parm = Trim(.Cells(i, 6))
                         Parm2 = Trim(.Cells(i, 1))
                         Parm3 = Format(.Cells(i, 4), "yyyyMM")
                         Parm4 = Format(.Cells(i, 4), "yyyyMM")
                         Parm5 = "0"
                         Char = "CURR"
                         up_ValidateUpload
                        
                         If rsSp.EOF = True Then
                            IInvalidPrice = IInvalidPrice + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "Invalid Price"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", Invalid Price"
                            End If
                            grid.TextMatrix(iGrdRow, bteColCurr) = Trim(.Cells(i, 6))
                         Else
                             grid.TextMatrix(iGrdRow, bteColCurr) = Trim(rsSp!Description)
                         End If
                         rsSp.Close
                        
                        '8. Cek Price
                         Parm = Trim(.Cells(i, 6))
                         Parm2 = Trim(.Cells(i, 1))
                         Parm3 = Format(.Cells(i, 4), "yyyyMM")
                         Parm4 = Format(.Cells(i, 4), "yyyyMM")
                         Parm5 = "0"
                         Char = "Price"
                         up_ValidateUpload
                        
                         If rsSp.EOF = True Then
                            IInvalidPrice = IInvalidPrice + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "Invalid Price"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", Invalid Price"
                            End If
                            grid.TextMatrix(iGrdRow, bteColPrice) = Trim(.Cells(i, 6))
                         Else
                             grid.TextMatrix(iGrdRow, bteColPrice) = Format(CDbl(rsSp!Price), gs_formatPrice)
                         End If
                         rsSp.Close
                         
                         '9. Cek Price Contract Cls
                         Parm = Trim(.Cells(i, 6))
                         Parm2 = Trim(.Cells(i, 1))
                         Parm3 = Format(.Cells(i, 4), "yyyyMM")
                         Parm4 = Format(.Cells(i, 4), "yyyyMM")
                         Parm5 = "0"
                         Char = "ContractCls"
                         up_ValidateUpload
                        
                         If rsSp.EOF = True Then
                            IInvalidPrice = IInvalidPrice + 1
                            StatusRow = False
                            If ls_invalidMsg = "" Then
                                ls_invalidMsg = "Invalid Price"
                            Else
                                ls_invalidMsg = ls_invalidMsg & ", Invalid Price"
                            End If
                            grid.TextMatrix(iGrdRow, bteColPriceContractCls) = ""
                         Else
                             grid.TextMatrix(iGrdRow, bteColPriceContractCls) = Trim(rsSp!PriceContractCls)
                         End If
                         rsSp.Close
                         
                         If grid.TextMatrix(iGrdRow, bteColPriceContractCls) <> "0" Then
                            '10. Cek Qty PO Price Contract
                            
                            If ItemCode <> "" Then
                                If ItemCode = grid.TextMatrix(iGrdRow, bteColItemCode) Then
                                    pQty = pQty + CDbl(grid.TextMatrix(iGrdRow, bteColQty))
'                                Else
'                                    pQty = 0
'                                    pQty = pQty + Grid.TextMatrix(iGrdRow, bteColItemCode)
                                End If
                            Else
                                pQty = 0
                            End If
                            
                             Parm = Trim(.Cells(i, 6))
                             Parm2 = Trim(.Cells(i, 1))
                             Parm3 = Format(.Cells(i, 4), "yyyyMM")
                             Parm4 = Format(.Cells(i, 4), "yyyyMM")
                             Parm5 = pQty
                             Char = "CekQtyContract"
                             up_ValidateUpload
    
                             If CDbl(rsSp!QtyRemaining) < 0 Then
                                IInvalidQtyContract = IInvalidQtyContract + 1
                                StatusRow = False
                                If ls_invalidMsg = "" Then
                                    ls_invalidMsg = "Invalid PO Qty Contract"
                                Else
                                    ls_invalidMsg = ls_invalidMsg & ", Invalid PO Qty Contract"
                                End If
                                grid.TextMatrix(iGrdRow, bteColQty) = "0"
                             End If
                             rsSp.Close
                             
                             ItemCode = grid.TextMatrix(iGrdRow, bteColItemCode)
                         End If
                                                  
                         grid.TextMatrix(iGrdRow, bteColAmount) = Format(CDbl(.Cells(i, 7)) * grid.TextMatrix(iGrdRow, bteColPrice), gs_formatAmount)
                        

                        If StatusRow = False Then
                            grid.Cell(flexcpBackColor, iGrdRow, bteColSupplierCode, iGrdRow, bteColRemark) = vbRed
                            grid.TextMatrix(iGrdRow, bteColRemark) = ls_invalidMsg
                            btnSubmit.Enabled = False
                        End If
                        iGrdRow = iGrdRow + 1

                    DoEvents
                    i = i + 1

                    JmlTran = JmlTran + 1
            Loop
            
            lblSupplierCode.Caption = "( " & IInvalidSupplier & " )"
            lblWHCode.Caption = "( " & IInvalidWHCode & " )"
            lblPONo.Caption = "( " & IInvalildPONo & " )"
            lblPODate.Caption = "( " & IInvalildPODate & " )"
            lblDelDate.Caption = "( " & IInvalidDelDate & " )"
            lblItemCode.Caption = "( " & IInvalidItemCode & " )"
            lblQty.Caption = "( " & IInvalidQty & " )"
            lblQtyContract.Caption = "( " & IInvalidQtyContract & " )"
            
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

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    HakU = hakUpdate(Me.Name)
    
    If newDb.State <> adStateClosed Then newDb.Close
    newDb.Open Db.ConnectionString
    
    up_GridHeader
    
     With Anchor1
      .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
      .DoInit
    End With
    
End Sub

Private Sub btnCancel_Click()
    up_Clear
End Sub

Private Sub CmdMenu_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Public Sub up_ValidateUpload()
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_Validateupload"
    
    cmd.Parameters.append cmd.CreateParameter("Char", adVarChar, adParamInput, 55, Char)
    cmd.Parameters.append cmd.CreateParameter("Param1", adVarChar, adParamInput, 100, Parm)
    cmd.Parameters.append cmd.CreateParameter("Param2", adVarChar, adParamInput, 100, Parm2)
    cmd.Parameters.append cmd.CreateParameter("Param3", adVarChar, adParamInput, 100, Parm3)
    cmd.Parameters.append cmd.CreateParameter("Param4", adVarChar, adParamInput, 100, Parm4)
    cmd.Parameters.append cmd.CreateParameter("Param5", adVarChar, adParamInput, 100, Parm5)
    
    Set rsSp = cmd.Execute

End Sub

Public Sub up_UploadPOMaster()
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
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_POUploadHeader"
    
    Set prm1 = cmd.CreateParameter("PONo", adVarChar, adParamInput, 25, PONO)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("SupplierCode", adVarChar, adParamInput, 15, SupplierCode)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("Period", adVarChar, adParamInput, 6, Period)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("PODate", adDate, adParamInput, , PODate)
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("DeliveryDate", adDate, adParamInput, , DelDate)
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("WHTo", adVarChar, adParamInput, 15, WHCode)
    cmd.Parameters.append prm6
    Set prm7 = cmd.CreateParameter("Amount", adDouble, adParamInput, , TotalAmount)
    prm7.Precision = 18
    prm7.NumericScale = 2
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("TotalAmount", adDouble, adParamInput, , TotalAmount)
    prm8.Precision = 18
    prm8.NumericScale = 2
    cmd.Parameters.append prm8
    Set prm10 = cmd.CreateParameter("PriceContract", adChar, adParamInput, 1, PriceContract)
    cmd.Parameters.append prm10
    Set prm9 = cmd.CreateParameter("UserID", adVarChar, adParamInput, 15, userLogin)
    cmd.Parameters.append prm9
        
    cmd.Execute

End Sub

Public Sub up_UploadPODetail()
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
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_POUploadDetail"
    
    Set prm1 = cmd.CreateParameter("PONo", adChar, adParamInput, 25, PONO)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("ItemCode", adChar, adParamInput, 25, ItemCode)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("DeliveryDate", adDBDate, adParamInput, , DelDate)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("Price", adDouble, adParamInput, , Price)
    prm4.Precision = 18
    prm4.NumericScale = 2
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("UnitCls", adVarChar, adParamInput, 10, UnitCls)
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("Qty", adDouble, adParamInput, , Qty)
    prm6.Precision = 18
    prm6.NumericScale = 2
    cmd.Parameters.append prm6
    Set prm7 = cmd.CreateParameter("Amount", adDouble, adParamInput, , Amount)
    prm7.Precision = 22
    prm7.NumericScale = 5
    cmd.Parameters.append prm7
    Set prm9 = cmd.CreateParameter("Curr", adChar, adParamInput, 25, Curr)
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("PriceContract", adChar, adParamInput, 1, PriceContract)
    cmd.Parameters.append prm10
    Set prm8 = cmd.CreateParameter("UserID", adChar, adParamInput, 15, userLogin)
    cmd.Parameters.append prm8
        
    cmd.Execute
    
End Sub

Public Sub up_UpdatePriceContract()
    Dim cmd As ADODB.Command
    Dim prm1 As ADODB.Parameter
    Dim prm2 As ADODB.Parameter
    Dim prm3 As ADODB.Parameter
    Dim prm4 As ADODB.Parameter
    Dim prm5 As ADODB.Parameter
    Dim prm6 As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_PriceMasterContract_Update"

    Set prm1 = cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 25, RTrim(ItemCode))
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("TradeCode", adVarChar, adParamInput, 15, RTrim(SupplierCode))
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("StartDate", adVarChar, adParamInput, 8, Format(DelDate, "YYYYMMDD"))
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("Type", adVarChar, adParamInput, 1, "1")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("Qty", adDouble, adParamInput, , Qty)
    prm5.Precision = 18
    prm5.NumericScale = 2
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("User", adVarChar, adParamInput, 15, RTrim(userLogin))
    cmd.Parameters.append prm6

    cmd.Execute
End Sub

