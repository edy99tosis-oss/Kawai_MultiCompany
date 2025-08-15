VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPOContract_Mst 
   BackColor       =   &H00FDDFE3&
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15465
   Icon            =   "FrmPOContract_Mst.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   15465
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExcel 
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   26
      Tag             =   "FFTT*/"
      Top             =   10080
      Width           =   1140
   End
   Begin EZRunnerv3.Anchor Anchor1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin VB.CommandButton CmdSubmit 
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
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "FFTT*/"
      Top             =   10080
      Width           =   1140
   End
   Begin VB.CommandButton CmdClear 
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "FFTT*/"
      Top             =   10080
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "FFTT*/"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdSubMenu 
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
      TabIndex        =   18
      Tag             =   "TFFT*/"
      Top             =   10080
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   120
      TabIndex        =   16
      Tag             =   "TFTT*/"
      Top             =   9360
      Width           =   15165
      Begin VB.Label lblErrMsg 
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
         Left            =   120
         TabIndex        =   17
         Tag             =   "TFTF*/"
         Top             =   195
         Width           =   14925
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   1635
      Left            =   120
      TabIndex        =   2
      Tag             =   "TTTF*/"
      Top             =   1080
      Width           =   15165
      Begin VB.TextBox lblCust 
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
         Height          =   240
         Left            =   4920
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "TTFF*/"
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFFFFF&
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
         Left            =   11280
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   2500
      End
      Begin VB.ComboBox cboStatus 
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
         ItemData        =   "FrmPOContract_Mst.frx":0E42
         Left            =   135
         List            =   "FrmPOContract_Mst.frx":0E4C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker DtContractDate 
         Height          =   315
         Left            =   11280
         TabIndex        =   6
         Tag             =   "TTFF*/"
         Top             =   645
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   124715011
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DTPOFrom 
         Height          =   315
         Left            =   3240
         TabIndex        =   22
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   124715011
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DTPOTo 
         Height          =   315
         Left            =   5280
         TabIndex        =   24
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   124715011
         CurrentDate     =   37798
      End
      Begin VB.CommandButton CmdCreate 
         BackColor       =   &H0080FFFF&
         Caption         =   "Create"
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
         Left            =   13890
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label lblCaption 
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
         Index           =   2
         Left            =   4920
         TabIndex        =   25
         Tag             =   "TTFF*/"
         Top             =   290
         Width           =   165
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO Date"
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
         Left            =   1680
         TabIndex        =   23
         Tag             =   "TTFF*/"
         Top             =   290
         Width           =   705
      End
      Begin VB.Line Line2 
         X1              =   4920
         X2              =   9720
         Y1              =   960
         Y2              =   960
      End
      Begin MSForms.ComboBox cboCust 
         Height          =   315
         Left            =   3240
         TabIndex        =   14
         Tag             =   "TTFF*/"
         Top             =   645
         Width           =   1530
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "2699;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
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
         Left            =   1680
         TabIndex        =   13
         Tag             =   "TTFF*/"
         Top             =   720
         Width           =   1350
      End
      Begin MSForms.ComboBox cboContractNo 
         Height          =   315
         Left            =   11280
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   2500
         VariousPropertyBits=   612386843
         MaxLength       =   50
         DisplayStyle    =   3
         Size            =   "4410;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contract Date"
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
         Index           =   11
         Left            =   9960
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   720
         Width           =   1200
      End
      Begin MSForms.ComboBox cboPONo 
         Height          =   315
         Left            =   3240
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   3570
         VariousPropertyBits=   612386843
         MaxLength       =   50
         DisplayStyle    =   3
         Size            =   "6297;556"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remark"
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
         Index           =   9
         Left            =   9960
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   1160
         Width           =   675
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contract No."
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
         Index           =   8
         Left            =   9960
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   290
         Width           =   1080
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO No."
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
         Index           =   7
         Left            =   1680
         TabIndex        =   7
         Tag             =   "TTFF*/"
         Top             =   1160
         Width           =   585
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13440
      TabIndex        =   0
      Tag             =   "FTTF*/"
      Top             =   120
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   6390
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   2880
      Width           =   15195
      _cx             =   26802
      _cy             =   11271
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
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Contract"
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
      Tag             =   "TTTF*/"
      Top             =   480
      Width           =   15165
   End
End
Attribute VB_Name = "FrmPOContract_Mst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Const STATUS_CREATE As String = "Create"
Private Const CONTRACT_STATUS_NEW As String = "0"
Private Const CONTRACT_STATUS_EXISTING As String = "1"

Dim bteColSelect As Byte
Dim bteColParentItemCode As Byte
Dim bteColItemCode As Byte
Dim bteColHSCode As Byte
Dim bteColItemName As Byte
Dim bteColReceiptDate As Byte
Dim bteColCurrCode As Byte
Dim bteColCurrName As Byte
Dim bteColRate As Byte
Dim bteColQtyBOM As Byte
Dim bteColQtyPO As Byte
Dim bteColQtySet As Byte
Dim bteColUnitCode As Byte
Dim bteColUnitName As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColBctype As Byte
Dim bteColBcNo As Byte
Dim bteColBCDate As Byte
Dim bteColTot As Byte

Private Sub headerGrid()
    bteColSelect = 0
    bteColParentItemCode = 1
    bteColItemCode = 2
    bteColHSCode = 3
    bteColItemName = 4
    bteColReceiptDate = 5
    bteColCurrCode = 6
    bteColCurrName = 7
    bteColRate = 8
    bteColQtyBOM = 9
    bteColQtyPO = 10
    bteColQtySet = 11
    bteColUnitCode = 12
    bteColUnitName = 13
    bteColPrice = 14
    bteColAmount = 15
    bteColBctype = 16
    bteColBcNo = 17
    bteColBCDate = 18
    bteColTot = 19

    With grid
        .ColS = bteColTot
        .clear
        
        .ColDataType(bteColSelect) = flexDTBoolean
    
        .Rows = 1

        .Cell(flexcpChecked, 0, bteColSelect) = flexUnchecked
        .TextMatrix(0, bteColSelect) = " "
        .TextMatrix(0, bteColParentItemCode) = "Parent Item Code"
        .TextMatrix(0, bteColItemCode) = "Item Code"
        .TextMatrix(0, bteColHSCode) = "HS Code"
        .TextMatrix(0, bteColItemName) = "Item Name"
        .TextMatrix(0, bteColReceiptDate) = "Receipt Date"
        .TextMatrix(0, bteColCurrCode) = "Curr. Code"
        .TextMatrix(0, bteColCurrName) = "Curr. Name"
        .TextMatrix(0, bteColRate) = "Exc. Rate"
        .TextMatrix(0, bteColQtyBOM) = "Qty BOM"
        .TextMatrix(0, bteColQtyPO) = "Qty PO"
        .TextMatrix(0, bteColQtySet) = "Qty Set"
        .TextMatrix(0, bteColUnitCode) = "Unit Code"
        .TextMatrix(0, bteColUnitName) = "Unit Name"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColBctype) = "BC Type"
        .TextMatrix(0, bteColBcNo) = "BC No."
        .TextMatrix(0, bteColBCDate) = "BC Date"

        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColParentItemCode) = 1750
        .ColWidth(bteColItemCode) = 1750
        .ColWidth(bteColHSCode) = 1400
        .ColWidth(bteColItemName) = 3500
        .ColWidth(bteColReceiptDate) = 1400
        .ColWidth(bteColCurrCode) = 1150
        .ColWidth(bteColCurrName) = 1150
        .ColWidth(bteColRate) = 1100
        .ColWidth(bteColQtyBOM) = 1100
        .ColWidth(bteColQtyPO) = 1000
        .ColWidth(bteColQtySet) = 1500
        .ColWidth(bteColUnitCode) = 1150
        .ColWidth(bteColUnitName) = 1150
        .ColWidth(bteColPrice) = 1200
        .ColWidth(bteColAmount) = 1200
        .ColWidth(bteColBctype) = 1000
        .ColWidth(bteColBcNo) = 1000
        .ColWidth(bteColBCDate) = 1200
        
        For i = 0 To .ColS - 1
            .ColAlignment(i) = flexAlignLeftTop ' default: kiri
        Next i

        .ColAlignment(bteColParentItemCode) = flexAlignLeftTop
        .ColAlignment(bteColItemCode) = flexAlignLeftTop
        .ColAlignment(bteColHSCode) = flexAlignLeftTop
        .ColAlignment(bteColItemName) = flexAlignLeftTop
        .ColAlignment(bteColReceiptDate) = flexAlignCenterCenter
        .ColAlignment(bteColCurrCode) = flexAlignCenterCenter
        .ColAlignment(bteColCurrName) = flexAlignLeftTop
        .ColAlignment(bteColRate) = flexAlignRightCenter
        .ColAlignment(bteColQtyBOM) = flexAlignRightCenter
        .ColAlignment(bteColQtyPO) = flexAlignRightCenter
        .ColAlignment(bteColQtySet) = flexAlignRightCenter
        .ColAlignment(bteColUnitCode) = flexAlignCenterCenter
        .ColAlignment(bteColUnitName) = flexAlignLeftCenter
        .ColAlignment(bteColPrice) = flexAlignRightCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
        .ColAlignment(bteColBctype) = flexAlignCenterCenter
        .ColAlignment(bteColBcNo) = flexAlignCenterCenter
        .ColAlignment(bteColBCDate) = flexAlignCenterCenter
        
        For i = 0 To .ColS - 1
            .Row = 0
            .Col = i
            .CellAlignment = flexAlignCenterCenter
        Next i
        
        .Editable = flexEDKbdMouse
        
    End With
End Sub

Private Sub cboContractNo_Change()
Dim RS As New ADODB.Recordset
    lblErrMsg.Caption = ""
    
    '**** GetDate
    sql = "EXEC dbo.sp_POContract_FillCombo @Param1 = '" & cboContractNo.Text & "', " & _
        "@Param2 = '',  @Param3 = '', @Param4 = '', @Type = 'GetDate'  "
        
    Set RS = Db.Execute(sql)
    
    If Not RS.EOF Then
        If Not IsNull(RS("Contract_Date")) Then
            DtContractDate.Value = RS("Contract_Date")
        End If
    End If
    
    
    headerGrid
End Sub

Private Sub CboCust_Change()
    If cboCust.ListIndex >= 0 Then
        lblCust.Text = cboCust.List(cboCust.ListIndex, 1)
        lblErrMsg.Caption = ""
        
        isiCboPO
    Else
        lblCust.Text = ""
    End If
End Sub

Private Sub CboPOnO_Change()
    Dim i As Integer
    Dim matchFound As Boolean

    lblErrMsg.Caption = ""
    matchFound = False

    If cboPONo.ListCount > 0 Then
        For i = 0 To cboPONo.ListCount - 1
            If cboPONo.Text = cboPONo.List(i) Then
                matchFound = True
                Exit For
            End If
        Next i

        If Not matchFound Then
            lblErrMsg.Caption = "PO Number tidak tersedia di list."
            cboPONo.Text = ""
            Exit Sub
        End If
    End If

    headerGrid
    isiCboContractNo
End Sub

Private Sub cboStatus_Click()
    If cboStatus.Text = "Create" Then
        CmdCreate.Caption = "Create"
    Else
        CmdCreate.Caption = "Search"
     End If
End Sub

Private Sub cmdClear_Click()
    clear
End Sub

Private Sub CmdCreate_Click()
    lblErrMsg.Caption = ""
    
    If Not IsValidInput() Then Exit Sub
    Me.MousePointer = vbHourglass
    gridLoad
    Me.MousePointer = vbDefault
End Sub

Private Sub CmdExcel_Click()
Dim xlapp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim iRow As Long, iCol As Long
    Dim iGridRow As Long, iGridCol As Long
    Dim actualCol As Long
    Dim lastCol As Long, lastRow As Long

    On Error GoTo errHandler
    
    If grid.Rows <= 1 Then
        MsgBox "Tidak ada data untuk diekspor.", vbExclamation
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Set xlapp = CreateObject("Excel.Application")
    Set xlBook = xlapp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)

    ' Header export
    actualCol = 1
    For iGridCol = 0 To grid.ColS - 1
        If iGridCol <> bteColSelect Then
            xlSheet.Cells(1, actualCol).Value = grid.TextMatrix(0, iGridCol)
            actualCol = actualCol + 1
        End If
    Next iGridCol
    lastCol = actualCol - 1

    ' Isi data
    For iGridRow = 1 To grid.Rows - 1
        actualCol = 1
        For iGridCol = 0 To grid.ColS - 1
            If iGridCol <> bteColSelect Then
                xlSheet.Cells(iGridRow + 1, actualCol).Value = grid.TextMatrix(iGridRow, iGridCol)
                actualCol = actualCol + 1
            End If
        Next iGridCol
    Next iGridRow
    lastRow = grid.Rows

    ' Rata tengah
    With xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, lastCol))
        .Font.Bold = True
        .horizontalAlignment = -4108 ' xlCenter
    End With

    With xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(lastRow, lastCol)).Borders
        .LineStyle = 1 ' xlContinuous
        .Weight = 2   ' xlThin
    End With

    xlapp.Visible = True
    
    Me.MousePointer = vbDefault
    
    lblErrMsg = "[0000] Export to Excel completed successfully" ' "Data berhasil disimpan"
    

    Exit Sub

errHandler:
    Me.MousePointer = vbDefault
    MsgBox "Error: " & err.Description, vbCritical
    On Error Resume Next
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlapp = Nothing
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Sub clear()
    AddToComboSupplier
    cboStatus.Text = "Update"
    DTPOFrom.Value = Format(Now, "dd MMM yyyy")
    DTPOTo.Value = Format(Now, "dd MMM yyyy")
    
    DtContractDate.Value = Format(Now, "dd MMM yyyy")
    
    isiCboPO
    
    headerGrid
End Sub

Private Sub CmdSubmit_Click()
    If Not IsValidInput(True) Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    SaveContract
    
    Me.MousePointer = vbDefault
End Sub

Private Sub DTPOFrom_Change()
    lblErrMsg.Caption = ""
    headerGrid
    
    isiCboPO
End Sub

Private Sub DTPOTo_Change()
    lblErrMsg.Caption = ""
    headerGrid
    
    isiCboPO
End Sub

Private Sub Form_Load()
 If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    clear

With Anchor1
  .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
  .DoInit
End With

End Sub

Private Sub gridLoad()
    Dim RsIsiG As New ADODB.Recordset
    Dim sql As String
    Dim i As Long
    Dim lsStatus As String

    If cboStatus.Text = "Create" Then
        lsStatus = "0"
    Else
        lsStatus = "1"
    End If
    
    Call headerGrid
    
    sql = "EXEC dbo.sp_POContract_GridLoad " & vbCrLf & _
          " @PONo = '" & cboPONo.Text & "'," & vbCrLf & _
          " @Period = '" & DtContractDate.Value & "'," & vbCrLf & _
          " @ExplosionType = 0, @Status = '" & lsStatus & "'"

    If RsIsiG.State = 1 Then RsIsiG.Close
    
    ' Tambahkan timeout (dalam detik)
    RsIsiG.ActiveConnection = Db
    RsIsiG.ActiveConnection.CommandTimeout = 600  ' 5 menit

    RsIsiG.Open sql, , adOpenKeyset, adLockOptimistic

    If Not RsIsiG.EOF Then
        i = 0
        Do While Not RsIsiG.EOF
            i = i + 1
            grid.Rows = i + 1
            
            If RsIsiG("Idxx") = "0" Then
                grid.Cell(flexcpChecked, i, bteColSelect) = flexChecked
            Else
                grid.Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
            End If
            
            grid.TextMatrix(i, bteColParentItemCode) = Trim(RsIsiG("Parent_ItemCode") & "")
            grid.TextMatrix(i, bteColItemCode) = Trim(RsIsiG("Item_Code") & "")
            grid.TextMatrix(i, bteColHSCode) = Trim(RsIsiG("HS_Code") & "")
            grid.TextMatrix(i, bteColItemName) = Trim(RsIsiG("Item_Name") & "")
            grid.TextMatrix(i, bteColReceiptDate) = Trim(RsIsiG("Receipt_Date") & "")
            grid.TextMatrix(i, bteColCurrCode) = Trim(RsIsiG("Currency_Code") & "")
            grid.TextMatrix(i, bteColCurrName) = Trim(RsIsiG("Currency_Name") & "")
            grid.TextMatrix(i, bteColRate) = Format(RsIsiG("Rate"), gs_formatPriceIDR)
            grid.TextMatrix(i, bteColQtyBOM) = Format(RsIsiG("QtyBOM"), gs_formatQtyBOM)
            grid.TextMatrix(i, bteColQtyPO) = Format(RsIsiG("QtyPO"), gs_formatQty)
            grid.TextMatrix(i, bteColQtySet) = Format(RsIsiG("QtySet"), gs_formatQtyBOM)
            grid.TextMatrix(i, bteColUnitCode) = Trim(RsIsiG("Unit_Code") & "")
            grid.TextMatrix(i, bteColUnitName) = Trim(RsIsiG("Unit_Name") & "")
            grid.TextMatrix(i, bteColPrice) = Format(RsIsiG("Price"), gs_formatPrice)
            grid.TextMatrix(i, bteColAmount) = Format(RsIsiG("Amount"), gs_formatAmount)
            grid.TextMatrix(i, bteColBctype) = Trim(RsIsiG("BC_Type") & "")
            grid.TextMatrix(i, bteColBcNo) = Trim(RsIsiG("BC40_No") & "")
            grid.TextMatrix(i, bteColBCDate) = Trim(RsIsiG("BC40_Date") & "")
            
            With grid
                .Row = i
                .Col = bteColQtyBOM
                .CellBackColor = vbWhite
            End With
            
            RsIsiG.MoveNext
        Loop
        grid.Cell(flexcpChecked, 0, bteColSelect) = flexUnchecked
    End If
End Sub

Sub AddToComboSupplier()
    
    Dim sqlcust As String
    Dim RsCust As New Recordset

    sqlcust = "SELECT RTRIM(Trade_Code) Trade_Code, RTRIM(Trade_Name) Trade_Name FROM Trade_Master " & _
        "WHERE Trade_Cls = '2' OR Trade_Cls = '3'"
        
    Set RsCust = Db.Execute(sqlcust)
    
    With cboCust
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;275pt"
        .ListWidth = 325
        .ListRows = 15
        i = 0
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("Trade_Code"))
            .List(i, 1) = IIf(IsNull(RsCust("Trade_Name")), " ", Trim(RsCust("Trade_Name")))
            
            RsCust.MoveNext
            i = i + 1
        Loop
        RsCust.Close
    End With
    
End Sub

Sub isiCboPO()
Dim rscbo As New ADODB.Recordset
    
With cboPONo
    .clear
    .columnCount = 1
    .TextColumn = 1
    
    '**** PONo
    sql = "EXEC dbo.sp_POContract_FillCombo @Param1 = '" & cboCust.Text & "', " & _
        "@Param2 = '" & DTPOFrom.Value & "',  @Param3 = '" & DTPOTo & "', @Param4 = '" & cboStatus.Text & "', @Type = 'PONo'  "
        
    Set rscbo = Db.Execute(sql)

    i = 0
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo("PO_No"))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 175
    .ColumnWidths = "175 pt"
    Set rscbo = Nothing
End With
End Sub

Sub isiCboContractNo()
Dim rscbo As New ADODB.Recordset
    
With cboContractNo
    .clear
    .columnCount = 1
    .TextColumn = 1
    
    '**** DO Master
    sql = "EXEC dbo.sp_POContract_FillCombo @Param1 = '" & cboPONo.Text & "', " & _
        "@Param2 = '',  @Param3 = '', @Param4 = '', @Type = 'ContractNo'  "
        
    Set rscbo = Db.Execute(sql)

    i = 0
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo("Contract_No"))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 125
    .ColumnWidths = "125 pt"
    Set rscbo = Nothing
End With
End Sub


Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = bteColQtyBOM Then
        If grid.Cell(flexcpChecked, Row, bteColSelect) = flexChecked Then
            Cancel = False
        Else
            Cancel = True
            MsgBox "Please select a row before revising the BOM quantity", vbExclamation
        End If
    ElseIf Col = bteColSelect Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     Dim i As Long

    ' === Handle kolom Select (checkbox) ===
    If Col = bteColSelect Then
        ' Jika Row 0 (header), maka Select All / Unselect All
        If Row = 0 Then
            Dim isChecked As Boolean
            isChecked = (grid.Cell(flexcpChecked, 0, bteColSelect) = flexChecked)

            For i = 1 To grid.Rows - 1
                grid.Cell(flexcpChecked, i, bteColSelect) = IIf(isChecked, flexChecked, flexUnchecked)
            Next i
        Else
            ' Jika salah satu detail baris di-uncheck, header juga harus uncheck
            If grid.Cell(flexcpChecked, Row, bteColSelect) = flexUnchecked Then
                grid.Cell(flexcpChecked, 0, bteColSelect) = flexUnchecked
            Else
                ' Cek apakah semua baris detail dicheck
                Dim semuaChecked As Boolean: semuaChecked = True
                For i = 1 To grid.Rows - 1
                    If grid.Cell(flexcpChecked, i, bteColSelect) <> flexChecked Then
                        semuaChecked = False
                        Exit For
                    End If
                Next i

                If semuaChecked Then
                    grid.Cell(flexcpChecked, 0, bteColSelect) = flexChecked
                End If
            End If
        End If
    End If

    ' === Handle Qty BOM ===
    If Col = bteColQtyBOM Then
        If grid.Cell(flexcpChecked, Row, bteColSelect) = flexChecked Then
            ' Validasi apakah inputan adalah numeric
            If Not IsNumeric(grid.TextMatrix(Row, bteColQtyBOM)) Then
                MsgBox "Input Qty BOM harus berupa angka.", vbExclamation
                grid.TextMatrix(Row, bteColQtyBOM) = ""
                Exit Sub
            End If
    
            Dim qtyBOM As Double
            qtyBOM = Val(grid.TextMatrix(Row, bteColQtyBOM))
    
            grid.TextMatrix(Row, bteColQtyBOM) = Format(qtyBOM, "0.0000")
            grid.TextMatrix(Row, bteColQtySet) = Format(qtyBOM * Val(grid.TextMatrix(Row, bteColQtyPO)), gs_formatQtyBOM)
    
            Call HitungAmount(Row)
        Else
            MsgBox "Silakan checklist baris ini terlebih dahulu sebelum mengubah Qty BOM.", vbExclamation
            grid.TextMatrix(Row, bteColQtyBOM) = ""
        End If
    End If

End Sub

Private Sub HitungAmount(ByVal Row As Long)
    Dim qtySet As Double
    Dim qtyBOM As Double
    Dim Price As Double
    Dim Amount As Double

    qtySet = Val(grid.TextMatrix(Row, bteColQtySet))
    qtyBOM = Val(grid.TextMatrix(Row, bteColQtyBOM))
    Price = Val(grid.TextMatrix(Row, bteColPrice))

    Amount = (qtySet * qtyBOM) * Price
    grid.TextMatrix(Row, bteColAmount) = Format(Amount, "#,##0.00")
End Sub

Private Function IsFirstRowChecked() As Boolean
    Dim i As Long
    For i = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
            IsFirstRowChecked = True
            Exit Function
        End If
    Next i
    IsFirstRowChecked = False
End Function

Private Sub SaveContract()
    Dim i As Long
    Dim isFirstRow As Boolean
    Dim hasCheckedRow As Boolean
    Dim cmd As ADODB.Command
    Dim contractNo As String
    Dim statusFlag As String
    Dim answer As VbMsgBoxResult

    On Error GoTo errHandler

    contractNo = Trim(cboContractNo.Text)
    statusFlag = IIf(cboStatus.Text = "Create", "0", "1")
    isFirstRow = True
    hasCheckedRow = False

    Db.BeginTrans

    ' Cek apakah ada baris yang dicentang
    For i = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
            hasCheckedRow = True
            Exit For
        End If
    Next i

    ' Jika tidak ada yang dicentang, hapus semua master & detail
    If Not hasCheckedRow Then
        answer = MsgBox("Are you sure you want to delete all data for this contract?", vbYesNo + vbQuestion, "Confirm Delete")
        If answer = vbYes Then
            Set cmd = New ADODB.Command
            With cmd
                .ActiveConnection = Db
                .CommandType = adCmdStoredProc
                .CommandText = "dbo.sp_POContract_Delete"

                .Parameters.append .CreateParameter("@ContractNo", adVarChar, adParamInput, 50, cboContractNo.Text)
                .Parameters.append .CreateParameter("@PONo", adVarChar, adParamInput, 50, cboPONo.Text)
                .Parameters.append .CreateParameter("@User", adVarChar, adParamInput, 50, userLogin)

                .Execute
            End With

            lblErrMsg = DisplayMsg(1201)
            Db.CommitTrans
        Else
            Db.RollbackTrans
            lblErrMsg = DisplayMsg(1203)
        End If
        Exit Sub
    End If

    ' Loop untuk insert/update detail
    For i = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
            Set cmd = New ADODB.Command
            With cmd
                .ActiveConnection = Db
                .CommandType = adCmdStoredProc
                .CommandText = "dbo.sp_POContract_InsUpd"
                .CommandTimeout = 300
                
                ' Assign parameters
                .Parameters.append .CreateParameter("@ContractNo", adVarChar, adParamInput, 50, contractNo)
                .Parameters.append .CreateParameter("@PONo", adVarChar, adParamInput, 50, Trim(cboPONo.Text))
                .Parameters.append .CreateParameter("@ContractDate", adDate, adParamInput, , DtContractDate.Value)
                .Parameters.append .CreateParameter("@Remark", adVarChar, adParamInput, 255, Trim(txtRemark.Text))
                .Parameters.append .CreateParameter("@SeqNo", adInteger, adParamInput, , i)
                .Parameters.append .CreateParameter("@ParentItemCode", adVarChar, adParamInput, 50, grid.TextMatrix(i, bteColParentItemCode))
                .Parameters.append .CreateParameter("@ItemCode", adVarChar, adParamInput, 50, grid.TextMatrix(i, bteColItemCode))
                .Parameters.append .CreateParameter("@HSCode", adVarChar, adParamInput, 50, grid.TextMatrix(i, bteColHSCode))
                .Parameters.append .CreateParameter("@ItemName", adVarChar, adParamInput, 100, grid.TextMatrix(i, bteColItemName))

                If IsDate(grid.TextMatrix(i, bteColReceiptDate)) Then
                    .Parameters.append .CreateParameter("@ReceiptDate", adDate, adParamInput, , CDate(grid.TextMatrix(i, bteColReceiptDate)))
                Else
                    .Parameters.append .CreateParameter("@ReceiptDate", adDate, adParamInput, , Null)
                End If

                .Parameters.append .CreateParameter("@CurrencyCode", adVarChar, adParamInput, 10, grid.TextMatrix(i, bteColCurrCode))
                .Parameters.append .CreateParameter("@Rate", adDouble, adParamInput, , IIf(IsNumeric(grid.TextMatrix(i, bteColRate)), CDbl(grid.TextMatrix(i, bteColRate)), 0))
                .Parameters.append .CreateParameter("@QtyBOM", adDouble, adParamInput, , IIf(IsNumeric(grid.TextMatrix(i, bteColQtyBOM)), CDbl(grid.TextMatrix(i, bteColQtyBOM)), 0))
                .Parameters.append .CreateParameter("@QtyPO", adDouble, adParamInput, , IIf(IsNumeric(grid.TextMatrix(i, bteColQtyPO)), CDbl(grid.TextMatrix(i, bteColQtyPO)), 0))
                .Parameters.append .CreateParameter("@QtySet", adDouble, adParamInput, , IIf(IsNumeric(grid.TextMatrix(i, bteColQtySet)), CDbl(grid.TextMatrix(i, bteColQtySet)), 0))
                .Parameters.append .CreateParameter("@Unit", adVarChar, adParamInput, 10, grid.TextMatrix(i, bteColUnitCode))
                .Parameters.append .CreateParameter("@Price", adDouble, adParamInput, , IIf(IsNumeric(grid.TextMatrix(i, bteColPrice)), CDbl(grid.TextMatrix(i, bteColPrice)), 0))
                .Parameters.append .CreateParameter("@Amount", adDouble, adParamInput, , IIf(IsNumeric(grid.TextMatrix(i, bteColAmount)), CDbl(grid.TextMatrix(i, bteColAmount)), 0))
                .Parameters.append .CreateParameter("@BCType", adVarChar, adParamInput, 5, grid.TextMatrix(i, bteColBctype))
                .Parameters.append .CreateParameter("@BCNo", adVarChar, adParamInput, 30, grid.TextMatrix(i, bteColBcNo))

                If IsDate(grid.TextMatrix(i, bteColBCDate)) Then
                    .Parameters.append .CreateParameter("@BCDate", adDate, adParamInput, , CDate(grid.TextMatrix(i, bteColBCDate)))
                Else
                    .Parameters.append .CreateParameter("@BCDate", adDate, adParamInput, , Null)
                End If

                .Parameters.append .CreateParameter("@User", adVarChar, adParamInput, 50, userLogin)
                .Parameters.append .CreateParameter("@Status", adChar, adParamInput, 1, statusFlag)
                .Parameters.append .CreateParameter("@IsFirstRow", adInteger, adParamInput, , IIf(isFirstRow, 1, 0))

                .Execute
            End With
            isFirstRow = False
        End If
    Next i

    Db.CommitTrans

    cboStatus.Text = "Update"
    isiCboContractNo
    cboContractNo.Text = contractNo
    gridLoad
    lblErrMsg = DisplayMsg(1000)
    Exit Sub

errHandler:
    Db.RollbackTrans
    lblErrMsg = "Terjadi kesalahan saat menyimpan data: " & err.Description
End Sub

Private Function IsValidInput(Optional includeGridCheck As Boolean = False) As Boolean
    lblErrMsg = ""
    
    If cboCust.Text = "" Then
        cboCust.SetFocus
        lblErrMsg = DisplayMsg(9017) & " Customer Code "
        IsValidInput = False
        Exit Function
    End If

    If cboPONo.Text = "" Then
        cboPONo.SetFocus
        lblErrMsg = DisplayMsg(9017) & " PO No. "
        IsValidInput = False
        Exit Function
    End If

    If cboContractNo.Text = "" Then
        If cboStatus = "Create" Then
            cboContractNo.SetFocus
            lblErrMsg = DisplayMsg("0001") & " Contract No. "
            IsValidInput = False
            Exit Function
        ElseIf cboStatus = "Update" Then
            cboContractNo.SetFocus
            lblErrMsg = DisplayMsg(9017) & " Contract No. "
            IsValidInput = False
            Exit Function
        End If
    End If

    If includeGridCheck Then
        If grid.Rows <= 1 Then
            lblErrMsg = DisplayMsg(9017) & " Data First "
            IsValidInput = False
            Exit Function
        End If
    End If

    IsValidInput = True
End Function

Private Sub grid_KeyPress(KeyAscii As Integer)
Dim Col As Long
    Col = grid.Col

    If Col = bteColQtyBOM Then
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack) Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

