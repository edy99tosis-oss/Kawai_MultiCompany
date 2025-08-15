VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPartsRecSupInquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Parts (Material) Receipt/Supply Inquiry"
   ClientHeight    =   10695
   ClientLeft      =   150
   ClientTop       =   570
   ClientWidth     =   15120
   Icon            =   "frmPartsRecSupInquiry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10695
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNormalize 
      BackColor       =   &H0080FFFF&
      Caption         =   "Normalize"
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
      Left            =   12555
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtstock 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.#0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Index           =   5
      Left            =   8970
      MaxLength       =   15
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2355
      Width           =   1650
   End
   Begin VB.TextBox txtstock 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.#0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Index           =   4
      Left            =   7245
      MaxLength       =   15
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2355
      Width           =   1650
   End
   Begin VB.TextBox txtstock 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.#0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Index           =   6
      Left            =   10695
      MaxLength       =   15
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2355
      Width           =   1650
   End
   Begin VB.TextBox txtstock 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.#0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Index           =   3
      Left            =   5520
      MaxLength       =   15
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2355
      Width           =   1650
   End
   Begin VB.CommandButton cmdBrowseItem 
      Caption         =   "..."
      Height          =   300
      Left            =   4770
      TabIndex        =   1
      Top             =   975
      Width           =   315
   End
   Begin MSComDlg.CommonDialog CDExcel 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save As"
      Filter          =   "Text File ( *.txt )"
   End
   Begin VB.CommandButton cmdExcel 
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
      Left            =   13763
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9840
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13035
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.CommandButton cmdSubmit 
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
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2415
      Width           =   1170
   End
   Begin VB.TextBox txtstock 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.#0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Index           =   2
      Left            =   3795
      MaxLength       =   15
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2355
      Width           =   1650
   End
   Begin VB.TextBox txtstock 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.#0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Index           =   1
      Left            =   2070
      MaxLength       =   15
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2355
      Width           =   1650
   End
   Begin VB.TextBox txtstock 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.#0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Index           =   0
      Left            =   345
      MaxLength       =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2355
      Width           =   1650
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   240
      TabIndex        =   15
      Top             =   9150
      Width           =   14655
      Begin VB.Label lblErrMsg 
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
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   180
         Width           =   14400
      End
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9840
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6270
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   14655
      _cx             =   25850
      _cy             =   11060
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
      GridColor       =   8421504
      GridColorFixed  =   8421504
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   2
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
      RowHeightMin    =   275
      RowHeightMax    =   0
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
   Begin MSComCtl2.DTPicker dtpMonth 
      Height          =   315
      Left            =   7755
      TabIndex        =   3
      Top             =   1395
      Width           =   1260
      _ExtentX        =   2223
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
      Format          =   141230083
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin VB.Line Line2 
      X1              =   3585
      X2              =   6720
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Label lblWarehouse 
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
      Left            =   3585
      TabIndex        =   30
      Top             =   1455
      Width           =   3150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment"
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
      Left            =   9315
      TabIndex        =   29
      Top             =   1950
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End Month"
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
      Left            =   7628
      TabIndex        =   27
      Top             =   1950
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "After Inventory"
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
      Left            =   10860
      TabIndex        =   25
      Top             =   1950
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loss/Reject"
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
      Index           =   6
      Left            =   5850
      TabIndex        =   23
      Top             =   1950
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supply Total"
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
      Index           =   5
      Left            =   4088
      TabIndex        =   22
      Top             =   1950
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Total"
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
      Index           =   4
      Left            =   2340
      TabIndex        =   21
      Top             =   1950
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Premonth"
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
      Left            =   758
      TabIndex        =   20
      Top             =   1950
      Width           =   825
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Left            =   240
      Top             =   2235
      Width           =   12330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
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
      Left            =   6990
      TabIndex        =   19
      Top             =   1455
      Width           =   510
   End
   Begin MSForms.ComboBox cboWarehouse 
      Height          =   315
      Left            =   1890
      TabIndex        =   2
      Top             =   1395
      Width           =   1650
      VariousPropertyBits=   746604571
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "2910;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblItem 
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
      Left            =   5145
      TabIndex        =   18
      Top             =   1035
      Width           =   3870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      TabIndex        =   17
      Top             =   1020
      Width           =   915
   End
   Begin VB.Line Line1 
      X1              =   5145
      X2              =   9000
      Y1              =   1275
      Y2              =   1275
   End
   Begin MSForms.ComboBox cboItem 
      Height          =   315
      Left            =   1890
      TabIndex        =   0
      Top             =   960
      Width           =   2850
      VariousPropertyBits=   612386843
      MaxLength       =   30
      DisplayStyle    =   3
      Size            =   "5027;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   1455
      Width           =   1470
   End
   Begin VB.Label lblTittle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Parts (Material) Receipt / Supply Inquiry"
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
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   14655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   240
      Top             =   1875
      Width           =   12330
   End
End
Attribute VB_Name = "frmPartsRecSupInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dateUp As Date
Dim selisih As Integer

Dim bytColDate As Byte
Dim bytColCls As Byte
Dim bytColClsDesc As Byte
Dim bytColTradeCode As Byte
Dim bytColTradeName As Byte
Dim bytColReceipt As Byte
Dim bytColSupply As Byte
Dim bytColLossReject As Byte
Dim bytColAdjust As Byte
Dim bytColCurr As Byte
Dim bytColPrice As Byte
Dim bytColAmount As Byte
Dim bytColRemark As Byte
Dim bytColPONo As Byte
Dim bytColDONo As Byte
Dim bytColDocNo As Byte

Private Sub SetComboItem()
    Dim ls_sql As String
    Dim RsStock As New ADODB.Recordset
    Dim i As Long
        
    cboitem.columnCount = 3
    cboitem.clear
        
    ls_sql = "select item_code, makeritem_code, item_name from item_master where use_endday > convert(char(8), getdate(), 112) "
    RsStock.Open ls_sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    i = 0
    Do While Not RsStock.EOF
        cboitem.AddItem ""
        cboitem.List(i, 0) = Trim(RsStock("item_code"))
        cboitem.List(i, 1) = Trim(RsStock("makeritem_code"))
        cboitem.List(i, 2) = Trim(RsStock("item_name"))
        i = i + 1
        RsStock.MoveNext
    Loop
    RsStock.Close
    Set RsStock = Nothing
    
    cboitem.ColumnWidths = "120 pt; 120 pt; 260 pt"
    cboitem.ListWidth = 500
    cboitem.ListRows = 15
End Sub

Private Sub SetComboWarehouse()
    Dim ls_sql As String
    Dim RsStock As New ADODB.Recordset
    Dim i As Integer
        
    cboWarehouse.columnCount = 2
    cboWarehouse.clear
        
    ls_sql = " select wh_code, wh_name, stockcontrol_cls, adm_group  " & vbCrLf & _
                " from( " & vbCrLf & _
                "   select wh_code, wh_name, stockcontrol_cls, adm_group, 2 idx  " & vbCrLf & _
                "   from warehouse_master  " & vbCrLf & _
                "   where use_endday >= convert(varchar, getdate(), 112) " & vbCrLf & _
                "   union all " & vbCrLf & _
                "   select trade_code, trade_name, '01', '' adm_group, 1 idx " & vbCrLf & _
                "   from trade_master " & vbCrLf & _
                "   where trade_code in (select manufacture_code from manufacture_line) " & vbCrLf & _
                " )wh  order by wh_code  "
                
    RsStock.Open ls_sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    i = 0
    Do While Not RsStock.EOF
        cboWarehouse.AddItem ""
        cboWarehouse.List(i, 0) = Trim(RsStock("wh_code"))
        cboWarehouse.List(i, 1) = Trim(RsStock("wh_name"))
        i = i + 1
        RsStock.MoveNext
    Loop
    RsStock.Close
    Set RsStock = Nothing
    
    cboWarehouse.ColumnWidths = "75 pt; 225 pt"
    cboWarehouse.ListWidth = 300
    cboWarehouse.ListRows = 15
End Sub

Private Sub SetDataStock()
    Dim adoRs As New ADODB.Recordset

    Me.MousePointer = vbHourglass
    On Error GoTo errHandler
    
    sql = " select item_code, warehouse_code, sum(premonth) premonth, sum(receipt) receipt, sum(supply) supply, sum(lossreject) lossreject, sum([current]) [current], sum(inventory) inventory " & vbCrLf & _
                " from( " & vbCrLf & _
                "   select * from( " & vbCrLf & _
                "       select cast(rtrim(sh.stock_year) + right('00' + rtrim(sh.stock_month), 2) + '01' as datetime) period, " & vbCrLf & _
                "       sh.warehouse_code, sh.item_code, sh.premonth, sh.receipt, sh.supply, sh.lossreject, sh.[current], sh.inventory " & vbCrLf & _
                "       from stock_history sh " & vbCrLf & _
                "       union all " & vbCrLf & _
                "       select (select cast(max(cast(inventory_year as varchar) + right('00' + cast(inventory_month as varchar), 2)) + '01' as datetime) from inventory_control) period, " & vbCrLf & _
                "       sm.warehouse_code, sm.item_code, sm.lm_premonth premonth, sm.lm_receipt receipt, sm.lm_supply supply, sm.lm_lossreject lossreject, sm.lm_current [current], sm.lm_inventory inventory " & vbCrLf & _
                "       from stock_master sm " & vbCrLf & _
                "       union all " & vbCrLf & _
                "       select dateadd(m, 1, (select cast(max(cast(inventory_year as varchar) + right('00' + cast(inventory_month as varchar), 2)) + '01' as datetime) from inventory_control)) period, " & vbCrLf & _
                "       sm.warehouse_code, sm.item_code, sm.tm_premonth premonth, sm.tm_receipt receipt, sm.tm_supply supply, sm.tm_lossreject lossreject, sm.tm_current [current], isnull(sm.tm_inventory, sm.tm_current) inventory " & vbCrLf & _
                "       from stock_master sm " & vbCrLf & _
                "       union all " & vbCrLf & _
                "       select dateadd(m, 2, (select cast(max(cast(inventory_year as varchar) + right('00' + cast(inventory_month as varchar), 2)) + '01' as datetime) from inventory_control)) period, " & vbCrLf & _
                "       sm.warehouse_code, sm.item_code, sm.nm_premonth premonth, sm.nm_receipt receipt, sm.nm_supply supply, sm.nm_lossreject lossreject, sm.nm_current [current], isnull(sm.nm_inventory, sm.nm_current) inventory " & vbCrLf & _
                "       from stock_master sm " & vbCrLf & _
                "   )beg " & vbCrLf
        
    sql = sql + " )beg " & vbCrLf & _
                " where item_code = '" & cboitem.Text & "' " & vbCrLf & _
                " and warehouse_code = '" & cboWarehouse.Text & "' " & vbCrLf & _
                " and year(period) = '" & Year(dtpMonth.Value) & "' " & vbCrLf & _
                " and month(period) = '" & Month(dtpMonth.Value) & "' " & vbCrLf & _
                " group by item_code, warehouse_code "
    
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    'add trace 20220414
    LblErrMsg = "Get Value to header"
    If Not adoRs.EOF Then
        txtstock(0).Text = Format(adoRs.Fields("premonth"), gs_formatQty)
        txtstock(1).Text = Format(adoRs.Fields("receipt"), gs_formatQty)
        txtstock(2).Text = Format(adoRs.Fields("supply"), gs_formatQty)
        txtstock(3).Text = Format(adoRs.Fields("lossreject"), gs_formatQty)
        txtstock(4).Text = Format(adoRs.Fields("current"), gs_formatQty)
        txtstock(5).Text = Format(adoRs.Fields("inventory") - adoRs.Fields("current"), gs_formatQty)
        txtstock(6).Text = Format(adoRs.Fields("inventory"), gs_formatQty)
    Else
        txtstock(0).Text = Format(0, gs_formatQty)
        txtstock(1).Text = Format(0, gs_formatQty)
        txtstock(2).Text = Format(0, gs_formatQty)
        txtstock(3).Text = Format(0, gs_formatQty)
        txtstock(4).Text = Format(0, gs_formatQty)
        txtstock(5).Text = Format(0, gs_formatQty)
        txtstock(6).Text = Format(0, gs_formatQty)
    End If
    adoRs.Close
    
'    sql = " UPDATE dbo.Stock_Master " & vbCrLf & _
'            " SET LM_Supply=5,TM_Supply=5,NM_Supply=5 " & vbCrLf & _
'            " WHERE Item_Code='798758' AND Warehouse_Code='wh-001' " & vbCrLf & _
'            "  "
'     Db.Execute sql

ErrExit:
    Set adoRs = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    Resume ErrExit
End Sub

Private Sub SetGridHeader()
    Dim bytHakPrice As Byte
    
    bytColDate = 0
    bytColCls = 1
    bytColClsDesc = 2
    bytColReceipt = 3
    bytColSupply = 4
    bytColLossReject = 5
    bytColAdjust = 6
    bytColTradeCode = 7
    bytColTradeName = 8
    bytColCurr = 9
    bytColPrice = 10
    bytColAmount = 11
    bytColRemark = 12
    bytColPONo = 13
    bytColDONo = 14
    bytColDocNo = 15
    
    bytHakPrice = hakPrice(Me.Name)
    
    With grid
        .Redraw = flexRDNone
        .clear
        .Rows = 1
        .ColS = 16
        .FrozenCols = bytColClsDesc + 1
        
        .TextMatrix(0, bytColDate) = "Date"
        .TextMatrix(0, bytColCls) = "Cls"
        .TextMatrix(0, bytColClsDesc) = "Cls"
        .TextMatrix(0, bytColTradeCode) = "Location Code"
        .TextMatrix(0, bytColTradeName) = "Location Name"
        .TextMatrix(0, bytColReceipt) = "Receipt"
        .TextMatrix(0, bytColSupply) = "Supply"
        .TextMatrix(0, bytColLossReject) = "Loss/Reject"
        .TextMatrix(0, bytColAdjust) = "Differences"
        .TextMatrix(0, bytColCurr) = "Curr"
        .TextMatrix(0, bytColPrice) = "Price"
        .TextMatrix(0, bytColAmount) = "Amount"
        .TextMatrix(0, bytColRemark) = "Remark"
        .TextMatrix(0, bytColPONo) = "PO No."
        .TextMatrix(0, bytColDONo) = "DN No./SJ No."
        .TextMatrix(0, bytColDocNo) = "No. LPB/No. KSTB"
        
        .ColWidth(bytColDate) = 1200
        .ColWidth(bytColClsDesc) = 1500
        .ColWidth(bytColTradeCode) = 1400
        .ColWidth(bytColTradeName) = 3000
        .ColWidth(bytColReceipt) = 1500
        .ColWidth(bytColSupply) = 1500
        .ColWidth(bytColLossReject) = 1500
        .ColWidth(bytColAdjust) = 1500
        .ColWidth(bytColCurr) = 600
        .ColWidth(bytColPrice) = 1500
        .ColWidth(bytColAmount) = 1500
        .ColWidth(bytColRemark) = 3000
        .ColWidth(bytColPONo) = 1800
        .ColWidth(bytColDONo) = 1800
        .ColWidth(bytColDocNo) = 1800
        
        .ColHidden(bytColCls) = True
        .ColHidden(bytColCurr) = (bytHakPrice <> 1)
        .ColHidden(bytColPrice) = (bytHakPrice <> 1)
        .ColHidden(bytColAmount) = (bytHakPrice <> 1)
        .ColHidden(bytColDocNo) = True
        
        .ColAlignment(bytColDate) = flexAlignLeftCenter
        .ColAlignment(bytColCls) = flexAlignCenterCenter
        .ColAlignment(bytColClsDesc) = flexAlignLeftCenter
        .ColAlignment(bytColTradeCode) = flexAlignLeftCenter
        .ColAlignment(bytColTradeName) = flexAlignLeftCenter
        .ColAlignment(bytColReceipt) = flexAlignRightCenter
        .ColAlignment(bytColSupply) = flexAlignRightCenter
        .ColAlignment(bytColLossReject) = flexAlignRightCenter
        .ColAlignment(bytColAdjust) = flexAlignRightCenter
        .ColAlignment(bytColCurr) = flexAlignCenterCenter
        .ColAlignment(bytColPrice) = flexAlignRightCenter
        .ColAlignment(bytColAmount) = flexAlignRightCenter
        .ColAlignment(bytColRemark) = flexAlignLeftCenter
        .ColAlignment(bytColPONo) = flexAlignLeftCenter
        .ColAlignment(bytColDONo) = flexAlignLeftCenter
        .ColAlignment(bytColDocNo) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, 0, grid.ColS - 1) = flexAlignCenterCenter
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub SetGridData()
    Dim adoRs As New ADODB.Recordset
    Dim RsPeriod As New ADODB.Recordset
    Dim dblTotalReceipt As Double
    Dim dblTotalSupply As Double
    Dim dblTotalLossReject As Double
    Dim dblTotalAdjust As Double
    Dim ls_sql As String
    Dim sqlx As String
    Dim selisih As Integer
    
    'On Error GoTo ErrHandler
    
    SetDataStock
    SetGridHeader
        
    Me.MousePointer = vbHourglass
                
'    sql = "  select * from(  " & vbCrLf & _
'                "    select 0 idx, pr.item_code, pr.warehouse_code, pr.receipt_date trans_date,  " & vbCrLf & _
'                "    pr.receipt_cls cls, case pr.receipt_cls when 'P1' then 'Prod. Result' when 'R' then 'Receipt' when 'R1' then 'Return' end cls_desc,  " & vbCrLf & _
'                "    case when pr.receipt_cls = 'P1' then pr.po_no else pr.supplier_code end trade_code, wh.trade_name,  " & vbCrLf & _
'                "    pr.qty receipt, 0 supply, 0 lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(pr.price, 0) price, isnull(pr.amount, 0) amount,  " & vbCrLf & _
'                "    pr.remarks, case when pr.receipt_cls in('R', 'R1') then case when pr.po_no = '0' then '[NON PO]' else pr.po_no end else '' end po_no,  " & vbCrLf & _
'                "    case when pr.receipt_cls in('R', 'R1') then pr.suratjalan_no else '' end do_no,  " & vbCrLf & _
'                "    --case when pr.receipt_cls in('R', 'R1') then pr.no_lpb else '' end no_lpb  " & vbCrLf & _
'                "    '' no_lpb  " & vbCrLf & _
'                "    from part_receipt pr  " & vbCrLf & _
'                "    left join curr_cls cc on pr.currency_code = cc.curr_cls  " & vbCrLf
'
'    sql = sql + "    left join(  " & vbCrLf & _
'                "        select trade_code, trade_name from trade_master  " & vbCrLf & _
'                "        union all  " & vbCrLf & _
'                "        select wh_code, wh_name from warehouse_master  " & vbCrLf & _
'                "        union all  " & vbCrLf & _
'                "        SELECT Line_Code, Line_Name FROM Manufacture_Line  " & vbCrLf & _
'                "    )wh on case when pr.receipt_cls = 'P1' then pr.po_no else pr.supplier_code end = wh.trade_code  " & vbCrLf & _
'                "    where pr.receipt_cls in ('P1')  " & vbCrLf & _
'                "    union all  " & vbCrLf & _
'                "    select 0 idx, pr.item_code, pr.warehouse_code, pr.receipt_date trans_date,  " & vbCrLf & _
'                "    pr.receipt_cls cls, case pr.receipt_cls when 'P1' then 'Prod. Result' when 'R' then 'Receipt' when 'R1' then 'Return' end cls_desc,  " & vbCrLf & _
'                "    case when pr.receipt_cls = 'P1' then pr.po_no else pr.supplier_code end trade_code, wh.trade_name,  " & vbCrLf & _
'                "    case when pr.receipt_cls = 'R1' then -pr.qty else pr.qty end receipt, 0 supply, 0 lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(pr.price, 0) price, isnull(pr.amount, 0) amount,  " & vbCrLf
'
'    sql = sql + "    pr.remarks, case when pr.receipt_cls in('R', 'R1') then case when pr.po_no = '0' then '[NON PO]' else pr.po_no end else '' end po_no,  " & vbCrLf & _
'                "    case when pr.receipt_cls in('R', 'R1') then pr.suratjalan_no else '' end do_no,  " & vbCrLf & _
'                "    --case when pr.receipt_cls in('R', 'R1') then pr.no_lpb else '' end no_lpb  " & vbCrLf & _
'                "    '' no_lpb  " & vbCrLf & _
'                "    from part_receipt pr  " & vbCrLf & _
'                "    left join curr_cls cc on pr.currency_code = cc.curr_cls  " & vbCrLf & _
'                "    left join(  " & vbCrLf & _
'                "        select trade_code, trade_name from trade_master  " & vbCrLf & _
'                "        union all  " & vbCrLf & _
'                "        select wh_code, wh_name from warehouse_master  " & vbCrLf & _
'                "    )wh on case when pr.receipt_cls = 'P1' then pr.po_no else pr.supplier_code end = wh.trade_code  " & vbCrLf
'
'    sql = sql + "    where pr.receipt_cls in ('R', 'R1')  " & vbCrLf & _
'                "    and not exists(select po_no from purchaseorder_master where others_cls = '1' and po_no = pr.po_no)  " & vbCrLf & _
'                "    union all  " & vbCrLf & _
'                "    select 0 idx, ps.childitem_code item_code, ps.towarehouse_code warehouse_code, ps.childsupply_date trans_date,  " & vbCrLf & _
'                "    ps.supply_cls cls, case ps.supply_cls when 'S1' then 'Supply' end cls_desc, ps.fromwarehouse_code trade_code, wh.trade_name,  " & vbCrLf & _
'                "    ps.childrequirement_qty receipt, 0 supply, 0 lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(ps.price, 0) price, isnull(ps.amount, 0) amount,  " & vbCrLf & _
'                "    ps.remarks, '' po_no, CASE WHEN COALESCE(SupplyRec_No,'')='' THEN COALESCE(SJNo,'') ELSE RTRIM(COALESCE(SupplyRec_No,'')) + '-' + RTRIM(imm.Item_Name) end do_no, " & vbCrLf & _
'                "    '' no_kstb  " & vbCrLf & _
'                "    from part_supply ps  " & vbCrLf & _
'                "    left join curr_cls cc on ps.currency_code = cc.curr_cls  " & vbCrLf & _
'                "    LEFT JOIN dbo.Item_Master imm ON imm.Item_Code=ps.ParentItem_Code " & vbCrLf
'
'    sql = sql + "    left join(  " & vbCrLf & _
'                "        select trade_code, trade_name from trade_master where trade_cls = '1'  " & vbCrLf & _
'                "        union all  " & vbCrLf & _
'                "        select wh_code, wh_name from warehouse_master  " & vbCrLf & _
'                "    )wh on ps.fromwarehouse_code = wh.trade_code  " & vbCrLf & _
'                "    where supply_cls in ('S1') " & vbCrLf & _
'                "    union all  " & vbCrLf & _
'                "    select 0 idx, ps.childitem_code item_code, ps.towarehouse_code warehouse_code, ps.childsupply_date trans_date,  " & vbCrLf & _
'                "    ps.supply_cls cls, case ps.supply_cls when 'TR' then 'Transfer' end cls_desc, ps.fromwarehouse_code trade_code, wh.trade_name,  " & vbCrLf & _
'                "    ps.childrequirement_qty receipt, 0 supply, 0 lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(ps.price, 0) price, isnull(ps.amount, 0) amount,  " & vbCrLf & _
'                "    ps.remarks, '' po_no, '' do_no, '' no_kstb  " & vbCrLf
'
'    sql = sql + "    from part_supply ps  " & vbCrLf & _
'                "    left join curr_cls cc on ps.currency_code = cc.curr_cls  " & vbCrLf & _
'                "    left join(  " & vbCrLf & _
'                "        select trade_code, trade_name from trade_master  " & vbCrLf & _
'                "        union all  " & vbCrLf & _
'                "        select wh_code, wh_name from warehouse_master  " & vbCrLf & _
'                "    )wh on ps.fromwarehouse_code = wh.trade_code  " & vbCrLf & _
'                "    where supply_cls in ('TR')  " & vbCrLf & _
'                "    union all  " & vbCrLf & _
'                "    select 1 idx, ps.childitem_code item_code, ps.fromwarehouse_code warehouse_code, ps.childsupply_date trans_date,  " & vbCrLf & _
'                "    ps.supply_cls cls, case ps.supply_cls when 'S1' then 'Supply' when 'D' then 'Delivery' when 'S' then 'Consumption' end cls_desc,  " & vbCrLf
'
'    sql = sql + "    ps.towarehouse_code  trade_code, wh.trade_name,  " & vbCrLf & _
'                "    0 receipt, ps.childrequirement_qty supply, 0 lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(ps.price, 0) price, isnull(ps.amount, 0) amount,  " & vbCrLf & _
'                "    ps.remarks, '' po_no, case when ps.supply_cls in ('D') then ps.do_no else CASE WHEN COALESCE(SupplyRec_No,'')='' THEN COALESCE(SJNo,'') ELSE RTRIM(COALESCE(SupplyRec_No,'')) + '-' + RTRIM(imm.Item_Name) end end do_no, " & vbCrLf & _
'                "   --case when ps.supply_cls in ('S1') then ps.no_kstb else '' end no_kstb  " & vbCrLf & _
'                "    '' no_kstb  " & vbCrLf & _
'                "    from part_supply ps  " & vbCrLf & _
'                "    left join curr_cls cc on ps.currency_code = cc.curr_cls  " & vbCrLf & _
'                "    LEFT JOIN dbo.Item_Master imm ON imm.Item_Code=ps.ParentItem_Code " & vbCrLf & _
'                "    left join(  " & vbCrLf & _
'                "        select trade_code, trade_name from trade_master  " & vbCrLf & _
'                "        union all  " & vbCrLf & _
'                "        select wh_code, wh_name from warehouse_master  " & vbCrLf & _
'                "    )wh on ps.towarehouse_code  = wh.trade_code  " & vbCrLf
'
'    sql = sql + "    where supply_cls in ('S1', 'S', 'D')  " & vbCrLf & _
'                "    union all  " & vbCrLf & _
'                "    select 1 idx, ps.childitem_code item_code, ps.fromwarehouse_code warehouse_code, ps.childsupply_date trans_date,  " & vbCrLf & _
'                "    ps.supply_cls cls, case ps.supply_cls when 'L' then 'Loss' when 'Rj' then 'Reject' end cls_desc,  " & vbCrLf & _
'                "    ps.towarehouse_code  trade_code, wh.trade_name,  " & vbCrLf & _
'                "    0 receipt, 0 supply, ps.childrequirement_qty lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(ps.price, 0) price, isnull(ps.amount, 0) amount,  " & vbCrLf & _
'                "    ps.remarks, '' po_no, COALESCE(ps.SJNo,'') do_no, '' no_kstb  " & vbCrLf & _
'                "    from part_supply ps  " & vbCrLf & _
'                "    left join curr_cls cc on ps.currency_code = cc.curr_cls  " & vbCrLf & _
'                "    left join(  " & vbCrLf & _
'                "        select trade_code, trade_name from trade_master  " & vbCrLf
'
'    sql = sql + "        union all  " & vbCrLf & _
'                "        select wh_code, wh_name from warehouse_master  " & vbCrLf & _
'                "    )wh on ps.towarehouse_code  = wh.trade_code  " & vbCrLf & _
'                "    where supply_cls in ('L', 'Rj')  " & vbCrLf & _
'                "    union all  " & vbCrLf & _
'                "    select 1 idx, ps.parentitem_code item_code, ps.fromwarehouse_code warehouse_code, ps.childsupply_date trans_date,  " & vbCrLf & _
'                "    ps.supply_cls cls, case ps.supply_cls when 'TR' then 'Transfer' end cls_desc, ps.towarehouse_code trade_code, wh.trade_name,  " & vbCrLf & _
'                "    0 receipt, ps.childrequirement_qty supply, 0 lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(ps.price, 0) price, isnull(ps.amount, 0) amount,  " & vbCrLf & _
'                "    ps.remarks, '' po_no, '' do_no, '' no_kstb  " & vbCrLf & _
'                "    from part_supply ps  " & vbCrLf & _
'                "    left join curr_cls cc on ps.currency_code = cc.curr_cls  " & vbCrLf
'
'    sql = sql + "    left join(  " & vbCrLf & _
'                "        select trade_code, trade_name from trade_master  " & vbCrLf & _
'                "        union all  " & vbCrLf & _
'                "        select wh_code, wh_name from warehouse_master  " & vbCrLf & _
'                "    )wh on ps.towarehouse_code = wh.trade_code  " & vbCrLf & _
'                "    where supply_cls in ('TR')  " & vbCrLf & _
'                "      union all  " & vbCrLf & _
'                "    select 2 idx, item_code, warehouse_code, dateadd(m, 1, period) - 1 trans_date, 'ADJ' cls, 'Adjustment' cls_desc, warehouse_code trade_code, trade_name,  " & vbCrLf & _
'                "    0 receipt, 0 supply, 0 lossreject, adjust, '' curr, 0 price, 0 amount, '' remarks, '' po_no, '' do_no, '' no_kstb  " & vbCrLf & _
'                "    from(  " & vbCrLf & _
'                "        select cast(rtrim(sh.stock_year) + right('00' + rtrim(sh.stock_month), 2) + '01' as datetime) period,  " & vbCrLf
'
'    sql = sql + "        sh.warehouse_code, sh.item_code, sh.premonth, sh.receipt, sh.supply, sh.lossreject, isnull(sh.inventory, sh.[current]) inventory, sh.[current], isnull(sh.inventory, sh.[current]) - sh.[current] adjust, wh.trade_name  " & vbCrLf & _
'                "        from stock_history sh  " & vbCrLf & _
'                "        left join(  " & vbCrLf & _
'                "            select trade_code, trade_name from trade_master  " & vbCrLf & _
'                "            union all  " & vbCrLf & _
'                "            select wh_code, wh_name from warehouse_master  " & vbCrLf & _
'                "        )wh on sh.warehouse_code = wh.trade_code  " & vbCrLf & _
'                "        where isnull(sh.inventory, sh.[current]) <> sh.[current]  " & vbCrLf & _
'                "        union all  " & vbCrLf & _
'                "        select (select cast(max(cast(inventory_year as varchar) + right('00' + cast(inventory_month as varchar), 2)) + '01' as datetime) from inventory_control) period,  " & vbCrLf & _
'                "        sm.warehouse_code, sm.item_code, sm.lm_premonth premonth, sm.lm_receipt receipt, sm.lm_supply supply, sm.lm_lossreject lossreject, isnull(sm.lm_inventory, sm.lm_current) inventory, sm.lm_current [current], isnull(sm.lm_inventory, sm.lm_current) - sm.lm_current adj, wh.trade_name  " & vbCrLf
'
'    sql = sql + "        from stock_master sm  " & vbCrLf & _
'                "        left join(  " & vbCrLf & _
'                "            select trade_code, trade_name from trade_master  " & vbCrLf & _
'                "            union all  " & vbCrLf & _
'                "            select wh_code, wh_name from warehouse_master  " & vbCrLf & _
'                "        )wh on sm.warehouse_code = wh.trade_code  " & vbCrLf & _
'                "        where isnull(sm.lm_inventory, sm.lm_current) <> sm.lm_current  " & vbCrLf & _
'                "    )adj  " & vbCrLf
'
'    sql = sql + " )trans " & vbCrLf & _
'                " where item_code = '" & cboitem.Text & "' " & vbCrLf & _
'                " and warehouse_code = '" & cboWarehouse.Text & "' " & vbCrLf & _
'                " and year(trans_date) = '" & Year(dtpMonth.Value) & "' " & vbCrLf & _
'                " and month(trans_date) = '" & Month(dtpMonth.Value) & "' " & vbCrLf
    'Optimasi Query Load data to grid 20230706
    sql = "   select * from(   " & vbCrLf & _
            "     select 0 idx, pr.item_code, pr.warehouse_code, pr.receipt_date trans_date,   " & vbCrLf & _
            "     pr.receipt_cls cls, case pr.receipt_cls when 'P1' then 'Prod. Result' when 'R' then 'Receipt' when 'R1' then 'Return' end cls_desc,   " & vbCrLf & _
            "     case when pr.receipt_cls = 'P1' then pr.po_no else pr.supplier_code end trade_code, wh.trade_name,   " & vbCrLf & _
            "     pr.qty receipt, 0 supply, 0 lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(pr.price, 0) price, isnull(pr.amount, 0) amount,   " & vbCrLf & _
            "     pr.remarks, case when pr.receipt_cls in('R', 'R1') then case when pr.po_no = '0' then '[NON PO]' else pr.po_no end else '' end po_no,   " & vbCrLf & _
            "     case when pr.receipt_cls in('R', 'R1') then pr.suratjalan_no else '' end do_no,   " & vbCrLf & _
            "     '' no_lpb   " & vbCrLf & _
            "     from part_receipt pr   " & vbCrLf & _
            "     left join curr_cls cc on pr.currency_code = cc.curr_cls   " & vbCrLf & _
            "     left join(   "

    sql = sql + "         select trade_code, trade_name from trade_master   " & vbCrLf & _
                "         union all   " & vbCrLf & _
                "         select wh_code, wh_name from warehouse_master   " & vbCrLf & _
                "         union all   " & vbCrLf & _
                "         SELECT Line_Code, Line_Name FROM Manufacture_Line   " & vbCrLf & _
                "     )wh on case when pr.receipt_cls = 'P1' then pr.po_no else pr.supplier_code end = wh.trade_code   " & vbCrLf & _
                "     where pr.receipt_cls in ('P1')  " & vbCrLf & _
                "   AND pr.Item_Code = '" & cboitem.Text & "'  " & vbCrLf & _
                "   AND pr.Warehouse_Code= '" & cboWarehouse.Text & "'  " & vbCrLf & _
                "   AND YEAR(pr.Receipt_Date)= '" & Year(dtpMonth.Value) & "' AND MONTH(pr.Receipt_Date)= '" & Month(dtpMonth.Value) & "' " & vbCrLf & _
                "     UNION ALL  "
    
    sql = sql + "     select 0 idx, pr.item_code, pr.warehouse_code, pr.receipt_date trans_date,   " & vbCrLf & _
                "     pr.receipt_cls cls, case pr.receipt_cls when 'P1' then 'Prod. Result' when 'R' then 'Receipt' when 'R1' then 'Return' end cls_desc,   " & vbCrLf & _
                "     case when pr.receipt_cls = 'P1' then pr.po_no else pr.supplier_code end trade_code, wh.trade_name,   " & vbCrLf & _
                "     case when pr.receipt_cls = 'R1' then -pr.qty else pr.qty end receipt, 0 supply, 0 lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(pr.price, 0) price,  " & vbCrLf & _
                "   ISNULL(pr.amount, 0) amount, pr.remarks,  " & vbCrLf & _
                "   CASE when pr.receipt_cls in('R', 'R1') then case when pr.po_no = '0' then '[NON PO]' else pr.po_no end else '' end po_no,   " & vbCrLf & _
                "     case when pr.receipt_cls in('R', 'R1') then pr.suratjalan_no else '' end do_no, '' no_lpb   " & vbCrLf & _
                "     from part_receipt pr   " & vbCrLf & _
                "     left join curr_cls cc on pr.currency_code = cc.curr_cls   " & vbCrLf & _
                "     left join(   " & vbCrLf & _
                "         select trade_code, trade_name from trade_master   "
    
    sql = sql + "         union all   " & vbCrLf & _
                "         select wh_code, wh_name from warehouse_master   " & vbCrLf & _
                "     )wh on case when pr.receipt_cls = 'P1' then pr.po_no else pr.supplier_code end = wh.trade_code   " & vbCrLf & _
                "     where pr.receipt_cls in ('R', 'R1')   " & vbCrLf & _
                "     and not exists(select po_no from purchaseorder_master where others_cls = '1' and po_no = pr.po_no) " & vbCrLf & _
                "   AND pr.Item_Code = '" & cboitem.Text & "'  " & vbCrLf & _
                "   AND pr.Warehouse_Code='" & cboWarehouse.Text & "'  " & vbCrLf & _
                "   AND YEAR(pr.Receipt_Date)='" & Year(dtpMonth.Value) & "' AND MONTH(pr.Receipt_Date)='" & Month(dtpMonth.Value) & "' AND COALESCE(pr.Receipt_Status,'01') = '01' " & vbCrLf & _
                "     UNION ALL    " & vbCrLf & _
                "     select 0 idx, ps.childitem_code item_code, ps.towarehouse_code warehouse_code, ps.childsupply_date trans_date,   " & vbCrLf & _
                "     ps.supply_cls cls, case ps.supply_cls when 'S1' then 'Supply' end cls_desc, ps.fromwarehouse_code trade_code, wh.trade_name,   "
    
    sql = sql + "     ps.childrequirement_qty receipt, 0 supply, 0 lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(ps.price, 0) price, isnull(ps.amount, 0) amount,   " & vbCrLf & _
                "     ps.remarks, '' po_no, CASE WHEN COALESCE(SupplyRec_No,'')='' THEN COALESCE(SJNo,'') ELSE RTRIM(COALESCE(SupplyRec_No,'')) + '-' + RTRIM(imm.Item_Name) end do_no, '' no_kstb   " & vbCrLf & _
                "     from part_supply ps   " & vbCrLf & _
                "     left join curr_cls cc on ps.currency_code = cc.curr_cls   " & vbCrLf & _
                "     LEFT JOIN dbo.Item_Master imm ON imm.Item_Code=ps.ParentItem_Code  " & vbCrLf & _
                "     left join(   " & vbCrLf & _
                "         select trade_code, trade_name from trade_master where trade_cls = '1'   " & vbCrLf & _
                "         union all   " & vbCrLf & _
                "         select wh_code, wh_name from warehouse_master   " & vbCrLf & _
                "     )wh on ps.fromwarehouse_code = wh.trade_code   " & vbCrLf & _
                "     where supply_cls in ('S1')  "
    
    sql = sql + "   AND ps.childitem_code = '" & cboitem.Text & "'  " & vbCrLf & _
                "   AND ps.towarehouse_code='" & cboWarehouse.Text & "'  " & vbCrLf & _
                "   AND YEAR(ps.childsupply_date)='" & Year(dtpMonth.Value) & "' AND MONTH(ps.childsupply_date)='" & Month(dtpMonth.Value) & "' " & vbCrLf & _
                "     UNION ALL  " & vbCrLf & _
                "     select 0 idx, ps.childitem_code item_code, ps.towarehouse_code warehouse_code, ps.childsupply_date trans_date,   " & vbCrLf & _
                "     ps.supply_cls cls, case ps.supply_cls when 'TR' then 'Transfer' end cls_desc, ps.fromwarehouse_code trade_code, wh.trade_name,   " & vbCrLf & _
                "     ps.childrequirement_qty receipt, 0 supply, 0 lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(ps.price, 0) price, isnull(ps.amount, 0) amount,   " & vbCrLf & _
                "     ps.remarks, '' po_no, '' do_no, '' no_kstb   " & vbCrLf & _
                "     from part_supply ps   " & vbCrLf & _
                "     left join curr_cls cc on ps.currency_code = cc.curr_cls   " & vbCrLf & _
                "     left join(   "
    
    sql = sql + "         select trade_code, trade_name from trade_master   " & vbCrLf & _
                "         union all   " & vbCrLf & _
                "         select wh_code, wh_name from warehouse_master   " & vbCrLf & _
                "     )wh on ps.fromwarehouse_code = wh.trade_code   " & vbCrLf & _
                "     where supply_cls in ('TR')   " & vbCrLf & _
                "   AND ps.childitem_code = '" & cboitem.Text & "'  " & vbCrLf & _
                "   AND ps.towarehouse_code='" & cboWarehouse.Text & "'  " & vbCrLf & _
                "   AND YEAR(ps.childsupply_date)='" & Year(dtpMonth.Value) & "' AND MONTH(ps.childsupply_date)='" & Month(dtpMonth.Value) & "' " & vbCrLf & _
                "     UNION ALL  " & vbCrLf & _
                "     select 1 idx, ps.childitem_code item_code, ps.fromwarehouse_code warehouse_code, ps.childsupply_date trans_date,   " & vbCrLf & _
                "     ps.supply_cls cls, case ps.supply_cls when 'S1' then 'Supply' when 'D' then 'Delivery' when 'S' then 'Consumption' end cls_desc,   "
    
    sql = sql + "     ps.towarehouse_code  trade_code, wh.trade_name,   " & vbCrLf & _
                "     0 receipt, ps.childrequirement_qty supply, 0 lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(ps.price, 0) price, isnull(ps.amount, 0) amount,   " & vbCrLf & _
                "     ps.remarks, '' po_no,  " & vbCrLf & _
                "   CASE WHEN ps.supply_cls in ('D') THEN ps.do_no ELSE  " & vbCrLf & _
                "       CASE WHEN COALESCE(SupplyRec_No,'')='' THEN COALESCE(SJNo,'') ELSE RTRIM(COALESCE(SupplyRec_No,'')) + '-' + RTRIM(imm.Item_Name) END  " & vbCrLf & _
                "   END do_no, '' no_kstb   " & vbCrLf & _
                "     from part_supply ps   " & vbCrLf & _
                "     left join curr_cls cc on ps.currency_code = cc.curr_cls   " & vbCrLf & _
                "     LEFT JOIN dbo.Item_Master imm ON imm.Item_Code=ps.ParentItem_Code  " & vbCrLf & _
                "     left join(   " & vbCrLf & _
                "         select trade_code, trade_name from trade_master   "
    
    sql = sql + "         union all   " & vbCrLf & _
                "         select wh_code, wh_name from warehouse_master   " & vbCrLf & _
                "     )wh on ps.towarehouse_code  = wh.trade_code   " & vbCrLf & _
                "     where supply_cls in ('S1', 'S', 'D')   " & vbCrLf & _
                "   AND ps.childitem_code = '" & cboitem.Text & "'  " & vbCrLf & _
                "   AND ps.fromwarehouse_code='" & cboWarehouse.Text & "'  " & vbCrLf & _
                "   AND YEAR(ps.childsupply_date)='" & Year(dtpMonth.Value) & "' AND MONTH(ps.childsupply_date)='" & Month(dtpMonth.Value) & "' " & vbCrLf & _
                "     UNION ALL " & vbCrLf & _
                "     select 1 idx, ps.childitem_code item_code, ps.fromwarehouse_code warehouse_code, ps.childsupply_date trans_date,   " & vbCrLf & _
                "     ps.supply_cls cls, case ps.supply_cls when 'L' then 'Loss' when 'Rj' then 'Reject' end cls_desc,   " & vbCrLf & _
                "     ps.towarehouse_code  trade_code, wh.trade_name,   "
    
    sql = sql + "     0 receipt, 0 supply, ps.childrequirement_qty lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(ps.price, 0) price, isnull(ps.amount, 0) amount,   " & vbCrLf & _
                "     ps.remarks, '' po_no, COALESCE(ps.SJNo,'') do_no, '' no_kstb   " & vbCrLf & _
                "     from part_supply ps   " & vbCrLf & _
                "     left join curr_cls cc on ps.currency_code = cc.curr_cls   " & vbCrLf & _
                "     left join(   " & vbCrLf & _
                "         select trade_code, trade_name from trade_master   " & vbCrLf & _
                "         union all   " & vbCrLf & _
                "         select wh_code, wh_name from warehouse_master   " & vbCrLf & _
                "     )wh on ps.towarehouse_code  = wh.trade_code   " & vbCrLf & _
                "     where supply_cls in ('L', 'Rj')   " & vbCrLf & _
                "   AND ps.childitem_code = '" & cboitem.Text & "'  "
    
    sql = sql + "   AND ps.fromwarehouse_code ='" & cboWarehouse.Text & "'  " & vbCrLf & _
                "   AND YEAR(ps.childsupply_date)='" & Year(dtpMonth.Value) & "' AND MONTH(ps.childsupply_date)='" & Month(dtpMonth.Value) & "' " & vbCrLf & _
                "     UNION ALL    " & vbCrLf & _
                "     select 1 idx, ps.parentitem_code item_code, ps.fromwarehouse_code warehouse_code, ps.childsupply_date trans_date,   " & vbCrLf & _
                "     ps.supply_cls cls, case ps.supply_cls when 'TR' then 'Transfer' end cls_desc, ps.towarehouse_code trade_code, wh.trade_name,   " & vbCrLf & _
                "     0 receipt, ps.childrequirement_qty supply, 0 lossreject, 0 adjust, isnull(cc.[description], '') curr, isnull(ps.price, 0) price, isnull(ps.amount, 0) amount,   " & vbCrLf & _
                "     ps.remarks, '' po_no, '' do_no, '' no_kstb   " & vbCrLf & _
                "     from part_supply ps   " & vbCrLf & _
                "     left join curr_cls cc on ps.currency_code = cc.curr_cls   " & vbCrLf & _
                "     left join(   " & vbCrLf & _
                "         select trade_code, trade_name from trade_master   "
    
    sql = sql + "         union all   " & vbCrLf & _
                "         select wh_code, wh_name from warehouse_master   " & vbCrLf & _
                "     )wh on ps.towarehouse_code = wh.trade_code   " & vbCrLf & _
                "     where supply_cls in ('TR')   " & vbCrLf & _
                "   AND ps.ParentItem_Code = '" & cboitem.Text & "'  --childitem_code " & vbCrLf & _
                "   AND ps.fromwarehouse_code ='" & cboWarehouse.Text & "'  " & vbCrLf & _
                "   AND YEAR(ps.childsupply_date)='" & Year(dtpMonth.Value) & "' AND MONTH(ps.childsupply_date)='" & Month(dtpMonth.Value) & "' " & vbCrLf & _
                "     UNION ALL  " & vbCrLf & _
                "     select 2 idx, item_code, warehouse_code, dateadd(m, 1, period) - 1 trans_date, 'ADJ' cls, 'Adjustment' cls_desc, warehouse_code trade_code, trade_name,   " & vbCrLf & _
                "     0 receipt, 0 supply, 0 lossreject, adjust, '' curr, 0 price, 0 amount, '' remarks, '' po_no, '' do_no, '' no_kstb   " & vbCrLf & _
                "     from(   "
    
    sql = sql + "         select cast(rtrim(sh.stock_year) + right('00' + rtrim(sh.stock_month), 2) + '01' as datetime) period,   " & vbCrLf & _
                "         sh.warehouse_code, sh.item_code, sh.premonth, sh.receipt, sh.supply, sh.lossreject, isnull(sh.inventory, sh.[current]) inventory, sh.[current],  " & vbCrLf & _
                "       ISNULL(sh.inventory, sh.[current]) - sh.[current] adjust, wh.trade_name   " & vbCrLf & _
                "         from stock_history sh   " & vbCrLf & _
                "         left join(   " & vbCrLf & _
                "             select trade_code, trade_name from trade_master   " & vbCrLf & _
                "             union all   " & vbCrLf & _
                "             select wh_code, wh_name from warehouse_master   " & vbCrLf & _
                "         )wh on sh.warehouse_code = wh.trade_code   " & vbCrLf & _
                "         where isnull(sh.inventory, sh.[current]) <> sh.[current]   " & vbCrLf & _
                "         union all   "
    
    sql = sql + "         select (select cast(max(cast(inventory_year as varchar) + right('00' + cast(inventory_month as varchar), 2)) + '01' as datetime) from inventory_control) period,   " & vbCrLf & _
                "         sm.warehouse_code, sm.item_code, sm.lm_premonth premonth, sm.lm_receipt receipt, sm.lm_supply supply, sm.lm_lossreject lossreject, isnull(sm.lm_inventory,  " & vbCrLf & _
                "       sm.lm_current) inventory, sm.lm_current [current], isnull(sm.lm_inventory, sm.lm_current) - sm.lm_current adj, wh.trade_name   " & vbCrLf & _
                "         from stock_master sm   " & vbCrLf & _
                "         left join(   " & vbCrLf & _
                "             select trade_code, trade_name from trade_master   " & vbCrLf & _
                "             union all   " & vbCrLf & _
                "             select wh_code, wh_name from warehouse_master   " & vbCrLf & _
                "         )wh on sm.warehouse_code = wh.trade_code   " & vbCrLf & _
                "         where isnull(sm.lm_inventory, sm.lm_current) <> sm.lm_current   " & vbCrLf & _
                "     )adj   "
    
    sql = sql + "   where item_code = '" & cboitem.Text & "'  " & vbCrLf & _
                "   AND warehouse_code ='" & cboWarehouse.Text & "'  " & vbCrLf & _
                "   AND YEAR(period)='" & Year(dtpMonth.Value) & "' AND MONTH(period)='" & Month(dtpMonth.Value) & "' " & vbCrLf & _
                "  )trans  "

    sql = sql + " order by trans_date, idx "
    
    'add trace 20220414
    LblErrMsg = "get value for bind to grid"
    
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    
    If Not adoRs.EOF Then
        With grid
            .Redraw = flexRDNone
            While Not adoRs.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, bytColDate) = Format(adoRs.Fields("trans_date"), "dd MMM yyyy")
                .TextMatrix(.Rows - 1, bytColCls) = Trim(adoRs.Fields("cls"))
                .TextMatrix(.Rows - 1, bytColClsDesc) = Trim(adoRs.Fields("cls_desc"))
                .TextMatrix(.Rows - 1, bytColTradeCode) = Trim(adoRs.Fields("trade_code"))
                .TextMatrix(.Rows - 1, bytColTradeName) = Trim(adoRs.Fields("trade_name") & "")
                .TextMatrix(.Rows - 1, bytColReceipt) = Format(adoRs.Fields("receipt"), gs_formatQty)
                .TextMatrix(.Rows - 1, bytColSupply) = Format(adoRs.Fields("supply"), gs_formatQty)
                .TextMatrix(.Rows - 1, bytColLossReject) = Format(adoRs.Fields("lossreject"), gs_formatQty)
                .TextMatrix(.Rows - 1, bytColAdjust) = Format(adoRs.Fields("adjust"), gs_formatQty)
                .TextMatrix(.Rows - 1, bytColCurr) = Trim(adoRs.Fields("curr"))
                .TextMatrix(.Rows - 1, bytColPrice) = Format(adoRs.Fields("price"), gs_formatPrice)
                .TextMatrix(.Rows - 1, bytColAmount) = Format(adoRs.Fields("amount"), gs_formatAmount)
                .TextMatrix(.Rows - 1, bytColRemark) = Trim(adoRs.Fields("remarks")) & ""
                .TextMatrix(.Rows - 1, bytColPONo) = Trim(adoRs.Fields("po_no") & "")
                .TextMatrix(.Rows - 1, bytColDONo) = Trim(adoRs.Fields("do_no") & "")
                .TextMatrix(.Rows - 1, bytColDocNo) = Trim(adoRs.Fields("no_lpb") & "")
                
                dblTotalReceipt = dblTotalReceipt + adoRs.Fields("receipt")
                dblTotalSupply = dblTotalSupply + adoRs.Fields("supply")
                dblTotalLossReject = dblTotalLossReject + adoRs.Fields("lossreject")
                dblTotalAdjust = dblTotalAdjust + adoRs.Fields("adjust")
                
                adoRs.MoveNext
                
                'add trace 20220414
                LblErrMsg = "Process Looping"
            Wend
                'add trace 20220414
                LblErrMsg = ""
            .AddItem ""
            .TextMatrix(.Rows - 1, bytColReceipt) = Format(dblTotalReceipt, gs_formatQty)
            .TextMatrix(.Rows - 1, bytColSupply) = Format(dblTotalSupply, gs_formatQty)
            .TextMatrix(.Rows - 1, bytColLossReject) = Format(dblTotalLossReject, gs_formatQty)
            .TextMatrix(.Rows - 1, bytColAdjust) = Format(dblTotalAdjust, gs_formatQty)
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .ColS - 1) = &H8000000F
            
            .Redraw = flexRDDirect
        End With
    Else
        LblErrMsg.Caption = DisplayMsg("0013")
    End If
    adoRs.Close
    
    If CDbl(dblTotalReceipt) <> CDbl(txtstock(1).Text) Or Format(CDbl(dblTotalSupply), "###.##") <> CDbl(txtstock(2).Text) Or CDbl(dblTotalLossReject) <> CDbl(txtstock(3).Text) Then 'Jika selisih baru update
'
'        ls_sql = " DECLARE @MaxYear AS CHAR(4) " & vbCrLf & _
'                      " DECLARE @Period AS DATETIME " & vbCrLf & _
'                      " SET @Period='" & Format(dtpMonth.Value, "yyyy-mm-dd") & "' " & vbCrLf & _
'                      "  " & vbCrLf & _
'                      " SET @MaxYear=(SELECT MAX(Inventory_Year)FROM dbo.Inventory_Control WHERE Fix_Cls='1') " & vbCrLf & _
'                      " SELECT Selisih=DATEDIFF(MOnth,cast(@MaxYear+RIGHT(100+MAX(Inventory_Month),2)+'01' AS DATETIME), " & vbCrLf & _
'                      "        CAST(CONVERT(CHAR(4),YEAR(@Period))+RIGHT(100+MONTH(@Period),2)+'01' AS datetime)) " & vbCrLf & _
'                      " FROM dbo.Inventory_Control " & vbCrLf & _
'                      " WHERE Inventory_Year=@MaxYear " & vbCrLf & _
'                      "  "
'        RsPeriod.Open ls_sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
'        If Not RsPeriod.EOF Then
'
'            selisih = RsPeriod("selisih")
'
'            sqlx = " DECLARE @selisih AS INT " & vbCrLf & _
'              " DECLARE @WHCode AS CHAR(15) " & vbCrLf & _
'              " DECLARE @Item_Code AS CHAR(25) " & vbCrLf & _
'              " DECLARE @Receipt AS NUMERIC(18,2) " & vbCrLf & _
'              " DECLARE @Supply AS NUMERIC(18,2) " & vbCrLf & _
'              " DECLARE @Loss AS NUMERIC(18,2) " & vbCrLf & _
'              "  " & vbCrLf & _
'              " SET @selisih=" & RsPeriod("selisih") & "" & vbCrLf & _
'              " SET @WHCode='" & Trim(cboWarehouse.Text) & "' " & vbCrLf & _
'              " SET @Item_Code='" & Trim(cboItem.Text) & "' " & vbCrLf & _
'              " SET @Receipt=" & CDbl(dblTotalReceipt) & " " & vbCrLf
'
'            sqlx = sqlx + " SET @Supply=" & CDbl(dblTotalSupply) & " " & vbCrLf & _
'                          " SET @Loss=" & CDbl(dblTotalLossReject) & "" & vbCrLf & _
'                          "  " & vbCrLf & _
'                          " IF @selisih=0  " & vbCrLf & _
'                          "     BEGIN    " & vbCrLf & _
'                          "         UPDATE dbo.Stock_Master " & vbCrLf & _
'                          "         SET LM_Receipt=@Receipt,LM_Supply=@Supply,LM_LossReject=@Loss " & vbCrLf & _
'                          "         WHERE Warehouse_Code=@WHCode AND Item_Code=@Item_Code " & vbCrLf & _
'                          "     END      " & vbCrLf & _
'                          " IF @selisih=1 " & vbCrLf & _
'                          "     BEGIN  " & vbCrLf
'
'            sqlx = sqlx + "         UPDATE dbo.Stock_Master " & vbCrLf & _
'                          "         SET TM_Receipt=@Receipt,TM_Supply=@Supply,TM_LossReject=@Loss " & vbCrLf & _
'                          "         WHERE Warehouse_Code=@WHCode AND Item_Code=@Item_Code " & vbCrLf & _
'                          "     END  " & vbCrLf & _
'                          " IF @selisih=2 " & vbCrLf & _
'                          "     BEGIN  " & vbCrLf & _
'                          "         UPDATE dbo.Stock_Master " & vbCrLf & _
'                          "         SET NM_Receipt=@Receipt,NM_Supply=@Supply,NM_LossReject=@Loss " & vbCrLf & _
'                          "         WHERE Warehouse_Code=@WHCode AND Item_Code=@Item_Code " & vbCrLf & _
'                          "     END  " & vbCrLf & _
'                          "  " & vbCrLf
'
'            sqlx = sqlx + "  " & vbCrLf & _
'                          "     BEGIN  " & vbCrLf & _
'                          "         UPDATE dbo.Stock_Master " & vbCrLf & _
'                          "         SET LM_Current=LM_PreMonth+LM_Receipt-LM_Supply-LM_LossReject,TM_PreMonth=COALESCE(LM_Inventory,(LM_PreMonth+LM_Receipt-LM_Supply-LM_LossReject)) " & vbCrLf & _
'                          "         WHERE Warehouse_Code=@WHCode AND Item_Code=@Item_Code " & vbCrLf & _
'                          "  " & vbCrLf & _
'                          "         UPDATE dbo.Stock_Master " & vbCrLf & _
'                          "         SET TM_Current=TM_PreMonth+TM_Receipt-TM_Supply-TM_LossReject,NM_PreMonth=COALESCE(TM_Inventory,(TM_PreMonth+TM_Receipt-TM_Supply-TM_LossReject)) " & vbCrLf & _
'                          "         WHERE Warehouse_Code=@WHCode AND Item_Code=@Item_Code " & vbCrLf & _
'                          "  " & vbCrLf & _
'                          "         UPDATE dbo.Stock_Master " & vbCrLf
'
'            sqlx = sqlx + "         SET NM_Current=NM_PreMonth+NM_Receipt-NM_Supply-NM_LossReject " & vbCrLf & _
'                          "         WHERE Warehouse_Code=@WHCode AND Item_Code=@Item_Code " & vbCrLf & _
'                          "     END " & vbCrLf
'
'           Db.Execute sqlx
'
'           If RsPeriod("selisih") > 0 Then
'           SetGridData
'           End If
'        End If
'        RsPeriod.Close
'        Set RsPeriod = Nothing
    End If
ErrExit:
    Set adoRs = Nothing
    
    
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    Resume ErrExit
End Sub

Private Sub PrintToExcel()
    Dim xlapp As New Excel.application
    Dim xlBook As New Excel.Workbook
    Dim xlSheet As New Excel.Worksheet
    Dim lngCol As Long
    Dim lngRow As Long
    Dim strCol As String
    Dim lngColXL As Long
    Dim lngRowXL As Long
    
    Me.MousePointer = vbHourglass
    On Error GoTo errHandler
    
    LblErrMsg.Caption = ""
    If grid.Rows = 2 Then
        LblErrMsg.Caption = DisplayMsg("0013")
        GoTo ErrExit
    End If
        
    Set xlapp = CreateObject("Excel.Application")
    Set xlBook = xlapp.Workbooks.Add
    Set xlSheet = xlBook.ActiveSheet
    
    With xlSheet
        'header
        .Range("A1") = "Parts (Material) Receipt / Supply Inquiry"
        .Range("A2") = "Item Code"
        .Range("A3") = "Warehouse Code"
        .Range("F3") = "Month"
        
        .Range("B2") = ": " & cboitem.Text & " " & lblitem.Caption
        .Range("B3") = ": " & cboWarehouse.Text & " " & lblWarehouse.Caption
        .Range("G3") = ": " & Format(dtpMonth.Value, "MMM yyyy")
        
        'gridheader
        lngRowXL = 5
        lngColXL = 0
        For lngCol = 0 To grid.ColS - 1
            If grid.ColHidden(lngCol) = False Then
                lngColXL = lngColXL + 1
                .Range(GetExcelColumn(lngColXL) & lngRowXL) = grid.TextMatrix(0, lngCol)
            End If
        Next
                
        'griddata
        lngRowXL = 6
        lngColXL = 0
        For lngRow = 1 To grid.Rows - 1
            For lngCol = 0 To grid.ColS - 1
                If grid.ColHidden(lngCol) = False Then
                    lngColXL = lngColXL + 1
                    If grid.Cell(flexcpChecked, lngRow, lngCol) = flexChecked Then
                        .Range(GetExcelColumn(lngColXL) & lngRowXL) = "Yes"
                    ElseIf grid.Cell(flexcpChecked, lngRow, lngCol) = flexUnchecked Then
                        .Range(GetExcelColumn(lngColXL) & lngRowXL) = "No"
                    Else
                        .Range(GetExcelColumn(lngColXL) & lngRowXL) = grid.TextMatrix(lngRow, lngCol)
                    End If
                End If
            Next
            lngColXL = 0
            lngRowXL = lngRowXL + 1
        Next
        
        'format
        .Cells.Font.Name = "Tahoma"
        .Cells.Font.Size = "8"
        
        .Range("A1").Font.Size = "10"
        .Range("A1").Font.Bold = True
        
        .Range("A1").columnWidth = 12
        .Range("B1").columnWidth = 10
        .Range("C1").columnWidth = 12
        .Range("D1").columnWidth = 12
        .Range("E1").columnWidth = 12
        .Range("F1").columnWidth = 12
        .Range("G1").columnWidth = 10
        .Range("H1").columnWidth = 30
        .Range("I1").columnWidth = 3
        .Range("J1").columnWidth = 12
        .Range("K1").columnWidth = 12
        .Range("L1").columnWidth = 20
        
        .Range("A5").RowHeight = 15
        
        .Range("A5:" & GetExcelColumn(grid.ColS - 1) & 5).horizontalAlignment = xlHAlignCenter
        .Range("A5:" & GetExcelColumn(grid.ColS - 1) & 5).verticalAlignment = xlVAlignCenter
        
        .Range("A5:" & GetExcelColumn(grid.ColS - 1) & 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("A5:" & GetExcelColumn(grid.ColS - 1) & 5).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("A5:" & GetExcelColumn(grid.ColS - 1) & 5).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("A5:" & GetExcelColumn(grid.ColS - 1) & 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A5:" & GetExcelColumn(grid.ColS - 1) & 5).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range("A5:" & GetExcelColumn(grid.ColS - 1) & 5).Borders(xlInsideVertical).LineStyle = xlContinuous
        
        .Range("A6:" & GetExcelColumn(grid.ColS - 1) & lngRowXL - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("A6:" & GetExcelColumn(grid.ColS - 1) & lngRowXL - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("A6:" & GetExcelColumn(grid.ColS - 1) & lngRowXL - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("A6:" & GetExcelColumn(grid.ColS - 1) & lngRowXL - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A6:" & GetExcelColumn(grid.ColS - 1) & lngRowXL - 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range("A6:" & GetExcelColumn(grid.ColS - 1) & lngRowXL - 1).Borders(xlInsideHorizontal).Weight = xlHairline
        .Range("A6:" & GetExcelColumn(grid.ColS - 1) & lngRowXL - 1).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range("A6:" & GetExcelColumn(grid.ColS - 1) & lngRowXL - 1).Borders(xlInsideVertical).Weight = xlHairline
                
    End With
    
    xlapp.Visible = True
ErrExit:
    Me.MousePointer = vbDefault
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlapp = Nothing
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub

Private Function GetExcelColumn(ByVal lngColParse As Long) As String
    Dim strColTemp As String
    Dim lngColTemp As Long
    
    If lngColParse <= 26 Then
        strColTemp = Chr(64 + lngColParse)
    Else
        If lngColParse Mod 26 = 0 Then
            lngColTemp = 1
        Else
            lngColTemp = Fix(lngColParse / 26)
        End If
        strColTemp = Chr(64 + lngColTemp)
        
        lngColParse = lngColParse - (lngColTemp * 26)
        strColTemp = strColTemp & Chr(64 + lngColParse)
    End If
            
    GetExcelColumn = strColTemp
End Function

Private Sub CboItem_Change()
    LblErrMsg.Caption = ""
    If cboitem.MatchFound Then
        lblitem.Caption = cboitem.Column(2)
    Else
        lblitem.Caption = ""
    End If
    
        txtstock(0).Text = Format(0, gs_formatQty)
        txtstock(1).Text = Format(0, gs_formatQty)
        txtstock(2).Text = Format(0, gs_formatQty)
        txtstock(3).Text = Format(0, gs_formatQty)
        txtstock(4).Text = Format(0, gs_formatQty)
        txtstock(5).Text = Format(0, gs_formatQty)
        txtstock(6).Text = Format(0, gs_formatQty)
    'SetDataStock
    SetGridHeader
End Sub

Private Sub cbowarehouse_Change()
    LblErrMsg.Caption = ""
    If cboWarehouse.MatchFound Then
        lblWarehouse.Caption = cboWarehouse.Column(1)
    Else
        lblWarehouse.Caption = ""
    End If
      txtstock(0).Text = Format(0, gs_formatQty)
        txtstock(1).Text = Format(0, gs_formatQty)
        txtstock(2).Text = Format(0, gs_formatQty)
        txtstock(3).Text = Format(0, gs_formatQty)
        txtstock(4).Text = Format(0, gs_formatQty)
        txtstock(5).Text = Format(0, gs_formatQty)
        txtstock(6).Text = Format(0, gs_formatQty)
    'SetDataStock
    SetGridHeader
End Sub

Private Sub cmdBrowseItem_Click()
    LblErrMsg.Caption = ""
    Me.MousePointer = vbHourglass
    frm_BrowseItem.getItemCode = cboitem.Text
    frm_BrowseItem.Show 1
    cboitem.Text = frm_BrowseItem.getItemCode
    Me.MousePointer = vbDefault
End Sub

Private Sub CmdExcel_Click()
    LblErrMsg.Caption = ""
    PrintToExcel
End Sub

Private Sub cmdNormalize_Click()
    Me.MousePointer = vbHourglass
    On Error GoTo errHandler
        
    Db.Execute "exec sp_normalize_receipt_supply"
    
ErrExit:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub

Private Sub CmdSubMenu_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub CmdSubmit_Click()
Dim strSQL As String
Me.MousePointer = vbHourglass
'    If cboitem.MatchFound = True Then
'        strSQL = "exec [sp_normalize_receipt_supply_BY_Warehouse_Item] '" & Trim(cboWarehouse.Text) & "','" & Trim(cboitem.Text) & "'"
'        Db.Execute strSQL
'    End If
    LblErrMsg.Caption = ""
    SetGridData
Me.MousePointer = vbDefault
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblErrMsg.Caption = ErrMsg
    End If
End Sub

Private Sub dtpMonth_Change()
    LblErrMsg.Caption = ""
    If Format(dtpMonth.Value, "MM") < Format(dateUp, "MM") And Val(Format(dtpMonth.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then
        dtpMonth.Year = dtpMonth.Year + 1
        GoTo pass
    End If
    If Format(dtpMonth.Value, "MM") > Format(dateUp, "MM") And Val(Format(dtpMonth.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then dtpMonth.Year = dtpMonth.Year - 1
    
pass:
    dateUp = Format(dtpMonth.Value, "dd MMM yyyy")
        txtstock(0).Text = Format(0, gs_formatQty)
        txtstock(1).Text = Format(0, gs_formatQty)
        txtstock(2).Text = Format(0, gs_formatQty)
        txtstock(3).Text = Format(0, gs_formatQty)
        txtstock(4).Text = Format(0, gs_formatQty)
        txtstock(5).Text = Format(0, gs_formatQty)
        txtstock(6).Text = Format(0, gs_formatQty)
    'SetDataStock
    SetGridHeader
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    CtrlMenu1.FormName = Me.Name
    
    dateUp = Format(Date, "MMM-yyyy")
    dtpMonth.Value = Format(Date, "MMM-yyyy")
    
    SetComboItem
    SetComboWarehouse
    SetGridHeader
    
    txtstock(0).Text = Format(0, gs_formatQty)
    txtstock(1).Text = Format(0, gs_formatQty)
    txtstock(2).Text = Format(0, gs_formatQty)
    txtstock(3).Text = Format(0, gs_formatQty)
    txtstock(4).Text = Format(0, gs_formatQty)
    txtstock(5).Text = Format(0, gs_formatQty)
    txtstock(6).Text = Format(0, gs_formatQty)
End Sub
