VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmInterfaceLink 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Interface Link Setup"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmInterfaceLink.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Delete"
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
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "FFTT*/"
      Top             =   9945
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "C&ancel"
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
      Left            =   12540
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "FFTT*/"
      Top             =   9945
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13020
      TabIndex        =   32
      Tag             =   "FTTF*/"
      Top             =   180
      Width           =   1875
      _extentx        =   3307
      _extenty        =   767
   End
   Begin VB.TextBox TxtProfit 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   13500
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "TTFF*/"
      Top             =   8640
      Width           =   1320
   End
   Begin VB.TextBox TxtCostCenter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   12180
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "TTFF*/"
      Top             =   8640
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   960
      Left            =   270
      TabIndex        =   24
      Tag             =   "TFTF*/"
      Top             =   1170
      Width           =   14640
      Begin VB.Label LblType 
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
         Left            =   3240
         TabIndex        =   26
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   60
      End
      Begin VB.Line Line1 
         X1              =   3240
         X2              =   5340
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
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
         Left            =   300
         TabIndex        =   25
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   1470
      End
      Begin MSForms.ComboBox CboType 
         Height          =   345
         Left            =   2040
         TabIndex        =   0
         Tag             =   "TTFF*/"
         Top             =   300
         Width           =   990
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "1746;609"
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
      TabIndex        =   22
      Tag             =   "TTTF*/"
      Top             =   9120
      Width           =   14640
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
         TabIndex        =   23
         Tag             =   "TTFF*/"
         Top             =   180
         Width           =   14325
      End
   End
   Begin VB.TextBox TxtPost 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   8745
      MaxLength       =   2
      TabIndex        =   4
      Tag             =   "TTFF*/"
      Top             =   8640
      Width           =   885
   End
   Begin VB.TextBox TxtAccount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   9690
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "TTFF*/"
      Top             =   8640
      Width           =   1455
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "FFTT*/"
      Top             =   9945
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
      TabIndex        =   14
      Tag             =   "TFFT*/"
      Top             =   9945
      Width           =   1125
   End
   Begin VB.TextBox TxtTax 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   11220
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "TTFF*/"
      Top             =   8640
      Width           =   900
   End
   Begin VB.CommandButton Cmd_save 
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "FFTT*/"
      Top             =   9945
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5460
      Left            =   270
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "TFTF*/"
      Top             =   2415
      Width           =   14640
      _cx             =   25823
      _cy             =   9631
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
      ExplorerBar     =   1
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
   Begin VB.Label lblTradeName 
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
      Left            =   2760
      TabIndex        =   34
      Tag             =   "TTFF*/"
      Top             =   8700
      Width           =   60
   End
   Begin VB.Line Line3 
      X1              =   2760
      X2              =   6960
      Y1              =   8940
      Y2              =   8940
   End
   Begin VB.Line Line2 
      X1              =   8040
      X2              =   8640
      Y1              =   8940
      Y2              =   8940
   End
   Begin VB.Label LblCurr 
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
      Left            =   8040
      TabIndex        =   33
      Tag             =   "TTFF*/"
      Top             =   8700
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profit Center"
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
      Index           =   10
      Left            =   13455
      TabIndex        =   31
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Center"
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
      Left            =   12195
      TabIndex        =   30
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Currency"
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
      Left            =   7095
      TabIndex        =   29
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   795
   End
   Begin MSForms.ComboBox CboCurr 
      Height          =   345
      Left            =   7080
      TabIndex        =   3
      Tag             =   "TTFF*/"
      Top             =   8640
      Width           =   870
      VariousPropertyBits=   746604571
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "1535;609"
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
      Caption         =   "Trade Code"
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
      Left            =   1275
      TabIndex        =   28
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   1005
   End
   Begin MSForms.ComboBox CboTrade 
      Height          =   345
      Left            =   1260
      TabIndex        =   2
      Tag             =   "TTFF*/"
      Top             =   8640
      Width           =   1410
      VariousPropertyBits=   746604571
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "2487;609"
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
      Caption         =   "Tr. Code"
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
      Left            =   375
      TabIndex        =   27
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   750
   End
   Begin MSForms.ComboBox CboTrans 
      Height          =   345
      Left            =   360
      TabIndex        =   1
      Tag             =   "TTFF*/"
      Top             =   8640
      Width           =   870
      VariousPropertyBits=   746604571
      MaxLength       =   5
      DisplayStyle    =   3
      Size            =   "1535;609"
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
      Caption         =   "Account No"
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
      Left            =   9690
      TabIndex        =   21
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   960
   End
   Begin MSForms.ComboBox cboAdd 
      Height          =   345
      Left            =   13710
      TabIndex        =   6
      Tag             =   "TTFF*/"
      Top             =   8640
      Visible         =   0   'False
      Width           =   870
      VariousPropertyBits=   746604571
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1535;609"
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
      Caption         =   "Tax Code"
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
      Left            =   11235
      TabIndex        =   20
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Post Key"
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
      Left            =   8760
      TabIndex        =   19
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   270
      Tag             =   "TTTF*/"
      Top             =   8205
      Width           =   14625
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Interface Link Setup"
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
      Left            =   120
      TabIndex        =   18
      Tag             =   "TTTF*/"
      Top             =   240
      Width           =   14610
   End
   Begin VB.Label LblCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   420
      TabIndex        =   17
      Tag             =   "TTFF*/"
      Top             =   8265
      Width           =   975
   End
   Begin VB.Label LblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   1500
      TabIndex        =   16
      Tag             =   "TTFF*/"
      Top             =   8265
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   270
      Tag             =   "TTTF*/"
      Top             =   8520
      Width           =   14625
   End
End
Attribute VB_Name = "FrmInterfaceLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim sql As String
Dim RS As New Recordset
Dim ubah As Boolean, hapus As Boolean, gavalid As Boolean, ubahedate As Boolean
Dim SDate, EDate, sdateawal, edateakhir
Dim i As Integer

Dim bteColSelect As Byte
Dim bteColJournalCode As Byte
Dim bteColTransCode As Byte
Dim bteColTransDesc As Byte
Dim bteColTradeCode As Byte
Dim bteColTradeName As Byte
Dim bteColCurrCode As Byte
Dim bteColCurrDesc As Byte
Dim bteColPostKey As Byte
Dim bteColAccountNo As Byte
Dim bteColTaxCode As Byte
Dim bteColCostCenter As Byte
Dim bteColProfitCenter As Byte

Sub Header()

    Dim X As Integer

    bteColSelect = 0
    bteColJournalCode = 1
    bteColTransCode = 2
    bteColTransDesc = 3
    bteColTradeCode = 4
    bteColTradeName = 5
    bteColCurrCode = 6
    bteColCurrDesc = 7
    bteColPostKey = 8
    bteColAccountNo = 9
    bteColTaxCode = 10
    bteColCostCenter = 11
    bteColProfitCenter = 12

    With grid
        .clear

        .Rows = 1
        .ColS = 13

        .TextMatrix(0, bteColSelect) = "S"
        .TextMatrix(0, bteColJournalCode) = "JournalCode"
        .TextMatrix(0, bteColTransCode) = "Tr. Code"
        .TextMatrix(0, bteColTransDesc) = "Trans. Code"
        .TextMatrix(0, bteColTradeCode) = "Trade Code"
        .TextMatrix(0, bteColTradeName) = "Trade Name"
        .TextMatrix(0, bteColCurrCode) = "CurrCode"
        .TextMatrix(0, bteColCurrDesc) = "Currency"
        .TextMatrix(0, bteColPostKey) = "Post Key"
        .TextMatrix(0, bteColAccountNo) = "Account No"
        .TextMatrix(0, bteColTaxCode) = "Tax Code"
        .TextMatrix(0, bteColCostCenter) = "Cost Center"
        .TextMatrix(0, bteColProfitCenter) = "Profit Center"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColJournalCode) = 0
        .ColWidth(bteColTransCode) = 0
        .ColWidth(bteColTransDesc) = 1550
        .ColWidth(bteColTradeCode) = 1150
        .ColWidth(bteColTradeName) = 4250
        .ColWidth(bteColCurrCode) = 0
        .ColWidth(bteColCurrDesc) = 1000
        .ColWidth(bteColPostKey) = 1000
        .ColWidth(bteColAccountNo) = 1500
        .ColWidth(bteColTaxCode) = 1000
        .ColWidth(bteColCostCenter) = 1250
        .ColWidth(bteColProfitCenter) = 1250

        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColTransCode) = flexAlignCenterCenter
        .ColAlignment(bteColTransDesc) = flexAlignLeftCenter
        .ColAlignment(bteColTradeCode) = flexAlignLeftCenter
        .ColAlignment(bteColTradeName) = flexAlignLeftCenter
        .ColAlignment(bteColCurrDesc) = flexAlignCenterCenter
        .ColAlignment(bteColPostKey) = flexAlignCenterCenter
        .ColAlignment(bteColAccountNo) = flexAlignCenterCenter
        .ColAlignment(bteColTaxCode) = flexAlignCenterCenter
        .ColAlignment(bteColCostCenter) = flexAlignCenterCenter
        .ColAlignment(bteColProfitCenter) = flexAlignCenterCenter

        .ColHidden(bteColSelect) = False
        .ColHidden(bteColJournalCode) = False
        .ColHidden(bteColCurrCode) = False
        .ColHidden(bteColTransCode) = False

        .EditMaxLength = 1
    End With

End Sub

Sub AddTransCode()
    Dim StrAdd As String
    Dim RsAdd As New ADODB.Recordset
    Dim NN As Integer

    StrAdd = "SELECT TR_Code, TR_CodeDesc FROM TR_Cls " & vbCrLf & _
                "           WHERE TR_Type='" & CboType.Text & "'" & vbCrLf

    CboTrans.clear

    If RsAdd.State <> adStateClosed Then RsAdd.Close

    Set RsAdd = Db.Execute(StrAdd)
    NN = 0

    Do While Not RsAdd.EOF
        CboTrans.AddItem ""
        CboTrans.List(NN, 0) = Trim(RsAdd("TR_Code") & "")
        CboTrans.List(NN, 1) = Trim(RsAdd("TR_CodeDesc") & "")
        NN = NN + 1
        RsAdd.MoveNext
    Loop

    CboTrans.columnCount = 2
    CboTrans.ColumnWidths = "50 pt;100 pt"
    CboTrans.ListWidth = "150 pt"
    'CboTrans.clear
    
'    If CboType = "AR" Then
'        CboTrans.AddItem ""
'        CboTrans.List(0, 0) = "AR"
'        CboTrans.List(0, 1) = "Account Receivable"
'
'        CboTrans.AddItem ""
'        CboTrans.List(1, 0) = "SL"
'        CboTrans.List(1, 1) = "Sales"
'
'        CboTrans.AddItem ""
'        CboTrans.List(2, 0) = "ST"
'        CboTrans.List(2, 1) = "Sales Tax"
'
'    ElseIf CboType = "MT" Then
'        CboTrans.AddItem ""
'        CboTrans.List(0, 0) = "AR"
'        CboTrans.List(0, 1) = "Account Receivable"
'
'        CboTrans.AddItem ""
'        CboTrans.List(1, 0) = "SL"
'        CboTrans.List(1, 1) = "Sales"
'
'        CboTrans.AddItem ""
'        CboTrans.List(2, 0) = "ST"
'        CboTrans.List(2, 1) = "Sales Tax"
'
'    ElseIf CboType = "WP" Then
'        CboTrans.AddItem ""
'        CboTrans.List(0, 0) = "AR"
'        CboTrans.List(0, 1) = "Account Receivable"
'
'        CboTrans.AddItem ""
'        CboTrans.List(1, 0) = "SL"
'        CboTrans.List(1, 1) = "Sales"
'
'        CboTrans.AddItem ""
'        CboTrans.List(2, 0) = "ST"
'        CboTrans.List(2, 1) = "Sales Tax"
'
'    ElseIf CboType = "FG" Then
'        CboTrans.AddItem ""
'        CboTrans.List(0, 0) = "AR"
'        CboTrans.List(0, 1) = "Account Receivable"
'
'        CboTrans.AddItem ""
'        CboTrans.List(1, 0) = "SL"
'        CboTrans.List(1, 1) = "Sales"
'
'        CboTrans.AddItem ""
'        CboTrans.List(2, 0) = "ST"
'        CboTrans.List(2, 1) = "Sales Tax"
'
'    ElseIf CboType = "AP" Then
'        CboTrans.AddItem ""
'        CboTrans.List(0, 0) = "AP"
'        CboTrans.List(0, 1) = "Account Payable"
'
'        CboTrans.AddItem ""
'        CboTrans.List(1, 0) = "PC"
'        CboTrans.List(1, 1) = "Purchase"
'
'        CboTrans.AddItem ""
'        CboTrans.List(2, 0) = "PT"
'        CboTrans.List(2, 1) = "Purchase Tax"
'
'        CboTrans.AddItem ""
'        CboTrans.List(3, 0) = "PH"
'        CboTrans.List(3, 1) = "Purchase Tax (PPH)"
'
'    End If
'
'
'    CboTrans.ColumnCount = 2
'    CboTrans.ColumnWidths = "50 pt;100 pt"
'    CboTrans.ListWidth = "150 pt"
    
End Sub

Sub AddtoTrade()
    Dim StrAdd As String
    Dim RsAdd As New ADODB.Recordset
    Dim NN As Integer
    
    StrAdd = "Select '0000' Trade_Code, 'Common' Trade_Name " & vbCrLf & _
                "   UNION ALL " & vbCrLf
                
    If CboType = "AR" Then
        StrAdd = StrAdd & "Select Trade_Code, Trade_Name From Trade_Master " & vbCrLf & _
                    "   Where Trade_Cls=2 AND left(trade_Code,1)='C' " & vbCrLf & _
                    "            ORDER BY Trade_Code " & vbCrLf
        
    ElseIf CboType = "AP" Then
        StrAdd = StrAdd & "Select Trade_Code, Trade_Name From Trade_Master" & vbCrLf & _
                    "   Where Trade_Cls=2 AND left(trade_Code,1)='S' " & vbCrLf & _
                    "            ORDER BY Trade_Code " & vbCrLf
                    
    ElseIf CboType = "MT" Then
        StrAdd = StrAdd & "Select Trade_Code, Trade_Name From Trade_Master" & vbCrLf & _
                    "   Where Trade_Cls=1 AND left(trade_Code,1)='K' " & vbCrLf & _
                    "            ORDER BY Trade_Code " & vbCrLf
        
    ElseIf CboType = "WP" Then
        StrAdd = StrAdd & "Select Trade_Code, Trade_Name From Trade_Master" & vbCrLf & _
                    "   Where Trade_Cls=1 AND left(trade_Code,1)='K' " & vbCrLf & _
                    "            ORDER BY Trade_Code " & vbCrLf
        
    Else
        StrAdd = StrAdd & "Select Trade_Code, Trade_Name From Trade_Master" & vbCrLf & _
                    "   Where Trade_Cls=999  " & vbCrLf & _
                    "            ORDER BY Trade_Code " & vbCrLf
    End If
    
    cbotrade.clear
    
    If RsAdd.State <> adStateClosed Then RsAdd.Close
    
    Set RsAdd = Db.Execute(StrAdd)
    NN = 0
    
    Do While Not RsAdd.EOF
        cbotrade.AddItem ""
        cbotrade.List(NN, 0) = Trim(RsAdd("Trade_Code") & "")
        cbotrade.List(NN, 1) = Trim(RsAdd("Trade_Name") & "")
        NN = NN + 1
        RsAdd.MoveNext
    Loop
    
    cbotrade.columnCount = 2
    cbotrade.ColumnWidths = "50 pt;250 pt"
    cbotrade.ListWidth = "300 pt"
End Sub

Sub AddtoCurr()
    Dim StrAdd As String
    Dim RsAdd As New ADODB.Recordset
    Dim NN As Integer
    
    StrAdd = "Select Curr_Cls, Description From Curr_Cls" & vbCrLf & _
                "            ORDER BY Curr_Cls " & vbCrLf
    
    cbocurr.clear
    
    If RsAdd.State <> adStateClosed Then RsAdd.Close
    
    Set RsAdd = Db.Execute(StrAdd)
    NN = 0
    
    Do While Not RsAdd.EOF
        cbocurr.AddItem ""
        cbocurr.List(NN, 0) = Trim(RsAdd("Curr_Cls") & "")
        cbocurr.List(NN, 1) = Trim(RsAdd("Description") & "")
        NN = NN + 1
        RsAdd.MoveNext
    Loop
    
    cbocurr.columnCount = 2
    cbocurr.ColumnWidths = "50 pt;100 pt"
    cbocurr.ListWidth = "150 pt"
End Sub

Sub AddTRType()
    Dim StrAdd As String
    Dim RsAdd As New ADODB.Recordset
    Dim NN As Integer

    StrAdd = "Select Distinct TR_Type, TR_TypeDesc from TR_Cls" & vbCrLf & _
                "            ORDER BY TR_Type Asc " & vbCrLf

    CboType.clear

    If RsAdd.State <> adStateClosed Then RsAdd.Close

    Set RsAdd = Db.Execute(StrAdd)
    NN = 0

    Do While Not RsAdd.EOF
        CboType.AddItem ""
        CboType.List(NN, 0) = Trim(RsAdd("TR_Type") & "")
        CboType.List(NN, 1) = Trim(RsAdd("TR_TypeDesc") & "")
        NN = NN + 1
        RsAdd.MoveNext
    Loop

    CboType.columnCount = 2
    CboType.ColumnWidths = "50 pt;100 pt"
    CboType.ListWidth = "150 pt"
End Sub

Sub ObjectEnable(setting As Boolean)
    CboType.Enabled = setting
    
    CboTrans.Enabled = setting
    cbotrade.Enabled = setting
    cbocurr.Enabled = setting
    
End Sub

Sub GridView()
    Dim StrSearch As String
    Dim RsSearch As New ADODB.Recordset
    
    On Error GoTo ErrSearch
    
    LblErrMsg = ""
    
    Call Header
                    
    StrSearch = " Select JournalType, LI.TransCode, TC.TRansDesc, LI.Trade_Code, Trade_Name, Curr_Code, cc.Description Curr_Desc, Posting_Key,  " & vbCrLf & _
                            "   Account_No, Tax_Code, Cost_Center, Profit_Center " & vbCrLf & _
                            "  From LinkInterface LI " & vbCrLf & _
                            "   LEFT JOIN  " & vbCrLf & _
                            "       (SELECT Trade_Code, Trade_Name FROM Trade_Master " & vbCrLf & _
                            "           UNION ALL " & vbCrLf & _
                            "        SELECT '0000', 'Common' " & vbCrLf & _
                            "        ) TM ON LI.Trade_Code=TM.Trade_Code " & vbCrLf & _
                            "   LEFT JOIN Curr_Cls CC ON Li.Curr_Code=CC.Curr_Cls " & vbCrLf & _
                            "   LEFT JOIN " & vbCrLf & _
                            "       (SELECT 'AP' TransCode, 'Acc. Payable'  TransDesc" & vbCrLf & _
                            "           UNION ALL " & vbCrLf & _
                            "        SELECT 'PC', 'Purchase' " & vbCrLf & _
                            "           UNION ALL " & vbCrLf & _
                            "        SELECT 'PT', 'Purchase Tax' " & vbCrLf & _
                            "           UNION ALL " & vbCrLf & _
                            "        SELECT 'PH', 'Purchase Tax (PPH)' " & vbCrLf & _
                            "           UNION ALL " & vbCrLf & _
                            "        SELECT 'AR', 'Acc. Receivable' " & vbCrLf & _
                            "           UNION ALL " & vbCrLf & _
                            "        SELECT 'SL', 'Sales' " & vbCrLf & _
                            "           UNION ALL " & vbCrLf & _
                            "        SELECT 'ST', 'Sales Tax' " & vbCrLf & _
                            "        ) TC ON LI.TransCode=TC.TransCode " & vbCrLf & _
                            "    Where JournalType='" & Trim(CboType) & "'  Order by LI.Trade_Code, LI.TransCode "
                    
    If RsSearch.State <> adStateClosed Then RsSearch.Close
    
    Set RsSearch = Db.Execute(StrSearch)
    
    Do While Not RsSearch.EOF
        grid.AddItem ""
        
        grid.Cell(flexcpChecked, grid.Rows - 1, bteColSelect) = flexUnchecked
        grid.TextMatrix(grid.Rows - 1, bteColTransCode) = Trim(RsSearch("TransCode") & "")
        grid.TextMatrix(grid.Rows - 1, bteColTransDesc) = Trim(RsSearch("TransDesc") & "")
        grid.TextMatrix(grid.Rows - 1, bteColTradeCode) = Trim(RsSearch("Trade_Code") & "")
        grid.TextMatrix(grid.Rows - 1, bteColTradeName) = Trim(RsSearch("Trade_Name") & "")
        grid.TextMatrix(grid.Rows - 1, bteColCurrCode) = Trim(RsSearch("Curr_Code") & "")
        grid.TextMatrix(grid.Rows - 1, bteColCurrDesc) = Trim(RsSearch("Curr_Desc") & "")
        grid.TextMatrix(grid.Rows - 1, bteColPostKey) = Trim(RsSearch("Posting_Key") & "")
        grid.TextMatrix(grid.Rows - 1, bteColAccountNo) = Trim(RsSearch("Account_No") & "")
        grid.TextMatrix(grid.Rows - 1, bteColTaxCode) = Trim(RsSearch("Tax_Code") & "")
        grid.TextMatrix(grid.Rows - 1, bteColCostCenter) = Trim(RsSearch("Cost_Center") & "")
        grid.TextMatrix(grid.Rows - 1, bteColProfitCenter) = Trim(RsSearch("Profit_Center") & "")
        
        RsSearch.MoveNext
    Loop
    
    
    Me.MousePointer = vbDefault
    Exit Sub
    
ErrSearch:
    Me.MousePointer = vbDefault
    LblErrMsg = "[" & err.number & "]-" & err.Description
    err.clear
    
End Sub

Private Sub cboAdd_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub cbocurr_Change()
    Call cbocurr_Click
End Sub

Private Sub cbocurr_Click()
    LblErrMsg = ""
    
    If cbocurr.ListIndex < 0 Then
        LblCurr = ""
    Else
        LblCurr = cbocurr.Column(1)
    End If
End Sub

Private Sub cbotrade_Change()
    Call CboTrade_Click
End Sub

Private Sub CboTrade_Click()
    LblErrMsg = ""
    
    If cbotrade.ListIndex < 0 Then
        lblTradeName = ""
    Else
        lblTradeName = Trim(cbotrade.Column(1) & "")
    End If
End Sub

Private Sub cboTrans_Change()
    Call CboTrans_Click
End Sub

Private Sub CboTrans_Click()
    LblErrMsg = ""
End Sub

Private Sub cboType_Change()
    Call CboType_Click
End Sub

Private Sub CboType_Click()
    LblErrMsg = ""
    
    If CboType.ListIndex < 0 Then
         LblType = ""
    Else
        LblType = CboType.Column(1)
    End If

    Call AddTransCode
    Call AddtoTrade
    Call GridView
End Sub

Private Sub Cmd_Save_Click()
    Dim StrInsert As String
    Dim StrUpdate As String
    Dim RowAff As Integer
    
    'On Error GoTo ErrSave
    
    ' Insert Validation
    ' #################
    
    If hakUpdate(Me.Name) = 0 Then
        LblErrMsg = DisplayMsg(3008)
        Exit Sub
    End If
    
    If CboType.ListIndex < 0 Then
        LblErrMsg = "Please select valid Transaction Type !"
        Exit Sub
    End If
    
    If CboTrans.ListIndex < 0 Then
        LblErrMsg = "Please select valid Transaction Code !"
        Exit Sub
    End If
    
    If cbotrade.ListIndex < 0 Then
        LblErrMsg = "Please select Trade Code !"
        Exit Sub
    End If
    
    If cbocurr.ListIndex < 0 Then
        LblErrMsg = "Please select valid Currency !"
        Exit Sub
    End If
    
    ' #################
    
    LblErrMsg = ""
        
    StrUpdate = " UPDATE LinkInterface " & vbCrLf & _
                    "   SET  Posting_Key='" & Trim(TxtPost) & "', " & vbCrLf & _
                    "           Account_No='" & Trim(TxtAccount) & "', " & vbCrLf & _
                    "           Tax_Code='" & Trim(TxtTax) & "', " & vbCrLf & _
                    "           Cost_Center='" & Trim(TxtCostCenter) & "', " & vbCrLf & _
                    "           Profit_Center='" & Trim(TxtProfit) & "' " & vbCrLf & _
                    "   Where JournalType='" & Trim(CboType) & "' " & vbCrLf & _
                    "       And TransCode='" & Trim(CboTrans) & "' " & vbCrLf & _
                    "       And Trade_Code='" & Trim(cbotrade) & "' " & vbCrLf & _
                    "       And Curr_Code='" & Trim(cbocurr) & "' " & vbCrLf
    
    Db.Execute StrUpdate, RowAff
    
    If RowAff <= 0 Then
    
        StrInsert = " INSERT INTO LinkInterface " & vbCrLf & _
                                "         ( JournalType , TransCode, Trade_Code ,Curr_Code , " & vbCrLf & _
                                "               Posting_Key ,Account_No ,Tax_Code ,Cost_Center , Profit_Center) " & vbCrLf & _
                                "       VALUES  ( '" & Trim(CboType) & "' , '" & Trim(CboTrans) & "' , '" & Trim(cbotrade) & "' , '" & Trim(cbocurr) & "' , " & vbCrLf & _
                                "                       '" & Trim(TxtPost) & "' , '" & Trim(TxtAccount) & "' , '" & Trim(TxtTax) & "' , '" & Trim(TxtCostCenter) & "' , '" & Trim(TxtProfit) & "'  ) "
    
        Db.Execute StrInsert
        
    End If
        
    Call GridView
    Call cmdCancel_Click
    
    LblErrMsg = DisplayMsg(1000)

    Exit Sub

ErrSave:
    LblErrMsg = "[" & err.number & "]-" & err.Description
    err.clear

End Sub

Private Sub cmdCancel_Click()
    Call Kosong
    Call GridView
End Sub

Private Sub cmdDelete_Click()
    Dim StrDelete As String
    Dim bteConfirm As Byte
    
    On Error GoTo ErrDelete
    
    LblErrMsg = ""
    
    ' Delete Validation
    ' ################
    
    If hakUpdate(Me.Name) = 0 Then
        LblErrMsg = DisplayMsg(3008)
        Exit Sub
    End If
    
    If CboTrans.Enabled = True Then
        LblErrMsg = "Please select valid data !"
        Exit Sub
    End If
    
    ' ################
    
    bteConfirm = MsgBox("Are you sure want to delete this link setting ?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete Confirm")
    
    If bteConfirm = vbYes Then
        StrDelete = "Delete From LinkInterface" & vbCrLf & _
                        "   Where JournalType='" & Trim(CboType) & "' " & vbCrLf & _
                        "       And TransCode='" & Trim(CboTrans) & "' " & vbCrLf & _
                        "       And Trade_Code='" & Trim(cbotrade) & "' " & vbCrLf & _
                        "       And Curr_Code='" & Trim(cbocurr) & "' " & vbCrLf
        
        Db.Execute (StrDelete)
        
        LblErrMsg = DisplayMsg(1201)
        Call cmdCancel_Click
        Call GridView
        
    End If
        
    Exit Sub

ErrDelete:
    LblErrMsg = "[" & err.number & "]-" & err.Description
    err.clear
    
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    Kosong
    Call AddtoCurr
    Call AddTRType
    Header
    
'    CboType.clear
'    CboType.AddItem ""
'    CboType.List(0, 0) = "AR"
'    CboType.List(0, 1) = "Account Receivable"
'
'    CboType.AddItem ""
'    CboType.List(1, 0) = "AP"
'    CboType.List(1, 1) = "Account Payable"
'
'    CboType.AddItem ""
'    CboType.List(2, 0) = "MT"
'    CboType.List(2, 1) = "Material"
'
'    CboType.AddItem ""
'    CboType.List(3, 0) = "WP"
'    CboType.List(3, 1) = "Working Process"
'
'    CboType.AddItem ""
'    CboType.List(4, 0) = "FG"
'    CboType.List(4, 1) = "Finshed Good"
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
'    With Anchor1
'      .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
'      .DoInit
'    End With
End Sub

Sub Kosong()

    CboTrans = ""
    cbotrade = ""
    cbocurr = ""
    
    TxtPost = ""
    TxtAccount = ""
    TxtTax = ""
    TxtCostCenter = ""
    TxtProfit = ""
    
    Call ObjectEnable(True)

End Sub

Private Sub Cmd_SubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub cmd_clear_Click()
    Kosong
    Header
    CboType = ""
    CboType.SetFocus
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblErrMsg.Caption = ErrMsg
    End If
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    Dim X As Long
    
    LblErrMsg = ""
    
    If grid.Cell(flexcpChecked, Row, bteColSelect) = flexChecked Then
        For X = 1 To grid.Rows - 1
            If X <> Row Then
                grid.Cell(flexcpChecked, X, bteColSelect) = flexUnchecked
            End If
        Next
        
        CboTrans = grid.TextMatrix(Row, bteColTransCode)
        cbotrade = grid.TextMatrix(Row, bteColTradeCode)
        cbocurr = grid.TextMatrix(Row, bteColCurrCode)
        TxtPost = grid.TextMatrix(Row, bteColPostKey)
        TxtAccount = grid.TextMatrix(Row, bteColAccountNo)
        TxtTax = grid.TextMatrix(Row, bteColTaxCode)
        TxtCostCenter = grid.TextMatrix(Row, bteColCostCenter)
        TxtProfit = grid.TextMatrix(Row, bteColProfitCenter)
        
        Call ObjectEnable(False)
    Else
        'Call BtnClear_Click
    End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Col <> bteColSelect Then Cancel = True
End Sub

Private Sub txtcost_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 And KeyAscii <> 8 Then KeyAscii = 0
End Sub


