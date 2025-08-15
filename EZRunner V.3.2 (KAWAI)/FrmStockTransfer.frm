VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmStockTransfer 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Stock Transfer"
   ClientHeight    =   7470
   ClientLeft      =   2190
   ClientTop       =   945
   ClientWidth     =   10695
   Icon            =   "FrmStockTransfer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_SubMenu 
      BackColor       =   &H00A6D2FF&
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
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   6840
      Width           =   1035
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   2835
      Left            =   390
      TabIndex        =   19
      Top             =   975
      Width           =   9885
      Begin VB.TextBox lblLocationName 
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
         Height          =   285
         Index           =   3
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2100
         Width           =   4815
      End
      Begin VB.TextBox lblLocationName 
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
         Index           =   2
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1260
         Width           =   4815
      End
      Begin VB.TextBox lblLocationName 
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
         Height          =   270
         Index           =   1
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1680
         Width           =   4815
      End
      Begin VB.TextBox lblLocationName 
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
         Index           =   0
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   840
         Width           =   4815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   4380
         TabIndex        =   3
         Top             =   1245
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   4380
         TabIndex        =   6
         Top             =   2115
         Width           =   330
      End
      Begin MSComCtl2.DTPicker DMonth 
         Height          =   315
         Left            =   1785
         TabIndex        =   0
         Top             =   345
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   138543107
         CurrentDate     =   37798
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   4845
         X2              =   9540
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   4845
         X2              =   9510
         Y1              =   2385
         Y2              =   2385
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Item Code"
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
         TabIndex        =   24
         Top             =   1305
         Width           =   1410
      End
      Begin MSForms.ComboBox cbolocationcd 
         Height          =   315
         Index           =   2
         Left            =   1785
         TabIndex        =   2
         Top             =   1245
         Width           =   2550
         VariousPropertyBits=   612386843
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "4498;556"
         ListRows        =   15
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
         Caption         =   "Transfer Date"
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
         TabIndex        =   23
         Top             =   405
         Width           =   1185
      End
      Begin MSForms.ComboBox cbolocationcd 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   1
         Top             =   810
         Width           =   2550
         VariousPropertyBits=   612386843
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "4498;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From WareHouse"
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
         TabIndex        =   22
         Top             =   870
         Width           =   1470
      End
      Begin MSForms.ComboBox cbolocationcd 
         Height          =   315
         Index           =   1
         Left            =   1785
         TabIndex        =   4
         Top             =   1680
         Width           =   2550
         VariousPropertyBits=   612386843
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "4498;556"
         ListRows        =   15
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To WareHouse"
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
         TabIndex        =   21
         Top             =   1740
         Width           =   1245
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4845
         X2              =   9525
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4845
         X2              =   9525
         Y1              =   1950
         Y2              =   1950
      End
      Begin MSForms.ComboBox cbolocationcd 
         Height          =   315
         Index           =   3
         Left            =   1785
         TabIndex        =   5
         Top             =   2115
         Width           =   2550
         VariousPropertyBits=   612386843
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "4498;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Item Code"
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
         TabIndex        =   20
         Top             =   2175
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   353
      TabIndex        =   16
      Top             =   6150
      Width           =   9855
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
         Left            =   105
         TabIndex        =   17
         Top             =   195
         Width           =   9615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2205
      Left            =   360
      TabIndex        =   14
      Top             =   3855
      Width           =   9900
      Begin VB.TextBox txtLot 
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
         Left            =   3519
         TabIndex        =   8
         Top             =   300
         Width           =   1590
      End
      Begin VB.TextBox Text6 
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
         Left            =   6180
         MaxLength       =   50
         TabIndex        =   9
         Top             =   315
         Width           =   3390
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Left            =   7260
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1665
         Width           =   1845
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Left            =   7260
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1845
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Left            =   5010
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1665
         Width           =   1890
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Left            =   5010
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1890
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   1548
         TabIndex        =   7
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot No"
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
         Left            =   2751
         TabIndex        =   36
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label19 
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
         Index           =   6
         Left            =   5307
         TabIndex        =   35
         Top             =   375
         Width           =   675
      End
      Begin VB.Label Label3 
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
         Left            =   270
         TabIndex        =   34
         Top             =   1695
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   270
         TabIndex        =   33
         Top             =   1245
         Width           =   435
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   9960
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Label LFrom 
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
         Left            =   1380
         TabIndex        =   32
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "After Transfer"
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
         Left            =   7260
         TabIndex        =   31
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label LTo 
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
         Left            =   1380
         TabIndex        =   28
         Top             =   1710
         Width           =   915
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Before Transfer"
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
         Left            =   5025
         TabIndex        =   26
         Top             =   930
         Width           =   1365
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer Qty"
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
         Left            =   270
         TabIndex        =   15
         Top             =   345
         Width           =   1080
      End
   End
   Begin VB.CommandButton Cmd_Clear 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clea&r"
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
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6855
      Width           =   1035
   End
   Begin VB.CommandButton Cmd_Cancel 
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
      Left            =   8025
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6855
      Width           =   1035
   End
   Begin VB.CommandButton Cmd_Submit 
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
      Left            =   9165
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6855
      Width           =   1035
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   8595
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   180
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   870
      Left            =   345
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   7725
      Visible         =   0   'False
      Width           =   9855
      _cx             =   17383
      _cy             =   1535
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
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Transfer"
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
      Left            =   360
      TabIndex        =   18
      Top             =   270
      Width           =   9645
   End
End
Attribute VB_Name = "FrmStockTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BteColTgl As Byte
Dim BteColFromWH As Byte
Dim BteColToWH As Byte
Dim bteColLot As Byte
Dim bteColQty As Byte
Dim bteColItem As Byte
Dim bteColDescription  As Byte
Dim bteColRemark As Byte
Dim bteColItemTo As Byte

Sub GridHeader()
    With grid
        .Rows = 1
        .ColS = 9
        
        BteColTgl = 0
        bteColItem = 1
        bteColItemTo = 2
        bteColDescription = 3
        bteColLot = 4
        bteColQty = 5
        BteColFromWH = 6
        BteColToWH = 7
        bteColRemark = 8
        
        .TextMatrix(0, BteColTgl) = "Tanggal"
        .TextMatrix(0, bteColItem) = "From Item Code"
        .TextMatrix(0, bteColItemTo) = "To Item Code"
        .TextMatrix(0, bteColDescription) = "Description"
        .TextMatrix(0, bteColLot) = "Lot"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, BteColFromWH) = "From Warehouse"
        .TextMatrix(0, BteColToWH) = "To Warehouse"
        .TextMatrix(0, bteColRemark) = "Remarks"
        
        .ColWidth(BteColTgl) = 1500
        .ColWidth(bteColItem) = 2050
        .ColWidth(bteColDescription) = 3230
        .ColWidth(bteColLot) = 1500
        .ColWidth(bteColQty) = 1500
        .ColWidth(BteColFromWH) = 2500
        .ColWidth(BteColToWH) = 2500
        
        .ColAlignment(BteColTgl) = flexAlignLeftCenter
        .ColAlignment(bteColItem) = flexAlignLeftCenter
        .ColAlignment(bteColDescription) = flexAlignLeftCenter
        .ColAlignment(bteColLot) = flexAlignLeftCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(BteColFromWH) = flexAlignLeftCenter
        .ColAlignment(BteColToWH) = flexAlignLeftCenter
        
    End With
End Sub

Sub IsiGrid()
'    Dim NewRsGrid As New ADODB.Recordset
'    Dim RowGrid As Integer
'    If NewRsGrid.State <> adStateClosed Then NewRsGrid.Close
'    If CboLocationCD(2).Text = "" Then
'    NewRsGrid.Open " select PS.ChildSupply_Date, PS.ChildItem_Code,PS.parentItem_Code ,PS.remarks, (Select IM.Item_Name from Item_Master IM Where IM.Item_Code = ChildItem_Code)Deskripsi,PS.Lot_No,PS.ChildRequireMent_qty, " & _
'                                " FromWarehouse_Code = rtrim(PS.FromWarehouse_Code) + ' / ' + (Select  WM.Wh_name from Warehouse_master WM where WM.wh_Code =PS.FromWarehouse_Code)  " & _
'                                " ,toWareHouse_Code =rtrim(PS.toWareHouse_Code) + ' / ' + (Select  WM.Wh_name from Warehouse_master WM where WM.wh_Code =PS.toWareHouse_Code)  " & _
'                                " from Part_supply PS wHERE PS.SUPPLY_cLS='TR' and " & _
'                                " PS.ChildSupply_date='" & Format(DMonth, "yyyy/mm/dd") & "' and PS.FromWarehouse_Code = '" & CboLocationCD(0).Text & "' " & _
'                                " ", Db, adOpenDynamic, adLockOptimistic
'
'    Else
'    NewRsGrid.Open " select PS.ChildSupply_Date,PS.ChildItem_Code ,PS.parentItem_Code ,PS.remarks, (Select IM.Item_Name from Item_Master IM Where IM.Item_Code = PS.ChildItem_Code)Deskripsi,PS.Lot_No,PS.ChildRequireMent_qty, " & _
'                                " FromWarehouse_Code = rtrim(PS.FromWarehouse_Code) + ' / ' + (Select  WM.Wh_name from Warehouse_master WM where WM.wh_Code =PS.FromWarehouse_Code) " & _
'                                " ,toWareHouse_Code =rtrim(PS.toWareHouse_Code) + ' / ' + (Select  WM.Wh_name from Warehouse_master WM where WM.wh_Code =PS.toWareHouse_Code)  " & _
'                                " from Part_supply PS wHERE PS.SUPPLY_cLS='TR' and " & _
'                                " PS.ChildSupply_date='" & Format(DMonth, "yyyy/mm/dd") & "' and PS.FromWarehouse_Code = '" & CboLocationCD(0).Text & "' " & _
'                                " And PS.ChildItem_Code='" & CboLocationCD(2).Text & "'", Db, adOpenDynamic, adLockOptimistic
'    End If
'    RowGrid = 0
'    Do While Not NewRsGrid.EOF
'             RowGrid = RowGrid + 1
'             Grid.AddItem ""
'             Grid.TextMatrix(RowGrid, BteColTgl) = Format(NewRsGrid("ChildSupply_date"), "dd/mmm/yyyy")
'             Grid.TextMatrix(RowGrid, bteColItemTo) = NewRsGrid("ChildItem_Code")
'             Grid.TextMatrix(RowGrid, bteColItem) = NewRsGrid("ParentItem_Code")
'             Grid.TextMatrix(RowGrid, bteColDescription) = NewRsGrid("Deskripsi")
'             Grid.TextMatrix(RowGrid, bteColLot) = NewRsGrid("Lot_no")
'             Grid.TextMatrix(RowGrid, bteColQty) = Format(NewRsGrid("ChildRequirement_qty"), gs_formatQty)
'             Grid.TextMatrix(RowGrid, BteColFromWH) = NewRsGrid("FromWarehouse_Code")
'             Grid.TextMatrix(RowGrid, BteColToWH) = NewRsGrid("toWareHouse_Code")
'             Grid.TextMatrix(RowGrid, bteColRemark) = NewRsGrid("Remarks")
'    NewRsGrid.MoveNext
'    Loop
 
    
End Sub


Sub AfterStock()
On Error Resume Next
If Text1 = "" Then
    Text1 = 0
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End If

Text4 = Format(CDbl(Text2) - CDbl(Text1), gs_formatQty)
Text5 = Format(CDbl(Text3) + CDbl(Text1), gs_formatQty)
End Sub

Sub ClearingForm()
Dim X As Byte

For X = 0 To 3
    CboLocationCD(X) = ""
Next X

Text1 = 0
Text2 = 0
Text3 = 0
Text4 = 0
Text5 = 0

Text6 = ""
LblErrMsg = ""
LFrom = ""
LTo = ""

End Sub



Private Sub CboLocationCD_Change(Index As Integer)
                
                

If CboLocationCD(Index) <> "" Then
        If CboLocationCD(Index).MatchFound Then
                If Index = 0 Then
                    CboLocationCD(1) = CboLocationCD(0)
                    CboLocationCD(1).locked = True
                End If
                LblLocationName(Index) = CboLocationCD(Index).List(CboLocationCD(Index).ListIndex, 1)
                If Index = 0 Or Index = 2 Then
                        Text2 = Format(uf_showStock(CboLocationCD(0), CboLocationCD(2)), gs_formatQty)
                ElseIf Index = 1 Or Index = 3 Then
                        Text3 = Format(uf_showStock(CboLocationCD(1), CboLocationCD(3)), gs_formatQty)
                End If
                LblErrMsg = ""
            
            LFrom = CboLocationCD(0) & " / " & CboLocationCD(2)
            LTo = CboLocationCD(1) & " / " & CboLocationCD(3)
                If Index = 2 Or Index = 0 Then
                    Call GridHeader
                    Call IsiGrid
                    
                End If
                
        Else
                LblLocationName(Index) = ""
                If Index > 1 Then
                        LblErrMsg = DisplayMsg(4003) '"Invalid Item code !": Exit Sub
                Else
                        LblErrMsg = DisplayMsg(4018) '"Invalid warehouse code !": Exit Sub
                End If
                
        End If
Else
    
                If Index = 2 Or Index = 0 Then
                    Call GridHeader
                    Call IsiGrid
                    
                End If
    LblLocationName(Index) = ""
    
End If

'If cbolocationcd(2) <> "" And cbolocationcd(3) <> "" Then
'mengecek hubungan parent and child
'Dim sqlCekBom As String
'sqlCekBom = "select Item_code FROM Bom_master WHERE Item_code='" & cbolocationcd(2) & " ' and parent_ItemCode='" & cbolocationcd(3) & "'"
'sqlCekBom = sqlCekBom & vbCrLf & " Union "
'sqlCekBom = sqlCekBom & vbCrLf & " select Item_code FROM Bom_master WHERE Parent_Itemcode='" & cbolocationcd(2) & "' and Item_Code='" & cbolocationcd(3) & "'"
'If Get_Record(sqlCekBom) = "" Then
'    LblErrMsg = DisplayMsg(1014)
'    cbolocationcd(3) = ""
'     LTo = ""
'    Exit Sub
'End If


'End If

        
Call AfterStock
'        If Index = 5 Then Call up_fillComboItemFrom
'        If Index = 6 Then Call up_fillComboItemTo
        
End Sub


Private Sub CboLocationCD_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
        If KeyCode = 13 Then
               
               Dim li_count As Integer
               Dim li_pos As Integer
                
                li_count = 0
                
                For li_pos = 0 To CboLocationCD(Index).ListCount - 1
                        If UCase(Trim(CboLocationCD(Index))) = UCase(Trim(CboLocationCD(Index).List(li_pos, 0))) Then
                                CboLocationCD(Index) = Trim(CboLocationCD(Index).List(li_pos, 0))
                                LblLocationName(Index) = Trim(CboLocationCD(Index).List(li_pos, 1))
                                li_count = 1: Exit For
                        End If
                Next
                
                If li_count = 0 Then
                        If Index > 1 Then
                                LblErrMsg = DisplayMsg(4003) '"Invalid Item code !": Exit Sub
                        Else
                                LblErrMsg = DisplayMsg(4018) '"Invalid warehouse code !": Exit Sub
                        End If
                Else
'                        If Index = 5 Then Call up_fillComboItemFrom
'                        If Index = 6 Then Call up_fillComboItemTo
                        LblErrMsg = ""
                End If
        End If
        
End Sub

Private Sub CboLocationCD_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
        If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


Private Sub Cmd_SubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub Command1_Click(Index As Integer)
 Me.MousePointer = vbHourglass
 
 Select Case Index
  Case 0:
   frm_BrowseItem.getItemCode = CboLocationCD(2).Text
   frm_BrowseItem.Show 1
   CboLocationCD(2).Text = frm_BrowseItem.getItemCode
  Case 1:
   frm_BrowseItem.getItemCode = CboLocationCD(3).Text
   frm_BrowseItem.Show 1
   CboLocationCD(3).Text = frm_BrowseItem.getItemCode
 End Select
 
 Me.MousePointer = vbDefault

End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
        If ErrMsg = "" Then
            Unload Me
        Else
            LblErrMsg.Caption = ErrMsg
        End If
End Sub


Private Sub cmd_Cancel_Click()
Call ClearingForm
End Sub

Private Sub cmd_clear_Click()
Call ClearingForm
End Sub

Private Sub Cmd_Submit_Click()

Dim X As Byte
Dim rsclosing As ADODB.Recordset
Dim lrs_ss As New ADODB.Recordset
Dim ls_sql As String, ls_sql2 As String, CloseThn As Integer, CloseBln As Integer
Dim selisih As Integer, FPilih As String


LblErrMsg = ""
LblErrMsg = up_ValidateDateRange(Format(DMonth, "yyyy-MM-dd"), True)
If LblErrMsg <> "" Then
    Exit Sub
End If

For X = 0 To 3
    If CboLocationCD(X) = "" Then
        LblErrMsg = "Uncomplete Data Item !"
        CboLocationCD(X).SetFocus
        Exit Sub
    Else
        LblErrMsg = ""
    End If
Next X

If CboLocationCD(2) = CboLocationCD(3) Then
    LblErrMsg = "Choose Different Item Code !"
    Exit Sub
End If

If IsNumeric(Text1) = False Then
    LblErrMsg = "Please input valid Qyy !"
    Exit Sub
Else
    LblErrMsg = ""
End If

If (Text2 - Text1) < 0 Then
    LblErrMsg = "Qty must greater than Stock !"
    Exit Sub
Else
    LblErrMsg = ""
End If
    

If CboLocationCD(2) <> "" And CboLocationCD(3) <> "" Then
'mengecek hubungan parent and child
Dim sqlCekBom As String
sqlCekBom = "select Item_code FROM Bom_master WHERE Item_code='" & CboLocationCD(2) & " ' and parent_ItemCode='" & CboLocationCD(3) & "'"
sqlCekBom = sqlCekBom & vbCrLf & " Union "
sqlCekBom = sqlCekBom & vbCrLf & " select Item_code FROM Bom_master WHERE Parent_Itemcode='" & CboLocationCD(2) & "' and Item_Code='" & CboLocationCD(3) & "'"
If Get_Record(sqlCekBom) = "" Then
    If MsgBox("The items you selected are not linked at BOM Master. Do You want to continue?", vbOKCancel + vbQuestion, "Confirmation") = vbCancel Then
        Exit Sub
    End If
End If
End If
If LblErrMsg <> "" Then
    Exit Sub
End If



If LblErrMsg = "" Then
    
    If MsgBox("Are you sure want to process ?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
                    
                            ls_sql = " INSERT INTO [Part_Supply] " & _
                                              " ( " & _
                                              "     [ChildSupply_Date],  " & _
                                              "     [Supply_Cls],  " & _
                                              "     [FromWarehouse_Code],  " & _
                                              "     [From_Address],  " & _
                                              "     [ParentItem_Code],  " & _
                                              "     [ToWareHouse_Code],  " & _
                                              "     [ChildItem_Code], " & _
                                              "     [Do_No], " & _
                                              "     [Remarks], " & _
                                              "     [ChildRequirement_Qty],LOt_No,[last_update],[Last_User]) values("


                            ls_sql = ls_sql + "         '" & Format(DMonth, "yyyy-MM-dd") & "',  " & _
                                              " 'TR',  " & _
                                              " '" & Trim(CboLocationCD(0)) & "',  " & _
                                              " ''," & _
                                              " '" & Trim(CboLocationCD(2)) & "',  " & _
                                              " '" & Trim(CboLocationCD(1)) & "',  " & _
                                              " '" & Trim(CboLocationCD(3)) & "',  " & _
                                              " ''," & _
                                              " '" & IIf(IsNull(Text6), "-", Text6) & "'," & _
                                              " " & CDbl(Text1) & ",'" & txtLot & "','" & Now & "','" & userLogin & "') "
                                
                                Db.Execute ls_sql
                
                               'Db.BeginTrans
                
                
                            '## Begin - Check Stock Data #################################################################
                            ls_sql = " select * from stock_master " & _
                                          " where warehouse_code='" & CboLocationCD(0) & "' " & _
                                          " and item_code='" & CboLocationCD(2) & "' "

                            lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
                            If lrs_ss.EOF = True Then
                                        ls_sql = " INSERT INTO [Stock_Master] " & _
                                                        " ( " & _
                                                        "     [Warehouse_Code],  " & _
                                                        "     [Item_Code],  " & _
                                                        "     [LM_PreMonth], [LM_Receipt], [LM_Supply], [LM_LossReject], [LM_Current], [LM_Inventory],  " & _
                                                        "     [TM_PreMonth], [TM_Receipt], [TM_Supply], [TM_LossReject], [TM_Current], [TM_Inventory], " & _
                                                        "     [NM_PreMonth], [NM_Receipt], [NM_Supply], [NM_LossReject], [NM_Current], [NM_Inventory],  " & _
                                                        "     [LM_Reason], [TM_Reason], [NM_Reason]) " & _
                                                        " VALUES " & _
                                                        " ( " & _
                                                        "     '" & CboLocationCD(0) & "', "

                                        ls_sql = ls_sql + "     '" & CboLocationCD(2) & "',  " & _
                                                          "     0,0,0,0,0,0, " & _
                                                          "     0,0,0,0,0,0, " & _
                                                          "     0,0,0,0,0,0, " & _
                                                          "     '','','' " & _
                                                          " ) "

                                        Db.Execute ls_sql
                            End If

                            ls_sql = " select * from stock_master " & _
                                          " where warehouse_code='" & CboLocationCD(1) & "' " & _
                                          " and item_code='" & CboLocationCD(3) & "' "

                            If lrs_ss.State <> adStateClosed Then lrs_ss.Close
                            lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic

                            If lrs_ss.EOF = True Then
                                        ls_sql = " INSERT INTO [Stock_Master] " & _
                                                        " ( " & _
                                                        "     [Warehouse_Code],  " & _
                                                        "     [Item_Code],  " & _
                                                        "     [LM_PreMonth], [LM_Receipt], [LM_Supply], [LM_LossReject], [LM_Current], [LM_Inventory],  " & _
                                                        "     [TM_PreMonth], [TM_Receipt], [TM_Supply], [TM_LossReject], [TM_Current], [TM_Inventory], " & _
                                                        "     [NM_PreMonth], [NM_Receipt], [NM_Supply], [NM_LossReject], [NM_Current], [NM_Inventory],  " & _
                                                        "     [LM_Reason], [TM_Reason], [NM_Reason]) " & _
                                                        " VALUES " & _
                                                        " ( " & _
                                                        "     '" & CboLocationCD(1) & "', "

                                        ls_sql = ls_sql + "     '" & CboLocationCD(3) & "',  " & _
                                                          "     0,0,0,0,0,0, " & _
                                                          "     0,0,0,0,0,0, " & _
                                                          "     0,0,0,0,0,0, " & _
                                                          "     '','','' " & _
                                                          " ) "

                                        Db.Execute ls_sql
                            End If
                                        

       '### Update Stock #################################
                                        

        If lrs_ss.State <> adStateClosed Then lrs_ss.Close

        Set rsclosing = New ADODB.Recordset
        rsclosing.Open " SELECT MAX(Period) FROM  " & vbCrLf & _
                                "   (SELECT CONVERT(VARCHAR(4),Inventory_Year)+RIGHT(CONVERT(VARCHAR(3),100+inventory_month),2) Period " & vbCrLf & _
                                "       FROM Inventory_Control) IV ", Db

        CloseThn = Left(rsclosing(0), 4)
        CloseBln = Right(rsclosing(0), 2)

        selisih = (Year(DMonth) * 12 + Month(DMonth)) - (CloseThn * 12 + CloseBln)

        If selisih <= 0 Or selisih > 2 Then
            LblErrMsg = "Invalid Date Periode !"
        Else
           If selisih = 1 Then ' This Month
                                        ls_sql = " Update Stock_master set  " & _
                                                    "     tm_current=tm_current- " & CDbl(Text1) & ", " & _
                                                    "     tm_Supply=tm_Supply+ " & CDbl(Text1) & ", " & _
                                                    "     nm_premonth=nm_premonth- " & CDbl(Text1) & ", " & _
                                                    "     nm_current=nm_current- " & CDbl(Text1) & " " & _
                                                    " Where Warehouse_code='" & CboLocationCD(0) & "' " & _
                                                    " and item_code='" & CboLocationCD(2) & "' "

                                        ls_sql2 = " Update Stock_master set  " & _
                                                    "     tm_current=tm_current+" & CDbl(Text1) & ", " & _
                                                    "     tm_receipt=tm_receipt+ " & CDbl(Text1) & ", " & _
                                                    "     nm_premonth=nm_premonth+ " & CDbl(Text1) & ", " & _
                                                    "     nm_current=nm_current+ " & CDbl(Text1) & " " & _
                                                    " Where Warehouse_code='" & CboLocationCD(1) & "' " & _
                                                    " and item_code='" & CboLocationCD(3) & "' "
            Else ' Next Month
                                         ls_sql = " Update Stock_master set  " & _
                                                    "     nm_Supply=nm_Supply+ " & CDbl(Text1) & ", " & _
                                                    "     nm_current=nm_current- " & CDbl(Text1) & " " & _
                                                    " Where Warehouse_code='" & CboLocationCD(0) & "' " & _
                                                    " and item_code='" & CboLocationCD(2) & "' "

                                        ls_sql2 = " Update Stock_master set  " & _
                                                    "     nm_receipt=nm_receipt+ " & CDbl(Text1) & ", " & _
                                                    "     nm_current=nm_current+ " & CDbl(Text1) & " " & _
                                                    " Where Warehouse_code='" & CboLocationCD(1) & "' " & _
                                                    " and item_code='" & CboLocationCD(3) & "' "

           End If
           Db.Execute ls_sql
           Db.Execute ls_sql2
        End If
        
'        Db.CommitTrans
                                       
        Call ClearingForm
    End If
End If

End Sub

Private Sub DMonth_Change()
LblErrMsg = ""

LblErrMsg = up_ValidateDateRange(Format(DMonth, "yyyy-MM-dd"), False)
If LblErrMsg <> "" Then
    Exit Sub
End If

Text2 = Format(uf_showStock(CboLocationCD(0), CboLocationCD(2)), "#,##0.00")
Text3 = Format(uf_showStock(CboLocationCD(1), CboLocationCD(3)), "#,##0.00")
GridHeader
IsiGrid
End Sub

Private Sub Form_Load()
        Call GridHeader
        CtrlMenu1.FormName = Me.Name
        Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
               
        DMonth = Format(Now, "yyyy-MM-dd")
        
        Dim lrs_ss As New ADODB.Recordset
        Dim ls_sql As String
        Dim li_pos As Long
                
        '## Begin - Fill Combo #################################################################
        
        For li_pos = 0 To 1
                CboLocationCD(li_pos).columnCount = 2
                CboLocationCD(li_pos).clear
                CboLocationCD(li_pos).ColumnWidths = "150 pt; 200 pt"
                CboLocationCD(li_pos).ListWidth = 350
                CboLocationCD(li_pos).ListRows = 15
        Next
        

        If lrs_ss.State <> adStateClosed Then lrs_ss.Close
        
        '##Tampilkan Warehouse code dari warehouse_master
        ls_sql = "Select rtrim(wh_code) as Wh_code,wh_name from warehouse_master " & _
        "order by wh_code"
        
        lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
                
        li_pos = 0
        CboLocationCD(0).ColumnWidths = "80 pt; 200 pt"
        CboLocationCD(1).ColumnWidths = "80 pt; 200 pt"
        While lrs_ss.EOF = False
                CboLocationCD(0).AddItem ""
                CboLocationCD(0).List(li_pos, 0) = Trim(lrs_ss("wh_code"))
                CboLocationCD(0).List(li_pos, 1) = Trim(lrs_ss("wh_name"))
                
                CboLocationCD(1).AddItem ""
                CboLocationCD(1).List(li_pos, 0) = Trim(lrs_ss("wh_code"))
                CboLocationCD(1).List(li_pos, 1) = Trim(lrs_ss("wh_name"))
                li_pos = li_pos + 1
                lrs_ss.MoveNext
        Wend
        
        If lrs_ss.State <> adStateClosed Then lrs_ss.Close
        
        For li_pos = 2 To 3
                CboLocationCD(li_pos).columnCount = 2
                CboLocationCD(li_pos).clear
                CboLocationCD(li_pos).ColumnWidths = "130 pt; 150 pt"
                CboLocationCD(li_pos).ListWidth = 280
                CboLocationCD(li_pos).ListRows = 15
        Next
        
        If lrs_ss.State <> adStateClosed Then lrs_ss.Close
        ls_sql = " select * from item_master"       ' where  Use_Endday>=replace( convert(char(10),getdate(),120),'-','') "
        
        
        lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
                
        li_pos = 0
        While lrs_ss.EOF = False
                CboLocationCD(2).AddItem ""
                CboLocationCD(2).List(li_pos, 0) = Trim(lrs_ss("item_code"))
                CboLocationCD(2).List(li_pos, 1) = Trim(lrs_ss("item_name"))
                
                CboLocationCD(3).AddItem ""
                CboLocationCD(3).List(li_pos, 0) = Trim(lrs_ss("item_code"))
                CboLocationCD(3).List(li_pos, 1) = Trim(lrs_ss("item_name"))
                
                li_pos = li_pos + 1
                lrs_ss.MoveNext
        Wend
        If lrs_ss.State <> adStateClosed Then lrs_ss.Close
        
        Call cmd_clear_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        If UnloadMode = 0 Then Cancel = 1
End Sub


Private Sub Text1_Change()
Call AfterStock
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 45 Then KeyAscii = 0
End Sub

Private Function uf_showStock(ls_WhCode As String, ls_ItemCode As String) As Double
        Dim rsclosing As ADODB.Recordset
        Dim lrs_ss As New ADODB.Recordset
        Dim ls_sql As String, CloseThn As Integer, CloseBln As Integer
        Dim selisih As Integer, FPilih As String
        
        If lrs_ss.State <> adStateClosed Then lrs_ss.Close
        
        Set rsclosing = New ADODB.Recordset
        rsclosing.Open "select * from inventory_control " & _
                              " Where Inventory_Month=(Select Max(Inventory_Month) from Inventory_Control " & _
                              " Where inventory_Year=(Select Max(Inventory_Year) from Inventory_Control)) ", Db
    
        CloseThn = rsclosing(0)
        CloseBln = rsclosing(1)
            
        selisih = (Year(DMonth) * 12 + Month(DMonth)) - (CloseThn * 12 + CloseBln)
        
        If selisih < 0 And selisih > 2 Then
            LblErrMsg = "Invalid Date Periode !"
        Else
            If selisih = 0 Then
                FPilih = "LM_Current"
            ElseIf selisih = 1 Then
                FPilih = "TM_Current"
            Else
                FPilih = "NM_Current"
            End If
            
            ls_sql = "Select " & FPilih & " As StockQty From Stock_Master Where WareHouse_Code='" & ls_WhCode & "' And Item_Code='" & ls_ItemCode & "'"
            
            lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
            
            If lrs_ss.EOF = False Then
                uf_showStock = lrs_ss!Stockqty
            Else
                uf_showStock = 0
            End If
        
            If lrs_ss.State <> adStateClosed Then lrs_ss.Close
            
        End If
                
End Function

'Private Function uf_showStock(ls_whCode As String, ls_itemCode As String) As Double
'        Dim lrs_ss As New ADODB.Recordset
'        Dim ls_sql As String
'        If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'
'ls_sql = " select  " & _
'                  "     case datediff(month,tb.closing_date,'" & Format(DMonth, "yyyy-MM-dd") & "' ) " & _
'                  "     when 0 then isnull(lm_inventory,0) " & _
'                  "     when 1 then isnull(tm_current,0) " & _
'                  "     else   isnull(nm_current,0) " & _
'                  "     end StockQty " & _
'                  " from " & _
'                  " ( " & _
'                  "     select *, " & _
'                  "         ( " & _
'                  "             select top 1 cast (inventory_year as char(4)) +'-'+ "
'
'ls_sql = ls_sql + "                      case when inventory_month <10 then '0' else '' end + " & _
'                  "                      rtrim(cast (inventory_month as char(2)) )+'-01' " & _
'                  "             from inventory_control " & _
'                  "             where fix_cls='1' " & _
'                  "             order by inventory_year desc, inventory_month desc " & _
'                  "         ) closing_date " & _
'                  "     from stock_master " & _
'                    " where  warehouse_code='" & Trim(ls_whCode) & "'" & _
'                    " and item_code='" & Trim(ls_itemCode) & "'" & _
'                  " ) tb "
'
'        lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
'        If lrs_ss.EOF = False Then
'                uf_showStock = lrs_ss!Stockqty
'        Else
'                uf_showStock = 0
'        End If
'        If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'End Function


'Private Sub up_fillComboItemFrom()
'        Dim lrs_ss As New ADODB.Recordset
'        For li_pos = 2 To 2
'                cbolocationcd(li_pos).ColumnCount = 2
'                cbolocationcd(li_pos).clear
'                cbolocationcd(li_pos).ColumnWidths = "150 pt; 150 pt"
'                cbolocationcd(li_pos).ListWidth = 300
'                cbolocationcd(li_pos).ListRows = 15
'        Next
'
'        If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'        ls_sql = " select * from item_master where  Use_Endday>=replace( convert(char(10),getdate(),120),'-','') "
'
'        If Trim(cbolocationcd(5)) <> "" Then
'            ls_sql = ls_sql + " and group_cls ='" & Trim(cbolocationcd(5)) & "'"
'        End If
'
'        lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
'
'        li_pos = 0
'        While lrs_ss.EOF = False
'                cbolocationcd(2).AddItem ""
'                cbolocationcd(2).List(li_pos, 0) = Trim(lrs_ss("item_code"))
'                cbolocationcd(2).List(li_pos, 1) = Trim(lrs_ss("item_name"))
'
'                li_pos = li_pos + 1
'                lrs_ss.MoveNext
'        Wend
'        If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'End Sub
'
'Private Sub up_fillComboItemTo()
'        Dim lrs_ss As New ADODB.Recordset
'        For li_pos = 3 To 3
'                cbolocationcd(li_pos).ColumnCount = 2
'                cbolocationcd(li_pos).clear
'                cbolocationcd(li_pos).ColumnWidths = "150 pt; 150 pt"
'                cbolocationcd(li_pos).ListWidth = 300
'                cbolocationcd(li_pos).ListRows = 15
'        Next
'
'        If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'        ls_sql = " select * from item_master where  Use_Endday>=replace( convert(char(10),getdate(),120),'-','') "
'
'        If Trim(cbolocationcd(6)) <> "" Then
'            ls_sql = ls_sql + " and group_cls ='" & Trim(cbolocationcd(6)) & "'"
'        End If
'        lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
'
'        li_pos = 0
'        While lrs_ss.EOF = False
'                cbolocationcd(3).AddItem ""
'                cbolocationcd(3).List(li_pos, 0) = Trim(lrs_ss("item_code"))
'                cbolocationcd(3).List(li_pos, 1) = Trim(lrs_ss("item_name"))
'
'                li_pos = li_pos + 1
'                lrs_ss.MoveNext
'        Wend
'        If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'End Sub


'
'Private Sub Cmd_Save_Click(Index As Integer)
'        If Index = 0 Then
'                    LblErrMsg = validDateRange(Format(DMonth, "yyyy-MM-dd"))
'
'                    If IsNumeric(Text1) = False Then
'                            LblErrMsg = DisplayMsg(4096) 'Please input valid QTy !
'                            Exit Sub
'                    End If
'
'                    If CDbl(Text2) + CDbl(Text1) > 99999999 Or CDbl(Text2) + CDbl(Text1) < -99999999 Or _
'                        CDbl(Text3) + CDbl(Text1) > 99999999 Or CDbl(Text3) + CDbl(Text1) < -99999999 Then
'                            LblErrMsg = DisplayMsg(4119) 'Please input valid QTy !
'                            Exit Sub
'                    End If
'
'                    If hakUpdate(Me.Name) = 0 Then
'                            LblErrMsg = DisplayMsg(3008)
'                            Me.MousePointer = vbDefault
'                            Exit Sub
'                    End If
'
'                    If LblErrMsg = "" Then
'
'                            Dim lrs_ss As New ADODB.Recordset
'                            Dim ls_sql As String
'                            Dim li_diff As Integer
'                            If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'
'
'                            '## Begin - Check Period #################################################################
'                            ls_sql = " select top 1 datediff(month, cast (inventory_year as char(4)) +'-'+ " & _
'                                      "     case when inventory_month <10 then '0' else '' end + " & _
'                                      "     rtrim(cast (inventory_month as char(2)) )+'-01','" & Format(DMonth, "yyyy-MM-dd") & "' )  diff " & _
'                                      " from inventory_control " & _
'                                      " where fix_cls='1' " & _
'                                      " order by inventory_year desc, inventory_month desc "
'
'
'                            lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
'                            li_diff = lrs_ss!diff
'                            If li_diff = 0 Then
'                                    LblErrMsg = DisplayMsg(4117) 'Can Not Process Transfer for Previous Period !
'                                    If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'                                    Exit Sub
'                            End If
'                            If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'                            '## End  - Check Period #################################################################
'
'                            '## Begin - Check Warehouse & Item's Stock Control ##########################################
'
'                            Dim ls_FromWHStockControl As String
'                            Dim ls_ToWHStockControl As String
'                            Dim ls_FromItemStockControl As String
'                            Dim ls_ToItemStockControl As String
'
'                            Dim ls_FromItemUnit As String
'                            Dim ls_ToItemUnit As String
'
'                            ls_sql = " select * from warehouse_master " & _
'                                          " where wh_code='" & cbolocationcd(0) & "' "
'                             If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'                            lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
'                            If lrs_ss.EOF = False Then
'                                    ls_FromWHStockControl = lrs_ss!stockcontrol_cls
'                            End If
'
'                            ls_sql = " select * from warehouse_master " & _
'                                          " where wh_code='" & cbolocationcd(1) & "' "
'                             If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'                            lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
'                            If lrs_ss.EOF = False Then
'                                    ls_ToWHStockControl = lrs_ss!stockcontrol_cls
'                            End If
'
'                            ls_sql = " select * from item_master " & _
'                                          " where item_code='" & cbolocationcd(2) & "' "
'                            If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'                            lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
'                            If lrs_ss.EOF = False Then
'                                    ls_FromItemStockControl = lrs_ss!stockcontrol_cls
'                                    ls_FromItemUnit = lrs_ss!Unit_cls
'                            End If
'
'                            ls_sql = " select * from item_master " & _
'                                          " where item_code='" & cbolocationcd(3) & "' "
'                            If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'                            lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
'                            If lrs_ss.EOF = False Then
'                                    ls_ToItemStockControl = lrs_ss!stockcontrol_cls
'                                    ls_ToItemUnit = lrs_ss!Unit_cls
'                            End If
'
'
'                            If ls_FromItemUnit <> ls_ToItemUnit Then
'                                    LblErrMsg = DisplayMsg(4118) 'Can not transfer into item with different unit !
'                                    If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'                                    Exit Sub
'                            End If
'
'                            If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'                            '## End - Check Warehouse & Item's Stock Control ##########################################
'
'                            '## Begin - Check Stock Data #################################################################
'                            ls_sql = " select * from stock_master " & _
'                                          " where warehouse_code='" & cbolocationcd(0) & "' " & _
'                                          " and item_code='" & cbolocationcd(2) & "' "
'
'                            lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
'                            If lrs_ss.EOF = True Then
'                                        ls_sql = " INSERT INTO [Stock_Master] " & _
'                                                        " ( " & _
'                                                        "     [Warehouse_Code],  " & _
'                                                        "     [Item_Code],  " & _
'                                                        "     [LM_PreMonth], [LM_Receipt], [LM_Supply], [LM_LossReject], [LM_Current], [LM_Inventory],  " & _
'                                                        "     [TM_PreMonth], [TM_Receipt], [TM_Supply], [TM_LossReject], [TM_Current], [TM_Inventory], " & _
'                                                        "     [NM_PreMonth], [NM_Receipt], [NM_Supply], [NM_LossReject], [NM_Current], [NM_Inventory],  " & _
'                                                        "     [Inspection], [Repair], [Inspection_Inventory], [Repair_Inventory]) " & _
'                                                        " VALUES " & _
'                                                        " ( " & _
'                                                        "     '" & cbolocationcd(0) & "', "
'
'                                        ls_sql = ls_sql + "     '" & cbolocationcd(2) & "',  " & _
'                                                          "     0,0,0,0,0,0, " & _
'                                                          "     0,0,0,0,0,0, " & _
'                                                          "     0,0,0,0,0,0, " & _
'                                                          "     0,0,0,0 " & _
'                                                          " ) "
'
'                                        Db.Execute ls_sql
'                            End If
'
'                            ls_sql = " select * from stock_master " & _
'                                          " where warehouse_code='" & cbolocationcd(1) & "' " & _
'                                          " and item_code='" & cbolocationcd(3) & "' "
'
'                            If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'                            lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
'                            If lrs_ss.EOF = True Then
'                                        ls_sql = " INSERT INTO [Stock_Master] " & _
'                                                        " ( " & _
'                                                        "     [Warehouse_Code],  " & _
'                                                        "     [Item_Code],  " & _
'                                                        "     [LM_PreMonth], [LM_Receipt], [LM_Supply], [LM_LossReject], [LM_Current], [LM_Inventory],  " & _
'                                                        "     [TM_PreMonth], [TM_Receipt], [TM_Supply], [TM_LossReject], [TM_Current], [TM_Inventory], " & _
'                                                        "     [NM_PreMonth], [NM_Receipt], [NM_Supply], [NM_LossReject], [NM_Current], [NM_Inventory],  " & _
'                                                        "     [Inspection], [Repair], [Inspection_Inventory], [Repair_Inventory]) " & _
'                                                        " VALUES " & _
'                                                        " ( " & _
'                                                        "     '" & cbolocationcd(1) & "', "
'
'                                        ls_sql = ls_sql + "     '" & cbolocationcd(3) & "',  " & _
'                                                          "     0,0,0,0,0,0, " & _
'                                                          "     0,0,0,0,0,0, " & _
'                                                          "     0,0,0,0,0,0, " & _
'                                                          "     0,0,0,0 " & _
'                                                          " ) "
'
'                                        Db.Execute ls_sql
'                            End If
'
'                            If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'                            '## End - Check Stock Data #################################################################
'
'                          Db.BeginTrans
'                            '## Begin - Insert Result #################################################################
'
'                            ls_sql = " select isnull(max(seq_no),0)+1 newSeq from stock_transfer  " & _
'                              " where transfer_date='" & Format(DMonth, "yyyy-MM-dd") & "' "
'
'                            Dim ls_seqNo As String
'
'                            lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
'                            ls_seqNo = lrs_ss!newSeq
'                            If lrs_ss.State <> adStateClosed Then lrs_ss.Close
'
'                            On Error Resume Next
'
'retry1:
'                            ls_sql = " INSERT INTO [Stock_Transfer] " & _
'                                              " ( " & _
'                                              "     [Transfer_Date],  " & _
'                                              "     [Seq_No],  " & _
'                                              "     [FromWH_Code],  " & _
'                                              "     [FromWH_StockControl],  " & _
'                                              "     [FromItem_Code],  " & _
'                                              "     [FromItem_StockControl],  " & _
'                                              "     [ToWH_Code],  " & _
'                                              "     [ToWH_StockControl],  " & _
'                                              "     [ToItem_Code],  "
'
'                            ls_sql = ls_sql + "     [ToItem_StockControl],  " & _
'                                              "     [Qty],  " & _
'                                              "     [Register_Date] " & _
'                                              " ) values("
'
'
'                            ls_sql = ls_sql + "         '" & Format(DMonth, "yyyy-MM-dd") & "',  " & _
'                                              " '" & ls_seqNo & " ',  " & _
'                                              " '" & Trim(cbolocationcd(0)) & "',  " & _
'                                              " '" & ls_FromWHStockControl & "',  " & _
'                                              " '" & Trim(cbolocationcd(2)) & "',  " & _
'                                              " '" & ls_FromItemStockControl & "',  " & _
'                                              " '" & Trim(cbolocationcd(1)) & "',  " & _
'                                              " '" & ls_ToWHStockControl & "',  " & _
'                                              " '" & Trim(cbolocationcd(3)) & "',  " & _
'                                              " '" & ls_ToItemStockControl & "',  " & _
'                                              " " & CDbl(Text1) & " , getdate() " & _
'                                              "     ) "
'
'                            Db.Execute ls_sql
'
'                            If Trim(Err.Description) <> "" Then
'                                    If InStr(1, Err.Description, "Violation of PRIMARY KEY constraint") Then
'                                        ls_seqNo = ls_seqNo + 1
'                                        Err.Description = ""
'                                        GoTo retry1
'                                    Else
'                                        LblErrMsg = Err.Description
'                                        Db.RollbackTrans
'                                    End If
'                            End If
'
'                            '## End - Insert Result #################################################################
'
'                             '## Begin - Update Stock #################################################################
'
'                            Dim ls_sql2 As String
'                            If li_diff = 1 Then '#This Month
'                                        ls_sql = " Update Stock_master set  " & _
'                                                    "     tm_current=tm_current- " & CDbl(Text1) & ", " & _
'                                                    "     tm_Supply=tm_Supply+ " & CDbl(Text1) & ", " & _
'                                                    "     nm_premonth=nm_premonth- " & CDbl(Text1) & ", " & _
'                                                    "     nm_current=nm_current- " & CDbl(Text1) & " " & _
'                                                    " Where Warehouse_code='" & cbolocationcd(0) & "' " & _
'                                                    " and item_code='" & cbolocationcd(2) & "' "
'
'                                        ls_sql2 = " Update Stock_master set  " & _
'                                                    "     tm_current=tm_current+" & CDbl(Text1) & ", " & _
'                                                    "     tm_receipt=tm_receipt+ " & CDbl(Text1) & ", " & _
'                                                    "     nm_premonth=nm_premonth+ " & CDbl(Text1) & ", " & _
'                                                    "     nm_current=nm_current+ " & CDbl(Text1) & " " & _
'                                                    " Where Warehouse_code='" & cbolocationcd(1) & "' " & _
'                                                    " and item_code='" & cbolocationcd(3) & "' "
'                            Else '#Next Month
'
'                                         ls_sql = " Update Stock_master set  " & _
'                                                    "     nm_Supply=nm_Supply+ " & CDbl(Text1) & ", " & _
'                                                    "     nm_current=nm_current- " & CDbl(Text1) & " " & _
'                                                    " Where Warehouse_code='" & cbolocationcd(0) & "' " & _
'                                                    " and item_code='" & cbolocationcd(2) & "' "
'
'                                        ls_sql2 = " Update Stock_master set  " & _
'                                                    "     nm_receipt=nm_receipt+ " & CDbl(Text1) & ", " & _
'                                                    "     nm_current=nm_current+ " & CDbl(Text1) & " " & _
'                                                    " Where Warehouse_code='" & cbolocationcd(1) & "' " & _
'                                                    " and item_code='" & cbolocationcd(3) & "' "
'
'                            End If
'
'                            If ls_FromWHStockControl = "01" And ls_FromItemStockControl = "01" Then
'                                        Db.Execute ls_sql
'                            End If
'
'                            If ls_ToWHStockControl = "01" And ls_ToItemStockControl = "01" Then
'                                        Db.Execute ls_sql2
'                            End If
'
'                            If Trim(Err.Description) <> "" Then
'                                LblErrMsg = Err.Description
'                                Db.RollbackTrans
'                                Exit Sub
'                            End If
'
'                            '## End - Update Stock #################################################################
'                            Db.CommitTrans
'
'                            LblInfo = "Last Transfered Qty = " & Text1
'                            Text1 = ""
'                            Text2 = Format(uf_showStock(cbolocationcd(0), cbolocationcd(2)), "#,##0.00")
'                            Text3 = Format(uf_showStock(cbolocationcd(1), cbolocationcd(3)), "#,##0.00")
'
'                    End If
'        ElseIf Index = 8 Then
'                    frmMainMenu.Show
'                    Unload Me
'        End If
'End Sub


