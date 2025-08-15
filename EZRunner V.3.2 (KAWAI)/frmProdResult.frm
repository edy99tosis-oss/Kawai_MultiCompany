VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProdResult 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Production Result"
   ClientHeight    =   10950
   ClientLeft      =   135
   ClientTop       =   420
   ClientWidth     =   15120
   Icon            =   "frmProdResult.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdDetail 
      BackColor       =   &H0080FFFF&
      Caption         =   "Detail"
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
      Left            =   11550
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7950
      Width           =   915
   End
   Begin VB.TextBox TxtSerialTo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   10170
      MaxLength       =   7
      TabIndex        =   9
      Tag             =   "1"
      Top             =   7965
      Width           =   1275
   End
   Begin VB.TextBox TxtSerialFrom 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   8820
      MaxLength       =   7
      TabIndex        =   8
      Tag             =   "1"
      Top             =   7965
      Width           =   1275
   End
   Begin VB.CommandButton cmdScanBarcode 
      BackColor       =   &H0080FFFF&
      Caption         =   "Scan Barcode [F2]"
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
      Left            =   6375
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   9795
      Width           =   2130
   End
   Begin VB.TextBox txtLot 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   7470
      MaxLength       =   15
      TabIndex        =   7
      Tag             =   "1"
      Text            =   "AAAAAAAAAAAAAAA"
      Top             =   7965
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   533
      TabIndex        =   22
      Top             =   1215
      Width           =   14175
      Begin VB.ComboBox cboResultCls 
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
         ItemData        =   "frmProdResult.frx":0E42
         Left            =   2160
         List            =   "frmProdResult.frx":0E4F
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1170
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dt 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1590
         Width           =   1515
         _ExtentX        =   2672
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
         CurrentDate     =   37799
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FDDFE3&
         FillStyle       =   0  'Solid
         Height          =   285
         Left            =   4530
         Top             =   1200
         Width           =   1785
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   3600
         X2              =   8670
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Factory CD"
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
         Left            =   330
         TabIndex        =   31
         Top             =   390
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Warehouse CD"
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
         Left            =   330
         TabIndex        =   30
         Top             =   810
         Width           =   1560
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   2
         Left            =   2160
         TabIndex        =   1
         Top             =   750
         Width           =   1335
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2355;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   29
         Top             =   810
         Width           =   960
      End
      Begin VB.Line Line8 
         Index           =   2
         X1              =   3600
         X2              =   6690
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result Cls"
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
         Left            =   300
         TabIndex        =   28
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P1:Result   L:Loss   RJ:Reject"
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
         Left            =   3600
         TabIndex        =   27
         Top             =   1230
         Width           =   2520
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result Date"
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
         TabIndex        =   26
         Top             =   1650
         Width           =   990
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   11010
         X2              =   13785
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   1
         Left            =   11010
         TabIndex        =   25
         Top             =   390
         Width           =   960
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   1
         Left            =   9810
         TabIndex        =   4
         Top             =   330
         Width           =   1005
         VariousPropertyBits=   746604571
         MaxLength       =   3
         DisplayStyle    =   3
         Size            =   "1773;556"
         ListRows        =   15
         ShowDropButtonWhen=   2
         Value           =   "AAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   0
         Top             =   330
         Width           =   1335
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2355;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line CD :"
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
         Left            =   8910
         TabIndex        =   24
         Top             =   390
         Width           =   810
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   23
         Top             =   390
         Width           =   960
      End
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   12525
      MaxLength       =   12
      TabIndex        =   11
      Tag             =   "1"
      Text            =   "9,999,999.99"
      Top             =   7965
      Width           =   1245
   End
   Begin VB.TextBox txtRemarks 
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
      Height          =   285
      Left            =   10313
      MaxLength       =   35
      TabIndex        =   13
      Tag             =   "1"
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      Top             =   8640
      Width           =   4395
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   533
      TabIndex        =   20
      Top             =   9045
      Width           =   14175
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
         Height          =   285
         Left            =   105
         TabIndex        =   21
         Top             =   195
         Width           =   13950
      End
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   12354
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9795
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Left            =   13568
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9795
      Width           =   1140
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
      Left            =   548
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9795
      Width           =   1140
   End
   Begin VB.TextBox txtUnit 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   13823
      MaxLength       =   15
      TabIndex        =   12
      Tag             =   "1"
      Text            =   "Kg"
      Top             =   7965
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Material Consumption"
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
      Left            =   10110
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9795
      Width           =   2130
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Working Time"
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
      Index           =   3
      Left            =   8578
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9795
      Width           =   1500
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   3825
      Left            =   540
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3480
      Width           =   14175
      _cx             =   25003
      _cy             =   6747
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
      AllowUserResizing=   3
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
      Left            =   12870
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   390
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial To"
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
      Left            =   10425
      TabIndex        =   44
      Top             =   7500
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial From"
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
      Left            =   8925
      TabIndex        =   43
      Top             =   7500
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lot No."
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
      Left            =   7815
      TabIndex        =   41
      Top             =   7500
      Width           =   600
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Production Result"
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
      Height          =   330
      Left            =   6615
      TabIndex        =   40
      Top             =   405
      Width           =   2025
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   6255
      X2              =   9195
      Y1              =   8970
      Y2              =   8970
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
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
      Left            =   705
      TabIndex        =   39
      Top             =   7500
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
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
      Left            =   5370
      TabIndex        =   38
      Top             =   8730
      Width           =   690
   End
   Begin VB.Label lblNm 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   3
      Left            =   6255
      TabIndex        =   37
      Top             =   8730
      Width           =   2955
   End
   Begin MSForms.ComboBox cbo 
      Height          =   315
      Index           =   3
      Left            =   705
      TabIndex        =   6
      Top             =   7950
      Width           =   2025
      VariousPropertyBits=   746604569
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "3572;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "AAAAAAAAAAAAAAA"
      FontName        =   "Verdana"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   0
      Left            =   540
      Top             =   7440
      Width           =   14175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
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
      Left            =   13005
      TabIndex        =   36
      Top             =   7500
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   9405
      TabIndex        =   35
      Top             =   8685
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   5
      Left            =   2835
      TabIndex        =   34
      Top             =   7500
      Width           =   960
   End
   Begin VB.Label lblNm 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   4
      Left            =   2835
      TabIndex        =   33
      Top             =   8010
      Width           =   960
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2835
      X2              =   7395
      Y1              =   8250
      Y2              =   8250
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
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
      Left            =   13830
      TabIndex        =   32
      Top             =   7500
      Width           =   330
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   330
      Index           =   1
      Left            =   540
      Top             =   7440
      Width           =   14175
   End
End
Attribute VB_Name = "frmProdResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public is_LoadByItemCode As String

Dim sql As String
Dim HakU As Integer
Dim i As Long, nilKosong As Boolean, blnCancel As Boolean
Dim simpan As Boolean, ubah As Boolean, hapus As Boolean 'Status Ubah/Hapus
Dim newCls As New clsMRP
Dim blnFix As Integer, thnFix As Integer
Dim dbTransfer As New ADODB.Connection

Public QtyRemaining As Double
Dim tampungQty As Double
Dim KeyProd As Double
Public qtydiPake As Double
Public tglProd As String
Public Bagcode As String
Public seqNoBag As Double
Dim clsMRP As New clsMRP
Dim ClsProc As New ClsProc
Dim seqNoConsump  As Double, seqNoProdReceipt As Double
Public nilPrice As String, Curr As String, Price As Double
Public dailyseqno As Double

Public qtyDaily As Double, qtyAllResult As Double
Public factoryDaily As String, completeCls As Double
Public PONO As String, poSEqNo As Double
Public UnitCls As String

Dim strChildItemCD As String

Dim bteColSelect As Byte
Dim bteColProdCode As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColLotNo As Byte
Dim bteColQty As Byte
Dim bteColRemarks As Byte
Dim bteColStockItem As Byte
Dim bteColStockWH As Byte
Dim bteColSeqNo As Byte
Dim bteColDailySeqNo As Byte
Dim bteColProdDate As Byte
Dim bteColQtyAll As Byte
Dim bteColQtyDaily As Byte
Dim BteColSerialFrom As Byte
Dim BteColSerialTo As Byte
Dim bteColDailyComplete As Byte
Dim bteColConsumpSeqNo As Byte
Dim bteColSupply As Byte
Dim bteColWorkTime As Byte

Dim strCol3 As String
Dim strCol4 As String
Dim strCol5 As String

Private Sub headerGrid()
    
    bteColSelect = 0
    bteColProdCode = 1
    bteColPartNo = 2
    bteColDesc = 3
    bteColLotNo = 4
    BteColSerialFrom = 5
    BteColSerialTo = 6
    bteColQty = 5 + 2
    bteColRemarks = 6 + 2
    bteColStockItem = 7 + 2
    bteColStockWH = 8 + 2
    bteColSeqNo = 9 + 2
    bteColDailySeqNo = 10 + 2
    bteColProdDate = 11 + 2
    bteColQtyAll = 12 + 2
    bteColQtyDaily = 13 + 2
    bteColDailyComplete = 14 + 2
    bteColConsumpSeqNo = 15 + 2
    bteColSupply = 16 + 2
    bteColWorkTime = 17 + 2
    
    With grid
        .clear
        .ColS = 18 + 2
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = "S"
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColPartNo) = "Part No."
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColLotNo) = "Lot No"
        .TextMatrix(0, bteColQty) = "Qty"
        
        .TextMatrix(0, BteColSerialFrom) = "Serial From"        'Add 20090210
        .TextMatrix(0, BteColSerialTo) = "Serial To"              'Add 20090210
        
        .TextMatrix(0, bteColRemarks) = "Remarks"
        .TextMatrix(0, bteColStockItem) = "StockItem"
        .TextMatrix(0, bteColStockWH) = "StockWH"
        .TextMatrix(0, bteColSeqNo) = "SeqNo"
        .TextMatrix(0, bteColDailySeqNo) = "DailySeqNo"
        .TextMatrix(0, bteColProdDate) = "ProdDate"
        .TextMatrix(0, bteColQtyAll) = "QtyAll"
        .TextMatrix(0, bteColQtyDaily) = "QtyDaily"
        .TextMatrix(0, bteColDailyComplete) = "DailyComplete"
        .TextMatrix(0, bteColConsumpSeqNo) = "ConsumpSeqNo"
        .TextMatrix(0, bteColSupply) = "Supply"
        .TextMatrix(0, bteColWorkTime) = "WorkTime"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColProdCode) = 1500
        .ColWidth(bteColPartNo) = 1600
        .ColWidth(bteColDesc) = 2500
        .ColWidth(bteColLotNo) = 1400
        
        .ColWidth(BteColSerialFrom) = 1100
        .ColWidth(BteColSerialTo) = 1100
        
        .ColWidth(bteColQty) = 1500
        .ColWidth(bteColRemarks) = 4500
        
        .ColHidden(bteColStockItem) = True
        .ColHidden(bteColStockWH) = True
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColDailySeqNo) = True
        .ColHidden(bteColProdDate) = True
        .ColHidden(bteColQtyAll) = True
        .ColHidden(bteColQtyDaily) = True
        .ColHidden(bteColDailyComplete) = True
        .ColHidden(bteColConsumpSeqNo) = True
        .ColHidden(bteColSupply) = True
        .ColHidden(bteColWorkTime) = True
        
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColPartNo) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColLotNo) = flexAlignLeftCenter
        
        .ColAlignment(BteColSerialFrom) = flexAlignCenterCenter
        .ColAlignment(BteColSerialTo) = flexAlignCenterCenter
        
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColRemarks) = flexAlignLeftCenter
        
        .EditMaxLength = 1
    End With
End Sub


Sub kosongAtas()
    nilKosong = True
    For i = 0 To 2
        cbo(i) = ""
        lblNm(i) = ""
    Next i
    'cboResultCls.ListIndex = -1
    cboResultCls.ListIndex = 0 'Update for Kawai just Result (P1) 20090210
    nilKosong = False
    Call headerGrid
End Sub

Sub kosongBwh()
'    cbo(3) = ""
'    lblNm(3) = ""
'    lblNm(4) = ""
 '   txtLot = ""
    txtQty = Format(0, gs_formatQty)
    tampungQty = 0
    KeyProd = 0
    txtremarks = ""
    
    TxtSerialFrom = ""          'Add 20090210
    TxtSerialTo = ""             'Add 20090210
    
    Command1(3).Enabled = False
    Call nyalaBwh
    simpan = True: ubah = False: hapus = False
End Sub

'******** Combo Factory Code **********
Sub isiCboCust() 'Isi Combo Dealer CD dr Customer Master
Dim RsCust As New ADODB.Recordset 'Data Customer

With cbo(0)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    '******** Ambil dr Customer Master utk Combo Dealer CD
    sql = "select Trade_Code,Trade_Name from Trade_Master " & _
        "where trade_code in (select distinct manufacture_code from manufacture_line) " & _
        "order by Trade_Code"
    Set RsCust = Db.Execute(sql)
    
    i = 0
    Do While Not (RsCust.EOF)
        .AddItem ""
        .List(i, 0) = Trim(RsCust(0))
        .List(i, 1) = Trim(RsCust(1))
        i = i + 1
        RsCust.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 350
    .ColumnWidths = "50 pt;300 pt"
    
    Set RsCust = Nothing
End With
End Sub

'******** Filter Combo Line Code **********
Sub isiCboLine(factoryCD As String)
Dim rscbo As New ADODB.Recordset

With cbo(1)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    sql = "select Line_Code,Line_Name from Manufacture_line " & _
        "where Manufacture_Code = '" & factoryCD & _
        "' order by Line_Code"
    Set rscbo = Db.Execute(sql)
     
    i = 0
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo(0))
        .List(i, 1) = Trim(rscbo(1))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 200
    .ColumnWidths = "50 pt;150 pt"
    
    Set rscbo = Nothing
End With
End Sub

'******** Filter Combo WH **********
Sub isiCboWH()
Dim rscbo As New ADODB.Recordset

With cbo(2)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    sql = "select * from (select wh_code,wh_name,stockControl_cls from warehouse_master " & _
        "union all " & _
        "select distinct(manufacture_line.manufacture_code)wh_code, trade_name wh_name, stockControl_Cls = '01'from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code)tbJ order by wh_code"
    Set rscbo = Db.Execute(sql)
    
    i = 0
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo(0))
        .List(i, 1) = Trim(rscbo(1))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 250
    .ColumnWidths = "50 pt;200 pt"
    
    Set rscbo = Nothing
End With
End Sub

'******** Filter Combo WH **********
Sub isiCboItem()
Dim rscbo As New ADODB.Recordset

With cbo(3)
    .clear
    .columnCount = 6
    .TextColumn = 1
    
    
    
    sql = "select a.Item_Code,MakerItem_Code,Item_Name, " & _
        "a.StockControl_cls as StockItem, b.StockControl_cls as StockWH, Suply_Cls = ISNULL(a.Suply_Cls,'')  " & _
        "from Item_master a, Warehouse_Master b " & _
        "where a.WH_Code = b.WH_Code and a.use_endday > convert(char(8), getdate(), 112) "

    If Trim(is_LoadByItemCode) <> "" Then
        sql = sql + " and a.Item_code ='" & Trim(is_LoadByItemCode) & "'"
        is_LoadByItemCode = ""
    End If
    
    sql = sql + " order by Item_Code "

    Set rscbo = Db.Execute(sql)
    
    
    i = 0
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo(0))
        .List(i, 1) = Trim(rscbo(1))
        .List(i, 2) = Trim(rscbo(2))
        .List(i, 3) = Trim(rscbo(3))
        .List(i, 4) = Trim(rscbo(4))
        .List(i, 5) = Trim(rscbo(5))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 500
    .ColumnWidths = "100 pt;100 pt;300 pt;0pt;0pt"
    
    Set rscbo = Nothing
End With
End Sub

Private Sub cmdDetail_Click()
If cbo(3) <> "" And txtLot <> "" Then
DoEvents
    FrmProdResultDetail.lblname(0) = lblNm(0)
    FrmProdResultDetail.lblname(1) = lblNm(2)
    FrmProdResultDetail.lblname(2) = Format(dt, "dd-mmm-yyyy")
    FrmProdResultDetail.lblname(3) = lblNm(1)
    FrmProdResultDetail.TxtSeqNo = KeyProd
    FrmProdResultDetail.Show
    Me.Hide
DoEvents
End If
End Sub

Private Sub cmdScanBarcode_Click()
    LblErrMsg.Caption = ""
    frmProdScanBarcode.Show vbModal
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty.Text)
    txtQty.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then cmdScanBarcode_Click: KeyCode = 0
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    Command1(0).Enabled = False
    dt = Format(Now, "dd MMM yyyy")
    HakU = hakUpdate(Me.Name)
    txtQty = Format(0, gs_formatQty)
    nilKosong = True
    Call kosongAtas
    Call kosongBwh
    Call isiCboCust
    Call isiCboWH
    Call isiCboItem
    nilKosong = False
End Sub

'*********** Tampilkan Data *********
Private Sub cbo_Click(Index As Integer)
If nilKosong = True Then Exit Sub

If cbo(Index) <> "" Then
    cbo(Index) = cbo(Index)
    If cbo(Index).MatchFound = True Then
        lblNm(Index) = cbo(Index).Column(1)
        If Index = 0 Then 'panggil Manufacture Line
            Call isiCboLine(cbo(0)): lblNm(1) = ""
        ElseIf Index = 3 Then
            lblNm(4) = cbo(3).Column(2) 'Panggil Desc
        End If
        
        If Index <> 3 Then Call tampilData 'Jika bukan Item maka panggil Grid
        LblErrMsg = ""
    Else
        lblNm(Index) = ""
        If Index = 0 Then 'Hapus Manufacture Line & Desc Line
            cbo(1).clear: lblNm(1) = ""
        ElseIf Index = 3 Then 'Beda Err Msg en hapus Desc
            lblNm(4) = ""
            LblErrMsg = DisplayMsg(4003)
        End If
        If Index <> 3 Then Call tampilData: LblErrMsg = DisplayMsg(4016 + Index) 'Err Msg en Panggil Grid
    End If
Else
    lblNm(Index) = ""
    If Index = 0 Then 'Hapus Manufacture Line * Desc
        cbo(1).clear: lblNm(1) = ""
    ElseIf Index = 3 Then 'Hapus Desc
        lblNm(4) = ""
    End If
    
    If Index <> 3 Then Call headerGrid 'Hapus Grid
    LblErrMsg = ""
End If
End Sub

Private Sub cbo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
If nilKosong = True Then Exit Sub
    If KeyCode = 13 Then Call cbo_Click(Index)
End Sub

Private Sub cbo_Change(Index As Integer)
If nilKosong = True Then Exit Sub
    lblNm(Index) = ""
    If Index = 0 Then 'Hapus Manufacture Line * Desc
        cbo(1).clear: lblNm(1) = ""
    Else 'Hapus Desc
        If Index = 3 Then
            lblNm(4) = ""
            cbo(Index) = Trim(cbo(Index))
            If cbo(Index).MatchFound Then
                If cbo(Index).Column(5) = "01" Then
                    Command1(2).Enabled = False
                    Command1(0).Enabled = True
                End If
            End If
        End If
    
    End If
End Sub

Private Sub cbo_LostFocus(Index As Integer)
If nilKosong = True Then Exit Sub
    If lblNm(Index) = "" Then Call cbo_Click(Index)
End Sub

Private Sub cboResultCls_Click()
    Call tampilData
End Sub

Private Sub dt_change()
    Call kosongBwh
    
    Call tampilData
End Sub

Public Sub tampilData()
nilKosong = True
    cbo(0) = cbo(0): cbo(1) = cbo(1): cbo(2) = cbo(2)
'    Call kosongBwh
    If cbo(0).MatchFound And cbo(1).MatchFound And cbo(1).MatchFound And cboResultCls <> "" Then
        Call IsiGrid
    Else
        Call headerGrid
    End If
nilKosong = False
End Sub

'************ Isi Grid **********
Sub IsiGrid()
Dim rsTglAwal As String, rsTglAkhir As String
Dim rsProd As New ADODB.Recordset
Dim noS As Double
Dim adoCon As New ADODB.Connection
Dim adoCmd As New Command
'Dim rs As New ADODB.Recordset
Dim spwhat As String
Dim cek As String
Dim rsCek As New ADODB.Recordset

    adoCon.ConnectionString = Db.ConnectionString
    adoCon.CursorLocation = adUseClient
    adoCon.CommandTimeout = 12000
    adoCon.Open
    adoCon.BeginTrans

With grid
    
    Call headerGrid
    
    sql = "Select pr.Seq_No, pr.Item_Code, im.MakerItem_Code, im.Item_Name, isnull(SuratJalan_No,'') as Lot_No, pr.Qty,pr.SerialNoFrom,pr.SerialNoTo, " & vbCrLf & _
        "Remarks, im.StockControl_Cls as stockItem, wm.StockControl_Cls as stockWH, Schedule_Date, dp.seq_no dailyseqno, " & vbCrLf & _
        "isnull((Select sum(Qty) from Part_Receipt Where DailySeq_No = pr.DailySeq_No),0) QtyAllResult, dp.Qty QtyDaily, " & vbCrLf & _
        "isnull((select  top 1 seq_no from part_supply where do_no = convert(char, pr.seq_no)),'0') ConsumpSeqno, " & vbCrLf & _
        "isnull((select complete_cls from daily_production where seq_no = pr.dailyseq_no),0) completecls, isnull(im.suply_cls,'02') suply_cls, " & vbCrLf & _
        "WkCls = isnull((select top 1 ProductionSeq_No from WorkingTime_Master " & vbCrLf & _
        "Where ProductionSeq_No in (Select ProductionSeq_No from WorkingTime_Detail Where ProductionSeq_No = pr.Seq_No)), 0) " & vbCrLf & _
        "from Part_Receipt pr WITH (NOLOCK) " & vbCrLf & _
        "inner join Daily_production dp on pr.Dailyseq_No = dp.seq_No " & vbCrLf & _
        "inner join Item_master im on pr.Item_Code = im.Item_Code " & vbCrLf & _
        "inner join ( " & vbCrLf & _
            "select wh_code, stockControl_cls from warehouse_master " & vbCrLf & _
            "Union " & vbCrLf & _
            "select distinct(manufacture_line.manufacture_code) wh_code, stockControl_Cls = '01' from manufacture_line " & vbCrLf & _
            "inner join trade_master on manufacture_line.manufacture_code = trade_master.trade_code) wm on pr.Warehouse_Code = wm.wh_code " & vbCrLf & _
        "where pr.Supplier_Code = '" & cbo(0) & "' " & vbCrLf & _
        "and pr.PO_NO = '" & cbo(1) & "' " & vbCrLf & _
        "and pr.Warehouse_Code = '" & cbo(2) & "' " & vbCrLf & _
        "and pr.Receipt_cls = '" & cboResultCls & "' " & vbCrLf & _
        "and pr.Receipt_Date = '" & Format(dt, "yyyy-MM-dd") & "' " & vbCrLf & _
        "and pr.ProductionResult_Cls = 1 " & vbCrLf & _
        "order by pr.Item_Code"
        
    Set rsProd = Db.Execute(sql)
    
    i = 1
    Do While Not rsProd.EOF
        .Rows = .Rows + 1
        .TextMatrix(i, bteColSelect) = ""
        .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
        .TextMatrix(i, bteColProdCode) = Trim(rsProd("Item_Code"))
        .TextMatrix(i, bteColPartNo) = Trim(rsProd("MakerItem_Code"))
        .TextMatrix(i, bteColDesc) = Trim(rsProd("Item_Name"))
        .TextMatrix(i, bteColLotNo) = Trim(rsProd("Lot_No"))
        'Add 20090210
        .TextMatrix(i, BteColSerialFrom) = IIf(IsNull(rsProd("SerialNoFrom")), "", rsProd("SerialNoFrom"))
        .TextMatrix(i, BteColSerialTo) = IIf(IsNull(rsProd("SerialNoTo")), "", rsProd("SerialNoTo"))
        '---
        .TextMatrix(i, bteColQty) = Format(Trim(rsProd("Qty")), gs_formatQty)
        .TextMatrix(i, bteColRemarks) = IIf(IsNull(rsProd("Remarks")), "", rsProd("Remarks"))
        .TextMatrix(i, bteColStockItem) = Trim(rsProd("StockItem"))
        .TextMatrix(i, bteColStockWH) = Trim(rsProd("StockWH"))
        .TextMatrix(i, bteColSeqNo) = Trim(rsProd("Seq_No"))
        .TextMatrix(i, bteColDailySeqNo) = Trim(rsProd("Schedule_Date"))
        .TextMatrix(i, bteColProdDate) = Trim(rsProd("dailyseqno"))
        .TextMatrix(i, bteColQtyAll) = Trim(rsProd("QtyAllresult"))
        .TextMatrix(i, bteColQtyDaily) = Trim(rsProd("QtyDaily"))
        .TextMatrix(i, bteColDailyComplete) = Trim(rsProd("Completecls"))
        .TextMatrix(i, bteColConsumpSeqNo) = Trim(rsProd("Consumpseqno"))
        .TextMatrix(i, bteColSupply) = Trim(rsProd("suply_cls"))
        .TextMatrix(i, bteColWorkTime) = Trim(rsProd("WkCls"))
        i = i + 1
        rsProd.MoveNext
    Loop
    
    Set rsProd = Nothing
End With

ErrExit:
    Set rsProd = Nothing
    Set adoCmd = Nothing
    Set adoCon = Nothing
    'SetControl True
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    adoCon.RollbackTrans
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
    
End Sub

'******************
Sub nyalaBwh()
'    cbo(3).Enabled = True
'    txtLot.Enabled = True
    txtQty.Enabled = True
    txtremarks.Enabled = True
    TxtSerialFrom.Enabled = True
    TxtSerialTo.Enabled = True
    CmdDetail.Enabled = True
End Sub

Sub MatiinBwh()
    txtQty.Enabled = False
    txtremarks.Enabled = False
    TxtSerialFrom.Enabled = False
    TxtSerialTo.Enabled = False
    CmdDetail.Enabled = False
End Sub

Private Sub kosongColGrid(Optional kolumn As String)
Dim i As Integer
Dim jmlD As Double

With grid
    .Col = bteColSelect
    
    jmlD = 0
    If kolumn = "" Then 'hapus semua
        For i = 1 To .Rows - 1
          If .TextMatrix(i, bteColSelect) <> "" Then .TextMatrix(i, bteColSelect) = ""
        Next i
    Else 'Hapus yg ada S
        For i = 1 To .Rows - 1
            If .TextMatrix(i, bteColSelect) = kolumn Or .TextMatrix(i, bteColSelect) <> "D" Then .TextMatrix(i, bteColSelect) = ""
            If .TextMatrix(i, bteColSelect) = "D" Then
                jmlD = jmlD + 1
            Else 'Hapus yg selain D
                .TextMatrix(i, bteColSelect) = ""
            End If
        Next i
        If jmlD <> 0 Then hapus = True: simpan = False: ubah = False: Call MatiinBwh
    End If
End With
End Sub

Public Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim TextGrid As String

LblErrMsg = ""
nilKosong = True
With grid
    TextGrid = .TextMatrix(Row, Col)
    
    If TextGrid = "S" Then
        Call kosongColGrid
        Call nyalaBwh
        
        cbo(3) = Trim(.TextMatrix(Row, bteColProdCode))
'        cbo(3).Enabled = False
        lblNm(3) = Trim(.TextMatrix(Row, bteColPartNo))
        lblNm(4) = Trim(.TextMatrix(Row, bteColDesc))
        txtLot = Trim(.TextMatrix(Row, bteColLotNo))
        txtQty = Trim(.TextMatrix(Row, bteColQty))
        tampungQty = IIf(cboResultCls = "P1", CDbl(.TextMatrix(Row, bteColQty)), -(CDbl(.TextMatrix(Row, bteColQty))))
        txtremarks = Trim(.TextMatrix(Row, bteColRemarks))
        
        ' Add 20090210
        TxtSerialFrom = Trim(.TextMatrix(Row, BteColSerialFrom))
        TxtSerialTo = Trim(.TextMatrix(Row, BteColSerialTo))
        ' ----
        
        KeyProd = CDbl(.TextMatrix(Row, bteColSeqNo))
        tglProd = Format(.TextMatrix(Row, bteColDailySeqNo), "yyyy-MM-dd")
        qtydiPake = 0 'CDbl(.TextMatrix(Row, bteColProdCode)) 'Tot Dipake
        dailyseqno = (.TextMatrix(Row, bteColProdDate))
        qtyAllResult = (.TextMatrix(Row, bteColQtyAll))
        qtyDaily = (.TextMatrix(Row, bteColQtyDaily))
        completeCls = (.TextMatrix(Row, bteColDailyComplete))
        Call cek_cboColumn(cbo(3))
        ' If cbo(3).Column(5) = "01" Then 'Auto
        If strCol5 = "01" Then 'Auto
            Command1(0).Enabled = True 'Submit
            Command1(2).Enabled = False
        Else
            Command1(0).Enabled = False
            Command1(2).Enabled = True 'Consump
        End If
        Command1(3).Enabled = True
        ubah = True
        simpan = False
        hapus = False
    Else
        Call kosongBwh
        Call kosongColGrid("S")
        Command1(0).Enabled = True
        Command1(3).Enabled = False
    End If
    
    .TextMatrix(Row, Col) = TextGrid
End With
nilKosong = False
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Col <> bteColSelect Then Cancel = True: Exit Sub
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
With grid
    If .Col = bteColSelect Then
        If KeyAscii = Asc(".") Then KeyAscii = 0: Exit Sub
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And _
            KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And _
            KeyAscii <> vbKeyReturn Then _
            KeyAscii = 0: Exit Sub
    End If
End With
End Sub

Private Sub Command1_Click(Index As Integer)
Dim rsCek As New ADODB.Recordset
Dim tampungBln As String
Dim pesanTgl As String
Dim tanya

Dim X As Long
Dim awal As Long, akhir As Long, Panjang As Integer
Dim TempSerial As String, Depan As String
Dim rsCeks As New ADODB.Recordset, sql1 As String

LblErrMsg.Caption = ""
Me.MousePointer = vbHourglass

If Index <> 1 Then

    ' Validation of Serial No
    ' Update 20090210
        If Trim(TxtSerialFrom) <> "" And Trim(TxtSerialTo) <> "" Then
            Panjang = Len(Trim(TxtSerialFrom))
            Depan = Left(TxtSerialFrom, 1)
                                
            awal = Val(Mid(TxtSerialFrom, 2, Panjang - 1))
            akhir = Val(Mid(Trim(TxtSerialTo), 2, Panjang - 1))
            
            ' Check For Valid Serial No
            
            If awal > akhir Then
                LblErrMsg = "[000] - Invalid Serial Number ! "
                Me.MousePointer = vbDefault
                Exit Sub
            End If
            
            If Depan <> Left(TxtSerialTo, 1) Then
                LblErrMsg = "[000] - Invalid Serial Number ! "
                Me.MousePointer = vbDefault
                Exit Sub
            End If
            
            If (akhir - awal) + 1 <> CDbl(txtQty) Then
                If MsgBox("Data doesn't match between Qty and Serial No !" & Chr(13) & _
                    "Continue this Process ? ", vbYesNo + vbDefaultButton2, "Attention") = vbNo Then
                    LblErrMsg = "[000] - Data doesn't match between Qty and Serial No ! "
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            
            'Cek Serial Number List on Table - 20090210
            ' Cek for not Exist
                For X = awal To akhir
                    TempSerial = Depan & Format(X, String(Panjang - 1, "0"))
                    sql1 = "Select * from Serial_Detail where item_code='" & cbo(3).Text & "' and " & _
                    " Serial_NO='" & TempSerial & "' And Product_No ='" & dailyseqno & "'"
                    Set rsCek = Db.Execute(sql1)
                    If rsCek.EOF Then
                        LblErrMsg = "[000] - Serial Number " & TempSerial & " Doesn't exist at Production Schedule !"
                        Me.MousePointer = vbDefault
                        Exit Sub
                    End If
                Next X
            
            ' Cek for Already has Result - 20090211
                For X = awal To akhir
                    TempSerial = Depan & Format(X, String(Panjang - 1, "0"))
                    
                    sql = " select * From ( " & vbLf & _
                        " Select * From Serial_Detail Where Product_No is not null ) n " & vbLf & _
                        " Where Item_Code='" & cbo(3).Text & "' and Serial_No='" & TempSerial & "' and result_No <>'" & KeyProd & "'"
                    

                    Set rsCek = Db.Execute(sql)
                    If Not rsCek.EOF Then
                        LblErrMsg = "[000] - Serial Number " & TempSerial & " already has Result at Result No : " _
                                        & Trim(rsCek("Result_No"))
                        
                        Me.MousePointer = vbDefault
                        Exit Sub
                    End If
                Next X
                
                
                
            ' -------------------------
        End If
        ' ----------------------------
End If

Select Case Index
Case 0: 'Submit
        If HakU = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
        
        pesanTgl = up_ValidateDateRange(Format(dt, "yyyy-MM-dd"), True)
        
        If pesanTgl <> "" Then LblErrMsg = pesanTgl: Me.MousePointer = vbDefault: Exit Sub
        
        pesanTgl = HolidayCheck(cbo(0), Format(dt, "yyyy-MM-dd"))
        If pesanTgl <> "" Then LblErrMsg = pesanTgl: Me.MousePointer = vbDefault: Exit Sub
        
        'Cek semua Combo, Jika hapus tidak perlu cek Cbo Item
        For i = 0 To IIf(hapus = True, 2, 3)
            If cbo(i) = "" Then
                If i = 3 Then LblErrMsg = DisplayMsg(1009) Else LblErrMsg = DisplayMsg(1040 + i)
                If cbo(i).Enabled Then cbo(i).SetFocus
                Me.MousePointer = vbDefault: Exit Sub
            Else '**** Jika data tdk kosong cek Data cocok / tdk dgn Combo
                cbo(i) = cbo(i)
                If cbo(i).MatchFound = False Then
                    If cbo(i).Enabled Then cbo(i).SetFocus
                    If i = 3 Then LblErrMsg = DisplayMsg(4003) Else LblErrMsg = DisplayMsg(4016 + i)
                    Me.MousePointer = vbDefault: Exit Sub
                End If
            End If
        Next i
        
        If Not hapus Then
            If CDbl(txtQty) = 0 Then
                LblErrMsg = DisplayMsg(1012)
                txtQty.SetFocus
                Me.MousePointer = vbDefault: Exit Sub
            ElseIf CDbl(txtQty) < qtydiPake Then 'Harus Lebih Besar dr Supp yg memakai receipt tersebut
                LblErrMsg = DisplayMsg(4043) & " " & Format(qtydiPake, gs_formatQty) & " (Already Being Supplied)"
                txtQty.SetFocus
                Me.MousePointer = vbDefault: Exit Sub
            ElseIf CDbl(txtQty.Text) > gd_MaxQty Then
                LblErrMsg = DisplayMsg(4045) & " " & gd_MaxQty & " !"
                txtQty.SetFocus
                Me.MousePointer = vbDefault: Exit Sub
            End If
            If cboResultCls = "" Then
                cboResultCls.SetFocus
                LblErrMsg = DisplayMsg(1043)
                Me.MousePointer = vbDefault: Exit Sub
            End If
        End If
        
        tampungBln = newCls.blnAkhir()
        blnFix = Split(tampungBln, ",")(0)
        thnFix = Split(tampungBln, ",")(1)
        
        'Jika belum ada Data Stock Inventory Closing
        If blnFix = 0 Then LblErrMsg = DisplayMsg(4019): Me.MousePointer = vbDefault: Exit Sub
        
        blnCancel = False
        If hapus = True Then
            Call hapusData
                          
        Else
'            If txtLot = "" Then
'                txtLot.SetFocus
'                LblErrMsg = DisplayMsg(1044)
'                Me.MousePointer = vbDefault: Exit Sub
'            Else
                tanya = MsgBox("Do you really want to Process Production Result?", vbQuestion & vbYesNo, "Confirmation")
                If tanya = vbYes Then
                    If completeCls = 1 Then LblErrMsg = DisplayMsg(1110): Me.MousePointer = vbDefault: Exit Sub 'Daily Completed
                    
                    If Command1(2).Enabled = True Then
                        sql = "Select Seq_No from Part_Supply where DO_No = '" & KeyProd & "'"
                        Set rsCek = Db.Execute(sql)
                        If rsCek.EOF Then
                            LblErrMsg = DisplayMsg(4078) 'Input Data Material
                            Me.MousePointer = vbDefault: Exit Sub
                        End If
                        Set rsCek = Nothing
                    End If
                   
                    dbTransfer.ConnectionTimeout = 0
                    dbTransfer.CommandTimeout = 0
                    dbTransfer.Open Db.ConnectionString
                    dbTransfer.BeginTrans

                    Call simpanUbah(dbTransfer)
                                        
                    dbTransfer.CommitTrans
                    dbTransfer.Close
                    
                    Call newCls.UpdateCompleteReq(Db, cbo(3), txtLot, tglProd)
                    If strChildItemCD <> "" Then Call clsMRP.UpdateRequirementResult(Db, tglProd, "'" & cbo(3) & "'", txtLot, Left(strChildItemCD, Len(strChildItemCD) - 1))
                    Call IsiGrid
                    
                    Call cek_cboColumn(cbo(3))
                    ' If cbo(3).Column(5) = "01" Then
                    If strCol5 = "01" Then
                        sql = "Select ProductionSeq_No from WorkingTime_Master where ProductionSeq_No = '" & KeyProd & "'"
                        Set rsCek = Db.Execute(sql)
                        If rsCek.EOF Then Call Command1_Click(3)
                        Set rsCek = Nothing
                    End If
                    Call kosongBwh
                End If
'            End If
        End If
        
Case 1: 'cancel
        If MsgBox("Are you sure want to cancel transfer proccess?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation") = vbYes Then
            blnCancel = True
            Call kosongBwh
            Call IsiGrid
            LblErrMsg = ""
        Else
           blnCancel = False
        End If

Case 2: 'Material Consumption
        If hakAkses("frmProdMaterialComp") = 0 Then Me.MousePointer = vbDefault: LblErrMsg = DisplayMsg(3007): Exit Sub
        pesanTgl = up_ValidateDateRange(Format(dt, "yyyy-MM-dd"), True)
        If pesanTgl <> "" Then LblErrMsg = pesanTgl: Me.MousePointer = vbDefault: Exit Sub
        pesanTgl = HolidayCheck(cbo(0), Format(dt, "yyyy-MM-dd"))
        If pesanTgl <> "" Then LblErrMsg = pesanTgl: Me.MousePointer = vbDefault: Exit Sub
        
        If cbo(3) = "" Then
            LblErrMsg = DisplayMsg(5002)             'Please Select Data
            Me.MousePointer = vbDefault
            Exit Sub
        ElseIf CDbl(txtQty) < qtydiPake Then 'Harus Lebih Besar dr Supp yg memakai receipt tersebut
            LblErrMsg = DisplayMsg(4043) & " " & Format(qtydiPake, gs_formatQty) & " (Already Being Supplied)"
            txtQty.SetFocus
            Me.MousePointer = vbDefault
            Exit Sub
        ElseIf CDbl(txtQty.Text) > gd_MaxQty Then
            LblErrMsg = DisplayMsg(4045) & " " & gd_MaxQty & " !"
            txtQty.SetFocus
            Me.MousePointer = vbDefault: Exit Sub
        Else
            If CDbl(txtQty) = 0 Then LblErrMsg = DisplayMsg(4044): Me.MousePointer = vbDefault: Exit Sub
            
            DoEvents
            tampungBln = newCls.blnAkhir()
            blnFix = Split(tampungBln, ",")(0)
            thnFix = Split(tampungBln, ",")(1)
            If blnFix = 0 Then LblErrMsg = DisplayMsg(4019): Me.MousePointer = vbDefault: Exit Sub
            
            With frmProdMaterialComp
                .factoryCD = cbo(0)   'factory CD
                .ZWHCode = cbo(2)   'factory CD
                .lblitem = cbo(3) 'Prod Code
                .lbldesc = lblNm(0)  'desc
                .lblLot = txtLot 'Lot No
                .lblDt = Format(dt, "dd MMM yyyy")  'Result Date
                .lblResultQty = Format(txtQty, gs_formatQty)  'Qty
                .lblDailyQty = Format(qtyDaily, gs_formatQty)   'Qty Daily
                
                .dailyseqno = dailyseqno   'Daily Seq
                .KeyProd = KeyProd   'Seq No Prod
                .thnFix = thnFix 'Thn Fix
                .blnFix = blnFix 'Bln Fix
                .schedule_date = tglProd
                .ResultSeq = KeyProd
                .completeCls = completeCls 'Complete Cls
                
'                If ubah Then
'                    Call .IsiGrid(KeyProd, False, dailyseqno)
'                Else
'                    Call .IsiGrid(KeyProd, True, dailyseqno)
'                End If
''
               Call .IsiGrid1(dailyseqno, Trim(cbo(0).Text), Trim(lblNm(0)))

               ' If completeCls = 1 Then .Command1(0).Enabled = False
                Call .Show
            End With
            DoEvents
            Me.Hide
        End If
    
Case 3:  'Working Time
    With frmProdWorkingTime
        ' Perubahan dari KeyProd --> dailySeqNo
        .ProdSeqNo = KeyProd
        .ViewDt (1)
        .cmdsubmenu.Caption = "&Back"
        .Show
    End With
    Me.Hide
End Select

Me.MousePointer = vbDefault
End Sub

Function isiPrice(ItemCode As String, tglDO As String, currCode As String) As String
Dim rsPrice As New ADODB.Recordset

    sql = "select top 1 currency_code,price from price_master where " & _
           "item_code='" & ItemCode & _
           "' and price_cls='01' " & _
           "and start_date<='" & Format(tglDO, "yyyymmdd") & _
           "' and end_date>='" & Format(tglDO, "yyyymmdd") & _
           "' order by trade_code desc, priority_cls desc"
    Set rsPrice = Db.Execute(sql)
    
    If rsPrice.EOF Then
        isiPrice = currCode & ",0"
    Else
        isiPrice = Trim(rsPrice(0)) & "," & Trim(rsPrice(1))
    End If
    Set rsPrice = Nothing
End Function

Sub simpanUbah(nmDb As Connection)

Dim UpdateQty As Double
Dim keupdate As Long

    LblErrMsg.Caption = ""
    Command1(0).Enabled = False
    
    'On Error Resume Next
'    dbTransfer.ConnectionTimeout = 0
'    dbTransfer.CommandTimeout = 0
'    dbTransfer.Open Db.ConnectionString
'    dbTransfer.BeginTrans
        
    Call inputReceipt(nmDb)
        
    'frmProdMaterialComp.KeyProd = KeyProd
    
    UpdateQty = IIf(cboResultCls = "P1", CDbl(txtQty), -(CDbl(txtQty)))
    '************* Update Complete Daily ******************
    If qtyDaily <= (qtyAllResult - tampungQty) + UpdateQty Then
        sql = "Update Daily_Production set Complete_Cls = '1', Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
            "where Seq_No = " & Val(dailyseqno) & " And (Complete_Cls = 0 Or Complete_Cls IS NULL)"
        nmDb.Execute sql, keupdate
        frmProdMaterialComp.completeCls = 1
        
    End If
    UpdateQty = IIf(cboResultCls = "P1", CDbl(txtQty), -(CDbl(txtQty)))
    If cboResultCls.ListIndex = 0 Then
        Call cek_cboColumn(cbo(3))
        ' If cbo(3).Column(3) = "01" And cbo(3).Column(4) = "01" Then
        If strCol3 = "01" And strCol4 = "01" Then _
            Call newCls.updateStock(cbo(2), cbo(3), UpdateQty, "", _
                Format(dt, "yyyy-MM-dd"), blnFix, thnFix, nmDb, "Receipt", tampungQty, 1)
                
            sql = "Warehouse_Code = '" & cbo(2) & _
                "' And Item_Code = '" & cbo(3) & "'"
    End If
    
    Call cek_cboColumn(cbo(3))
    'If cbo(3).Column(5) = "01" Then
    If strCol5 = "01" Then
        Call clsMRP.HapusDataSupp(nmDb, "'" & CStr(KeyProd) & "'", blnFix, thnFix, 0)
        Call inputConsump(nmDb)
    End If
    Command1(0).Enabled = True
    LblErrMsg = DisplayMsg(IIf(simpan = True, 1000, 1101))
    
End Sub

Sub hapusData()
Dim tanya
Dim stockItem As String, stockWH As String
Dim ItemCode As String, Qty As Double
Dim KeyProd As Double, strKey As String, stSupp As String

Dim X As Long
Dim awal As Long, akhir As Long, Panjang As Integer
Dim TempSerial As String, Depan As String
Dim rsCeks As New ADODB.Recordset, sql1 As String

With grid
    LblErrMsg.Caption = ""
    Command1(0).Enabled = False
    
    'On Error Resume Next
    dbTransfer.ConnectionTimeout = 0
    dbTransfer.CommandTimeout = 0
    dbTransfer.Open Db.ConnectionString
    dbTransfer.BeginTrans
    
    strKey = ""
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 0) = "D" Then
            If IsEmpty(tanya) Then tanya = MsgBox("Do you really want to Delete this Data?", vbQuestion & vbYesNo, "Confirmation")
            If tanya = vbYes Then
                '********Hapus data
                ItemCode = .TextMatrix(i, 1)
                Qty = IIf(cboResultCls = "P1", CDbl(CDbl(.TextMatrix(i, bteColQty))), -(CDbl(CDbl(.TextMatrix(i, bteColQty)))))
                stockItem = .TextMatrix(i, bteColStockItem)
                stockWH = .TextMatrix(i, bteColStockWH)
                KeyProd = CDbl(.TextMatrix(i, bteColSeqNo))
                completeCls = .TextMatrix(i, bteColDailyComplete)
                seqNoConsump = .TextMatrix(i, bteColConsumpSeqNo)
                stSupp = .TextMatrix(i, bteColSupply)
                
                If completeCls = 1 Then 'Daily Completed
                    Command1(0).Enabled = True
                    Me.MousePointer = vbDefault
                    
                    dbTransfer.CommitTrans
                    dbTransfer.Close
                    
                    LblErrMsg = DisplayMsg(1110)
                    Exit Sub
                
                ElseIf seqNoConsump <> 0 And stSupp = "02" Then  'Sudah DiConsump
                
                    Command1(0).Enabled = True
                    Me.MousePointer = vbDefault
                    
                    dbTransfer.CommitTrans
                    dbTransfer.Close
                    
                    LblErrMsg = DisplayMsg(1209) 'Already Consump
                    Exit Sub
                    
'                ElseIf .TextMatrix(i, bteColWorkTime) <> 0 Then 'Working Time
'                    command1(0).Enabled = True
'                    Me.MousePointer = vbDefault
'
'                    dbTransfer.CommitTrans
'                    dbTransfer.Close
'
'                    LblErrMsg = DisplayMsg(8013) 'Can't delete Record. Please delete Working Time first
'                    Exit Sub
                    
                Else
                    sql = "delete Part_Receipt where Seq_No  = " & KeyProd
                    dbTransfer.Execute sql
                    
                    sql = "delete WorkingTime_Master where ProductionSeq_No = " & KeyProd
                    dbTransfer.Execute sql
                    
                    ' Update Serial Number Back to Production Data
                    ' Update 20090210
                    
                    If .TextMatrix(i, BteColSerialFrom) <> "" And .TextMatrix(i, BteColSerialTo) <> "" Then
                        Panjang = Len(Trim(.TextMatrix(i, BteColSerialFrom)))
                        
                        awal = Val(Mid(.TextMatrix(i, BteColSerialFrom), 2, (Panjang - 1)))
                        akhir = Val(Mid(.TextMatrix(i, BteColSerialTo), 2, (Panjang - 1)))
                        
                        For X = awal To akhir
                            TempSerial = Left(.TextMatrix(i, BteColSerialFrom), 1) & Format(X, String(Panjang - 1, "0"))
'                            sql1 = "Update Serial_Detail " & _
'                                " Set Result_No=Null, Serial_Status='2' where Result_No='" & .TextMatrix(i, bteColSeqNo) & _
'                                "'  and item_code='" & .TextMatrix(i, bteColProdCode) & "' and " & _
'                                " Serial_No='" & TempSerial & "'"
'                            dbTransfer.Execute sql1
                        Next X
                    End If
                    
                    ' --------------------------------------
                    
                    
                    
                    
                    If cboResultCls.ListIndex = 0 Then
                        If stockItem = "01" And stockWH = "01" Then _
                            Call newCls.updateStock(cbo(2), ItemCode, Qty, "", _
                                Format(dt, "yyyy-MM-dd"), blnFix, thnFix, dbTransfer, "Receipt", 0, 0)
                    End If
                                                    
                    strKey = strKey & "'" & KeyProd & "',"
                End If
            Else
                Command1(0).Enabled = True
                dbTransfer.CommitTrans
                dbTransfer.Close
                Me.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    Next i
    
    If strKey <> "" Then Call newCls.HapusDataSupp(dbTransfer, Left(strKey, Len(strKey) - 1), blnFix, thnFix)
    
    Command1(0).Enabled = True
        
    dbTransfer.CommitTrans
    dbTransfer.Close
    
    Call kosongBwh
    Call IsiGrid
    
    LblErrMsg = DisplayMsg(1201)
End With
End Sub

'********* Validate ******
Private Sub txtLot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub txtLot_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
LblErrMsg.Caption = ""
    If KeyCode = 13 Then SendKeys vbTab
    If txtQty.Text <> "" Then ' Qty Tidak boleh lebih besar dari Remaning
        If CDbl(txtQty.Text) > QtyRemaining Then txtQty.Text = QtyRemaining: LblErrMsg.Caption = "Qty must be equal or lower than" & " " & QtyRemaining: Exit Sub
        
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
        And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
        Then KeyAscii = 0
    If txtQty <> "" And IsNumeric(txtQty & Chr(KeyAscii)) Then
       If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    End If
    
End Sub

Private Sub txtQty_LostFocus()

    LblErrMsg.Caption = ""
    If txtQty.Text <> "" Then ' Qty Tidak boleh lebih besar dari Remaning
        If CDbl(txtQty.Text) > QtyRemaining Then txtQty.Text = 0: LblErrMsg.Caption = "Qty must be equal or lower than" & " " & QtyRemaining: Exit Sub
        
    End If


    If IsNumeric(txtQty.Text) Then
        txtQty.Text = Format(txtQty.Text, gs_formatQty)
    Else
        txtQty.Text = Format(0, gs_formatQty)
    End If

End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


'************ Unload **********
Private Sub CmdSubMenu_Click()
    DoEvents
    If cmdsubmenu.Caption = "Sub &Menu" Then
        frmMainMenu.Show
    Else
        Call frmProdResultInquiry.cmdSearch_Click
        frmProdResultInquiry.Show
    End If
    DoEvents
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub
'**************

Sub inputReceipt(nmDb As ADODB.Connection)
Dim rsSimpan As New ADODB.Recordset
Dim UpdateQty As Double
Dim nilPriceReceipt As String, CurrReceipt As String, PriceReceipt As Double

Dim X As Long
Dim awal As Long, akhir As Long, Panjang As Integer
Dim TempSerial As String, TempCancel As String, Depan As String
Dim rsCeks As New ADODB.Recordset, sql1 As String


    On Error Resume Next
    
    nilPriceReceipt = isiPriceReceipt(cbo(3).Text, dt.Value)
    CurrReceipt = Split(nilPriceReceipt, ",")(0)
    PriceReceipt = CDbl(Split(nilPriceReceipt, ",")(1))
    

    UpdateQty = IIf(cboResultCls = "P1", CDbl(txtQty), -(CDbl(txtQty)))

    sql = "select Seq_No,Supplier_Code,PO_No as Line_Code,Warehouse_Code,SerialNoFrom,SerialNoTo," & _
        "Receipt_cls,Receipt_Date,Item_Code,Qty,SuratJalan_No, " & _
        "ProductionResult_Cls as Cls,Remarks, dailyseq_no,unit_cls " & _
        "from Part_Receipt where " & "Seq_No = " & KeyProd
    rsSimpan.Open sql, nmDb, adOpenDynamic, adLockOptimistic
    
    With rsSimpan
        If Not (.EOF) And simpan = True Then
            LblErrMsg = DisplayMsg(1023)
            Command1(0).Enabled = True
            'nmdb.CommitTrans
            'nmdb.Close
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        If .EOF Then .AddNew: KeyProd = newCls.keyReceipt
        !Seq_no = KeyProd
        !Supplier_Code = Trim(cbo(0))
        !line_code = Trim(cbo(1))
        !Warehouse_Code = Trim(cbo(2))
        !receipt_cls = Trim(cboResultCls)
        !Receipt_Date = Trim(Format(dt, "yyyy-MM-dd"))
        !Item_Code = Trim(cbo(3))
        !Qty = UpdateQty
        !SuratJalan_No = Trim(txtLot)
        !dailyseq_no = dailyseqno
        !Unit_cls = UnitCls
        !Remarks = Trim(txtremarks)
        
        ' Add 20090210
        !SerialNoFrom = Trim(TxtSerialFrom)
        !SerialNoTo = Trim(TxtSerialTo)
        '---
        
        !Cls = "1"
        !Last_Update = Now
        !last_user = userLogin
        
        '**** Handle No yg Sama saat yg Sama
        If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
            Command1(0).Enabled = True
            nmDb.CommitTrans
            nmDb.Close
            err.clear
            Me.MousePointer = vbDefault
            Call simpanUbah(nmDb)
            Exit Sub
        End If
        .update
    
    End With
'    If UpdateQty = 0 Then
'        Sql = "delete Part_Receipt where Seq_No  = " & KeyProd
'        Db.Execute Sql
'
'        Sql = "delete WorkingTime_Master where ProductionSeq_No = " & KeyProd
'        Db.Execute Sql
'        Exit Sub
'    End If
    
    ' Save Result Seq No And Update Serial Number Status  - 20090210
    Dim IntSelisih As Integer, IntCek As Integer
    Dim StrSerialCek As String
    
    Panjang = Len(Trim(TxtSerialFrom))
    Depan = Left(Trim(TxtSerialFrom), 1)
    
    awal = Val(Mid(TxtSerialFrom, 2, (Panjang - 1)))
    akhir = Val(Mid(TxtSerialTo, 2, (Panjang - 1)))
    
    If Trim(TxtSerialFrom) <> "" And Trim(TxtSerialTo) <> "" Then
        
        ' Add Serial No Data base on Order Data
        If (akhir + 1) - awal <> txtQty Then
            IntSelisih = 1
        Else
            IntSelisih = 0
        End If
        
       For X = awal To akhir
            TempSerial = Depan & Format(X, String(Panjang - 1, "0"))
            IntCek = 1
            If IntSelisih = 1 Then
                IntCek = IntCek + 7
                Do While IntCek < Len(Trim(txtremarks)) - 7
                    StrSerialCek = Mid(Trim(txtremarks), IntCek, 7)
                    If TempSerial <> StrSerialCek Then
                        sql1 = "Update Serial_Detail Set Result_No='" & KeyProd & "' , " & _
                                 " Serial_Status = '3' Where item_Code='" & Trim(cbo(3)) & "' And " & _
                                 " Serial_No='" & TempSerial & "'"
                        Db.Execute (sql1)
                    End If
                    IntCek = IntCek + 8
                Loop
            Else
                sql1 = "Update Serial_Detail Set Result_No='" & KeyProd & "' , " & _
                         " Serial_Status = '3' Where item_Code='" & Trim(cbo(3)) & "' And " & _
                         " Serial_No='" & TempSerial & "'"
                Db.Execute (sql1)
            End If
        Next X
    End If
    
' --------------------
    
    
    
    frmProdMaterialComp.KeyProd = KeyProd
    If rsSimpan.State <> adStateClosed Then rsSimpan.Close
End Sub

Function isiPriceReceipt(ItemCode As String, resultdate As Date) As String
    Dim rsPrice As New ADODB.Recordset
    
    sql = "select top 1 currency_code, Price from price_master where item_code ='" & ItemCode & "' and priority_cls ='1' and start_date <= '" & Format(resultdate, "YYYYMMDD") & "' and end_date >= '" & Format(resultdate, "YYYYMMDD") & "'"
    Set rsPrice = Db.Execute(sql)
    
    If rsPrice.EOF Then
        isiPriceReceipt = "" & ",0"
    Else
        isiPriceReceipt = Trim(rsPrice(0)) & "," & Trim(rsPrice(1))
    End If
    Set rsPrice = Nothing
End Function

Function HolidayCheck(Factory As String, dt As Date) As String
Dim rscal As New Recordset
sql = "select * from calendar_Master where Factory_Code ='" & Factory & "' and cal_date = '" & Format(dt, "YYYY-MM-DD") & "'"
Set rscal = Db.Execute(sql)
If rscal.EOF Then
    HolidayCheck = ""
Else
    HolidayCheck = "Please select a valid time range !"
End If
End Function

Sub inputConsump(nmDb As Connection)
Dim rsAnak As New ADODB.Recordset
Dim itemAnak As String, UnitCls As String
Dim qtyAnak As Double, qtyConvert As Double
Dim QtyResult As Double
    
    strChildItemCD = ""
    
    '*********Update Supply Anak2nya diambil dr BOM Master ***********
    sql = "Select BOM.Item_Code, BOM.Qty as QtyAnak, BOM.Unit_Cls as UnitBOM, IC.Unit_Cls as UnitItem, " & _
            "IC.WH_Code, IC.StockControl_Cls as StockItem, WH.Stockcontrol_Cls as stockWH " & _
        "from BOM_Master BOM, Item_Master IC, Warehouse_Master WH " & _
        "where BOM.Item_Code = IC.Item_Code " & _
            "And IC.WH_Code = WH.WH_Code " & _
            "and BOM.Parent_ItemCode = '" & cbo(3) & "' " & _
            "and BOM.Start_Date <='" & Format(dt, "yyyyMMdd") & "' " & _
            "and BOM.End_Date >= '" & Format(dt, "yyyyMMdd") & "'"
    Set rsAnak = Db.Execute(sql)

    Do While Not rsAnak.EOF
    
        If blnCancel Then
            nmDb.RollbackTrans: blnCancel = False
            nmDb.Close
            Call IsiGrid
            Command1(0).Enabled = True
            LblErrMsg = DisplayMsg(8097)
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        itemAnak = Trim(rsAnak!Item_Code)
        
        QtyResult = CDbl(txtQty)
        qtyAnak = rsAnak!qtyAnak * QtyResult
        qtyConvert = ClsProc.nilConvertUnit(rsAnak("QtyAnak"), rsAnak!unitBOM, rsAnak!UnitItem) * QtyResult
        
        UnitCls = rsAnak("unitBOM")
        
        nilPrice = isiPriceReceipt(itemAnak, Format(dt, "yyyy-MM-dd"))
        Curr = Split(nilPrice, ",")(0)
        Price = Split(nilPrice, ",")(1)
                        
        Call clsMRP.inputSupply(nmDb, 0, rsAnak!wh_code, "", rsAnak!wh_code, Format(dt, "yyyy-MM-dd"), _
            itemAnak, "", "S", qtyConvert, qtyAnak, UnitCls, _
            cbo(3), txtLot, CStr(KeyProd), tglProd)
                                        
        If rsAnak("StockWH") = "01" And rsAnak("StockItem") = "01" Then
            Call clsMRP.updateStock(rsAnak!wh_code, itemAnak, qtyConvert, "", _
                Format(dt, "yyyy-MM-dd"), blnFix, thnFix, dbTransfer, "Supply", 0, 1)
        End If
        strChildItemCD = strChildItemCD & "'" & itemAnak & "',"
        rsAnak.MoveNext
    Loop
    '******************
End Sub

Private Sub cek_cboColumn(psItem As String)
   On Local Error GoTo err_handler
   Screen.MousePointer = vbHourglass
   
   Dim adoRs   As New ADODB.Recordset
   Dim strSQL  As String
   
   strSQL = "select a.Item_Code,MakerItem_Code,Item_Name, " & _
            "a.StockControl_cls as StockItem, b.StockControl_cls as StockWH, Suply_Cls = ISNULL(a.Suply_Cls,'')  " & _
            "from Item_master a, Warehouse_Master b " & _
            "where a.WH_Code = b.WH_Code and a.Item_code = '" & psItem & "' order by Item_Code"
   
   Set adoRs = Db.Execute(strSQL)
      
   If Not adoRs.EOF Then
      strCol3 = Trim(adoRs(3))
      strCol4 = Trim(adoRs(4))
      strCol5 = Trim(adoRs(5))
   End If
   adoRs.Close
   
err_exit:
   Screen.MousePointer = vbDefault
   Set adoRs = Nothing
   Exit Sub
err_handler:
   MsgBox "Err. Number : " & err.number & vbCrLf & "Err. Description : " & err.Description, vbCritical, "Error"
   err.clear
   Resume err_exit
End Sub

Private Sub TxtSerialFrom_LostFocus()
If TxtSerialFrom <> "" Then TxtSerialTo.Text = GetSerialTo(Trim(TxtSerialFrom), txtQty)
If TxtSerialFrom = "" Then TxtSerialTo = ""
End Sub
