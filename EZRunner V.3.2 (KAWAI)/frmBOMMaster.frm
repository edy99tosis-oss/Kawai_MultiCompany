VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmBOMMaster 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BOM Master"
   ClientHeight    =   11025
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   15240
   Icon            =   "frmBOMMaster.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11025
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "..."
      Height          =   300
      Index           =   1
      Left            =   2880
      TabIndex        =   43
      Top             =   8722
      Width           =   300
   End
   Begin VB.TextBox txtRevision 
      Height          =   300
      Left            =   13020
      MaxLength       =   50
      TabIndex        =   40
      Top             =   8715
      Width           =   1980
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
      Height          =   315
      Left            =   7695
      MaxLength       =   11
      TabIndex        =   3
      Tag             =   "1"
      Text            =   "99,999.9999"
      Top             =   8715
      Width           =   1155
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
      Index           =   0
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9870
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   795
      Left            =   150
      TabIndex        =   32
      Top             =   1320
      Width           =   14940
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   4110
         TabIndex        =   42
         Top             =   307
         Width           =   300
      End
      Begin VB.TextBox txtDocNo 
         Height          =   315
         Left            =   12345
         MaxLength       =   50
         TabIndex        =   38
         Top             =   300
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. Number"
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
         TabIndex        =   39
         Top             =   360
         Width           =   1125
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   7140
         X2              =   10815
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblNm 
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
         Left            =   7140
         TabIndex        =   35
         Top             =   360
         Width           =   3690
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   1500
         TabIndex        =   0
         Top             =   300
         Width           =   2550
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "4498;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAAAAAAAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblNm 
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
         Left            =   4500
         TabIndex        =   34
         Top             =   360
         Width           =   2565
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   4500
         X2              =   7050
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parent Code"
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
         TabIndex        =   33
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   150
      TabIndex        =   30
      Top             =   9165
      Width           =   14940
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
         TabIndex        =   31
         Top             =   195
         Width           =   14715
      End
   End
   Begin VB.TextBox txtNm 
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
      Height          =   210
      Left            =   5445
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "AAAAAAA"
      Top             =   8775
      Width           =   2175
   End
   Begin VB.CommandButton cmdPage 
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Page"
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
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9870
      Width           =   1140
   End
   Begin VB.CommandButton cmdPage 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last Page"
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
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9870
      Width           =   1140
   End
   Begin VB.CommandButton cmdPage 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Next Page"
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
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9870
      Width           =   1140
   End
   Begin VB.CommandButton cmdPage 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Prev Page"
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
      Left            =   5805
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9870
      Width           =   1140
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9870
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "BOM Inquiry"
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
      Left            =   1365
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9870
      Width           =   1365
   End
   Begin VB.TextBox txtRQty 
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
      Height          =   315
      Left            =   7695
      MaxLength       =   11
      TabIndex        =   2
      Tag             =   "1"
      Text            =   "AAAAAAA"
      Top             =   8715
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSMask.MaskEdBox dtAkhir 
      Height          =   315
      Left            =   11415
      TabIndex        =   5
      Top             =   8715
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13245
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   510
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
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
      Left            =   12735
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9870
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
      Left            =   13950
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9870
      Width           =   1140
   End
   Begin MSComCtl2.DTPicker dtAwal 
      Height          =   315
      Left            =   9810
      TabIndex        =   4
      Top             =   8715
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
      Format          =   141230083
      CurrentDate     =   37891
   End
   Begin MSComCtl2.DTPicker dtAkhir1 
      Height          =   315
      Left            =   11415
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8715
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
      Format          =   141230083
      CurrentDate     =   37799
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5985
      Left            =   150
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2205
      Width           =   14940
      _cx             =   26352
      _cy             =   10557
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
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
      Left            =   13020
      TabIndex        =   41
      Top             =   8340
      Width           =   720
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
      Index           =   4
      Left            =   9060
      TabIndex        =   24
      Top             =   8340
      Width           =   330
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
      Left            =   8970
      TabIndex        =   17
      Top             =   8775
      Width           =   660
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   8970
      X2              =   9615
      Y1              =   9015
      Y2              =   9015
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
      Index           =   2
      Left            =   8115
      TabIndex        =   22
      Top             =   8340
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      Height          =   855
      Index           =   0
      Left            =   150
      Top             =   8280
      Width           =   14940
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Master"
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
      Left            =   6900
      TabIndex        =   37
      Top             =   510
      Width           =   1440
   End
   Begin MSForms.ComboBox cbo 
      Height          =   300
      Index           =   2
      Left            =   8925
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   8715
      Visible         =   0   'False
      Width           =   810
      VariousPropertyBits=   746604575
      BackColor       =   14737632
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "1429;529"
      BoundColumn     =   2
      TextColumn      =   2
      ListRows        =   0
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R Qty"
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
      Left            =   8025
      TabIndex        =   27
      Top             =   8340
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   5445
      X2              =   7605
      Y1              =   9015
      Y2              =   9015
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3255
      X2              =   5340
      Y1              =   9015
      Y2              =   9015
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
      Left            =   5445
      TabIndex        =   26
      Top             =   8340
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cls"
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
      Left            =   8985
      TabIndex        =   23
      Top             =   8340
      Visible         =   0   'False
      Width           =   660
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
      Left            =   420
      TabIndex        =   21
      Top             =   8340
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
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
      Left            =   3255
      TabIndex        =   20
      Top             =   8340
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
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
      Left            =   10140
      TabIndex        =   19
      Top             =   8340
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
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
      Left            =   11790
      TabIndex        =   18
      Top             =   8340
      Width           =   780
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
      Left            =   3255
      TabIndex        =   16
      Top             =   8760
      Width           =   2100
   End
   Begin MSForms.ComboBox cbo 
      Height          =   315
      Index           =   1
      Left            =   270
      TabIndex        =   1
      Top             =   8715
      Width           =   2565
      VariousPropertyBits=   746604571
      MaxLength       =   25
      DisplayStyle    =   3
      Size            =   "4524;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "AAAAAAAAAAAAAAA"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   330
      Index           =   1
      Left            =   150
      Top             =   8280
      Width           =   14940
   End
End
Attribute VB_Name = "FrmBOMMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Long
Dim simpan As Boolean, ubah As Boolean, hapus As Boolean 'Status Ubah/Hapus
Dim parentItem As String 'Parent Item
Dim gridTglAwal As String, gridTglAkhir As String 'Klik Grid
Dim sama As Boolean 'cek Parent
Dim nilKosong As Boolean
Dim kondisi As String
Dim tglAwal As String, tglAkhir As String
Dim tglSesdh As String, tglSeblm As String

Dim bteColSelect As Byte
Dim bteColProdCode As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColQtyR As Byte
Dim bteColQty As Byte
Dim bteColUnit As Byte
Dim bteColUnitDesc As Byte
Dim bteColDateStart As Byte
Dim bteColDateEnd As Byte
Dim bteColRevision As Byte

Private Sub headerGrid()
    Dim i As Integer
    
    bteColSelect = 0
    bteColProdCode = 1
    bteColPartNo = 2
    bteColDesc = 3
    bteColQtyR = 4
    bteColQty = 5
    bteColUnit = 6
    bteColUnitDesc = 7
    bteColDateStart = 8
    bteColDateEnd = 9
    bteColRevision = 10
    
    With grid
        .clear
        .ColS = 11
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = "S"
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColPartNo) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColQtyR) = "R Qty"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColUnitDesc) = "Desc"
        .TextMatrix(0, bteColDateStart) = "Start Date"
        .TextMatrix(0, bteColDateEnd) = "End Date"
        .TextMatrix(0, bteColRevision) = "Revision"
        
        .ColWidth(bteColSelect) = 300 'besar kolom
        .ColWidth(bteColProdCode) = 2200
        .ColWidth(bteColPartNo) = 2200
        .ColWidth(bteColDesc) = 3000
        .ColWidth(bteColQtyR) = 1300
        .ColWidth(bteColQty) = 1300
        .ColWidth(bteColUnit) = 500
        .ColWidth(bteColUnitDesc) = 600
        .ColWidth(bteColDateStart) = 1500
        .ColWidth(bteColDateEnd) = 1500
        .ColWidth(bteColRevision) = 1500
        
        .ColHidden(bteColQtyR) = True
        
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColPartNo) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColQtyR) = flexAlignRightCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColUnit) = flexAlignCenterCenter
        .ColAlignment(bteColUnitDesc) = flexAlignLeftCenter
        .ColAlignment(bteColDateStart) = flexAlignCenterCenter
        .ColAlignment(bteColDateEnd) = flexAlignCenterCenter
        .ColAlignment(bteColRevision) = flexAlignLeftCenter
        
        .EditMaxLength = 1
    End With
End Sub

Sub Kosong(Optional mulai As Integer)
nilKosong = True
    '****** Jika mulai <> 0 maka Parent Code ngga dihapus
    For i = mulai To 2
        cbo(i) = ""
        lblNm(i) = ""
    Next i
    
    If mulai = 0 Then lblNm(3) = "": dtAwal = Now
    txtNm = ""
    cbo(1).Enabled = True
    Call filterCboUnit("00", cbo(2))
    txtQty = 0
    txtRQty = 0
    dtAkhir = "99/99/9999"
    txtRevision = ""
    simpan = True: ubah = False: hapus = False
    '***** Jika mulai = 0 berarti grid kosong
    If mulai = 0 Then Call headerGrid
    Call nyalaBwh
nilKosong = False
End Sub

Sub isiCboItem()
Dim rscbo As New ADODB.Recordset 'Isi Combo

    '******** Isi Combo *********
   sql = "select Item_Code as a,MakerItem_Code b,Item_Name c,Unit_Cls " & _
           "from Item_Master " & _
           "where use_endday >= convert(char(8), getdate(), 112) " & _
           "order by Item_Code"
    Set rscbo = Db.Execute(sql)
    
    '**** Isi combo Parent dan Item Code
    cbo(0).clear
    cbo(0).columnCount = 3
    cbo(0).TextColumn = 1
    
    cbo(1).clear
    cbo(1).columnCount = 4
    cbo(1).TextColumn = 1
    
    i = 0
    Do While Not (rscbo.EOF)
        cbo(0).AddItem ""
        cbo(0).List(i, 0) = Trim(rscbo("a"))
        cbo(0).List(i, 1) = Trim(rscbo("b"))
        cbo(0).List(i, 2) = Trim(rscbo("c"))
        
        cbo(1).AddItem ""
        cbo(1).List(i, 0) = Trim(rscbo("a"))
        cbo(1).List(i, 1) = Trim(rscbo("b"))
        cbo(1).List(i, 2) = Trim(rscbo("c"))
        cbo(1).List(i, 3) = Trim(rscbo("Unit_Cls"))
        
        i = i + 1
        rscbo.MoveNext
    Loop
    cbo(0).ListWidth = 480
    cbo(0).ColumnWidths = "120 pt;120 pt;240 pt"
    cbo(0).ListRows = 15
    
    cbo(1).ListWidth = 480
    cbo(1).ColumnWidths = "120 pt;120 pt;240 pt;0pt"
    cbo(1).ListRows = 15
    
    '************************
    Set rscbo = Nothing
End Sub

Private Sub cmdBrowser_Click(Index As Integer)
 Me.MousePointer = vbHourglass
 Select Case Index
  Case 0:
   frm_BrowseItem.getItemCode = cbo(0).Text
   frm_BrowseItem.Show 1
   cbo(0).Text = frm_BrowseItem.getItemCode
  Case 1:
   If cbo(1).Enabled = True Then
    frm_BrowseItem.getItemCode = cbo(1).Text
    frm_BrowseItem.Show 1
    cbo(1).Text = frm_BrowseItem.getItemCode
   End If
  End Select
 Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    Call filterCboUnit("00", cbo(2))  'utk Unit
    Call Kosong
    Call isiCboItem
    
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 And Index = 0 Then Call cbo_Click(0)
End Sub

Private Sub cbo_Change(Index As Integer)
If Index = 0 Then Call headerGrid
If nilKosong = True Then Exit Sub
    lblNm(Index) = ""
    LblErrMsg = ""
    If Index = 0 Then
        lblNm(Index + 3) = ""
    ElseIf Index = 1 Then
        txtNm = ""
    End If
End Sub

Private Sub cbo_LostFocus(Index As Integer)
If nilKosong = True Then Exit Sub
If lblNm(Index) = "" Then
    If cbo(Index) <> "" Then
        If Index <> 0 Then
            Call cbo_Click(Index)
        End If
    
    End If
End If
End Sub

Private Sub cbo_Click(Index As Integer)
If nilKosong = True Then Exit Sub

Dim rscbo As New ADODB.Recordset

If Index = 0 Then parentItem = Trim(cbo(0))
If cbo(Index) <> "" Then
    cbo(Index) = cbo(Index)
    If cbo(Index).MatchFound = True Then
        lblNm(Index) = cbo(Index).Column(1)
        
        If Index = 0 Then 'Cbo Parent
            Call Kosong(1)
            Call IsiGrid
            lblNm(3) = cbo(0).Column(2)
            
            With grid
                If kirimPar <> "" Then 'DR BOM Inquiry
                    For i = 1 To .Rows - 1
                        If kirimPar = .TextMatrix(i, bteColProdCode) & .TextMatrix(i, bteColDateStart) Then
                            .Row = i
                            .SetFocus
                            .TextMatrix(i, bteColSelect) = "S"
                            
                            Call Grid_AfterEdit(i, bteColSelect)
                            kirimPar = ""
                            Exit Sub
                        End If
                    Next i
                    kirimPar = ""
                End If
            End With
            
        ElseIf Index = 1 Then 'cbo Child
            Call up_FillCombo(cbo(2), "unit_cls")
           ' Call filterCboUnit(cbo(1).Column(3), cbo(2))
            Dim rsUnit As New ADODB.Recordset
            rsUnit.Open "select * from Item_master where item_code='" & Trim(cbo(1)) & "'", Db, adOpenKeyset, adLockOptimistic
            If rsUnit.EOF = False Then
                cbo(2).Text = Trim(rsUnit!Unit_cls)
                lblNm(2) = uf_GetUnitDescription(Trim(rsUnit!Unit_cls))
                txtNm = Trim(rsUnit!item_name)
            Else
                cbo(2).Text = ""
                lblNm(2) = ""
                txtNm = ""
            End If
            If rsUnit.State <> adStateClosed Then rsUnit.Close
        End If
        '**************
        
        LblErrMsg = ""
    Else
        lblNm(Index) = ""
        
        If Index = 0 Then
            lblNm(Index + 3) = ""
            Call Kosong(1): Call headerGrid
        ElseIf Index = 1 Then
            txtNm = ""
        End If
        LblErrMsg = DisplayMsg(4002 + Index)
    End If
Else
    Call headerGrid
    lblNm(Index) = ""
    LblErrMsg = ""
End If
End Sub

'************ Isi Grid **********
Sub IsiGrid()
Dim rsTglAwal As String, rsTglAkhir As String
Dim rsBOM As New ADODB.Recordset
    
With grid
    Call headerGrid
    
    '******** Untuk BOM Master *********
    sql = "select a.*,b.MakerItem_Code from " & _
        "BOM_Master a,Item_Master b " & _
        "where a.Item_Code = b.Item_Code and " & _
        "Parent_ItemCode = '" & parentItem & _
        "' order by Parent_ItemCode,a.Item_Code,Start_Date"
    Set rsBOM = Db.Execute(sql)

    i = 1
    Do While Not rsBOM.EOF
        .Rows = .Rows + 1
        rsTglAwal = Trim(rsBOM("Start_Date"))
        rsTglAkhir = Trim(rsBOM("End_Date"))
        .TextMatrix(i, bteColSelect) = ""
        .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
        .TextMatrix(i, bteColProdCode) = Trim(rsBOM("Item_Code"))
        .TextMatrix(i, bteColPartNo) = Trim(rsBOM("MakerItem_Code"))
        .TextMatrix(i, bteColDesc) = Trim(rsBOM("Item_Name"))
        .TextMatrix(i, bteColQtyR) = Format(rsBOM("R_Qty"), gs_formatQtyBOM)
        .TextMatrix(i, bteColQty) = Format(Trim(rsBOM("Qty")), gs_formatQtyBOM)
        .TextMatrix(i, bteColUnit) = IIf(IsNull(Trim(rsBOM("Unit_Cls"))), "", Trim(rsBOM("Unit_Cls")))
        .TextMatrix(i, bteColUnitDesc) = uf_GetUnitDescription(Trim(rsBOM("Unit_Cls")))
        .TextMatrix(i, bteColDateStart) = Format(Left(rsTglAwal, 4) & "-" & CInt(Mid(rsTglAwal, 5, 2)) & "-" & _
                                          Right(rsTglAwal, 2), "dd mmm yyyy")
            
        If rsTglAkhir = "99999999" Then
            .TextMatrix(i, bteColDateEnd) = "99/99/9999"
        Else
            .TextMatrix(i, bteColDateEnd) = Format(Left(rsTglAkhir, 4) & "-" & CInt(Mid(rsTglAkhir, 5, 2)) & "-" & _
                                            Right(rsTglAkhir, 2), "dd mmm yyyy")
        End If
        .TextMatrix(i, bteColRevision) = Trim(rsBOM("Revision_No") & "")
        txtDocNo = Trim(rsBOM("doc_no") & "")
        i = i + 1
        rsBOM.MoveNext
    Loop
    Set rsBOM = Nothing
End With
End Sub

'******** Validate Grid ********

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

'******************
Sub nyalaBwh()
    cbo(1).Enabled = True
    txtQty.Enabled = True
    cbo(2).Enabled = True
    dtAwal.Enabled = True
    dtAkhir.Enabled = True
    dtAkhir1.Enabled = True
    txtRevision.Enabled = True
End Sub

Sub MatiinBwh()
    cbo(1).Enabled = False
    txtQty.Enabled = False
    cbo(2).Enabled = False
    dtAwal.Enabled = False
    dtAkhir.Enabled = False
    dtAkhir1.Enabled = False
    txtRevision.Enabled = False
End Sub

Private Sub kosongColGrid(Optional kolumn As String)
Dim i As Integer
Dim jmlD As Double

With grid
    .Col = 0
    
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
With grid
    TextGrid = grid.Text
    
    If TextGrid = "S" Then
        Call kosongColGrid
        Call nyalaBwh
        cbo(1) = Trim(.TextMatrix(Row, bteColProdCode))
        cbo(1).Enabled = False
        txtRQty = Trim(.TextMatrix(Row, bteColQtyR))
        txtQty = Trim(.TextMatrix(Row, bteColQty))
        cbo(2) = Trim(.TextMatrix(Row, bteColUnit))
        dtAwal = Trim(.TextMatrix(Row, bteColDateStart))
        gridTglAwal = Format(CDate(dtAwal), "yyyyMMdd")
        dtAkhir = IIf(IsDate(.TextMatrix(Row, bteColDateEnd)) = False, _
            "99/99/9999", Format(.TextMatrix(Row, bteColDateEnd), "MM/dd/yyyy"))
        dtAkhir1 = IIf(IsDate(.TextMatrix(Row, bteColDateEnd)), .TextMatrix(Row, bteColDateEnd), .TextMatrix(Row, bteColDateStart))
        gridTglAkhir = IIf(IsDate(dtAkhir), Format(dtAkhir, "yyyyMMdd"), "99999999")
        txtRevision = Trim(.TextMatrix(Row, bteColRevision))
        ubah = True
        simpan = False
        hapus = False
    Else
        Call Kosong(1)
        Call kosongColGrid("S")
    End If
    
    .TextMatrix(Row, Col) = TextGrid
End With
End Sub

'******************

'*** Klik button ****
Sub samaParent(anak As String)
Dim kdCekParent As String
Dim rsCekParent As New ADODB.Recordset
    '****** Cek apakah menjadi parent dari anaknya
    sql = "select Parent_ItemCode from BOM_Master where Item_Code='" & anak & "'"
    Set rsCekParent = Db.Execute(sql)
    
    If Not (rsCekParent.EOF) Then
        Do While Not rsCekParent.EOF
            kdCekParent = rsCekParent(0)
            If LCase(Trim(kdCekParent)) = LCase(Trim(cbo(1))) Then
                sama = True
                Exit Sub
            End If
            Call samaParent(kdCekParent)
            If Not (rsCekParent.EOF) Then rsCekParent.MoveNext
        Loop
    End If
    Set rsCekParent = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim rsUpdate As New ADODB.Recordset
Dim tanya

Me.MousePointer = vbHourglass
Select Case Index
Case 0: 'Submit
    If hakUpdate(Me.Name) = 0 Then _
        LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
    '****** Validate ********
    If cbo(0) <> parentItem Then
        LblErrMsg = DisplayMsg(1007)
        cbo(0).SetFocus
    Else

        '******** Hapus *********
        If hapus = True Then
            Call hapusData
        Else
            '********* Simpan atau Ubah
            If LCase(Trim(cbo(1))) = LCase(Trim(parentItem)) Then 'Jika sama dgn Parentnya
                LblErrMsg = DisplayMsg(1014)
                cbo(1).SetFocus
            Else
                
                '****** Cek apakah menjadi parent dari anaknya
                sama = False
                Call samaParent(parentItem)
                If sama = True Then
                    LblErrMsg = DisplayMsg(1020)
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
                '******************
        
                '**** CEK Data Kosong dan Cocok / Tdk
                For i = 0 To 2
                    If cbo(i) = "" Then
                        LblErrMsg = DisplayMsg(1008 + i)
                        cbo(i).SetFocus
                        Me.MousePointer = vbDefault
                        Exit Sub
                    Else '**** Jika data tdk kosong cek Data cocok / tdk dgn Combo
                        If i <> 2 Then
                            cbo(i).MatchEntry = 0
                            cbo(i) = cbo(i)
                            If cbo(i).MatchFound = False Then
                                '**** cek Data cocok / tdk dgn Database
                                Dim rsUnit As New ADODB.Recordset
                                rsUnit.Open "select Item_Code from Item_Master where Item_Code='" & Trim(cbo(i)) & "'", Db, adOpenKeyset, adLockOptimistic
                                If rsUnit.EOF = True Then
                                    LblErrMsg = DisplayMsg(4002 + i)
                                    cbo(i).MatchEntry = 2
                                    Me.MousePointer = vbDefault
                                    Exit Sub
                                End If
                                rsUnit.Close
                            End If
                            cbo(i).MatchEntry = 2
                        Else
                                rsUnit.Open "select * from unit_Cls where unit_cls='" & Trim(cbo(2)) & "'", Db, adOpenKeyset, adLockOptimistic
                                If rsUnit.EOF = True Then
                                    LblErrMsg = DisplayMsg(4004)
                                    Me.MousePointer = vbDefault
                                    rsUnit.Close
                                    Exit Sub
                                End If
                                rsUnit.Close
                        End If
                    End If
                Next i
                '********************

                '***** Validate ******
                If CDbl(txtQty) = 0 Then 'Qty tdk boleh 0
                    LblErrMsg = DisplayMsg(1012)
                    txtQty.SetFocus
                ElseIf IsDate(dtAkhir) = False And dtAkhir <> "99/99/9999" Then
                    LblErrMsg = DisplayMsg(1021)
                    dtAkhir.SetFocus
                Else
                    '******** Format Tanggal Awal & Tanggal Akhir yg diinput
                    tglAwal = Format(dtAwal, "yyyyMMdd")
                    '****** Cek Start Date < End date
                    If IsDate(dtAkhir) Then
                        tglAkhir = Format(dtAkhir, "yyyyMMdd")
                        If Format(dtAwal, "yyyy-MM-dd") > _
                            Format(CDate(dtAkhir), "yyyy-MM-dd") Then _
                            LblErrMsg = DisplayMsg("0063") & " " & Format(dtAwal, "dd MMM yyyy"): _
                            dtAkhir1.SetFocus: Me.MousePointer = vbDefault: Exit Sub
                    Else
                        tglAkhir = "99999999"
                    End If
                    '***********
                '********************
                 
                    kondisi = "Parent_ItemCode ='" & parentItem & "' and " & _
                                "Item_Code ='" & cbo(1) & " '"
                                
                    sql = "update bom_master set doc_no = '" & txtDocNo & "' where parent_itemcode = '" & parentItem & "' "
                    Db.Execute sql
                    
                    If simpan = True Then
                        Call simpanData
                    ElseIf ubah = True Then
                        Call ubahData
                    End If
                    
                End If
            End If
        End If
    End If
    
Case 1: 'Cancel
    Call Kosong(1)
    Call kosongColGrid
End Select
Me.MousePointer = vbDefault
End Sub

Sub simpanData()
    
    Dim rsSimpan As New ADODB.Recordset
    Dim rsUpdate As New ADODB.Recordset 'utk Update Data
    
    sql = "Select Start_Date from BOM_Master where " & kondisi
    rsSimpan.Open sql, Db, adOpenDynamic, adLockReadOnly
    If Not rsSimpan.EOF Then 'jika telah ada data
        rsSimpan.Find "Start_Date = '" & tglAwal & "'"
        If Not rsSimpan.EOF Then 'jika telah ada data dan start date sama
            LblErrMsg = DisplayMsg(1015)
            dtAwal.SetFocus
            GoTo ErrExit
        Else
            'Yudha 2008-01-15 : Tambah dialog box konfirmasi lanjutkan proses atau tidak
            If MsgBox("Item already exists. Do you want to continue ?", vbQuestion + vbOKCancel) <> vbOK Then
                LblErrMsg = DisplayMsg(8097)
                GoTo ErrExit
            End If
        End If
    End If
    rsSimpan.Close
    
    '***** cek data max sebelonnya dan update end Date nya lebih kecil 1 hari dr baru yg diinput
    sql = "update  BOM_Master " & _
        "set End_Date ='" & Format(DateAdd("d", -1, dtAwal), "yyyyMMdd") & "', " & _
        "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
        "where " & kondisi & " and Start_Date = (" & _
        "select max(Start_Date) as tglSblm from BOM_Master where " & kondisi & "and Start_Date  < '" & tglAwal & "')"
    Db.Execute sql
        
    '****** cek data sesudahnya agar end date yg baru diinput menjadi start date data sesudahnya-1
    sql = "select min(Start_Date) as tglSsdh from BOM_Master " & _
        "where " & kondisi & "and " & _
        "Start_Date  > '" & tglAwal & "'"
    Set rsUpdate = Db.Execute(sql)
    If IsNull(rsUpdate("tglSsdh")) = False Then
        tglSesdh = CDate(Left(rsUpdate("tglSsdh"), 4) & "-" & Mid(rsUpdate("tglSsdh"), 5, 2) & "-" & Right(rsUpdate("tglSsdh"), 2))
        tglAkhir = Format(DateAdd("d", -1, tglSesdh), "yyyyMMdd")
    End If
    Set rsUpdate = Nothing
    
    '********** Simpan Data *********
    sql = "insert into BOM_Master (Parent_ItemCode,Description,Item_Code,Item_name,Qty,Unit_Cls,R_Qty,Start_Date,End_Date,Last_Update,Last_User,Doc_No, Revision_No) " & _
        "values ('" & Trim(parentItem) & "','" & lblNm(3) & "','" & cbo(1) & "','" & txtNm & "'," & CDbl(txtQty) & ",'" & Trim(cbo(2)) & "','" & CDbl(txtRQty) & "','" & tglAwal & "','" & tglAkhir & "', getdate(), '" & userLogin & "', '" & txtDocNo & "', '" & txtRevision & "')"
    Db.Execute sql
    
    Call IsiGrid
    Call Kosong(1)
    
    LblErrMsg = DisplayMsg(1000)
    
ErrExit:
    Set rsUpdate = Nothing
    Set rsSimpan = Nothing

End Sub

Sub ubahData()
Dim tanya
Dim rsUpdate As New ADODB.Recordset 'utk Update Data

    '***** cek data max sebelonnya syarat start date > start date seblomnya
    sql = "select max(Start_Date) as tglSblm from BOM_Master " & _
        "where " & kondisi & "and " & _
        "Start_Date  < '" & gridTglAwal & "'"
    Set rsUpdate = Db.Execute(sql)
    If IsNull(rsUpdate("tglSblm")) = False Then
        tglSeblm = rsUpdate("tglSblm")

        If Format(dtAwal, "yyyyMMdd") <= tglSeblm Then
            LblErrMsg = DisplayMsg(8055) & " " & _
                Format(Left(tglSeblm, 4) & "-" & CInt(Mid(tglSeblm, 5, 2)) & "-" & _
                Right(tglSeblm, 2), "dd mmmm yyyy")
            dtAwal.SetFocus
            Me.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    Set rsUpdate = Nothing
    '**************

    '****** cek data sesudahnya syarat end date < dr start date sesudahnya
    sql = "select min(Start_Date) as tglSsdh from BOM_Master " & _
        "where " & kondisi & "and " & _
        "Start_Date  > '" & gridTglAwal & "'"
    Set rsUpdate = Db.Execute(sql)
    If IsNull(rsUpdate("tglSsdh")) = False Then
        tglSesdh = CDate(Left(rsUpdate("tglSsdh"), 4) & "-" & Mid(rsUpdate("tglSsdh"), 5, 2) & "-" & Right(rsUpdate("tglSsdh"), 2))

        If IIf(IsDate(dtAkhir), Format(dtAkhir, "yyyyMMdd"), "99999999") >= Format(tglSesdh, "yyyyMMdd") Then
            LblErrMsg = DisplayMsg(8054) & " " & Format(tglSesdh, "dd MMM yyyy")
            dtAkhir1.SetFocus
            Me.MousePointer = vbDefault
            Exit Sub
        Else
            '****** Jika ada jarak yg kosong antara end date diinput dan start date sesdhnya
            If DateDiff("d", Format(dtAkhir, "yyyy-MM-dd"), tglSesdh) > 1 Then
                tanya = MsgBox("Data from " & Format(DateAdd("d", 1, Format(dtAkhir, "yyyy-MM-dd")), "dd MMM yyyy") & _
                    " until " & Format(DateAdd("d", -1, tglSesdh), "dd MMM yyyy") & _
                    " will have no BOM." & vbCrLf & " Are you sure want to update this data?", _
                    vbQuestion + vbOKCancel, "Confirmation")
                If tanya <> vbOK Then
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
    End If
    Set rsUpdate = Nothing

    '**********Update data yg ingin diUpdate *********
    sql = "update BOM_Master " & _
        "set qty =" & CDbl(txtQty) & ", " & _
        "Unit_Cls = '" & Trim(cbo(2)) & "', " & _
        "R_Qty = '" & CDbl(IIf(txtRQty.Text = "", 0, txtRQty.Text)) & "', " & _
        "Start_Date = '" & tglAwal & "', " & _
        "End_Date = '" & tglAkhir & "', " & _
        "Last_Update = getdate() , Last_User = '" & userLogin & "', " & _
        "Revision_No = '" & txtRevision & "' " & _
        "where " & kondisi & " and Start_Date = '" & gridTglAwal & "'"
    Db.Execute sql

    '**** Update end date sebelomnya menjadi start date sesdh yg dihapus
    sql = "update BOM_Master " & _
        "set End_Date = '" & Format(DateAdd("d", -1, dtAwal), "yyyyMMdd") & "', " & _
        "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
        "where " & kondisi & "and Start_Date =" & _
        "(select max(Start_Date) from BOM_Master where " & kondisi & " and start_Date < '" & tglAwal & "')"
    Db.Execute sql

    Call IsiGrid
    Call Kosong(1)
    LblErrMsg = DisplayMsg(1101)

End Sub

Sub hapusData()
Dim tanya
Dim rsUpdate As New Recordset

    With grid
        For i = 1 To .Rows - 1
            '********* Hapus ***********
            If .TextMatrix(i, bteColSelect) = "D" Then
                    kondisi = "Parent_ItemCode ='" & parentItem & "' and " & _
                            "Item_Code ='" & .TextMatrix(i, bteColProdCode) & "' "

                    If IsEmpty(tanya) Then tanya = MsgBox("Do you really want to Delete this Data?", vbQuestion & vbYesNo, "Confirmation")
                    If tanya = vbYes Then
                        '********Hapus data
                        sql = "delete BOM_Master where " & _
                            kondisi & "and " & _
                            "Start_Date ='" & Format(.TextMatrix(i, bteColDateStart), "yyyyMMdd") & _
                            "' and End_Date ='" & _
                            IIf(IsDate(.TextMatrix(i, bteColDateEnd)), Format(.TextMatrix(i, bteColDateEnd), "yyyyMMdd"), "99999999") & "'"
                        Db.Execute sql
                        '********

                        '***** cek data min data sesudahnya agar dpt ambil start date -1 utk update end date data seblmnya
                        sql = "select min(Start_Date) as tglSsdh from BOM_Master " & _
                            "where " & kondisi & "and " & _
                            "Start_Date  > '" & Format(.TextMatrix(i, bteColDateStart), "yyyyMMdd") & "'"
                        Set rsUpdate = Db.Execute(sql)
                        If IsNull(rsUpdate("tglSsdh")) = False Then
                            tglSesdh = CDate(Left(rsUpdate("tglSsdh"), 4) & "-" & Mid(rsUpdate("tglSsdh"), 5, 2) & "-" & Right(rsUpdate("tglSsdh"), 2))
                            tglAkhir = Format(DateAdd("d", -1, tglSesdh), "yyyyMMdd")
                        Else
                            tglAkhir = "99999999"
                        End If
                        Set rsUpdate = Nothing
                        '********

                        '**** Update end date sebelomnya menjadi start date sesdh yg dihapus
                        sql = "update BOM_Master " & _
                            "set End_Date = '" & tglAkhir & "', " & _
                            "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                            "where " & kondisi & "and Start_Date =" & _
                            "(select max(Start_Date) from BOM_Master where " & kondisi & "and start_Date < '" & Format(.TextMatrix(i, bteColDateStart), "yyyyMMdd") & "')"
                        Db.Execute sql
                        '******************
                        
                        '**** Update end date sebelomnya menjadi start date sesdh yg dihapus
                        sql = "update bom_master set doc_no = '" & txtDocNo & "' where parent_itemcode = '" & parentItem & "' "
                        Db.Execute sql
                        '******************
                    Else
                        Me.MousePointer = vbDefault
                        Exit Sub
                    End If
                'End If
            End If
        Next i
    End With

    Set rsUpdate = Nothing
    Call IsiGrid
    Call Kosong(1)
    LblErrMsg = DisplayMsg(1201)
'****************************
End Sub

Private Sub cmdPage_Click(Index As Integer)
With grid
Select Case Index
    Case 0: 'First
        .TopRow = 1
    Case 1: 'Prev
        .TopRow = IIf((.TopRow = 1), 1, .TopRow - 5)
    Case 2: 'Next
        .TopRow = IIf((.TopRow = .Rows), .Rows, .TopRow + 5)
    Case 3: 'Last
        .TopRow = .Rows
End Select
End With
End Sub

Private Sub cmdReport_Click()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim Rpt As New FrmRpt3

    Me.MousePointer = vbHourglass
    sql = "select a.*, substring(start_Date,5,2) as blnAwal," & _
          "substring(end_Date,5,2) as blnEnd," & _
          "b.MakerItem_Code as makerParent," & _
          "c.MakerItem_Code as makerItem " & _
          "from BOM_Master a " & _
          "left Outer JOIN Item_Master b " & _
          "ON a.parent_ItemCode  = b.Item_Code " & _
          "Left Outer JOIN Item_Master c " & _
          "ON a.Item_Code = c.Item_Code " & _
          "order by Parent_ItemCode, a.Item_Code"
    Set rsRpt = Db.Execute(sql)
    
    
    If rsRpt.EOF Then
        LblErrMsg.Caption = DisplayMsg(4002)
        cbo(0).SetFocus
    Else
        sqlprint = sql
        reportcode = "BOM"
        printorient = 2
        Set report = application.OpenReport(App.path & "\Reports\BOM.rpt")
        report.Database.Tables(1).SetDataSource rsRpt
        
        sql = "select Company_Name,Address1,Address2,City," & _
            "Postal_Code,Phone1,Fax from company_Profile"
        Set rsCompany = Db.Execute(sql)
        If Not (rsCompany.EOF) Then
            report.FormulaFields(bteColUnit).Text = "'" & Trim(rsCompany(0)) & "'"
            report.FormulaFields(7).Text = "'" & Trim(rsCompany(1)) & ", " & _
                Trim(rsCompany(2)) & ". " & Trim(rsCompany(3)) & " Indonesia'"
            report.FormulaFields(8).Text = "'Phone : " & Trim(rsCompany(4)) & " " & _
                                "Fax : " & Trim(rsCompany(5)) & "'"
        End If
        Set rsCompany = Nothing
        
        Rpt.CRViewer1.ReportSource = report
        Rpt.CRViewer1.ViewReport
        Rpt.CRViewer1.Zoom 1
        
        Rpt.WindowState = 2
        Rpt.Show 1
    End If
    Set rsRpt = Nothing
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
        And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
        Then KeyAscii = 0
End Sub

Private Sub txtQty_LostFocus()
    If IsNumeric(txtQty) = False Then txtQty = 0
    If CDec(txtQty) > gd_MaxQty Then txtQty = gd_MaxQty
    txtQty = Format(CDec(txtQty), gs_formatQtyBOM)
End Sub

Private Sub txtqty_GotFocus()
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty)
End Sub

Private Sub txtRQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
        And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
        Then KeyAscii = 0
End Sub

Private Sub txtRQty_LostFocus()
    If IsNumeric(txtRQty) = False Then txtRQty = 0
    If CDec(txtRQty) > gd_MaxQty Then txtRQty = gd_MaxQty
    txtRQty = Format(CDec(txtRQty), gs_formatQtyBOM)
End Sub

Private Sub txtRQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub txtRQty_GotFocus()
    txtRQty.SelStart = 0
    txtRQty.SelLength = Len(txtRQty)
End Sub

Private Sub dtAwal_Change()
    LblErrMsg = ""
    If IsDate(dtAkhir) Then
        If Format(dtAwal, "yyyy-MM-dd") > _
            Format(CDate(dtAkhir), "yyyy-MM-dd") Then _
        LblErrMsg = DisplayMsg(4068): Exit Sub
    End If
End Sub

Private Sub dtAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub dtAkhir1_Change()
    LblErrMsg = ""
    dtAkhir = Format(dtAkhir1, "MM/dd/yyyy")
    If Format(dtAwal, "yyyy-MM-dd") > _
        Format(CDate(dtAkhir), "yyyy-MM-dd") Then _
        LblErrMsg = DisplayMsg(4066): Exit Sub
End Sub

Private Sub dtAkhir1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

'************ Unload **********
Private Sub CmdSubMenu_Click(Index As Integer)
Select Case Index
    Case 0:
        DoEvents
        frmMainMenu.Show
        DoEvents
        Unload Me
    Case 1: 'BOM inquiry
        If hakAkses("frmBOMInquiry") = 0 Then LblErrMsg = DisplayMsg(3007)
        
        If lblNm(3).Caption <> "" Then
            frmBOMInquiry.cbo(0) = cbo(0)
            frmBOMInquiry.parent = Trim(cbo(0).Column(0))
        End If
        
        If grid.Rows > 1 Then
         frmBOMInquiry.command2.Enabled = False
        Else
         frmBOMInquiry.command2.Enabled = True
        End If
        
        frmBOMInquiry.cmdsubmenu(0).Caption = "&Back"
        frmBOMInquiry.Show
        Unload Me
End Select
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
