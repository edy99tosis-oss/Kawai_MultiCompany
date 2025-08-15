VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPajak_Create_New 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faktur Pajak Create"
   ClientHeight    =   11040
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   15165
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPajak_Create_New.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15165
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdInv 
      BackColor       =   &H0080FFFF&
      Caption         =   "To Inv. (Inquiry)"
      Height          =   375
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10317
      Width           =   1635
   End
   Begin VB.Frame FrameAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   4815
      TabIndex        =   51
      Top             =   8580
      Width           =   10170
      Begin VB.TextBox TxtRec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   600
         Width           =   1900
      End
      Begin VB.TextBox TxtRec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   600
         Width           =   1900
      End
      Begin VB.TextBox TxtRec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   600
         Width           =   1900
      End
      Begin VB.TextBox txtAmountIDR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6195
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   600
         Width           =   1900
      End
      Begin VB.TextBox txtPPNIDR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8175
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   600
         Width           =   1900
      End
      Begin VB.Shape Shape2 
         Height          =   465
         Left            =   0
         Top             =   510
         Width           =   6045
      End
      Begin VB.Label LblRec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
         Height          =   195
         Index           =   2
         Left            =   4050
         TabIndex        =   61
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label LblRec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPN"
         Height          =   195
         Index           =   1
         Left            =   2070
         TabIndex        =   60
         Top             =   210
         Width           =   330
      End
      Begin VB.Label LblRec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   59
         Top             =   210
         Width           =   1140
      End
      Begin VB.Shape Shape3 
         Height          =   465
         Left            =   6105
         Top             =   510
         Width           =   4065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount (IDR)"
         Height          =   195
         Left            =   6195
         TabIndex        =   58
         Top             =   210
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPN (IDR)"
         Height          =   195
         Left            =   8175
         TabIndex        =   57
         Top             =   210
         Width           =   870
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   435
         Left            =   0
         Top             =   90
         Width           =   6045
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   435
         Left            =   6105
         Top             =   90
         Width           =   4065
      End
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   8832
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10317
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Delete"
      Height          =   375
      Left            =   10074
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10317
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1215
      Left            =   165
      TabIndex        =   44
      Top             =   1047
      Width           =   14835
      Begin VB.TextBox LblPajak 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   1
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   315
         Width           =   3435
      End
      Begin VB.TextBox LblPajak 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   2
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   315
         Width           =   5955
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   705
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyy"
         Format          =   293666819
         CurrentDate     =   37860
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   315
         Left            =   4275
         TabIndex        =   2
         Top             =   705
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   293666819
         CurrentDate     =   37860
      End
      Begin MSForms.ComboBox cboInvoice 
         Height          =   315
         Left            =   8460
         TabIndex        =   3
         Top             =   705
         Width           =   3045
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "5371;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
         Height          =   195
         Index           =   1
         Left            =   7320
         TabIndex        =   50
         Top             =   765
         Width           =   975
      End
      Begin VB.Line Line2 
         X1              =   8460
         X2              =   14400
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line1 
         X1              =   3750
         X2              =   7170
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label LblPajakt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   195
         Index           =   4
         Left            =   3870
         TabIndex        =   49
         Top             =   765
         Width           =   165
      End
      Begin VB.Label LblPajaki 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
         Height          =   195
         Index           =   3
         Left            =   450
         TabIndex        =   48
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label LblPajak2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   2
         Left            =   7605
         TabIndex        =   47
         Top             =   330
         Width           =   690
      End
      Begin VB.Label LblPajakc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   46
         Top             =   330
         Width           =   1350
      End
      Begin MSForms.ComboBox CboCust 
         Height          =   315
         Left            =   2040
         TabIndex        =   0
         Top             =   270
         Width           =   1605
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2831;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblfix 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Fix"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   13410
         TabIndex        =   45
         Top             =   765
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
      Height          =   375
      Left            =   12558
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10317
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Left            =   11316
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10317
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   165
      TabIndex        =   26
      Top             =   9567
      Width           =   14835
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
         Height          =   240
         Left            =   105
         TabIndex        =   27
         Top             =   195
         Width           =   14610
      End
   End
   Begin VB.CommandButton CmdSubmit 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10317
      Width           =   1185
   End
   Begin VB.CommandButton CmdMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10317
      Width           =   1185
   End
   Begin VB.CommandButton CmdData 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last Page"
      Height          =   375
      Index           =   3
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   10317
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdData 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Next Page"
      Height          =   375
      Index           =   2
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   10317
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdData 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Prev Page"
      Height          =   375
      Index           =   1
      Left            =   3195
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   10317
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdData 
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Page"
      Height          =   375
      Index           =   0
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   10317
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ComboBox Cbo1 
      Height          =   315
      ItemData        =   "FrmPajak_Create_New.frx":0E42
      Left            =   210
      List            =   "FrmPajak_Create_New.frx":0E4C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2487
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3795
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2487
      Width           =   930
   End
   Begin VB.CommandButton CmdCreate 
      BackColor       =   &H0080FFFF&
      Caption         =   "Create"
      Height          =   375
      Left            =   8745
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2457
      Width           =   1200
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FDDFE3&
      Height          =   315
      Left            =   915
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8712
      Width           =   1410
   End
   Begin VB.TextBox txtstart 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3345
      MaxLength       =   10
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   8712
      Width           =   1410
   End
   Begin MSComCtl2.DTPicker TglPajak 
      Height          =   315
      Left            =   6870
      TabIndex        =   7
      Top             =   2490
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd MMM yyy"
      Format          =   293666819
      CurrentDate     =   37860
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13140
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   349
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5505
      Left            =   165
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3057
      Width           =   14835
      _cx             =   26167
      _cy             =   9710
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
   Begin MSForms.ComboBox cboCode 
      Height          =   315
      Left            =   2940
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2490
      Width           =   855
      VariousPropertyBits=   746604571
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "1508;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblUSD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   12270
      TabIndex        =   42
      Top             =   2547
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblTHB 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   14805
      TabIndex        =   41
      Top             =   2547
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblSD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   14805
      TabIndex        =   40
      Top             =   2322
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblEUR 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   12270
      TabIndex        =   39
      Top             =   2787
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblYen 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   12270
      TabIndex        =   38
      Top             =   2322
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblUSD1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "US$ : IDR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   10425
      TabIndex        =   37
      Top             =   2547
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblTHB1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "THB : IDR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   12885
      TabIndex        =   36
      Top             =   2547
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblSD1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S$ : IDR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   13020
      TabIndex        =   35
      Top             =   2322
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblEur1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EUR : IDR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   10425
      TabIndex        =   34
      Top             =   2787
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblYen1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "YEN : IDR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   10425
      TabIndex        =   33
      Top             =   2322
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Faktur Pajak Create"
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
      Left            =   165
      TabIndex        =   32
      Top             =   364
      Width           =   14835
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   315
      Left            =   4725
      TabIndex        =   6
      Top             =   2490
      Width           =   1380
      VariousPropertyBits=   746604571
      MaxLength       =   7
      DisplayStyle    =   3
      Size            =   "2434;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Faktur Pajak No."
      Height          =   195
      Index           =   0
      Left            =   1425
      TabIndex        =   31
      Top             =   2550
      Width           =   1425
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   165
      TabIndex        =   30
      Top             =   8775
      Width           =   690
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start From"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   3
      Left            =   2370
      TabIndex        =   29
      Top             =   8775
      Width           =   915
   End
   Begin VB.Label LblPajakd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Index           =   5
      Left            =   6345
      TabIndex        =   28
      Top             =   2550
      Width           =   405
   End
End
Attribute VB_Name = "FrmPajak_Create_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsPM As Recordset, RsPD As Recordset, RsI As Recordset, RsNF As Recordset, RsP As Recordset
Dim SqlPM As String, SqlPD As String, SqlI As String, SqlNF As String, sqlP As String, tempInvoice As String
Dim NoStart As Long, StartNo As Long
Dim HakU As Integer
Dim StFix As Boolean
Dim currConv As Double

Dim bteColSelect As Byte
Dim bteColInvNo As Byte
Dim bteColInvDate As Byte
Dim bteColProdCode As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColDate As Byte
Dim bteColQty As Byte
Dim bteColUnit As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColPPn As Byte
Dim bteColTAmount As Byte
Dim bteColDONo As Byte
Dim bteColPONo As Byte
Dim bteColSeqNo As Byte
Dim bteColDOSeqNo As Byte

Dim bteHakPrice As Byte
Function GetNumber(sql)
Dim rr As New ADODB.Recordset
Set rr = Nothing
rr.Open sql, Db, adOpenDynamic, adLockBatchOptimistic
If rr.EOF Then
GetNumber = rr.Fields(0)
Else
GetNumber = 1
End If
rr.Close
End Function
Sub Header()
        
    bteColSelect = 0
    bteColInvNo = 1
    bteColInvDate = 2
    bteColProdCode = 3
    bteColPartNo = 4
    bteColDesc = 5
    bteColDate = 6
    bteColQty = 7
    bteColUnit = 8
    bteColCurr = 9
    bteColPrice = 10
    bteColAmount = 11
    bteColPPn = 12
    bteColTAmount = 13
    bteColDONo = 14
    bteColPONo = 15
    bteColSeqNo = 16
    bteColDOSeqNo = 17
    
    With grid
        
        .ColS = 18
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColInvNo) = "Invoice No"
        .TextMatrix(0, bteColInvDate) = "Invoice Date"
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColPartNo) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColDate) = "Delivery Date"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColPPn) = "PPN"
        .TextMatrix(0, bteColTAmount) = "T Amount"
        .TextMatrix(0, bteColDONo) = "DO No"
        .TextMatrix(0, bteColPONo) = "PO_No"
        .TextMatrix(0, bteColSeqNo) = "SeqNo"
        .TextMatrix(0, bteColDOSeqNo) = "DOSeqNo"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColInvNo) = 2000
        .ColWidth(bteColInvDate) = 1400
        .ColWidth(bteColPartNo) = 1700
        .ColWidth(bteColDesc) = 2500
        .ColWidth(bteColDate) = 1400
        .ColWidth(bteColQty) = 1000
        .ColWidth(bteColUnit) = 500
        .ColWidth(bteColCurr) = 500
        .ColWidth(bteColPrice) = 1700
        .ColWidth(bteColAmount) = 1700
        
        .ColAlignment(bteColInvNo) = flexAlignLeftCenter
        .ColAlignment(bteColInvDate) = flexAlignLeftCenter
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColPartNo) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColDate) = flexAlignCenterCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColUnit) = flexAlignLeftCenter
        .ColAlignment(bteColCurr) = flexAlignLeftCenter
        .ColAlignment(bteColPrice) = flexAlignRightCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
        
        .ColHidden(bteColSelect) = True
        .ColHidden(bteColProdCode) = True
        .ColHidden(bteColPPn) = True
        .ColHidden(bteColTAmount) = True
        .ColHidden(bteColDONo) = True
        .ColHidden(bteColPONo) = True
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColDOSeqNo) = True
        
        .ColHidden(bteColCurr) = (bteHakPrice = 0)
        .ColHidden(bteColPrice) = (bteHakPrice = 0)
        .ColHidden(bteColAmount) = (bteHakPrice = 0)
        
        .EditMaxLength = 1
    
    End With
    
End Sub

Private Sub SetPajakCode()
    
    With cboCode
        
        .clear
        .columnCount = 2
        .ColumnWidths = "20pt;310pt"
        .ListWidth = "330pt"
        .ListRows = 9
        
        .AddItem ""
        .List(.ListCount - 1, 0) = "01"
        .List(.ListCount - 1, 1) = "Kepada Selain Pemungut PPN"
        .AddItem ""
        .List(.ListCount - 1, 0) = "02"
        .List(.ListCount - 1, 1) = "Kepada Pemungut Bendaharawan"
        .AddItem ""
        .List(.ListCount - 1, 0) = "03"
        .List(.ListCount - 1, 1) = "Kepada Pemungut PPN Lainnya"
        .AddItem ""
        .List(.ListCount - 1, 0) = "04"
        .List(.ListCount - 1, 1) = "Yang Menggunakan DPP Nilai Lain Kepada Selain Pemungut PPN"
        .AddItem ""
        .List(.ListCount - 1, 0) = "05"
        .List(.ListCount - 1, 1) = "Yang PM-nya di-deemed Kepada Selain Pemungut PPN"
        .AddItem ""
        .List(.ListCount - 1, 0) = "06"
        .List(.ListCount - 1, 1) = "Penyerahan Lainnya Kepada Selain Pemungut PPN"
        .AddItem ""
        .List(.ListCount - 1, 0) = "07"
        .List(.ListCount - 1, 1) = "Yang PPN-nya TDP Kepada Selain Pemungut PPN"
        .AddItem ""
        .List(.ListCount - 1, 0) = "08"
        .List(.ListCount - 1, 1) = "Yang Dibebaskan Dari Pengenaan PPN Kepada Selain Pemungut PPN"
        .AddItem ""
        .List(.ListCount - 1, 0) = "09"
        .List(.ListCount - 1, 1) = "Penyerahan Aktiva Pasal 16 D Kepada Selain Pemungut PPN"
        
        .ListIndex = 0
        
    End With
    
End Sub

Private Sub Cbo1_Click()
    
    Dim strFakturNo As String
    
    If Cbo1.Text = "Create" Then
        
        RsP.Requery
        RsP.filter = "TC='" & cboCust & "'"
        
        If Not RsP.EOF Then
            
            strFakturNo = GetFakturNo
            If strFakturNo <> "" Then nomorinvoice
            ComboBox1.locked = False
            
            sql = "Delete From FakturPajak_Master Where FakturPajak_No Not In (Select FakturPajak_No From FakturPajak_Detail)"
            Db.Execute sql
            
            RsNF.Requery
            RsNF.filter = " no >= '" & Format(txtstart, "00000000") & "' and Pajak_Year = " & TglPajak.Year
            If Not RsNF.EOF Then
                Text2.Text = Format(RsNF!NO, "00000000")
                ComboBox1 = Format(RsNF!NO + 1, "00000000")
            Else
                If NoStart <> Val(txtstart) Then
                    ComboBox1 = Format(txtstart, "00000000")
                Else
                    If Val(txtstart) <> 0 Then
                        If NoStart = Val(txtstart) Then
                            ComboBox1 = Format(txtstart, "00000000")
                        Else
                            ComboBox1 = Format(ComboBox1 + 1, "00000000")
                        End If
                    Else
                            ComboBox1 = "00000001"
                    End If
                End If
                NoStart = txtstart
            End If
            RsNF.Requery
            RsNF.filter = ""
            Header
            grid.Editable = flexEDNone
            ComboBox1.locked = True
            Kosong 2
            TxtRec(0) = 0
            TxtRec(1) = 0
            TxtRec(2) = 0
            
        End If
        
        CmdCreate.Caption = "Create"
        cboCust.locked = False
        ComboBox1.locked = True
        cmdSubmit.Enabled = False
        RecBtn False 'untuk meng-enable/non enable cmdData
        
    Else
    
        Header
        CmdCreate.Caption = "Update"
        grid.Editable = flexEDKbdMouse
        If ComboBox1 <> "" Then If CekNoPajak(cboCode & Text1 & "." & Trim(ComboBox1)) = False Then ComboBox1 = "": Kosong 2: NomorFaktur
        ComboBox1.locked = False
        cmdSubmit.Enabled = True
        RecBtn True
    
    End If
    
End Sub

Private Sub cboCode_Change()
    
    If Not cboCode.MatchFound Then cboCode.ListIndex = 0
    
End Sub

Private Sub CboCust_Change()
    
    cboCust_Click

End Sub

Private Sub cboCust_Click()

    '# Tampilkan Detil dari Combo Customer code
    nomorinvoice
    RsP.Requery
    RsP.filter = "TC='" & cboCust & "'"
    LblPajak(1) = ""
    LblPajak(2) = "Address :"
    
    If Not RsP.EOF Then LblPajak(1) = "" & RsP!TN: LblPajak(2) = "Address : " & RsP!a
        
    LblErr = ""
    If Cbo1.Text = "Update" Then
        
        Header
        NomorFaktur
        LblErr = ""
        Kosong 2
        cboCust.locked = False
    
    Else
        
        RsP.Requery
        RsP.filter = "TC='" & cboCust & "'"
        If Not RsP.EOF Then
            ComboBox1.locked = False
            RsNF.Requery
            RsNF.filter = " no >= '" & Format(txtstart, "00000000") & "' "
            If Not RsNF.EOF Then
                Text2.Text = Format(RsNF!NO, "00000000")
                ComboBox1 = Format(RsNF!NO + 1, "00000000")
            Else
                If NoStart <> Val(txtstart) Then
                    ComboBox1 = Format(txtstart, "00000000")
                Else
                    If Val(txtstart) <> 0 Then
                        If NoStart = Val(txtstart) Then ComboBox1 = Format(txtstart, "00000000") Else ComboBox1 = Format(ComboBox1 + 1, "00000000")
                    Else
                        ComboBox1 = "00000001"
                    End If
                End If
                NoStart = txtstart
            End If
            RsNF.Requery
            RsNF.filter = ""
        End If
        LblErr = ""
        ComboBox1.locked = True
        Kosong 2
    
    End If
    
    MousePointer = vbDefault
        
End Sub

Private Sub cbocust_KeyPress(KeyAscii As MSForms.ReturnInteger)
    
    If KeyAscii = Asc("'") Then KeyAscii = 0

End Sub

Private Sub cboInvoice_Click()
    
    Dim strFakturNo As String
    
    LblErr = ""
    strFakturNo = GetFakturNo
    If strFakturNo <> "" Then
        ComboBox1.Text = strFakturNo
        Cbo1.ListIndex = 1
    Else
        Cbo1.ListIndex = 0
    End If
    Header
    
End Sub

Private Sub cmdCancel_Click()
    
    nomorinvoice
    InvoiceDisplay
    
End Sub

Private Sub cmdClear_Click()
    
    Header
    cboCust = ""
    ComboBox1 = ""
    Tgl1 = Date
    Tgl2 = Date
    TglPajak = Date
    TxtRec(0) = 0
    TxtRec(1) = 0
    TxtRec(2) = 0
    lblfix.Visible = False
    hitungCurr

End Sub

Private Sub CmdCreate_Click()
    
    Dim RsCF As Recordset
    
    MousePointer = vbHourglass
    
    If HakU = 0 Then LblErr = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    If Text1 = "" Then LblErr = DisplayMsg("0051"): Exit Sub
    
    If Cbo1.Text = "Create" Then
                
        If cboInvoice = "" Then LblErr = DisplayMsg("0050"): cboInvoice.SetFocus: Me.MousePointer = vbDefault: Exit Sub
        
        cboCust = cboCust
        If cboCust.MatchFound = False Then
            LblErr = DisplayMsg(4072)
            cboCust.SetFocus
            MousePointer = vbDefault
            Exit Sub
        ElseIf Trim(ComboBox1.Text) = "" Then
            MousePointer = vbDefault
            Exit Sub
        End If
        
        savemaster
        inquiry
        CmdCreate.Caption = "Update"
        NomorFaktur
        Cbo1.ListIndex = 1
        RsNF.Requery
        Text2.Text = Format(RsNF!NO, "00000000")
        StFix = False
        lblfix.Visible = False
        InvoiceDisplay
        cmdSubmit.Enabled = True
        RecBtn (True)
    
    Else
    
        cboCust = Trim(cboCust)
        If cboCust.MatchFound = False Then
            LblErr = DisplayMsg(4072)
            cboCust.SetFocus
            MousePointer = vbDefault: Exit Sub
        ElseIf Trim(ComboBox1.Text) = "" Then
            MousePointer = vbDefault: Exit Sub
        End If
    
        LblErr = ""
        If Trim(ComboBox1.Text) <> "" Then
            
            RsPM.Requery
            RsPM.filter = "fakturpajak_no = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "'"
            If Not RsPM.EOF Then
                
                Set RsCF = Db.Execute("select * from fakturpajak_master where (fix_cls='0' or fix_cls is null) and fakturpajak_no='" & cboCode & Text1 & "." & Trim(ComboBox1) & "'")
                If Not RsCF.EOF Then
                    
                    lblfix.Visible = False
                    inquiry
                    updateMaster
                    LblErr.Caption = DisplayMsg(1101)
                    StFix = False
                    grid.Editable = flexEDKbdMouse
                    cmdSubmit.Enabled = True
                
                Else
                    
                    LblErr = DisplayMsg("0052"): inquiry: InvoiceDisplay: MousePointer = vbDefault
                    lblfix.Visible = True
                    StFix = True
                    cmdSubmit.Enabled = False
                    grid.Editable = flexEDNone
                    Exit Sub
                
                End If
                
                InvoiceDisplay
                
            Else
            
                LblErr = DisplayMsg("0053")
            
            End If
            
            RsPM.Requery
            RsPM.filter = ""
        
        End If
    
    End If
    
    cmdSubmit.Enabled = True
    MousePointer = vbDefault

End Sub

Private Sub cmdDelete_Click()
    
    Dim Sqlc As String
    MousePointer = vbHourglass
    
    If Cbo1.Text = "Create" Or ComboBox1.Text = "" Then Me.MousePointer = vbDefault: Exit Sub
    If HakU = 0 Then LblErr = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    If MsgBox("Are you sure want to delete this data ?", vbQuestion + vbYesNo, "Delete Confirmation") = vbYes Then
    
        ComboBox1 = Trim(ComboBox1)
        If ComboBox1.MatchFound = False Then LblErr = DisplayMsg("0054"): MousePointer = vbDefault: Exit Sub
        If StFix = False Then
            
            sql = "Delete From FakturPajak_Detail Where FakturPajak_No = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "'"
            Db.Execute sql
            sql = "Delete From FakturPajak_Master Where FakturPajak_No = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "'"
            Db.Execute sql

            ComboBox1.RemoveItem ComboBox1.ListIndex
            ComboBox1.Text = ""
            
            Header
            hitungCurr
            
            TxtRec(0) = 0
            TxtRec(1) = 0
            TxtRec(2) = 0
            
            RsNF.Requery
            RsNF.filter = ""
            If Not RsNF.EOF Then
                If IsNull(RsNF!NO) Then
                    Text2.Text = "00000000"
                    txtstart.Text = "00000001"
                Else
                    Text2.Text = Format(RsNF!NO, "00000000")
                    txtstart = Format(RsNF!NO + 1, "00000000")
                End If
            Else
                Text2.Text = "00000000"
                txtstart.Text = "00000001"
            End If
            
            lblfix.Visible = False
            LblErr = DisplayMsg(1201)
            
        Else
            LblErr = DisplayMsg("0052"): inquiry: InvoiceDisplay: MousePointer = vbDefault: Exit Sub
            lblfix.Visible = True
        End If
        
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub CmdInv_Click()

    If grid.Row <= 0 Then Exit Sub
    If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
    Me.MousePointer = vbHourglass
    With frm_invoice_inquiry
        
        .cmd_sub_menu.Caption = "&Back"
        .cmd_sub_menu.Tag = "pajak_new"
        .ComboBox1 = cboCust
        Call .ComboBox1_Click
        
        .DTPicker3 = grid.TextMatrix(grid.Row, bteColInvDate)
        .DTPicker4 = grid.TextMatrix(grid.Row, bteColInvDate)
        .combo1 = grid.TextMatrix(grid.Row, bteColInvNo)
        .CmdSearch(0).Value = True
        Call .cmdSearch_Click(0)
        
        .lbl_pesan = ""
        .Show
         Me.Hide
         
    End With
    Me.MousePointer = vbDefault

End Sub

Private Sub CmdMenu_Click()

    sql = "Delete From FakturPajak_Master Where FakturPajak_No Not In (Select FakturPajak_No From FakturPajak_Detail)"
    Db.Execute sql

    '# tutup recordset
    RsNF.Close
    Set RsNF = Nothing
    RsPM.Close
    Set RsPM = Nothing
    RsPD.Close
    Set RsPD = Nothing
    RsP.Close
    Set RsP = Nothing

    frmMainMenu.Show
    Unload Me
    DoEvents
    
End Sub

Private Sub CmdPreview_Click()

    MousePointer = vbHourglass
    
    If Trim(ComboBox1) <> "" Then
        RsPM.Requery
        RsPM.filter = "fakturpajak_no ='" & cboCode & Text1 & "." & Trim(ComboBox1) & "'"
        If Not RsPM.EOF Then
            pajak_No = cboCode & Text1 & "." & Trim(ComboBox1)
            ViewReport (cboCode & Text1 & "." & Trim(ComboBox1))
        Else
            LblErr = DisplayMsg("0053")
        End If
        RsPM.Requery
        RsPM.filter = ""
    End If
    
    MousePointer = vbDefault

End Sub

Private Sub CmdSubmit_Click()

    Dim Sqlc As String
    MousePointer = vbHourglass
    
    If HakU = 0 Then LblErr = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    ComboBox1 = Trim(ComboBox1)
    If ComboBox1.MatchFound = False Then LblErr = DisplayMsg("0054"): MousePointer = vbDefault: Exit Sub
    If StFix = False Then
        lblfix.Visible = False
        InvoiceDisplay
        savedetail
        inquiry
        InvoiceDisplay
    Else
        LblErr = DisplayMsg("0052"): inquiry: InvoiceDisplay: MousePointer = vbDefault: Exit Sub
        lblfix.Visible = True
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub ComboBox1_Change()
        
    cboCust = cboCust
    
    If cboCust.MatchFound Then
        RsP.Requery
        RsP.filter = "TC='" & cboCust & "'"
        Header
        LblErr = ""
        If RsP.EOF Then RsP.Requery: Exit Sub
        ComboBox1.SelStart = 0
        If Cbo1.Text = "Create" Then
            MousePointer = vbHourglass
            TglPajak = Date
            hitungCurr
            MousePointer = vbDefault
        Else
            If Trim(ComboBox1.Text) = "" Then Exit Sub
            MousePointer = vbHourglass
            cboInvoice.Text = GetInvoiceNo
            If FPajakDate(cboCode & Text1 & "." & Trim(ComboBox1)) <> "" Then
                TglPajak = FPajakDate(cboCode & Text1 & "." & Trim(ComboBox1))
            Else
                TglPajak = Date
            End If
            hitungCurr
            MousePointer = vbDefault
        End If
    End If
        
End Sub

Private Sub ComboBox1_Click()
    
    ComboBox1.SelStart = 0
    ComboBox1_Change

End Sub

Private Sub Combobox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
    
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0

End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    
    If ErrMsg = "" Then Unload Me Else LblErr.Caption = ErrMsg

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then Cancel = 1
    
End Sub

Private Sub CmdData_Click(Index As Integer)

    If RsPM.RecordCount = 0 Then Exit Sub

    Select Case Index
    Case 0
        RsPM.MoveFirst
        LblErr = DisplayMsg(4020)
    Case 1
        If Not RsPM.BOF And RsPM.AbsolutePosition > 1 Then
            RsPM.MovePrevious
            LblErr = ""
        Else
            RsPM.MoveFirst
            LblErr = DisplayMsg(4020)
        End If
    Case 2
        If Not RsPM.EOF And RsPM.AbsolutePosition < RsPM.RecordCount Then
            RsPM.MoveNext
            LblErr = ""
        Else
            RsPM.MoveLast
            LblErr = DisplayMsg(4021)
        End If
    Case 3
        RsPM.MoveLast
        LblErr = DisplayMsg(4021)
    End Select
    
    ComboBox1.Text = Right(Trim(RsPM!fakturpajak_no), 8)
    cboCust = RsPM!Cust_CodE
    TglPajak = Format(RsPM!fakturpajak_date, "dd mmm yyyy")
    hitungCurr
    InvoiceDisplay
    
    sql = "Delete From FakturPajak_Master Where FakturPajak_No Not In (Select FakturPajak_No From FakturPajak_Detail)"
    Db.Execute sql
    
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    Dim ir As Long
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    HakU = hakUpdate(Me.Name)
    
    bteHakPrice = hakPrice(Me.Name)
    FrameAmount.Visible = (bteHakPrice = 1)
    Header
    SetPajakCode
    
    '## Faktur Pajak Master
    SqlPM = "select * from fakturPajak_master"
    Set RsPM = New Recordset
    RsPM.CursorLocation = adUseClient
    RsPM.Open SqlPM, Db, adOpenDynamic, adLockOptimistic
    
    '## Faktur Pajak Detil
    SqlPD = "select * from fakturPajak_Detail"
    Set RsPD = New Recordset 'Db.Execute(SqlPD)
    RsPD.Open SqlPD, Db, adOpenDynamic, adLockOptimistic
        
    '## No faktur
    SqlNF = "select year(fakturpajak_date) As Pajak_Year, max(right(rtrim(fakturpajak_no), 8)) no from fakturpajak_master Group By year(fakturpajak_date)"
    
    Set RsNF = New Recordset
    RsNF.Open SqlNF, Db, adOpenDynamic, adLockOptimistic
    
    If Not RsNF.EOF Then
        If IsNull(RsNF!NO) Then
            Text2.Text = "00000000"
            txtstart.Text = "00000001"
        Else
            Text2.Text = Format(RsNF!NO, "00000000")
            txtstart = Format(RsNF!NO + 1, "00000000")
        End If
    Else
        Text2.Text = "00000000"
        txtstart.Text = "00000001"
    End If
    
    '##Tampilkan Combo Customer code dari trade_master
    sqlP = "Select rtrim(trade_code) as TC,Trade_name as TN, Address1 as A from trade_master where trade_cls='4' and   isnull(country_cls,0)<>'1'  order by trade_code"
    Set RsP = Db.Execute(sqlP)
    cboCust.clear
    cboCust.columnCount = 2
    cboCust.TextColumn = 1
    ir = 0
    While Not RsP.EOF
        cboCust.AddItem ""
        cboCust.List(ir, 0) = RsP!TC
        cboCust.List(ir, 1) = Trim(RsP!TN)
        ir = ir + 1
        RsP.MoveNext
    Wend
    cboCust.ColumnWidths = "60 pt; 300 pt"
    cboCust.ListWidth = 360
    cboCust.ListRows = 15
    Cbo1.ListIndex = 1
    
    cmdSubmit.Enabled = False
    Kosong 1
    NoStart = 0
    
    Tgl1 = Date
    Tgl2 = Date
    TglPajak = Date
    
    Text1 = "0.000-" & Right(TglPajak.Year, 2)
    hitungCurr

End Sub

Sub hitungCurr()

    Dim sql As String
    Dim RS As New ADODB.Recordset
    
    lblYen1.Visible = False
    lblYEN.Visible = False
    lblUSD1.Visible = False
    lblUSD.Visible = False
    lblEur1.Visible = False
    lblEur.Visible = False
    lblSD1.Visible = False
    lblSD.Visible = False
    lblTHB1.Visible = False
    lblTHB.Visible = False
    
    sql = "select * from tax_exchangerate where start_date <= " & Format(TglPajak, "YYYYMMDD") & " and end_date >= " & Format(TglPajak, "YYYYMMDD")
    If RS.State = adStateOpen Then RS.Close
    RS.Open sql, Db, adOpenDynamic, adLockOptimistic
    
    Do While Not RS.EOF
        Select Case RS("currency_code")
        Case "01"
            lblYen1.Visible = True
            lblYEN.Visible = True
            lblYEN.Caption = IIf(IsNull(RS("tax_exchangerate")), "0", Format(RS("tax_exchangerate"), gs_formatExchangeRate))
        Case "02"
            lblUSD1.Visible = True
            lblUSD.Visible = True
            lblUSD.Caption = IIf(IsNull(RS("tax_exchangerate")), "0", Format(RS("tax_exchangerate"), gs_formatExchangeRate))
        Case "04"
            lblEur1.Visible = True
            lblEur.Visible = True
            lblEur.Caption = IIf(IsNull(RS("tax_exchangerate")), "0", Format(RS("tax_exchangerate"), gs_formatExchangeRate))
        Case "05"
            lblSD1.Visible = True
            lblSD.Visible = True
            lblSD.Caption = IIf(IsNull(RS("tax_exchangerate")), "0", Format(RS("tax_exchangerate"), gs_formatExchangeRate))
        Case "06"
            lblTHB1.Visible = True
            lblTHB.Visible = True
            lblTHB.Caption = IIf(IsNull(RS("tax_exchangerate")), "0", Format(RS("tax_exchangerate"), gs_formatExchangeRate))
        End Select
        RS.MoveNext
    Loop

End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    If Row = 0 Then Exit Sub
    TxtRec(0) = 0
    TxtRec(1) = 0
    TxtRec(2) = 0
            
    Dim i As Long
    For i = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, bteColSelect) = flexChecked And i <> Row Then
            If grid.TextMatrix(Row, bteColCurr) <> grid.TextMatrix(i, bteColCurr) Then
                grid.Cell(flexcpChecked, Row, bteColSelect) = flexUnchecked
                LblErr = DisplayMsg(4084)
                Exit Sub
            End If
        End If
    Next i
        
    If Trim(tempInvoice) = "" Then
        For i = 1 To grid.Rows - 1
            If grid.Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
                tempInvoice = Trim(grid.TextMatrix(i, bteColInvNo))
                LblErr = ""
                Exit For
            End If
            tempInvoice = ""
        Next
    End If
        
    If tempInvoice <> Trim(grid.TextMatrix(Row, bteColInvNo)) Then LblErr = DisplayMsg("0055"): grid.Cell(flexcpChecked, Row, bteColSelect) = flexUnchecked: Exit Sub
    tempInvoice = ""
    For i = 1 To grid.Rows - 1
         If grid.Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
             tempInvoice = Trim(grid.TextMatrix(i, bteColInvNo))
             LblErr = ""
             Exit For
         End If
     Next
    LblErr = ""
    Dim icg As Long
    For icg = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, icg, bteColSelect) = flexChecked Then
            TxtRec(0) = Format(CDbl(TxtRec(0)) + CDbl(grid.TextMatrix(icg, bteColAmount)), gs_formatAmount)
            TxtRec(0) = IIf(InStr(1, TxtRec(0), "."), Format(TxtRec(0), gs_formatAmount), Format(TxtRec(0), gs_formatAmount))
        End If
    Next

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If grid.Col > bteColSelect Then Cancel = True

End Sub

Private Sub nomorinvoice()
    
    Dim adoRs As New ADODB.Recordset
    Dim intCount As Integer
    
    With cboInvoice
        
        .clear
        .columnCount = 2
        .ColumnWidths = "150pt;0pt"
        .ListWidth = 150

        If LCase(Cbo1.Text) = "create" Then
        
            sql = "Select Invoice_No, 0 As Invoice_Cls From Invoice_Master " & _
                "Where Invoice_Date >= '" & Format(Tgl1, "yyyy-mm-dd") & "' And Invoice_Date <= '" & Format(Tgl2, "yyyy-mm-dd") & "' And Cust_Code = '" & cboCust & "' And Fix_Cls = '1' " & _
                "And Invoice_No Not In (Select Distinct Invoice_No From FakturPajak_Detail) " '& _

        Else
        
            sql = "Select Invoice_No, 0 As Invoice_Cls From Invoice_Master " & _
                "Where Invoice_Date >= '" & Format(Tgl1, "yyyy-mm-dd") & "' And Invoice_Date <= '" & Format(Tgl2, "yyyy-mm-dd") & "' And Cust_Code = '" & cboCust & "' And Fix_Cls = '1' " '& _

        End If
        adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly
        .locked = adoRs.EOF
        While Not adoRs.EOF
            .AddItem
            .Column(0, intCount) = Trim(adoRs.Fields("Invoice_No") & "")
            .Column(1, intCount) = Trim(adoRs.Fields("Invoice_Cls") & "")
            intCount = intCount + 1
            adoRs.MoveNext
        Wend
        adoRs.Close
        Set adoRs = Nothing
        
    End With
    
End Sub

Sub NomorFaktur()
    
    RsPM.Requery
    RsPM.filter = "cust_code = '" & cboCust.Text & "'"
    If Not RsPM.EOF Then
        ComboBox1.clear
        ComboBox1.locked = False
        Do While Not RsPM.EOF
            ComboBox1.AddItem Right(Trim(RsPM!fakturpajak_no), 8)
            RsPM.MoveNext
        Loop
    Else
        ComboBox1.Text = ""
        ComboBox1.clear
        ComboBox1.locked = True
    End If
    RsPM.Requery
    RsPM.filter = ""

End Sub

Sub Kosong(nSection As Byte)
    
    Select Case nSection
    Case 1
        cboCust.ListIndex = -1
        ComboBox1 = ""
        NoStart = 0
        TxtRec(0) = ""
        TxtRec(1) = ""
        TxtRec(2) = ""
        Header
        cmdSubmit.Enabled = False
        CmdCreate.Caption = "Create"
        cboCust.locked = False
        ComboBox1.Text = ""
        ComboBox1.clear
        Cbo1.Text = "Update"
        RecBtn True
        tempInvoice = ""
    Case 2
        TxtRec(0) = ""
        TxtRec(1) = ""
        TxtRec(2) = ""
    End Select

End Sub

Function RecBtn(blnGerakdata As Boolean) As Boolean
    
    Dim kl As Byte
    For kl = 0 To 3
        CmdData(kl).Enabled = blnGerakdata
    Next

End Function

Sub InvoiceDisplay()

    Dim iv As Long, nQty As Long
    Dim jAmnt As Currency, jPPN As Currency, jTAmnt As Currency, nAmount As Currency
    Dim cInv As String, cPart As String, cCurr As String, cPrice As String, SQCP As String, NoInv As String, strUdah As String
    
    cInv = ""
    cPart = ""
    cCurr = ""
    cPrice = ""
    
    sql = "Select Seq_No, Cust_Code, PO_No, Unit_Cls, Invoice_No, Invoice_Date, Item_Code, Item_Name, IQty As Qty, Amount, Delivery_Date, Price, Currency_Code, DO_No, MakerItem_Code, QtyResult, FQty, AmountKu, DOSeq_No From( "
        
    sql = sql & " Select inv_d.Qty - Sum(Isnull(fak_d.Qty, 0)) As QtyResult, inv_d.DO_No, inv_m.Cust_Code, inv_d.Item_Code, RTrim(inv_d.MakerItem_Code) MakerItem_Code, " & _
        "im.Item_Name, inv_d.PO_No, inv_d.Unit_Cls, inv_d.Delivery_Date, inv_m.Invoice_Date, inv_d.Currency_Code, inv_d.Price As Price, inv_d.Amount As Amount, " & _
        "Abs(inv_d.Qty - Sum(Isnull(fak_d.Qty, 0))) * inv_d.Price  As AmountKu, inv_d.Qty As IQty, Sum(Isnull(fak_d.Qty, 0)) As FQty, inv_d.Seq_No, inv_d.DOSeq_No, inv_m.Invoice_No " & _
        "From Invoice_Detail inv_d " & _
        "Inner Join Item_Master im On inv_d.Item_Code = im.Item_Code " & _
        "Inner Join Invoice_Master inv_m On inv_d.Invoice_No = inv_m.Invoice_No " & _
        "Left Outer Join FakturPajak_Detail fak_d On inv_d.DO_No = fak_d.DO_No And inv_d.PO_No = fak_d.PO_No And inv_d.Seq_No = fak_d.Seq_No And inv_d.DOSeq_No = fak_d.DOSeq_No And inv_d.Invoice_No = fak_d.Invoice_No " & _
        "Where inv_m.Invoice_Date >= '" & Format(Tgl1.Value, "YYYY-MM-DD") & "' And (inv_m.Invoice_Date <= '" & Format(Tgl2.Value, "YYYY-MM-DD") & "') And inv_m.Cust_Code = '" & Trim(cboCust.Text) & "' And inv_m.Fix_Cls = '1' And inv_m.Invoice_No = '" & Trim(cboInvoice.Text) & "' " & _
        "Group By inv_d.Qty, inv_d.DO_No, inv_m.Cust_Code, im.Item_Name, inv_d.PO_No, inv_d.Unit_Cls, inv_d.Delivery_Date, inv_d.Currency_Code, inv_d.Price, " & _
        "inv_d.amount, inv_d.Item_Code, inv_d.MakerItem_Code, im.Item_Name, inv_m.Fix_Cls, inv_m.Invoice_Date, inv_d.Seq_No, inv_d.DOSeq_No, inv_d.Invoice_No, inv_m.Invoice_No " & _
        "Having inv_d.qty - Sum(IsNull(fak_d.qty, 0)) > 0 " & _
        "And (RTrim(inv_d.PO_No) + RTrim(inv_d.DO_No) + RTrim(inv_d.Invoice_No) + RTrim(inv_d.Seq_No) + RTrim(inv_d.DOSeq_No) Not In (" & _
            "Select (RTrim(fak_d.PO_No) + RTrim(fak_d.DO_No) + RTrim(fak_d.Invoice_No) + RTrim(fak_d.Seq_No) + RTrim(fak_d.DOSeq_No)) " & _
            "From FakturPajak_Detail fak_d " & _
            "Left Outer Join Invoice_Detail inv_d On fak_d.DO_No = inv_d.DO_No And fak_d.PO_No = inv_d.PO_No And fak_d.Seq_No = inv_d.Seq_No And fak_d.DOSeq_No = inv_d.DOSeq_No And inv_d.Invoice_No = fak_d.Invoice_No " & _
            "Where (fak_d.FakturPajak_No = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "')))"
            
    sql = sql & " Union "
    
    sql = sql & _
        "Select inv_d.Qty - Sum(Isnull(fak_d.Qty, 0)) As QtyResult, inv_d.DO_No, inv_m.Cust_Code, inv_d.Item_Code, RTrim(inv_d.MakerItem_Code) MakerItem_Code, " & _
        "im.Item_Name, inv_d.PO_No, inv_d.Unit_Cls, inv_d.Delivery_Date, inv_m.Invoice_Date, inv_d.Currency_Code,inv_d.Price As Price, inv_d.Amount As Amount, " & _
        "Abs(inv_d.Qty - Sum(Isnull(fak_d.Qty, 0))) * inv_d.Price As AmountKu, inv_d.Qty As IQty, Sum(Isnull(fak_d.Qty, 0)) As Qty, inv_d.Seq_No, inv_d.DOSeq_No, inv_d.Invoice_No " & _
        "From Invoice_Detail inv_d " & _
        "Inner Join Item_Master im On inv_d.Item_Code = im.Item_Code " & _
        "Inner Join Invoice_Master inv_m On inv_d.Invoice_No = inv_m.Invoice_No " & _
        "Inner Join FakturPajak_Detail fak_d On inv_d.DO_No = fak_d.DO_No And inv_d.PO_No = fak_d.PO_No And inv_d.Seq_No = fak_d.Seq_No And inv_d.DOSeq_No = fak_d.DOSeq_No And inv_d.Invoice_No = fak_d.Invoice_No " & _
        "Where fak_d.FakturPajak_No = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "' And inv_m.Invoice_No Like  '%" & Trim(cboInvoice.Text) & "' " & _
        "Group By inv_d.Qty, inv_d.DO_No, inv_m.Cust_Code, im.Item_Name, inv_d.PO_No, inv_d.Unit_Cls, inv_d.Delivery_Date, inv_d.Currency_Code, inv_d.Price, " & _
        "inv_d.amount, inv_d.Item_Code, inv_d.MakerItem_Code, im.Item_Name, inv_m.Fix_Cls, inv_m.Invoice_Date, inv_d.Seq_No, inv_d.DOSeq_No, inv_d.Invoice_No, inv_m.Invoice_No "
        
    sql = sql & ") a Order by Invoice_No, MakerItem_Code, Price"
        
    CmdInv.Enabled = True
    
    iv = 0
    Header
    tempInvoice = ""
    
    Set RsI = New Recordset
    RsI.Open sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
    
    If Not RsI.EOF Then cboInvoice.Text = Trim(RsI!Invoice_No)
    
    While Not RsI.EOF
        
        iv = iv + 1
        With grid
        
            cInv = Trim(RsI!Invoice_No)
            If Val(RsI!Price) = 0 Then
                cPrice = 0
            Else
                If Round(CDbl(RsI!Price)) / CDbl(RsI!Price) = 1 Then
                    cPrice = Format(CDbl(RsI!Price), gs_formatPrice)
                Else
                    cPrice = Format(CDbl(RsI!Price), gs_formatPrice)
                End If
            End If
            
            If IsNull(RsI!MakerItem_Code) Then
                cPart = ""
            Else
                cPart = Trim(RsI!MakerItem_Code)
            End If
            If IsNull(RsI!currency_code) Then
                cCurr = ""
            Else
                cCurr = uf_GetCurrencyDescription(Trim(RsI!currency_code))
            End If
            
            If (NoInv <> Trim(cInv) + Trim(cPart) + Trim(cCurr) + Trim(cPrice) And iv > 1) Or RsI.EOF Then
                .AddItem ""
                .Cell(flexcpBackColor, iv, bteColSelect) = vbWhite
                .TextMatrix(iv, bteColQty) = Format(nQty, gs_formatQty)
                .TextMatrix(iv, bteColAmount) = IIf(InStr(1, nAmount, "."), Format(nAmount, gs_formatAmount), Format(nAmount, gs_formatAmount))
                .Cell(flexcpBackColor, iv, 1, iv, bteColAmount) = &HE0E0E0
                iv = iv + 1
                nQty = 0
                nAmount = 0
            End If
            
            .AddItem ""
            .GridLines = flexGridRaised
            .Cell(flexcpBackColor, iv, bteColSelect) = vbWhite
            .TextMatrix(iv, bteColInvNo) = Trim(cInv)
            .TextMatrix(iv, bteColInvDate) = Format(RsI!Invoice_Date, "DD MMM YYYY") 'TglInvoice(Trim(RsI!invoice_no))
            .TextMatrix(iv, bteColProdCode) = Trim(RsI!Item_Code) & ""
            .TextMatrix(iv, bteColPartNo) = cPart
            .TextMatrix(iv, bteColDesc) = Trim(RsI!item_name) & ""
            .TextMatrix(iv, bteColDate) = Format(RsI!delivery_Date, "dd mmm yyyy")
            .TextMatrix(iv, bteColQty) = Format(RsI!Qty, gs_formatQty)
            .TextMatrix(iv, bteColUnit) = uf_GetUnitDescription(Trim(RsI!Unit_cls))
            .TextMatrix(iv, bteColCurr) = cCurr
            .TextMatrix(iv, bteColPrice) = cPrice
            .TextMatrix(iv, bteColAmount) = IIf(InStr(1, RsI!Amount, "."), Format(RsI!Amount, gs_formatAmount), Format(RsI!Amount, gs_formatAmount))
            .TextMatrix(iv, bteColPPn) = (CDbl(grid.TextMatrix(iv, bteColAmount)) * CDbl(tax(0))) / 100 'hide
            .TextMatrix(iv, bteColPPn) = IIf(InStr(1, .TextMatrix(iv, bteColPPn), "."), Format(.TextMatrix(iv, bteColPPn), gs_formatAmount), Format(.TextMatrix(iv, bteColPPn), gs_formatAmount))
            .TextMatrix(iv, bteColTAmount) = CDbl(.TextMatrix(iv, bteColAmount)) + CDbl(.TextMatrix(iv, bteColPPn)) 'hide
            .TextMatrix(iv, bteColTAmount) = IIf(InStr(1, .TextMatrix(iv, bteColTAmount), "."), Format(.TextMatrix(iv, bteColTAmount), gs_formatAmount), Format(.TextMatrix(iv, bteColTAmount), gs_formatAmount))
            .TextMatrix(iv, bteColDONo) = Trim(RsI!do_no) 'hide
            .TextMatrix(iv, bteColPONo) = Trim(RsI!po_no) 'hide
            .TextMatrix(iv, bteColSeqNo) = Trim(RsI!Seq_no) 'hide
            .TextMatrix(iv, bteColDOSeqNo) = Trim(RsI!DOSeq_No) 'hide
            
            '#Check All
            .Cell(flexcpChecked, iv, bteColSelect) = flexChecked
            jAmnt = jAmnt + CDbl(grid.TextMatrix(iv, bteColAmount))
            jPPN = jPPN + CDbl(grid.TextMatrix(iv, bteColPPn))
            jTAmnt = jTAmnt + CDbl(grid.TextMatrix(iv, bteColTAmount))
            tempInvoice = Trim(RsI!Invoice_No)
            
            nQty = nQty + RsI!Qty
            nAmount = nAmount + RsI!Amount
            NoInv = Trim(cInv) + Trim(cPart) + Trim(cCurr) + Trim(cPrice)
            
            .Cell(flexcpBackColor, iv, bteColInvNo, iv, bteColAmount) = &H80000018
            .Cell(flexcpBackColor, iv, bteColSelect) = vbWhite

            RsI.MoveNext
        
            If RsI.EOF Then
                iv = iv + 1
                .AddItem ""
                .Cell(flexcpBackColor, iv, bteColSelect) = vbWhite
                .TextMatrix(iv, bteColQty) = Format(nQty, gs_formatQty)
                .TextMatrix(iv, bteColAmount) = IIf(InStr(1, nAmount, "."), Format(nAmount, gs_formatAmount), Format(nAmount, gs_formatAmount))
                .Cell(flexcpBackColor, iv, bteColInvNo, iv, bteColAmount) = &HE0E0E0
                nQty = 0
                nAmount = 0
            End If
        
        End With
    
    Wend
    RsI.Close
    
    TxtRec(0) = IIf(InStr(1, jAmnt, "."), Format(jAmnt, gs_formatAmount), Format(jAmnt, gs_formatAmount))
    hitungKonversi
    
    Set RsI = Nothing
    
End Sub

Function TglInvoice(invno As String) As String
    
    Dim RsTI As Recordset
    
    Set RsTI = Db.Execute("Select invoice_date from invoice_master where invoice_no='" & invno & "'")
    If Not RsTI.EOF Then TglInvoice = Format(RsTI!Invoice_Date, "dd mmm yyyy") Else TglInvoice = ""

End Function

Private Sub tgl1_Change()
    
    nomorinvoice

End Sub

Private Sub tgl1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then SendKeys vbTab

End Sub

Private Sub Tgl2_Change()
    
    nomorinvoice

End Sub

Private Sub Tgl2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then SendKeys vbTab

End Sub

Sub savemaster()

    With RsPM
        
        RsPM.Requery
        .filter = "faktuRPajak_NO = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "'"
        
        If Not .EOF Then
            
            .Fields("FakturPajak_No") = cboCode & Text1 & "." & Trim(ComboBox1)
            .Fields("FakturPajak_date") = Format(TglPajak, "yyyy-mm-dd")
            .Fields("Cust_code") = cboCust.Text
            
            If Val(TxtRec(0)) = 0 Then .Fields("Amount") = Val(TxtRec(0)) Else .Fields("Amount") = CDbl(TxtRec(0))
            If Val(TxtRec(1)) = 0 Then .Fields("PPN") = Val(TxtRec(1)) Else .Fields("PPN") = CDbl(TxtRec(1))
            If Val(TxtRec(bteColInvDate)) = 0 Then .Fields("Total_amount") = Val(TxtRec(2)) Else .Fields("Total_amount") = CDbl(TxtRec(2))
            .Fields("Last_Update") = Now
            .Fields("Last_User") = userLogin
            .update
            LblErr.Caption = DisplayMsg(1101)
    
    Else
    
        .AddNew
        .Fields("FakturPajak_No") = cboCode & Text1 & "." & Trim(ComboBox1)
        .Fields("FakturPajak_date") = Format(TglPajak, "yyyy-mm-dd")
        .Fields("Cust_code") = cboCust.Text
        .Fields("Amount") = Val(TxtRec(0))
        .Fields("PPN") = Val(TxtRec(1))
        .Fields("Total_amount") = Val(TxtRec(2))
        .Fields("Last_Update") = Now
        .Fields("Last_User") = userLogin
        
        On Error Resume Next
        .update
    
Errhandle:
        
        If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
            err.clear
            ComboBox1.locked = False
        
            ' generate number
            RsNF.Requery
            RsNF.filter = " no >= '" & Format(txtstart, "00000000") & "' "
            If Not RsNF.EOF Then
                Text2.Text = Format(RsNF!NO, "00000000")
                ComboBox1 = Format(RsNF!NO + 1, "00000000")
            Else
                If NoStart <> Val(txtstart) Then
                    ComboBox1 = Format(txtstart, "00000000")
                Else
                    If Val(txtstart) <> 0 Then
                        If NoStart = Val(txtstart) Then
                            ComboBox1 = Format(txtstart, "00000000")
                        Else
                            ComboBox1 = Format(ComboBox1 + 1, "00000000")
                        End If
                    Else
                            ComboBox1 = "00000001"
                    End If
                End If
                NoStart = txtstart
            End If
            RsNF.filter = ""
            RsNF.Requery
            .Fields("fakturpajak_no") = cboCode & Text1 & "." & Trim(ComboBox1)
            If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then GoTo Errhandle
            .Fields("Last_Update") = Now
            .Fields("Last_User") = userLogin
            .update
            ComboBox1.locked = False
            End If
            LblErr.Caption = DisplayMsg(1000)
        End If
        .Requery
        .filter = ""
        
    End With

End Sub

Sub inquiry()

    Dim RsInq As Recordset
    Dim SQin As String
    
    SQin = "select FakturPajak_master.*, isnull(invoice_no,'') invoice_no from FakturPajak_master , FakturPajak_detail  where FakturPajak_master.FakturPajak_no =  FakturPajak_detail.FakturPajak_no and FakturPajak_detail.FakturPajak_no = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "'"
    Set RsInq = Db.Execute(SQin)
    
    If Not RsInq.EOF Then
        ComboBox1 = Right(Trim(RsInq!fakturpajak_no), 8)
        TglPajak = RsInq!fakturpajak_date
        cboCust = Trim(RsInq!Cust_CodE)
        TxtRec(0) = IIf(InStr(1, RsInq!Amount, "."), Format(RsInq!Amount, gs_formatAmount), Format(RsInq!Amount, gs_formatAmount))
        TxtRec(1) = IIf(InStr(1, RsInq!ppn, "."), Format(RsInq!ppn, gs_formatAmount), Format(RsInq!ppn, gs_formatAmount))
        TxtRec(2) = IIf(InStr(1, RsInq!total_amount, "."), Format(RsInq!total_amount, gs_formatAmount), Format(RsInq!total_amount, gs_formatAmount))
        hitungCurr
        cboInvoice.Text = Trim(RsInq!Invoice_No)
    End If

End Sub

Sub updateMaster()

Dim rsUpdate As Recordset

    sql = "select * from FakturPajak_master where FakturPajak_no = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "'"
    Set rsUpdate = New Recordset
    rsUpdate.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    With rsUpdate
        If Val(TxtRec(0)) = 0 And Val(TxtRec(1)) = 0 And Val(TxtRec(2)) = 0 Then Exit Sub
        If Not .EOF Then
            !Amount = CDbl(TxtRec(0))
            !ppn = CDbl(TxtRec(1))
            !total_amount = CDbl(TxtRec(2))
            !fakturpajak_date = Format(TglPajak, "yyyy-mm-dd")
            !amount_IDR = CDbl(txtAmountIDR)
            !ppn_IDR = CDbl(txtPPNIDR)
            !Last_Update = Now
            !last_user = userLogin
            .update
        End If
        rsUpdate.Close
        Set rsUpdate = Nothing
    End With

End Sub

Function tax(nRow As Long) As String

    Dim rtax As Recordset, tempidate As String

    tempidate = Format(TglPajak.Value, "yyyymmdd")
    
    sql = "SELECT Rate FROM tax_cls where start_date <= '" & tempidate & "'  and end_date >= '" & tempidate & "' and tax_code = 'PPN'"
    Set rtax = Db.Execute(sql)

    tax = 0
    If Not rtax.EOF Then
        If nRow > 0 Then
            tax = (CDbl(grid.TextMatrix(nRow, bteColQty)) * rtax!rate) / 100
            tax = IIf(InStr(1, tax, "."), Format(tax, gs_formatAmount), Format(tax, gs_formatAmount))
        Else
            tax = rtax!rate
        End If
    End If
    
End Function

Sub savedetail()

    Dim i As Long
    Dim rsdetail As Recordset
    
    sql = "Delete From FakturPajak_Detail Where FakturPajak_No = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "' And Invoice_No <> '" & cboInvoice.Text & "'"
    Db.Execute sql
    
    For i = 1 To grid.Rows - 1
        
        If grid.Cell(flexcpChecked, i, bteColSelect, i, bteColSelect) = flexChecked Then
            sql = " select * from fakturPajak_detail where FakturPajak_no = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "' and Do_no='" & grid.TextMatrix(i, bteColDONo) & "' and invoice_no ='" & grid.TextMatrix(i, bteColInvNo) & "' and PO_No='" & grid.TextMatrix(i, bteColPONo) & "' and Seq_no='" & Trim(grid.TextMatrix(i, bteColSeqNo)) & "' and DoSeq_no='" & Trim(grid.TextMatrix(i, bteColDOSeqNo)) & "' "
            Set rsdetail = New Recordset
            rsdetail.Open sql, Db, adOpenDynamic, adLockOptimistic
            With rsdetail
            If .EOF Then
                .AddNew
                .Fields("FakturPajak_no") = cboCode & Text1 & "." & Trim(ComboBox1)
                .Fields("invoice_no") = grid.TextMatrix(i, bteColInvNo)
                .Fields("DO_no") = grid.TextMatrix(i, bteColDONo)
                .Fields("Item_code") = grid.TextMatrix(i, bteColProdCode)
                .Fields("makerItem_code") = grid.TextMatrix(i, bteColPartNo)
                .Fields("Delivery_date") = Format(grid.TextMatrix(i, bteColDate), "yyyy-mm-dd")
                .Fields("Qty") = CDbl(grid.TextMatrix(i, bteColQty))
                .Fields("Price") = CDbl(grid.TextMatrix(i, bteColPrice))
                .Fields("Currency_code") = uf_GetCurrencyCode(grid.TextMatrix(i, bteColCurr))
                .Fields("Unit_cls") = uf_GetUnitCode(Trim(grid.TextMatrix(i, bteColUnit)))
                .Fields("Amount") = CDbl(grid.TextMatrix(i, bteColAmount))
                .Fields("PO_No") = Trim(grid.TextMatrix(i, bteColPONo))
                .Fields("Seq_No") = Trim(grid.TextMatrix(i, bteColSeqNo))
                .Fields("DoSeq_No") = Trim(grid.TextMatrix(i, bteColDOSeqNo))
                .Fields("Last_Update") = Now
                .Fields("Last_User") = userLogin
                .Fields("Packing_No") = GetNumber("select max(Packing_NO)+1 FROM fakturPajak_detail")
                .Fields("PackingSeq_No") = GetNumber("select max(PackingSeq_No)+1 FROM fakturPajak_detail")
                .update
                LblErr.Caption = DisplayMsg(1000)
            Else
                .Fields("FakturPajak_no") = cboCode & Text1 & "." & Trim(ComboBox1)
                .Fields("invoice_no") = grid.TextMatrix(i, bteColInvNo)
                .Fields("DO_no") = grid.TextMatrix(i, bteColDONo)
                .Fields("Item_code") = grid.TextMatrix(i, bteColProdCode)
                .Fields("makerItem_code") = grid.TextMatrix(i, bteColPartNo)
                .Fields("Delivery_date") = Format(grid.TextMatrix(i, bteColDate), "yyyy-mm-dd")
                .Fields("Qty") = CDbl(grid.TextMatrix(i, bteColQty))
                .Fields("Price") = CDbl(grid.TextMatrix(i, bteColPrice))
                .Fields("Currency_code") = uf_GetCurrencyCode(grid.TextMatrix(i, bteColCurr))
                .Fields("Unit_cls") = uf_GetUnitCode(Trim(grid.TextMatrix(i, bteColUnit)))
                .Fields("Amount") = CDbl(grid.TextMatrix(i, bteColAmount))
                .Fields("PO_No") = Trim(grid.TextMatrix(i, bteColPONo))
                .Fields("Seq_No") = Trim(grid.TextMatrix(i, bteColSeqNo))
                .Fields("DoSeq_No") = Trim(grid.TextMatrix(i, bteColDOSeqNo))
                .Fields("Last_Update") = Now
                .Fields("Last_User") = userLogin
                .update
                LblErr.Caption = DisplayMsg(1101)
            End If
            End With
            SummaryDetil cboCode & Text1 & "." & Trim(ComboBox1)
            updateMaster
        Else
            If Trim(grid.TextMatrix(i, bteColInvNo)) <> "" Then
                sql = "select * from fakturpajak_detail where FakturPajak_no = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "' and Do_no='" & grid.TextMatrix(i, bteColDONo) & "' and invoice_no ='" & grid.TextMatrix(i, bteColInvNo) & "' and PO_No='" & grid.TextMatrix(i, bteColPONo) & "' and Seq_no='" & Trim(grid.TextMatrix(i, bteColSeqNo)) & "' and doseq_no = '" & Trim(grid.TextMatrix(i, bteColDOSeqNo)) & "' "
                Set rsdetail = New Recordset
                rsdetail.Open sql, Db, adOpenDynamic, adLockOptimistic
                With rsdetail
                If Not .EOF Then
                    sql = "delete FakturPajak_detail where FakturPajak_no = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "' and Do_no='" & grid.TextMatrix(i, bteColDONo) & "' and invoice_no ='" & grid.TextMatrix(i, bteColInvNo) & "' and PO_No='" & grid.TextMatrix(i, bteColPONo) & "' and Seq_no='" & Trim(grid.TextMatrix(i, bteColSeqNo)) & "' and doseq_no = '" & Trim(grid.TextMatrix(i, bteColDOSeqNo)) & "' "
                    Db.Execute sql
                    SummaryDetil cboCode & Text1 & "." & Trim(ComboBox1)
                    updateMaster
                End If
                End With
            End If
        End If
    Next

End Sub

Sub SummaryDetil(strNofaktur As String)

    Dim RsSD As Recordset
    Dim sql1 As String

    sql1 = "SELECT SUM(Amount) AS TA, FakturPajak_No From FakturPajak_Detail where fakturpajak_no = '" & strNofaktur & "' GROUP BY FakturPajak_No"
    Set RsSD = Db.Execute(sql1)

    If Not RsSD.EOF Then TxtRec(0) = IIf(InStr(1, RsSD!TA, "."), Format(RsSD!TA, gs_formatAmount), Format(RsSD!TA, gs_formatAmount))
    TxtRec(1) = (CDbl(TxtRec(0)) * CDbl(tax(0))) / 100
    TxtRec(1) = IIf(InStr(1, TxtRec(1), "."), Format(TxtRec(1), gs_formatAmount), Format(TxtRec(1), gs_formatAmount))
    TxtRec(2) = CDbl(TxtRec(0)) + CDbl(TxtRec(1))
    TxtRec(2) = IIf(InStr(1, TxtRec(2), "."), Format(TxtRec(2), gs_formatAmount), Format(TxtRec(2), gs_formatAmount))

    RsSD.Close
    Set RsSD = Nothing

End Sub



Private Sub TglPajak_Change()
    
    LblErr = ""
'    If year(TglPajak) < 2007 Then
'        LblErr = "Faktur Pajak Year must be higher or equal than '2007'"
'        If year(Date) >= 2007 Then TglPajak = Date Else TglPajak = CDate("2007-01-01")
'    End If
    Text1 = "0.000-" & Right(TglPajak.Year, 2)
    hitungCurr
    Cbo1_Click

End Sub

Private Sub TxtRec_Change(Index As Integer)

    If Index = 0 Then
        If Trim(TxtRec(0)) <> "" Then
            TxtRec(1) = (CDbl(TxtRec(0)) * tax(0)) / 100
            TxtRec(1) = IIf(InStr(1, TxtRec(1), "."), Format(TxtRec(1), gs_formatAmount), Format(TxtRec(1), gs_formatAmount))
            TxtRec(2) = CDbl(TxtRec(0)) + CDbl(TxtRec(1))
            TxtRec(2) = IIf(InStr(1, TxtRec(2), "."), Format(TxtRec(2), gs_formatAmount), Format(TxtRec(2), gs_formatAmount))
            hitungKonversi
        Else
            TxtRec(1) = 0
            TxtRec(2) = 0
            txtAmountIDR = 0
            txtPPNIDR = 0
        End If
    End If
    
End Sub

Private Sub hitungKonversi()
    
    Dim i As Long
    For i = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
            Select Case RTrim(grid.TextMatrix(i, bteColCurr))
                Case "YEN"
                    currConv = IIf(lblYEN.Visible = False, 0, CDbl(lblYEN.Caption))
                Case "US$"
                    currConv = IIf(lblUSD.Visible = False, 0, CDbl(lblUSD.Caption))
                Case "IDR"
                    currConv = 1
                Case "EUR"
                    currConv = IIf(lblEur.Visible = False, 0, CDbl(lblEur.Caption))
                Case "S$"
                    currConv = IIf(lblSD.Visible = False, 0, CDbl(lblSD.Caption))
                Case "THB"
                    currConv = IIf(lblTHB.Visible = False, 0, CDbl(lblTHB.Caption))
                Case Else
                    currConv = 1
            End Select
            Exit For
        End If
    Next i
    
    txtAmountIDR = TxtRec(0) * currConv
    txtAmountIDR = IIf(InStr(1, txtAmountIDR, "."), Format(txtAmountIDR, gs_formatAmount), Format(txtAmountIDR, gs_formatAmount))
    txtPPNIDR = ((TxtRec(0) * currConv) * tax(0)) / 100
    txtPPNIDR = IIf(InStr(1, txtPPNIDR, "."), Format(txtPPNIDR, gs_formatAmount), Format(txtPPNIDR, gs_formatAmount))

End Sub

Private Sub txtstart_KeyPress(KeyAscii As Integer)

    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0

End Sub

Function InvDate(strNoInvoice As String) As String

    Dim RSiD As Recordset
    
    Set RSiD = Db.Execute("select Invoice_date from invoice_master where invoice_no='" & strNoInvoice & "'")
    If Not RSiD.EOF Then
        InvDate = RSiD!Invoice_Date
    Else
        InvDate = ""
    End If
    RSiD.Close
    Set RSiD = Nothing

End Function

Function FPajakDate(strFPajakNo As String) As String

    Dim rsFD As Recordset
    
    Set rsFD = Db.Execute("Select FakturPajak_date as FDate from fakturpajak_master where fakturpajak_no='" & strFPajakNo & "'")
    If Not rsFD.EOF Then FPajakDate = Format(rsFD!fdate, "dd mmm yyyy")
    rsFD.Close
    Set rsFD = Nothing

End Function

Function CekNoPajak(strNoPajak As String) As Boolean

    Dim RCn As Recordset
    
    Set RCn = Db.Execute("Select FakturPajak_No,FakturPajak_Date from fakturpajak_master where fakturpajak_no='" & strNoPajak & "'")
    If Not RCn.EOF Then CekNoPajak = True Else CekNoPajak = False

End Function

Sub ViewReport(NO$)
  
    Dim application As New CRAXDDRT.application
    Dim Repot As New CRAXDDRT.report
    Dim rsRpt As New ADODB.Recordset
    Dim Rpt As New FrmRpt3
    
    LblErr = ""
    Me.MousePointer = vbHourglass
    
    sql = "Select " & bteHakPrice & "As HakPrice, RTrim(cp.NPWP_No) As NPWP_No, RTrim(cp.NPWP_Name) As NPWP_Name, RTrim(cp.NPWP_Address) As NPWP_Address, RTrim(cp.NPWP_City) As NPWP_City, " & _
        "cp.TglPengukuhan, RTrim(cp.Tax_Position) As Tax_Position, RTrim(cp.Tax_Person) As Tax_Person, " & _
        "RTrim(inv_to.NPWP_No) As cust_NPWP_No, RTrim(inv_to.NPWP_Name) As cust_NPWP_Name, RTrim(inv_to.NPWP_Address) As cust_NPWP_Address, RTrim(inv_to.NPWP_City) As cust_NPWP_City, RTrim(inv_to.NPPKP_No) As cust_NPPKP_No, " & _
        "RTrim(fak_d.FakturPajak_No) As FakturPajak_No, RTrim(inv_m.Invoice_No) As Invoice_No, fak_m.FakturPajak_Date, inv_m.Invoice_Date, " & _
        "inv_m.Amount, inv_m.PPN, inv_m.Total_Amount, RTrim(inv_m.List_Do) As List_Do, RTrim(inv_m.List_Po) As List_Po, inv_m.Due_Date, '' As FakturPajak_Ref, " & _
        "RTrim(im.Item_Name) As Item_Name, RTrim(inv_d.MakerItem_Code) As MakerItem_Code, Sum(inv_d.Qty) Qty, inv_d.Price, " & _
        " (select description from unit_cls uc where uc.unit_cls=  inv_d.unit_cls) unit_desc," & _
        "Currency = (select description from curr_cls cc where cc.curr_cls= inv_d.Currency_Code), " & _
        "Ex_Rate = " & _
            "IsNull((Select Top 1 Tax_ExchangeRate From Tax_exchangerate " & _
            "Where Currency_Code = inv_d.Currency_Code And Start_Date <= inv_m.Invoice_Date And End_Date >= inv_m.Invoice_Date Order By Start_Date Desc), 1) " & _
        "From Company_Profile cp, Trade_Master inv_to " & _
        "Inner Join Trade_Master cust On  inv_to.Trade_Code = IsNull( cust.Trade_Code,cust.Invoice_To) " & _
        "Inner Join Invoice_Master inv_m On cust.Trade_Code = inv_m.Cust_Code " & _
        "Inner Join Invoice_Detail inv_d On inv_m.Invoice_No = inv_d.Invoice_No " & _
        "Inner Join Item_Master im On inv_d.Item_Code = im.Item_Code " & _
        "Left Outer Join FakturPajak_Detail fak_d On inv_d.Invoice_No = fak_d.Invoice_No And inv_d.DO_No = fak_d.DO_No And inv_d.PO_No = fak_d.PO_No And inv_d.Seq_No = fak_d.Seq_No And inv_d.DOSeq_No = fak_d.DOSeq_No " & _
        "Left Outer Join FakturPajak_Master fak_m On fak_d.FakturPajak_No = fak_m.FakturPajak_No "

    sql = sql & _
        "Where fak_m.FakturPajak_No = '" & NO$ & "' " & _
        "Group by cp.NPWP_No, cp.NPWP_Name , cp.NPWP_Address, cp.NPWP_City, cp.TglPengukuhan, " & _
        "cp.Tax_Position, cp.Tax_Person, inv_to.NPWP_No, inv_to.NPWP_Name , " & _
        "inv_to.NPWP_Address, inv_to.NPWP_City, inv_to.NPPKP_No, fak_d.FakturPajak_No, inv_m.Invoice_No, " & _
        "fak_m.FakturPajak_Date, inv_m.Invoice_Date, inv_m.Amount, inv_m.PPN, inv_m.Total_Amount, " & _
        "inv_m.List_Do, inv_m.List_Po, inv_m.Due_Date, fak_m.FakturPajak_No, im.Item_Name, inv_d.MakerItem_Code, inv_d.Price, " & _
        "inv_d.Unit_Cls, inv_d.Currency_Code " & _
        "Order By fak_m.FakturPajak_No, im.Item_Name, fak_d.MakerItem_Code"
    
    If rsRpt.State <> adStateClosed Then rsRpt.Close
    rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
    
    If rsRpt.EOF Then LblErr.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
    sqlprint = sql
    printorient = 1
    reportcode = "Faktur Pajak"
    Set Repot = application.OpenReport(App.path & "\Reports\faktur_pajak_std.rpt")
    Repot.Database.Tables(1).SetDataSource rsRpt
    
    '#####################################################################
    '# Qty Digit and decimal
    Repot.FormulaFields(39).Text = "" & gi_decimalDigitAmount & ""
    Repot.FormulaFields(40).Text = "" & gi_decimalDigitAmount & ""
    Repot.FormulaFields(42).Text = "" & gi_decimalDigitExchangeRate & ""
    Repot.FormulaFields(43).Text = "" & gi_decimalDigitExchangeRate & ""
    '#####################################################################
    
    Repot.ReportTitle = "Faktur Pajak"
    
    Rpt.CRViewer1.ReportSource = Repot
    Rpt.CRViewer1.ViewReport
    Rpt.CRViewer1.Zoom 1
    
    Rpt.WindowState = 2
    Rpt.Show 1
    
    Me.MousePointer = vbDefault

End Sub

Private Function GetFakturNo() As String
    
    Dim adoRs As New ADODB.Recordset
    
    GetFakturNo = ""
    sql = "Select FakturPajak_No From FakturPajak_Detail Where Invoice_No ='" & Trim(cboInvoice.Text) & "'"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        GetFakturNo = Right(Trim(adoRs.Fields("FakturPajak_No")), 8)
    End If
    adoRs.Close
    
End Function

Private Function GetInvoiceNo() As String
    
    Dim adoRs As New ADODB.Recordset
    
    GetInvoiceNo = ""
    sql = "Select  Invoice_No From FakturPajak_Detail Where  FakturPajak_No = '" & cboCode & Text1 & "." & Trim(ComboBox1) & "'"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        GetInvoiceNo = adoRs.Fields("Invoice_No")
    End If
    adoRs.Close
    
End Function

