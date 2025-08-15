VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAPInvoiceInformationEntry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "AP Invoice Information Entry"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15240
   Icon            =   "FrmAPInvoiceInformationEntry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10515
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtrate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   12570
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   8745
      Width           =   2355
   End
   Begin VB.TextBox txtkonversi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   10125
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   8745
      Width           =   2355
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "FrmAPInvoiceInformationEntry.frx":0E42
      Left            =   180
      List            =   "FrmAPInvoiceInformationEntry.frx":0E44
      Style           =   2  'Dropdown List
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1485
   End
   Begin VB.TextBox txtamount_service 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   5325
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   8745
      Width           =   2355
   End
   Begin VB.TextBox txtamount_material 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   2925
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   8745
      Width           =   2355
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13170
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   323
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   714
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
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10005
      Width           =   1290
   End
   Begin VB.TextBox txtPPN 
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
      Height          =   315
      Left            =   12570
      MaxLength       =   23
      TabIndex        =   10
      Top             =   10545
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox txtinvno 
      Appearance      =   0  'Flat
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
      Left            =   330
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   8745
      Width           =   2550
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   7725
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8745
      Width           =   2355
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
      Left            =   12660
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10005
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
      Left            =   11415
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10005
      Width           =   1140
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
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10005
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   180
      TabIndex        =   18
      Top             =   9255
      Width           =   14865
      Begin VB.Label lblerror 
         Alignment       =   2  'Center
         BackColor       =   &H00FDDFE3&
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
         Height          =   270
         Left            =   105
         TabIndex        =   19
         Top             =   195
         Width           =   14640
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   3675
      Left            =   180
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4440
      Width           =   14865
      _cx             =   26220
      _cy             =   6482
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
      FocusRect       =   5
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1305
      Left            =   180
      TabIndex        =   20
      Top             =   1020
      Width           =   14835
      Begin VB.TextBox txtSearch 
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
         Left            =   11130
         MaxLength       =   25
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   270
         Width           =   2430
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "Find [F3]"
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
         Left            =   13635
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox txt_name 
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
         Height          =   225
         Left            =   3555
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "txt_name"
         Top             =   360
         Width           =   4650
      End
      Begin MSComCtl2.DTPicker InvDateEnd 
         Height          =   315
         Left            =   3870
         TabIndex        =   2
         Top             =   765
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   149553155
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker InvDateSt 
         Height          =   315
         Left            =   1530
         TabIndex        =   1
         Top             =   765
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   149553155
         CurrentDate     =   37810
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
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
         Index           =   26
         Left            =   8280
         TabIndex        =   64
         Top             =   330
         Width           =   600
      End
      Begin MSForms.ComboBox cboSearch 
         Height          =   315
         Left            =   8970
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   270
         Width           =   2085
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   7
         Size            =   "3678;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox CboPoNo 
         Height          =   315
         Left            =   6870
         TabIndex        =   3
         Top             =   765
         Width           =   3015
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "5318;556"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Surat Jalan No"
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
         Left            =   5580
         TabIndex        =   45
         Top             =   825
         Width           =   1365
      End
      Begin VB.Label LblFix 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
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
         Left            =   12000
         TabIndex        =   31
         Top             =   960
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier CD"
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
         TabIndex        =   24
         Top             =   375
         Width           =   1035
      End
      Begin MSForms.ComboBox CboSupp 
         Height          =   315
         Left            =   1500
         TabIndex        =   0
         Top             =   315
         Width           =   1890
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3334;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         X1              =   3555
         X2              =   8205
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
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
         Left            =   3405
         TabIndex        =   23
         Top             =   825
         Width           =   195
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
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
         Left            =   330
         TabIndex        =   22
         Top             =   825
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   1695
      Left            =   180
      TabIndex        =   36
      Top             =   2700
      Width           =   14835
      Begin VB.TextBox TxtBlNo 
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
         Left            =   1410
         MaxLength       =   25
         TabIndex        =   52
         Top             =   765
         Width           =   2175
      End
      Begin VB.TextBox txtFakturPajakNo 
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
         Left            =   9420
         MaxLength       =   25
         TabIndex        =   51
         Top             =   765
         Width           =   2475
      End
      Begin VB.TextBox TxtVoucherDesc 
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
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   50
         Top             =   765
         Width           =   2475
      End
      Begin VB.TextBox txtairfreight 
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
         Height          =   315
         Left            =   9780
         MaxLength       =   15
         TabIndex        =   39
         Top             =   2820
         Width           =   1575
      End
      Begin VB.CommandButton BtnCreate 
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
         Left            =   7965
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1185
         Width           =   1005
      End
      Begin VB.TextBox TxtNmBank 
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
         Height          =   225
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "txt_name"
         Top             =   1260
         Width           =   1395
      End
      Begin VB.TextBox TxtBankAcc 
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
         Height          =   225
         Left            =   5445
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "txt_name"
         Top             =   1260
         Width           =   2220
      End
      Begin MSComCtl2.DTPicker DateInv 
         Height          =   315
         Left            =   5280
         TabIndex        =   5
         Top             =   255
         Width           =   2475
         _ExtentX        =   4366
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
         Format          =   108658691
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DateBL 
         Height          =   315
         Left            =   1230
         TabIndex        =   6
         Top             =   2370
         Visible         =   0   'False
         Width           =   2145
         _ExtentX        =   3784
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
         Format          =   108265475
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DateRecp 
         Height          =   315
         Left            =   9420
         TabIndex        =   46
         Top             =   255
         Width           =   2475
         _ExtentX        =   4366
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
         Format          =   152567811
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DateDue 
         Height          =   315
         Left            =   13140
         TabIndex        =   47
         Top             =   255
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   54853635
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DtFaktur 
         Height          =   315
         Left            =   13125
         TabIndex        =   53
         Top             =   765
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   152371203
         CurrentDate     =   37798
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Voucher No"
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
         Left            =   300
         TabIndex        =   57
         Top             =   825
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faktur Pajak No"
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
         Left            =   7860
         TabIndex        =   56
         Top             =   825
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   12600
         TabIndex        =   55
         Top             =   825
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Desc"
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
         Left            =   3840
         TabIndex        =   54
         Top             =   825
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date"
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
         Index           =   12
         Left            =   12195
         TabIndex        =   48
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B/C Date"
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
         Left            =   300
         TabIndex        =   44
         Top             =   2430
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv. Receipt"
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
         Left            =   8190
         TabIndex        =   43
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No"
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
         Left            =   435
         TabIndex        =   42
         Top             =   315
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv. Date"
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
         Left            =   3840
         TabIndex        =   41
         Top             =   315
         Width           =   810
      End
      Begin MSForms.ComboBox CboCurr 
         Height          =   315
         Left            =   1410
         TabIndex        =   7
         Top             =   1215
         Width           =   2175
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "3836;556"
         TextColumn      =   2
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Code"
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
         Left            =   300
         TabIndex        =   40
         Top             =   1275
         Width           =   945
      End
      Begin VB.Line Line1 
         X1              =   3780
         X2              =   5295
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Line Line3 
         X1              =   5430
         X2              =   7725
         Y1              =   1545
         Y2              =   1545
      End
      Begin MSForms.ComboBox CboInvNo 
         Height          =   315
         Left            =   1410
         TabIndex        =   4
         Top             =   255
         Width           =   2175
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3836;556"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
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
      Left            =   12600
      TabIndex        =   60
      Top             =   8340
      Width           =   390
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount Service"
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
      Left            =   5325
      TabIndex        =   35
      Top             =   8355
      Width           =   1845
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount Material"
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
      Left            =   2925
      TabIndex        =   33
      Top             =   8355
      Width           =   1875
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   525
      Left            =   6960
      TabIndex        =   30
      Top             =   5220
      Width           =   1245
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   180
      Top             =   8625
      Width           =   14865
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Konversi"
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
      Left            =   10125
      TabIndex        =   28
      Top             =   8355
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No"
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
      TabIndex        =   27
      Top             =   8355
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total Amount"
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
      Left            =   7725
      TabIndex        =   26
      Top             =   8355
      Width           =   1725
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
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
      Left            =   12570
      TabIndex        =   25
      Top             =   8805
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AP Invoice Information Entry"
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
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   330
      Width           =   14865
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   180
      Top             =   8265
      Width           =   14865
   End
End
Attribute VB_Name = "FrmAPInvoiceInformationEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim status As String
Dim rsSup As ADODB.Recordset
Dim RsInv As ADODB.Recordset
Dim RsAcc As ADODB.Recordset
Dim CtrCreate As Integer
Dim CtrCur As String
Dim Overseas_Cls As String
Dim tampung, rema As Long
Dim CtrNotAlow As Integer
Dim ColCls, ColPONo, colpodate, ColDoNo, ColBCNO, colpcode, colprodname, ColUnit, colpoqty, coldoqty As Integer
Dim colinvqty, colrem, ColCurr, ColPrice, ColAmount, colPriceMat, ColAmountMat, colPriceService As Integer
Dim ColAmountService, colqtyasli, ColReceiptSeqNo As Integer
Dim ColBantuUnit, ColBantuCurr, ColCurX, ColPoQtyReal As Long
Dim colAdj As Long
Dim RateOri As Double


Dim bteHakPrice As Byte
Dim InvDate As Date

Private Sub NotAlow()
Dim rsnotalow As New ADODB.Recordset
sql = "select * from ap_detail where supplier_code = '" & cboSupp & "' and invoice_no = '" & CboInvNo & "'"
rsnotalow.Open sql, Db, adOpenKeyset, adLockOptimistic
If rsnotalow.EOF And rsnotalow.BOF Then
    CtrNotAlow = 0
Else
    CtrNotAlow = 1
End If
End Sub

Private Sub BtnCreate_Click()
Dim Tamp As Double
Dim RS As New ADODB.Recordset
lblerror = ""
If hakUpdate(Me.Name) = 0 Then lblerror = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub


status = ""
status = Trim(combo1.Text)

CtrNotAlow = 0

NotAlow

MousePointer = vbHourglass
'SUPPLIER CHECK
    If cboSupp.Text = "" Then
        lblerror = DisplayMsg(1054)
        MousePointer = Default
        Exit Sub
    Else
        For i = 0 To cboSupp.ListCount - 1
            If Trim(cboSupp.List(i, 0)) = Trim(cboSupp.Text) Then
                GoTo Okay
            End If
        Next
        MousePointer = Default
        Exit Sub
    End If

Okay:
Header
'INVOICE NO CHECK
    If Trim(CboInvNo.Text) = "" Then
        lblerror = DisplayMsg(4082)
        MousePointer = Default
        Exit Sub
    End If
    
    If BtnCreate.Caption = "Update" Then
        For i = 0 To CboInvNo.ListCount - 1
            If Trim(CboInvNo.List(i, 0)) = Trim(CboInvNo.Text) Then
                GoTo Okay2
            End If
        Next
        MousePointer = Default
        Exit Sub
    End If

Okay2:
'CURRENCY CHECK
'    If cboCurr = "" Then
'        lblerror = DisplayMsg(4081)
'        MousePointer = Default
'        Exit Sub
'    Else
'        For i = 0 To cboCurr.ListCount - 1
'            If Trim(cboCurr.List(i, 1)) = Trim(cboCurr.Text) Then
'                GoTo Okay3
'            End If
'        Next
'        MousePointer = Default
'        Exit Sub
'    End If
Okay3:

If BtnCreate.Caption = "Create" Then
    
    lblerror = up_ValidateDateRange(DateInv, True)
    If lblerror <> "" Then Me.MousePointer = vbDefault: Exit Sub

    CtrCreate = 1
    'save Header doeloe
    txtamount_material.Text = "0"
    txtamount_service.Text = "0"
    txtamount.Text = "0"
    txtPPN.Text = "0"
            
    'INVOICE NO ALREADY IN USED OR NOT
    If RS.State = 1 Then RS.Close
'    rs.Open "select * from invoicesupplier_master where invoice_no = '" & CboInvNo.Text & "' and " & _
'        "supplier_code = '" & CboSupp.Text & "'", Db, adOpenKeyset, adLockOptimistic
    RS.Open "select * from invoicesupplier_master where invoice_no = '" & CboInvNo.Text & "' ", Db, adOpenKeyset, adLockOptimistic
    If Not RS.EOF And Not RS.BOF Then
        CboInvNo.SetFocus
        lblerror = DisplayMsg(4080)
        MousePointer = Default
        Exit Sub
    End If

    Tamp = 0
            
    
        
    sql = " INSERT INTO InvoiceSupplier_Master (Supplier_Code, Invoice_No, Invoice_Date, InvoiceReceipt_Date, Due_Date, " & _
            "BL_Date, BL_No, AirFreight_Amount, Bank_Code, Total_Amount, FakturPajak_NO, FakturPajak_Date, PPN, Exchange_Rate, Last_Update, Last_User, Total_Amount_Mat, Total_Amount_Service,PaymentVoucher_No,VoucherDesc) " & _
            "VALUES ('" & cboSupp.Text & "', '" & CboInvNo.Text & "', '" & Format(DateInv.Value, "yyyy-mm-dd") & "', '" & Format(DateRecp.Value, "yyyy-mm-dd") & "', " & _
            "'" & Format(DateDue.Value, "yyyy-mm-dd") & "', '" & Format(DateBL.Value, "yyyy-mm-dd") & "', '" & TxtBlNo.Text & "', " & CDbl(IIf(Trim(txtairfreight.Text) = "", 0, txtairfreight.Text)) & "," & _
            "'" & cbocurr.Text & "', " & CDbl(IIf(Trim(txtamount.Text) = "", 0, txtamount.Text)) & ", '" & IIf(Trim(txtFakturPajakNo.Text) = "", "", txtFakturPajakNo.Text) & "', '" & Format(DtFaktur.Value, "yyyy-mm-dd") & "', " & _
            "" & CDbl(IIf(Trim(txtPPN.Text) = "", 0, txtPPN.Text)) & ", " & Tamp & ", getdate(), '" & userLogin & "', " & _
            "" & CDbl(IIf(Trim(txtamount_material.Text) = "", 0, txtamount_material.Text)) & ", " & CDbl(IIf(Trim(txtamount_service.Text) = "", 0, txtamount_service.Text)) & ",'" & Trim(TxtBlNo.Text) & "','" & Trim(TxtVoucherDesc.Text) & "')"
            
    Db.Execute sql
   

    combo1.ListIndex = 1
    CboInvNo_Change
'    If Overseas_Cls = "1" Then
'        AddGridImport
'    ElseIf Overseas_Cls = "0" Then
        AddGridLocal
'    End If
    
    CtrCreate = 0
ElseIf BtnCreate.Caption = "Update" Then

    Header
'    If Overseas_Cls = "1" Then
'        AddGridImport
'    ElseIf Overseas_Cls = "0" Then
        AddGridLocal
'    End If
End If
MousePointer = Default
End Sub

Private Sub cbocurr_Change()
    'DONE
    RsAcc.Requery
    RsAcc.Find "Bank_code ='" & cbocurr & "'"
    If Not RsAcc.EOF Then
        TxtNmBank.Text = RsAcc!bank_name
        TxtBankAcc.Text = RsAcc!bank_Account
        rsSup.Requery
    End If
End Sub

Private Sub cbocurr_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then
KeyAscii = 0
End If
End Sub

Private Sub CboInvNo_Change()
'PROG : FILL DATE AND ANOTHER HEAD
lblerror = ""
If combo1.ListIndex = 1 Then
    Dim RsBank As New ADODB.Recordset
    
    Header
    cbocurr.Text = ""
    TxtNmBank.Text = ""
    TxtBankAcc.Text = ""
    TxtBlNo.Text = ""
    txtairfreight.Text = ""
    txtamount_material = ""
    txtamount_service = ""
    txtamount = ""
    txtFakturPajakNo = ""
    txtPPN = ""
    txtinvno = ""
    lblFix = ""
    InvDate = Date
    DtFaktur = Date     ' Tambahan 20110503
    DateInv.Value = Date
    DateRecp.Value = Date
    'DateDue.Value = Date
    DateBL.Value = Date
    
    
    sql = "SELECT          InvoiceSupplier_Master.Bank_Code, Bank_Master.Bank_Name, " & _
                " InvoiceSupplier_Master.Invoice_date, InvoiceSupplier_Master.InvoiceReceipt_Date, " & _
                " InvoiceSupplier_Master.Due_date, InvoiceSupplier_Master.BL_date, " & _
                " InvoiceSupplier_Master.BL_no, InvoiceSupplier_Master.AirFreight_Amount, " & _
                " Bank_Master.Bank_Account, InvoiceSupplier_Master.FakturPajak_No, InvoiceSupplier_Master.FakturPajak_Date,  " & _
                " InvoiceSupplier_Master.PPN, InvoiceSupplier_Master.fix_cls,VoucherDesc,exchange_Rate,Exchange_amount " & _
                "FROM            InvoiceSupplier_Master LEFT JOIN " & _
                "                Bank_Master ON " & _
                "                InvoiceSupplier_Master.Bank_Code = Bank_Master.Bank_Code " & _
                "Where           InvoiceSupplier_Master.Supplier_Code = '" & cboSupp & "' AND " & _
                "                InvoiceSupplier_Master.Invoice_no = '" & CboInvNo & "'"
                
    RsBank.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If Not RsBank.EOF And Not RsBank.BOF Then
        cbocurr.Text = RsBank!bank_Code
        TxtNmBank.Text = RsBank!bank_name & ""
        TxtBankAcc.Text = RsBank!bank_Account & ""
        InvDate = RsBank!Invoice_Date
        DateInv.Value = RsBank!Invoice_Date
        DateRecp.Value = RsBank!InvoiceReceipt_Date
        TxtRate.Text = Format(get_ExchangeRate("02"), gs_formatAmount)
        DateDue.Value = RsBank!due_date
        DateBL.Value = RsBank!BL_Date
        TxtBlNo.Text = Trim(RsBank!BL_no)
        txtairfreight.Text = RsBank!AirFreight_Amount
        txtinvno.Text = CboInvNo
        txtFakturPajakNo.Text = IIf(IsNull(RsBank!fakturpajak_no), "", Trim(RsBank!fakturpajak_no))
        DtFaktur = RsBank!fakturpajak_date      ' Tambahan 20110503
        txtPPN.Text = Format(RsBank!ppn, gs_formatAmount)
        TxtVoucherDesc.Text = IIf(IsNull(RsBank!VoucherDesc), "", Trim(RsBank!VoucherDesc))
        If Trim(RsBank!fix_cls & "") = "1" Then
            lblFix = "Status : Fixed"
        End If
    End If
    RsBank.Close
End If
End Sub

Private Sub CboInvNo_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then
KeyAscii = 0
End If
End Sub

Private Sub CboSupp_Change()
    lblerror = ""
    If CtrCreate = 0 Then
        ClearData
        CboInvNo.Text = ""
    End If
    combo1.ListIndex = 1
    txt_name.Text = ""
    Overseas_Cls = ""
    TxtBankAcc.Text = ""
    TxtNmBank.Text = ""
    ComboSuppUbah
    AddtoPO
    Header
    If cboSupp.MatchFound Then DateDue = DateAdd("d", Val(cboSupp.Column(4)), DateRecp) Else DateDue.Value = Date
End Sub

Private Sub CboSupp_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then
KeyAscii = 0
End If
End Sub

Private Sub cmdCancel_Click()
Dim suppno, invno As String
Dim dt1, dt2 As Date
suppno = cboSupp.Text
invno = CboInvNo.Text
dt1 = InvDateSt.Value
dt2 = InvDateEnd.Value
cmdClear_Click
cboSupp.Text = suppno
InvDateSt.Value = dt1
InvDateEnd.Value = dt2
CboInvNo.Text = invno
BtnCreate_Click
End Sub

Private Sub cmdClear_Click()
cboSupp.Text = ""
End Sub

Private Sub cmdSearch_Click()
    
    Dim i As Double
    
    lblerror = ""
    
    If txtSearch = "" Or grid.Rows = 2 Then txtSearch.SetFocus: Exit Sub
    'If grid.Row = grid.Rows - 1 Then i = 1 Else i = grid.Row - 1
    i = 1
    Do
        Select Case cboSearch.ListIndex
        Case 0
            grid.Col = colpcode
            If UCase(Mid(grid.TextMatrix(i, colpcode), 1, Len(txtSearch))) = UCase(txtSearch) Then
                Exit Do
            End If
        Case 1
            grid.Col = colprodname
            If InStr(UCase(grid.TextMatrix(i, colprodname)), UCase(txtSearch)) <> 0 Then
                Exit Do
            End If
        Case 2
            grid.Col = ColDoNo
            If InStr(UCase(grid.TextMatrix(i, ColDoNo)), UCase(txtSearch)) <> 0 Then
                Exit Do
            End If
            
        End Select
        i = i + 1
        
        If i > grid.Rows - 1 Then
            txtSearch = ""
            i = 2
            lblerror = DisplayMsg(8012)
            Exit Do
        End If
    Loop
    
    grid.Row = i
    grid.TopRow = i
    grid.SetFocus
    
End Sub

Private Sub CmdSubMenu_Click()
    ClearData
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub CmdSubmit_Click()
Dim i As Long
Dim Dono As String
Dim RS As New ADODB.Recordset
Dim EX As Double
Dim CtrEXc As Integer

lblerror = ""

MousePointer = vbHourglass
If hakUpdate(Me.Name) = 0 Then lblerror = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub

lblerror = up_ValidateDateRange(InvDate, True)
If lblerror <> "" Then Me.MousePointer = vbDefault: Exit Sub

If uf_checkFix(CboInvNo) = True Then lblerror = DisplayMsg(4046): Me.MousePointer = vbDefault: Exit Sub

If grid.Rows <= 1 Then
Me.MousePointer = vbDefault
Exit Sub
End If

If CtrNotAlow = 1 Then
Me.MousePointer = vbDefault
lblerror = DisplayMsg(4089)
Exit Sub
End If

CtrEXc = 0

    Db.BeginTrans
    'delete dulu semua detail dari detail
    sql = "delete from invoicesupplier_detail where supplier_code = '" & cboSupp & "' and " & _
        "invoice_no = '" & CboInvNo & "'"
    Db.Execute sql
    
    CtrCu
    
    For i = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
            CtrEXc = 1
            GoTo JOk
        End If
    Next
JOk:
    If CtrEXc = 1 Then
        If CtrCur <> "03" Then
            If RS.State = 1 Then RS.Close
            'cari nilai tukar
            RS.Open "SELECT daily_exchangerate from daily_exchangerate where exchangerate_date " & _
                    " = '" & Format(DateInv.Value, "yyyy-mm-dd") & "' and currency_code " & _
                    " = '" & CtrCur & "'", Db, adOpenKeyset, adLockOptimistic
            If RS.EOF And RS.BOF Then

                EX = 0
            Else
                EX = RS.Fields("daily_exchangerate")
                TxtRate.Text = Format(RS.Fields("daily_exchangerate"), gs_formatAmountIDR)
            End If
        Else
        EX = 1
        End If
        
    End If
    
    If EX = 0 And CtrEXc <> 0 Then
           
       lblerror.Caption = DisplayMsg("0073")
       Db.RollbackTrans
       Me.MousePointer = vbDefault
       Exit Sub
        
    End If
    
    
    For i = 1 To grid.Rows - 1
         'save karena di check kalo gak buang aja
         If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
            
'            If Overseas_Cls = 1 Then
'                Dono = ""
'            Else
                Dono = grid.TextMatrix(i, ColDoNo)
'            End If

            'perlu di cek apakah udah di insert apa belum kalo belum insert
            'tapi kalo sudah update saja
            'SAVE DI INVOICE DETAIL
            
'            If Overseas_Cls = 0 Then
                    sql = " INSERT INTO InvoiceSupplier_Detail (Supplier_Code, Invoice_No, PO_No, DO_No, BC_No, Item_Code, Qty, Unit_Cls, " & _
                        "Currency_Code, Price, Amount, Exchange_Amount, Last_Update, Last_User,ReceiptSeq_No," & _
                        "Price_Mat,Amount_Mat,Price_Service,Amount_Service) " & _
                        "VALUES ('" & cboSupp.Text & "', '" & CboInvNo.Text & "', '" & grid.TextMatrix(i, ColPONo) & "', '" & Dono & "', '" & grid.TextMatrix(i, ColBCNO) & "', '" & grid.TextMatrix(i, colpcode) & "', " & _
                        CDbl(grid.TextMatrix(i, colinvqty)) & ", '" & grid.TextMatrix(i, ColBantuUnit) & "', '" & grid.TextMatrix(i, ColBantuCurr) & "', " & CDbl(grid.TextMatrix(i, ColPrice)) & ", " & _
                        CDbl(CDbl(grid.TextMatrix(i, ColPrice)) * CDbl(grid.TextMatrix(i, colinvqty))) & ", " & _
                        CDbl((CDbl(grid.TextMatrix(i, ColPrice)) * CDbl(grid.TextMatrix(i, colinvqty))) * EX) & ", getdate(), '" & userLogin & "'," & grid.TextMatrix(i, ColReceiptSeqNo) & ", " & _
                        CDbl(grid.TextMatrix(i, colPriceMat)) & ", " & CDbl(grid.TextMatrix(i, colinvqty)) * CDbl(grid.TextMatrix(i, colPriceMat)) & ", " & _
                        CDbl(grid.TextMatrix(i, colPriceService)) & ", " & CDbl(grid.TextMatrix(i, colinvqty)) * CDbl(grid.TextMatrix(i, colPriceService)) & ") "
'            Else
'                    sql = " INSERT INTO InvoiceSupplier_Detail (Supplier_Code, Invoice_No, PO_No, DO_No, Item_Code, Qty, Unit_Cls, " & _
'                        "Currency_Code, Price, Amount, Exchange_Amount, Last_Update, Last_User,ReceiptSeq_No) " & _
'                        "VALUES ('" & CboSupp.Text & "', '" & CboInvNo.Text & "', '" & Grid.TextMatrix(i, ColPONo) & "', '" & Dono & "', '" & Grid.TextMatrix(i, colpcode) & "', " & _
'                        CDbl(Grid.TextMatrix(i, colinvqty)) & ", '" & Grid.TextMatrix(i, ColBantuUnit) & "', '" & Grid.TextMatrix(i, ColBantuCurr) & "', " & CDbl(Grid.TextMatrix(i, colPrice)) & ", " & _
'                        CDbl(CDbl(Grid.TextMatrix(i, colPrice)) * CDbl(Grid.TextMatrix(i, colinvqty))) & ", " & _
'                        CDbl((CDbl(Grid.TextMatrix(i, colPrice)) * CDbl(Grid.TextMatrix(i, colinvqty))) * EX) & ", getdate(), '" & userLogin & "'," & Grid.TextMatrix(i, ColReceiptSeqNo) & ")"
'            End If
            Db.Execute sql
            
         End If
    Next
    
    'save header lagi
    txtamount.Text = 0
    txtamount_material.Text = 0
    txtamount_service.Text = 0
    For i = 0 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
            txtamount.Text = Format(CDbl(txtamount.Text) + CDbl(grid.TextMatrix(i, ColAmount)), gs_formatAmount)
            txtamount_material.Text = Format(CDbl(txtamount_material.Text) + CDbl(grid.TextMatrix(i, ColAmountMat)), gs_formatAmount)
            txtamount_service.Text = Format(CDbl(txtamount_service.Text) + CDbl(grid.TextMatrix(i, ColAmountService)), gs_formatAmount)
        End If
    Next
    
    txtamount.Text = Format(CDbl(txtamount.Text), gs_formatAmount)
    txtamount_material.Text = Format(CDbl(txtamount_material.Text), gs_formatAmount)
    txtamount_service.Text = Format(CDbl(txtamount_service.Text), gs_formatAmount)
            
    
    If RS.State = 1 Then RS.Close
    RS.Open "Select * from invoicesupplier_master where supplier_code = '" & cboSupp.Text & "' and invoice_no = '" & CboInvNo & "'", Db, adOpenKeyset, adLockOptimistic
    If RS.EOF And RS.BOF Then
        'imposible seh
        lblerror = lblerror = DisplayMsg("0057")
        Db.RollbackTrans
        Exit Sub
    Else
        RS.Fields("Invoice_Date") = Format(DateInv, "yyyy-mm-dd")
        RS.Fields("InvoiceReceipt_Date") = Format(DateRecp, "yyyy-mm-dd")
        RS.Fields("Due_date") = Format(DateDue, "yyyy-mm-dd")
        RS.Fields("BL_date") = Format(DateBL, "yyyy-mm-dd")
        RS.Fields("BL_No") = TxtBlNo.Text
        RS.Fields("VoucherDesc") = Trim(TxtVoucherDesc.Text)
        RS.Fields("AirFreight_Amount") = CDbl(IIf(Trim(txtairfreight.Text) = "", 0, txtairfreight.Text))
        RS.Fields("Total_Amount") = CDbl(IIf(Trim(txtamount.Text) = "", 0, txtamount.Text))
        RS.Fields("Exchange_Rate") = EX
        RS.Fields("Exchange_amount") = EX * (CDbl(IIf(Trim(txtairfreight.Text) = "", 0, txtairfreight.Text)) + CDbl(IIf(Trim(txtamount.Text) = "", 0, txtamount.Text)))
        RS.Fields("FakturPajak_No") = txtFakturPajakNo.Text
        RS.Fields("FakturPajak_Date") = Format(DtFaktur, "yyyy-mm-dd")                    ' Tambahan 20110503
        RS.Fields("PPN") = CDbl(IIf(Trim(txtPPN.Text) = "", 0, txtPPN.Text))
        RS.Fields("Bank_Code") = cbocurr.Text
        RS.Fields("Last_Update") = Date
        RS.Fields("Last_User") = userLogin
        RS.Fields("Total_Amount_Mat") = CDbl(IIf(Trim(txtamount_material.Text) = "", 0, txtamount_material.Text))
        RS.Fields("Total_Amount_Service") = CDbl(IIf(Trim(txtamount_service.Text) = "", 0, txtamount_service.Text))
        
        RS.update
    End If
    
    Db.CommitTrans
    Header
'    If Overseas_Cls = "1" Then
'        AddGridImport
'    ElseIf Overseas_Cls = "0" Then
        AddGridLocal
'    End If
    lblerror = DisplayMsg(1000)
    MousePointer = Default
End Sub
Private Function uf_checkFix(ls_InvoiceNo As String) As Boolean
Dim RS As New ADODB.Recordset
If RS.State <> adStateClosed Then RS.Close
RS.CursorLocation = adUseClient
RS.Open " select fix_cls from invoiceSupplier_master where Invoice_no='" & Trim(ls_InvoiceNo) & "'", Db, adOpenKeyset, adLockOptimistic
If RS.EOF = False Then
    If Trim(RS!fix_cls & "") = "1" Then
        uf_checkFix = True
    Else
        uf_checkFix = False
    End If
Else
    uf_checkFix = False
End If
If RS.State <> adStateClosed Then RS.Close
End Function

Private Sub Combo1_Click()
lblerror = ""
If combo1.ListIndex = 1 Then
    BtnCreate.Caption = "Update"
    If CtrCreate = 0 Then
        InitSave
    End If
    CboSupp_Change
ElseIf combo1.ListIndex = 0 Then
    ClearData
    BtnCreate.Caption = "Create"
    InitSave
    Header
    txtinvno.Text = ""
    txtamount.Text = ""
    txtamount_material.Text = ""
    txtamount_service.Text = ""
    txtPPN.Text = ""
    txtFakturPajakNo.Text = ""
End If
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    lblerror.Caption = ErrMsg
End If
End Sub

Private Function get_ExchangeRate(Curr As String) As Double

Dim RaRate As New ADODB.Recordset

If Curr <> "03" Then
 RaRate.Open "SELECT daily_exchangerate from daily_exchangerate where exchangerate_date " & _
                    " = '" & Format(DateInv.Value, "yyyy-mm-dd") & "' and currency_code " & _
                    " = '" & Curr & "'", Db, adOpenKeyset, adLockOptimistic
            If RaRate.EOF And RaRate.BOF Then

                get_ExchangeRate = 0
            Else
                
                get_ExchangeRate = Format(RaRate.Fields("daily_exchangerate"), gs_formatAmountIDR)
            End If
Else
    get_ExchangeRate = 1
End If


End Function



Private Sub DateRecp_Change()
    Dim invTgl As Integer
    Dim DueDate As Integer
    
    invTgl = Format(DateRecp.Value, "dd")
    
    If invTgl >= 14 And invTgl <= 28 Then
        DueDate = 25
    Else
        DueDate = 10
    End If
    
    DateDue = DateAdd("M", 1, DateRecp)
    DateDue = Month(DateDue) & "/" & DueDate & "/" & Year(DateDue)
    
'    If cboSupp.MatchFound Then DateDue = DateAdd("d", Val(cboSupp.Column(4)), DateRecp) Else DateDue.Value = Date
    TxtRate.Text = Format(get_ExchangeRate("02"), gs_formatAmountIDR)
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    HakU = hakUpdate(Me.Name)
    bteHakPrice = hakPrice(Me.Name)
    Label2.Visible = (bteHakPrice = 1)
    txtamount.Visible = (bteHakPrice = 1)
    Label10.Visible = (bteHakPrice = 1)
    txtamount_material.Visible = (bteHakPrice = 1)
    Label11.Visible = (bteHakPrice = 1)
    txtamount_service.Visible = (bteHakPrice = 1)
'    Label5.Visible = (bteHakPrice = 1)
'    txtPPn.Visible = (bteHakPrice = 1)
    lblFix = ""
    DateDue.Value = Date
    DtFaktur = Date
    CtrNotAlow = 0
    adtocombo
    Init
    Header
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then
    If grid.Cell(flexcpChecked, Row, Col) = 1 Then
        grid.TextMatrix(Row, colqtyasli) = grid.TextMatrix(Row, colinvqty)
        grid.TextMatrix(Row, ColPoQtyReal) = 0
    Else
        grid.TextMatrix(Row, colinvqty) = Format(CDbl(grid.TextMatrix(Row, colqtyasli)) + CDbl(grid.TextMatrix(Row, ColPoQtyReal)), gs_formatQty)
        grid.TextMatrix(Row, colqtyasli) = grid.TextMatrix(Row, colinvqty)
        grid.TextMatrix(Row, ColPoQtyReal) = grid.TextMatrix(Row, colinvqty)
    End If
    GridUbahOk Row
    CtrCur = ""
ElseIf Col = colinvqty Then
    GridUbahOk Row
    CtrCur = ""
End If
End Sub

Sub CtrCu()
CtrCur = ""
For i = 1 To grid.Rows - 1
    If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
        CtrCur = grid.TextMatrix(i, ColBantuCurr)
        Exit Sub
    End If
Next
End Sub

Sub GridUbahOk(lngRow As Long)
If Not IsNumeric(grid.TextMatrix(lngRow, colinvqty)) Then grid.TextMatrix(lngRow, colinvqty) = 0
grid.TextMatrix(lngRow, colinvqty) = Format(grid.TextMatrix(lngRow, colinvqty), gs_formatQty)

'INI BISA INPUT BERAPA AJA
'    If Overseas_Cls = 1 Then
'        Grid.TextMatrix(lngRow, ColAmount) = Format(Grid.TextMatrix(lngRow, colinvqty) * Grid.TextMatrix(lngRow, colPrice), gs_formatAmount)
'        Grid.TextMatrix(lngRow, ColRem) = Format((CDbl(Grid.TextMatrix(lngRow, colqtyasli)) + CDbl(Grid.TextMatrix(lngRow, ColPoQtyReal))) - CDbl(Grid.TextMatrix(lngRow, colinvqty)), gs_formatQty)
'        If CDbl(Grid.TextMatrix(lngRow, ColRem)) < 0 Then Grid.TextMatrix(lngRow, ColRem) = Format(0, gs_formatQty)
'    ElseIf Overseas_Cls = 0 Then
        grid.TextMatrix(lngRow, ColAmount) = Format(grid.TextMatrix(lngRow, colinvqty) * grid.TextMatrix(lngRow, ColPrice), gs_formatAmount)
        grid.TextMatrix(lngRow, colrem) = Format((CDbl(grid.TextMatrix(lngRow, colqtyasli)) + CDbl(grid.TextMatrix(lngRow, ColPoQtyReal))) - CDbl(grid.TextMatrix(lngRow, colinvqty)), gs_formatQty)
        If CDbl(grid.TextMatrix(lngRow, colrem)) < 0 Then grid.TextMatrix(lngRow, colrem) = Format(0, gs_formatQty)
'    End If
    
    grid.TextMatrix(lngRow, ColAmount) = Format(grid.TextMatrix(lngRow, colinvqty) * grid.TextMatrix(lngRow, ColPrice), gs_formatAmount)
    grid.TextMatrix(lngRow, ColAmountMat) = Format(grid.TextMatrix(lngRow, colinvqty) * grid.TextMatrix(lngRow, colPriceMat), gs_formatAmount)
    grid.TextMatrix(lngRow, ColAmountService) = Format(grid.TextMatrix(lngRow, colinvqty) * grid.TextMatrix(lngRow, colPriceService), gs_formatAmount)
    
    txtamount.Text = 0
    txtamount_material.Text = 0
    txtamount_service.Text = 0
    For i = 0 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
            txtamount.Text = CDbl(txtamount.Text) + CDbl(grid.TextMatrix(i, ColAmount))
            txtamount_material.Text = CDbl(txtamount_material.Text) + CDbl(grid.TextMatrix(i, ColAmountMat))
            txtamount_service.Text = CDbl(txtamount_service.Text) + CDbl(grid.TextMatrix(i, ColAmountService))
        End If
    Next
    
            txtamount.Text = Format(CDbl(txtamount.Text), gs_formatAmountIDR)
            txtamount_material.Text = Format(CDbl(txtamount_material.Text), gs_formatAmount)
            txtamount_service.Text = Format(CDbl(txtamount_service.Text), gs_formatAmount)
    
    
    If txtamount.Text > 0 And txtamount.Text <> "" And RateOri > 0 And TxtRate.Text > 0 And TxtRate.Text <> "" Then
        txtkonversi.Text = Format((CDbl(txtamount) * CDbl(RateOri)) / CDbl(TxtRate.Text), gs_formatAmountIDR)
    End If
    
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim X As String
lblerror = ""

If CtrNotAlow <> 1 Then
    If grid.Col <> 0 And grid.Col <> colinvqty Then
        Cancel = True
        Exit Sub
    Else
        tampung = grid.TextMatrix(Row, colinvqty)
        rema = grid.TextMatrix(Row, colrem)
        If grid.Col = colinvqty Then
            If grid.Cell(flexcpChecked, Row, 0) = flexUnchecked Then
                Cancel = True
                Exit Sub
            End If
        Else
            X = grid.TextMatrix(Row, ColBantuCurr)
            For i = 1 To grid.Rows - 1
                If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
                    If X <> grid.TextMatrix(i, ColBantuCurr) Then
                        lblerror = DisplayMsg(4084)
                        Cancel = True
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If
Else
    lblerror = DisplayMsg(4089)
    Cancel = True
End If
End Sub

Private Sub grid_Click()
If grid.Col = 0 Then
    grid.FocusRect = flexFocusInset
ElseIf grid.Col = colinvqty Then
    If CtrNotAlow <> 1 Then
        grid.TextMatrix(grid.Row, colinvqty) = Format(CDbl(tampung), gs_formatQty)
        grid.FocusRect = flexFocusInset
    End If
Else
    grid.FocusRect = flexFocusNone
End If
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
lblerror = ""
  If grid.Col = colinvqty Then
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
  End If
End Sub

Private Sub InvDateEnd_Change()
If combo1.ListIndex <> 0 Then
    ClearData
    InvAdd
    AddtoPO
    If CboInvNo.MatchFound = False Then CboInvNo = ""
End If
End Sub

Private Sub InvDateSt_Change()
If combo1.ListIndex <> 0 Then
    ClearData
    InvAdd
    AddtoPO
    If CboInvNo.MatchFound = False Then CboInvNo = ""
End If
End Sub

Private Sub txtairfreight_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
          KeyAscii = 0
    End If
    If (txtairfreight.Text & Chr(KeyAscii)) > 999999999.99 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub


'================================================================================================================================================================================================

Sub Init()
    
    ColCls = 0
    ColPONo = 1
    colpodate = 2
    ColDoNo = 3
    ColBCNO = 4
    colpcode = 4 + 1
    colprodname = 5 + 1
    ColUnit = 6 + 1
    colpoqty = 7 + 1
    coldoqty = 8 + 1
    colinvqty = 9 + 1
    colrem = 10 + 1
    ColCurr = 11 + 1
    ColPrice = 16 + 1
    ColAmount = 17 + 1
    colPriceMat = 14 + 1
    ColAmountMat = 15 + 1
    colPriceService = 12 + 1
    ColAmountService = 13 + 1
    ColBantuUnit = 18 + 1
    ColBantuCurr = 19 + 1
    ColCurX = 20 + 1
    ColPoQtyReal = 21 + 1
    colqtyasli = 22 + 1
    ColReceiptSeqNo = 23 + 1

    CtrCreate = 0
    combo1.clear
    combo1.AddItem "Create"
    combo1.AddItem "Update"
    combo1.ListIndex = 1
    txt_name.Text = ""
    InvDateSt.Value = Date
    InvDateEnd.Value = Date
    InitSave
    Overseas_Cls = 0
    txtinvno.Text = ""
    txtamount.Text = ""
    txtamount_material.Text = ""
    txtamount_service.Text = ""
    txtPPN.Text = ""
    txtFakturPajakNo.Text = ""
    tampung = 0
    rema = 0

End Sub

Sub InitSave()
    CboInvNo = ""
    CboInvNo.clear
    cbocurr.Text = ""
    TxtBankAcc.Text = ""
    TxtNmBank.Text = ""
    TxtBlNo.Text = ""
    txtairfreight.Text = ""
    DateBL.Value = Date
    'DateDue.Value = Date
    InvDate = Date
    DateInv.Value = Date
    DateRecp.Value = Date
End Sub

Sub Conec()
    Db.Open "Provider=SQLOLEDB.1;Persist Security Info=true;User ID=sa;Initial Catalog=Santelindo_real;Data Source=.;pwd=;"
End Sub

Sub AddGridLocal()
        
    Dim Currancy_Code As String
    Dim i As Long
    Dim RsLC As New ADODB.Recordset
    Dim RSX As New ADODB.Recordset
    
    sql = "declare @InvoiceNo as char(50) " & vbCrLf & _
        "declare @SupplierCode as char(50) " & vbCrLf & _
        "declare @StartDate as char(10) " & vbCrLf & _
        "declare @EndDate as char(10) " & vbCrLf & _
        "set @InvoiceNo = '" & CboInvNo & "' " & vbCrLf & _
        "set @SupplierCode = '" & cboSupp & "' " & vbCrLf & _
        "set @StartDate = '" & Format(InvDateSt, "yyyy-MM-dd") & "' " & vbCrLf & _
        "set @EndDate = '" & Format(InvDateEnd, "yyyy-MM-dd") & "' "
    
    sql = sql & vbCrLf & _
        "select distinct * from ( " & vbCrLf & _
            "select inv_d.PO_No, Part_Receipt.Receipt_Date, inv_d.Item_Code, im.Item_Name, inv_d.Unit_Cls, Isnull(pd.Qty, 0) as Po_Qty, " & vbCrLf & _
            "( " & vbCrLf & _
                "select sum(isnull(qty,0) )from  Part_Receipt where " & vbCrLf & _
                "inv_d.Supplier_Code = Part_Receipt.Supplier_Code " & vbCrLf & _
                "and case when Part_Receipt.PO_No = '0' then '0' else inv_d.PO_No end = Part_Receipt.PO_No and inv_d.Item_Code = Part_Receipt.Item_Code " & vbCrLf & _
                "and inv_d.ReceiptSeq_No = Part_Receipt.Seq_No and inv_d.Do_No = Part_Receipt.SuratJalan_No " & vbCrLf & _
            ") as do_qty, " & vbCrLf & _
            "inv_d.Currency_Code, inv_d.Price, " & vbCrLf & _
            "inv_d.Supplier_Code as Trade_code, inv_d.Qty as QtyX, " & vbCrLf & _
            "inv_d.do_no , inv_d.BC_no , inv_d.ReceiptSeq_No, " & vbCrLf & _
            "inv_d.Price_Mat , inv_d.Price_Service " & vbCrLf & _
            "from (SELECT * fROM InvoiceSupplier_Detail where Invoice_No = @InvoiceNo and Supplier_Code = @SupplierCode ) inv_d " & vbCrLf & _
            "left join PurchaseOrder_Master pm on inv_d.PO_No = pm.PO_No " & vbCrLf & _
            "left join Item_Master im on   inv_d.Item_Code = im.Item_Code " & vbCrLf & _
            "left join PurchaseOrder_Detail pd on pm.PO_No = pd.PO_No and inv_d.Item_Code = pd.Item_Code and inv_d.PO_No = pd.PO_No " & vbCrLf & _
            "left join Part_Receipt on inv_d.Supplier_Code = Part_Receipt.Supplier_Code " & vbCrLf & _
            "and case when Part_Receipt.PO_No = '0' then '0' else inv_d.PO_No end = Part_Receipt.PO_No " & vbCrLf & _
            "and inv_d.Item_Code = Part_Receipt.Item_Code " & vbCrLf & _
            "and inv_d.ReceiptSeq_No = Part_Receipt.Seq_No " & vbCrLf & _
            "and inv_d.Do_No = Part_Receipt.SuratJalan_No " & vbCrLf & _
            " "
    
    sql = sql & vbCrLf & _
            "union all " & vbCrLf & _
            "select * from ( " & vbCrLf & _
                "select pm.PO_No, Part_Receipt.Receipt_Date, pd.Item_Code, im.Item_Name, " & vbCrLf & _
                "pd.Unit_Cls,  pd.Qty as Po_Qty, Part_Receipt.Qty as do_qty, pd.Currency_Code, " & vbCrLf & _
                "(case when isnull(pd.price_Adj,0)=0 then Isnull(part_receipt.Price, 0) + Isnull(part_receipt.Price_Service, 0) else pd.price_Adj end) Price, " & vbCrLf & _
                "tm.Trade_Code, 0 as QtyX, Part_Receipt.SuratJalan_No as Do_no, Part_Receipt.BC40_No as BC_No," & vbCrLf & _
                "Part_Receipt.Seq_No ReceiptSeq_No, " & vbCrLf & _
                "isnull(pd.Price, 0) Price_Mat, isnull(pd.Price_Service, 0) Price_Service " & vbCrLf & _
                "from (Select * From PurchaseOrder_Master Where Supplier_Code=@SupplierCode and Fix_Cls='1')pm " & vbCrLf & _
                "left join PurchaseOrder_Detail pd on  pm.PO_No = pd.PO_No " & vbCrLf & _
                "left join Trade_Master tm on  pm.Supplier_Code = tm.Trade_Code " & vbCrLf & _
                "left join Item_Master im on  pd.Item_Code = im.Item_Code " & vbCrLf & _
                "inner join (SELECT * fROM Part_Receipt where Receipt_Date >= @StartDate  and Receipt_Date <= @EndDate )Part_Receipt on  pm.PO_No = Part_Receipt.PO_No and  pd.Item_Code = Part_Receipt.Item_Code  " & vbCrLf & _
                "" & vbCrLf & _
                " " & vbCrLf & _
                " " & vbCrLf & _
                " "
    
    sql = sql & vbCrLf & _
                "union all " & vbCrLf & _
                "select pr.Remarks PO_No, pr.Receipt_Date, pr.Item_code, im.Item_Name, pr.Unit_Cls, 0 as Po_Qty, pr.Qty as do_qty, pr.Currency_Code, " & vbCrLf & _
                "price, tm.Trade_Code, 0 as QtyX, pr.SuratJalan_No as Do_no, pr.BC40_No as BC_No, pr.Seq_No  ReceiptSeq_No, " & vbCrLf & _
                "Price_Mat = isnull(( " & vbCrLf & _
                    "select price " & vbCrLf & _
                    "From price_master " & vbCrLf & _
                    "where price_cls = '01' and item_code = pr.item_code and trade_code = pr.supplier_code and priority_cls = 1 " & vbCrLf & _
                    "and start_date <= convert(varchar, pr.receipt_date, 112) and end_date >= convert(varchar, pr.receipt_date, 112)), 0), " & vbCrLf & _
                "Price_Service = isnull(( " & vbCrLf & _
                    "select price " & vbCrLf & _
                    "From price_master " & vbCrLf & _
                    "where price_cls = '05' and item_code = pr.item_code and trade_code = pr.supplier_code and priority_cls = 1 " & vbCrLf & _
                    "and start_date <= convert(varchar, pr.receipt_date, 112) and end_date >= convert(varchar, pr.receipt_date, 112)), 0) " & vbCrLf & _
                "from (SELECT * fROM Part_Receipt where Receipt_Date >= @StartDate  and Receipt_Date <= @EndDate and Supplier_Code=@SupplierCode and  PO_No = '0' and Receipt_Cls = 'R' )  pr  " & vbCrLf & _
                "inner join Item_Master im on pr.Item_Code = im.Item_Code " & vbCrLf & _
                "inner join Trade_Master tm on pr.Supplier_Code = tm.Trade_Code " & vbCrLf & _
                "--where PO_No = '0' and Receipt_Cls = 'R' " & vbCrLf & _
                "--and (tm.Trade_Code = @SupplierCode) " & vbCrLf & _
                "--and (pr.Receipt_Date >= @StartDate) " & vbCrLf & _
                "--and (pr.Receipt_Date <= @EndDate) "
                
    sql = sql & vbCrLf & _
            ") tb " & vbCrLf & _
            "where do_no <> '' and rtrim(tb.PO_NO) + rtrim(tb.Trade_code) + rtrim(tb.item_code) + rtrim(Do_No) + convert(char(10), receiptSeq_No) " & vbCrLf & _
            "not in ( " & vbCrLf & _
                "select rtrim(PO_No) + rtrim(supplier_code) + rtrim(item_code) + rtrim(Do_No) + convert(char(10), receiptSeq_No) " & vbCrLf & _
                "From InvoiceSupplier_Detail " & vbCrLf & _
                "where invoice_no = @InvoiceNo  and supplier_code = @SupplierCode) " & vbCrLf & _
        ") a "
        
    ' Add for KAWAI 20090623
    If Trim(CboPOnO) <> strAll Then sql = sql & " Where DO_NO='" & Trim(CboPOnO) & "' "
    
    sql = sql & " order by Receipt_Date, PO_No, DO_No, Item_Code"

    
    RsLC.Open sql, Db, adOpenKeyset, adLockOptimistic
        
    i = 1
    Do Until RsLC.EOF
        With grid
            .Rows = .Rows + 1
            If RsLC.Fields("qtyX") <> 0 Then
                .Cell(flexcpChecked, i, 0) = flexChecked
            Else
                .Cell(flexcpChecked, i, 0) = flexUnchecked
            End If
            .TextMatrix(i, ColPONo) = RsLC.Fields("PO_No")
            .TextMatrix(i, colpodate) = Format(RsLC.Fields("Receipt_Date"), "dd MMM yyyy")
            .TextMatrix(i, ColDoNo) = IIf(IsNull(RsLC.Fields("Do_No")), "", RsLC.Fields("Do_No"))
            .TextMatrix(i, ColBCNO) = IIf(IsNull(RsLC.Fields("BC_No")), "", RsLC.Fields("BC_No"))
            .TextMatrix(i, colpcode) = RsLC.Fields("Item_Code")
            .TextMatrix(i, colprodname) = RsLC.Fields("Item_Name")
            .TextMatrix(i, ColUnit) = uf_GetUnitDescription(Trim(RsLC.Fields("Unit_Cls")))
            
            If RSX.State = 1 Then RSX.Close
            RSX.Open "SELECT isnull(sum(qty),0) as Q from invoicesupplier_detail where " & _
            " po_no = '" & RsLC.Fields("PO_No") & "' and item_code = '" & RsLC.Fields("Item_Code") & "' and " & _
            " do_no = '" & RsLC.Fields("do_No") & "'", Db, adOpenKeyset, adLockOptimistic
            .TextMatrix(i, colpoqty) = Format(RsLC.Fields("Po_Qty"), gs_formatQty)
            .TextMatrix(i, coldoqty) = Format(IIf(IsNull(RsLC.Fields("Do_Qty")), "0", RsLC.Fields("Do_Qty")), gs_formatQty)
            .TextMatrix(i, ColPoQtyReal) = .TextMatrix(i, coldoqty) - RSX.Fields("Q")
            .TextMatrix(i, colrem) = Format(.TextMatrix(i, coldoqty) - RSX.Fields("Q"), gs_formatQty)
            If CDbl(.TextMatrix(i, colrem)) < 0 Then .TextMatrix(i, colrem) = Format(0, gs_formatQty)
           
           ' request jika pada saat posisi create jika sudah remaining 0 tidak usah di tampilkan  20170516
           If status = "Create" Then
           
                If CDbl(.TextMatrix(i, colrem)) = 0 Then
                
                    .RowHidden(i) = True
                
                End If
           
           End If
            
            
            If RSX.State = 1 Then RSX.Close
            RSX.Open "select qty from invoicesupplier_detail where po_no = '" & RsLC.Fields("PO_No") & "' " & _
            " and item_code = '" & RsLC.Fields("Item_Code") & "' and invoice_no = '" & CboInvNo & "' and " & _
            " do_no = '" & RsLC.Fields("do_No") & "' and receiptSeq_no ='" & RsLC.Fields("receiptSeq_no") & "'", Db, adOpenKeyset, adLockOptimistic
            If RSX.EOF And RSX.BOF Then
                .TextMatrix(i, colinvqty) = .TextMatrix(i, colrem)
                .TextMatrix(i, colqtyasli) = CDbl(.TextMatrix(i, colrem))
            Else
                .TextMatrix(i, colinvqty) = Format(RSX.Fields("qty"), gs_formatQty)
                .TextMatrix(i, colqtyasli) = RSX.Fields("qty")
            End If
            
            
            .TextMatrix(i, ColCurr) = uf_GetCurrencyDescription(Trim(RsLC.Fields("Currency_Code")))
            If Trim(RsLC.Fields("Currency_Code")) = "03" Then
                .TextMatrix(i, ColPrice) = Format(RsLC.Fields("Price"), gs_formatPriceIDR)
                .TextMatrix(i, ColAmount) = Format(.TextMatrix(i, colinvqty) * RsLC.Fields("Price"), gs_formatAmountIDR)
                
                .TextMatrix(i, colPriceMat) = Format(RsLC.Fields("Price_Mat"), gs_formatPriceIDR)
                .TextMatrix(i, ColAmountMat) = Format(.TextMatrix(i, colinvqty) * RsLC.Fields("Price_Mat"), gs_formatAmount)
                
                .TextMatrix(i, colPriceService) = Format(RsLC.Fields("Price_Service"), gs_formatPriceIDR)
                .TextMatrix(i, ColAmountService) = Format(.TextMatrix(i, colinvqty) * RsLC.Fields("Price_Service"), gs_formatAmountIDR)
            Else
                .TextMatrix(i, ColPrice) = Format(RsLC.Fields("Price"), gs_formatPrice)
                .TextMatrix(i, ColAmount) = Format(.TextMatrix(i, colinvqty) * RsLC.Fields("Price"), gs_formatAmount)
                
                .TextMatrix(i, colPriceMat) = Format(RsLC.Fields("Price_Mat"), gs_formatPrice)
                .TextMatrix(i, ColAmountMat) = Format(.TextMatrix(i, colinvqty) * RsLC.Fields("Price_Mat"), gs_formatAmount)
                
                .TextMatrix(i, colPriceService) = Format(RsLC.Fields("Price_Service"), gs_formatPrice)
                .TextMatrix(i, ColAmountService) = Format(.TextMatrix(i, colinvqty) * RsLC.Fields("Price_Service"), gs_formatAmount)
            End If
            
            .TextMatrix(i, ColBantuUnit) = RsLC.Fields("Unit_Cls")
            .TextMatrix(i, ColBantuCurr) = RsLC.Fields("Currency_Code")
            .TextMatrix(i, ColReceiptSeqNo) = Trim(RsLC!ReceiptSeq_No & "")
            Currancy_Code = RsLC.Fields("Currency_Code")
            
             AloEdit (i)
            i = i + 1
            RsLC.MoveNext
        End With
    Loop
    grid.Cell(flexcpAlignment, 0, ColUnit, grid.Rows - 1, ColUnit) = flexAlignCenterCenter
    grid.Cell(flexcpAlignment, 0, ColCurr, grid.Rows - 1, ColCurr) = flexAlignCenterCenter
    
    txtamount.Text = 0
    txtamount_material.Text = 0
    txtamount_service.Text = 0
    
    RateOri = Format(get_ExchangeRate(Currancy_Code), gs_formatAmount)
        
    For i = 0 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
            txtamount.Text = CDbl(txtamount.Text) + CDbl(grid.TextMatrix(i, ColAmount))
            txtamount_material.Text = CDbl(txtamount_material.Text) + CDbl(grid.TextMatrix(i, ColAmountMat))
            txtamount_service.Text = CDbl(txtamount_service.Text) + CDbl(grid.TextMatrix(i, ColAmountService))
        End If
    Next
    
        txtamount.Text = Format(CDbl(txtamount.Text), gs_formatAmountIDR)
        txtamount_material.Text = Format(CDbl(txtamount_material.Text), gs_formatAmount)
        txtamount_service.Text = Format(CDbl(txtamount_service.Text), gs_formatAmount)

     If txtamount.Text > 0 And txtamount.Text <> "" And RateOri > 0 And TxtRate.Text > 0 And TxtRate.Text <> "" Then
    
        txtkonversi.Text = Format((CDbl(txtamount) * CDbl(RateOri)) / CDbl(TxtRate.Text), gs_formatAmountIDR)
    End If
End Sub

'Sub AddGridImport()
'
'    Dim i As Long
'    Dim RsIm As New ADODB.Recordset
'    Dim RSX As New ADODB.Recordset
'    Dim EX As Double
'
'    sql = "Declare @InvoiceNo as char(50) " & _
'        "Declare @SupplierCode as char(50) " & _
'        "Declare @StartDate as char(10) " & _
'        "Declare @EndDate as char(10) " & _
'        "set @InvoiceNo = '" & CboInvNo & "' " & _
'        "set @SupplierCode = '" & CboSupp & "' " & _
'        "set @StartDate = '" & Format(InvDateSt, "yyyy-MM-dd") & "' " & _
'        "set @EndDate = '" & Format(InvDateEnd, "yyyy-MM-dd") & "' "
'
'    sql = sql & _
'        "SELECT InvoiceSupplier_Detail.PO_No, Isnull(PurchaseOrder_Master.PO_Date, pr.Receipt_Date) PO_Date, InvoiceSupplier_Detail.Item_Code, " & _
'        "Item_Master.Item_Name, InvoiceSupplier_Detail.Unit_Cls, Isnull(PurchaseOrder_Detail.Qty, pr.Qty) AS Po_Qty, " & _
'        "InvoiceSupplier_Detail.Currency_Code, InvoiceSupplier_Detail.Price, InvoiceSupplier_Detail.Amount, " & _
'        "InvoiceSupplier_Detail.Supplier_Code AS Trade_code, InvoiceSupplier_Detail.Qty AS QtyX, ReceiptSeq_No " & _
'        "From InvoiceSupplier_Detail " & _
'        "Left JOIN PurchaseOrder_Master ON InvoiceSupplier_Detail.PO_No = PurchaseOrder_Master.PO_No " & _
'        "Left JOIN Item_Master ON InvoiceSupplier_Detail.Item_Code = Item_Master.Item_Code " & _
'        "Left JOIN PurchaseOrder_Detail ON PurchaseOrder_Master.PO_No = PurchaseOrder_Detail.PO_No " & _
'        "AND InvoiceSupplier_Detail.Item_Code = PurchaseOrder_Detail.Item_Code " & _
'        "AND InvoiceSupplier_Detail.PO_No = PurchaseOrder_Detail.PO_No " & _
'        "Left JOIN Part_Receipt pr ON InvoiceSupplier_Detail.Supplier_Code = pr.Supplier_Code " & _
'        "AND CASE WHEN pr.PO_No = '0' THEN '0' else InvoiceSupplier_Detail.PO_No END = pr.PO_No " & _
'        "AND InvoiceSupplier_Detail.Item_Code = pr.Item_Code " & _
'        "AND InvoiceSupplier_Detail.ReceiptSeq_No = pr.Seq_No " & _
'        "WHERE InvoiceSupplier_Detail.Invoice_No = @InvoiceNo " & _
'        "AND InvoiceSupplier_Detail.Supplier_Code = @SupplierCode "
'
'    sql = sql & _
'        "Union All " & _
'        "SELECT * FROM " & _
'        "( " & _
'            "SELECT PurchaseOrder_Master.PO_No, PurchaseOrder_Master.PO_Date, PurchaseOrder_Detail.Item_Code, " & _
'            "Item_Master.Item_Name, PurchaseOrder_Detail.Unit_Cls, PurchaseOrder_Detail.Qty AS Po_Qty, " & _
'            "PurchaseOrder_Detail.Currency_Code, PurchaseOrder_Detail.Price, PurchaseOrder_Detail.Amount, " & _
'            "Trade_Master.Trade_Code, 0 AS QtyX, 0 Seq_No " & _
'            "From PurchaseOrder_Master " & _
'            "Left JOIN PurchaseOrder_Detail ON PurchaseOrder_Master.PO_No = PurchaseOrder_Detail.PO_No " & _
'            "Left JOIN Trade_Master ON PurchaseOrder_Master.Supplier_Code = Trade_Master.Trade_Code " & _
'            "Left JOIN Item_Master ON PurchaseOrder_Detail.Item_Code = Item_Master.Item_Code " & _
'            "WHERE Trade_Master.Trade_Code = @SupplierCode " & _
'            "AND PurchaseOrder_Master.PO_Date >= @StartDate " & _
'            "AND PurchaseOrder_Master.PO_Date <= @EndDate " & _
'            "AND PurchaseOrder_Master.fix_cls = '1' " & _
'            "AND PurchaseOrder_Detail.Qty > " & _
'            "( " & _
'                "SELECT isnull(sum(Qty), 0) from invoicesupplier_detail " & _
'                "Where PO_No = PurchaseOrder_Master.PO_No And Item_Code = PurchaseOrder_Detail.Item_Code " & _
'            ") "
'
'    sql = sql & _
'            "Union All " & _
'            "SELECT pr.Remarks PO_No, pr.Receipt_Date PO_Date, pr.Item_Code, im.Item_Name, pr.Unit_Cls, pr.Qty Po_Qty, " & _
'            "pr.Currency_Code, pr.Price, pr.Amount, tm.Trade_Code, 0 As QtyX, pr.Seq_No " & _
'            "FROM Part_Receipt pr " & _
'            "INNER JOIN Item_Master im ON pr.Item_Code = im.Item_Code " & _
'            "INNER JOIN Trade_Master tm ON pr.Supplier_Code = tm.Trade_Code " & _
'            "WHERE pr.Receipt_Cls = 'R' AND pr.PO_No = '0' " & _
'            "AND pr.Supplier_Code = @SupplierCode " & _
'            "AND pr.Receipt_Date >= @StartDate " & _
'            "AND pr.Receipt_Date <= @EndDate " & _
'        ")TB WHERE RTrim(TB.PO_NO) + RTrim(TB.Trade_code) + RTrim(TB.item_code) + Cast(Seq_No As varchar) NOT IN " & _
'        "( " & _
'            "SELECT RTrim(PO_No) + RTrim(supplier_code) + RTrim(item_code) + Cast(ReceiptSeq_No As varchar) " & _
'            "From InvoiceSupplier_Detail " & _
'            "WHERE invoice_no = @InvoiceNo AND supplier_code = @SupplierCode " & _
'        ") "
'
'    RsIm.Open sql, Db, adOpenKeyset, adLockOptimistic
'
'    i = 1
'    Do Until RsIm.EOF
'        With grid
'            .Rows = .Rows + 1
'            If RsIm.Fields("qtyX") <> 0 Then
'                .Cell(flexcpChecked, i, 0) = flexChecked
'            Else
'                .Cell(flexcpChecked, i, 0) = flexUnchecked
'            End If
'            .TextMatrix(i, ColPONo) = RsIm.Fields("PO_No")
'            .TextMatrix(i, colpodate) = Format(RsIm.Fields("PO_Date"), "dd MMM yyyy")
'            .TextMatrix(i, colpcode) = RsIm.Fields("Item_Code")
'            .TextMatrix(i, colprodname) = RsIm.Fields("Item_Name")
'            .TextMatrix(i, ColUnit) = uf_GetUnitDescription(Trim(RsIm.Fields("Unit_Cls")))
'
'
'            sql = "select isnull(sum(qty),0) as Q from invoicesupplier_detail " & _
'                "where po_no = '" & RsIm.Fields("PO_No") & "' and item_code = '" & RsIm.Fields("Item_Code") & "' and receiptseq_no = " & RsIm.Fields("ReceiptSeq_No")
'
'            If RSX.State = 1 Then RSX.Close
'            RSX.Open sql, Db, adOpenKeyset, adLockOptimistic
'            .TextMatrix(i, colpoqty) = Format(RsIm.Fields("Po_Qty"), gs_formatQty)
'            .TextMatrix(i, ColPoQtyReal) = .TextMatrix(i, colpoqty) - RSX.Fields("Q")
'            .TextMatrix(i, ColRem) = Format(.TextMatrix(i, colpoqty) - RSX.Fields("Q"), gs_formatQty)
'            If CDbl(.TextMatrix(i, ColRem)) < 0 Then .TextMatrix(i, ColRem) = Format(0, gs_formatQty)
'
'            sql = "select qty from invoicesupplier_detail where po_no = '" & RsIm.Fields("PO_No") & "' " & _
'                " and item_code = '" & RsIm.Fields("Item_Code") & "' and invoice_no = '" & CboInvNo & "' and receiptseq_no = " & RsIm.Fields("ReceiptSeq_No")
'
'            If RSX.State = 1 Then RSX.Close
'            RSX.Open sql, Db, adOpenKeyset, adLockOptimistic
'            If RSX.EOF And RSX.BOF Then
'                .TextMatrix(i, colinvqty) = .TextMatrix(i, ColRem)
'                .TextMatrix(i, colqtyasli) = CDbl(.TextMatrix(i, ColRem))
'            Else
'                .TextMatrix(i, colinvqty) = Format(RSX.Fields("qty"), gs_formatQty)
'                .TextMatrix(i, colqtyasli) = RSX.Fields("qty")
'            End If
'
'            .TextMatrix(i, ColCurr) = uf_GetCurrencyDescription(Trim(RsIm.Fields("Currency_Code")))
'            .TextMatrix(i, colPrice) = Format(RsIm.Fields("Price"), gs_formatPrice)
'            .TextMatrix(i, ColAmount) = Format(.TextMatrix(i, colinvqty) * RsIm.Fields("Price"), gs_formatAmount)
'            .TextMatrix(i, ColBantuUnit) = RsIm.Fields("Unit_Cls")
'            .TextMatrix(i, ColBantuCurr) = RsIm.Fields("Currency_Code")
'            .TextMatrix(i, ColReceiptSeqNo) = RsIm.Fields("ReceiptSeq_No")
'            AloEdit (i)
'            i = i + 1
'            RsIm.MoveNext
'        End With
'    Loop
'    grid.Cell(flexcpAlignment, 0, ColUnit, grid.Rows - 1, ColUnit) = flexAlignCenterCenter
'
'    txtamount.Text = 0
'    For i = 0 To grid.Rows - 1
'        If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
'            txtamount.Text = Format(CDbl(txtamount.Text) + CDbl(grid.TextMatrix(i, ColAmountInv)), gs_formatAmount)
'        End If
'    Next
'
'End Sub

Sub AloEdit(X As Long)
grid.Cell(flexcpBackColor, X, 0) = &HFFFFFF
grid.Cell(flexcpBackColor, X, colinvqty) = &HFFFFFF
End Sub

Sub ComboSuppUbah()
    Dim i As Long
    MousePointer = vbHourglass
    i = 0
    rsSup.Requery
    rsSup.Find "supp_code ='" & cboSupp & "'"
    If Not rsSup.EOF Then
        txt_name.Text = Trim(rsSup!supp_name)
        Overseas_Cls = rsSup!country_cls
        rsSup.Requery
    End If
    
    i = 0
    BankAcc
    If combo1.ListIndex <> 0 Then
        i = 0
        InvAdd
    End If
    
    MousePointer = Default
End Sub

Sub BankAcc()
    Dim i As Long
    
    sql = "SELECT Bank_Code, Bank_Name, Bank_Account, currency_code from Bank_master where trade_code = '" & cboSupp & "'"
    Set RsAcc = New Recordset
    RsAcc.Open sql, Db, adOpenKeyset, adLockOptimistic
    With cbocurr
        .clear
        
        .columnCount = 4
        .ColumnWidths = "30 pt; 40 pt;100 pt; 120 pt"
        .ListWidth = 300
        .ListRows = 15
        If RsAcc.RecordCount > 0 Then
            RsAcc.MoveFirst
        End If
        Do Until RsAcc.EOF
            .AddItem ""
            .List(i, 0) = uf_GetCurrencyDescription(RsAcc!currency_code)
            .List(i, 1) = RsAcc!bank_Code
            .List(i, 2) = Trim(RsAcc!bank_name)
            .List(i, 3) = Trim(RsAcc!bank_Account)
            i = i + 1
            RsAcc.MoveNext
        Loop
    End With
End Sub

Sub InvAdd()
    Dim sqlno As String
    Dim i As Long
    Dim rsno As New Recordset
    
    sql = "SELECT Invoice_no,exchange_Rate,Exchange_Amount from InvoiceSupplier_Master where Supplier_code = '" & Trim(cboSupp) & "' " & _
            "and invoice_date >='" & Format(InvDateSt.Value, "yyyy-mm-dd") & "' and invoice_date <='" & Format(InvDateEnd.Value, "yyyy-mm-dd") & "'"
    Set rsno = New Recordset
    rsno.Open sql, Db, adOpenKeyset, adLockOptimistic
    With CboInvNo
        .clear
        .columnCount = 1
        .ColumnWidths = "110 pt"
        .ListWidth = 110
        .ListRows = 15
        Do Until rsno.EOF
            .AddItem ""
            .List(i, 0) = Trim(rsno!Invoice_No)
            '.List(i, 1) = Trim(rsno!Exchange_Rate)
            '.List(i, 2) = Trim(rsno!Exchange_Amount)
            
            i = i + 1
            rsno.MoveNext
        Loop
    End With
End Sub

Sub AddtoPO()
    Dim strSQL As String
    Dim AdoPo As New ADODB.Recordset
    
   'Pak toha minta di rubah ngambilnya dari Surat Jalan
   '2017-03-07
   
'    Strsql = " select * From Purchaseorder_master " & vbCrLf & _
'                      "     Where PO_Date>='" & Format(InvDateSt, "dd-MMM-YYYY") & "' And PO_Date<='" & Format(InvDateEnd, "dd-MMM-YYYY") & "' " & vbCrLf & _
'                      "         And Supplier_Code='" & Trim(CboSupp) & "' "
                      
                      
    strSQL = " select distinct Suratjalan_no From Part_receipt " & vbCrLf & _
                      "     Where Receipt_Date>='" & Format(InvDateSt, "dd-MMM-YYYY") & "' And Receipt_Date<='" & Format(InvDateEnd, "dd-MMM-YYYY") & "' " & vbCrLf & _
                      "         And Supplier_Code='" & Trim(cboSupp) & "' "
                      
    Set AdoPo = Db.Execute(strSQL)
    
    With CboPOnO
        .clear
        .columnCount = 1
        .ColumnWidths = "150 pt"
        .ListWidth = 150
        .ListRows = 8
    
        .AddItem ""
        .List(.ListCount - 1, 0) = strAll
    
        Do While Not AdoPo.EOF
            .AddItem ""
            .List(.ListCount - 1, 0) = Trim(AdoPo!SuratJalan_No)
            AdoPo.MoveNext
        Loop
    End With
CboPOnO.ListIndex = 0
End Sub
Sub adtocombo()
    sql = "SELECT  rtrim(Trade_Master.trade_Code) supp_code, rtrim(Trade_Master.Trade_Name) supp_name, " & _
        "rtrim(Trade_Master.Address1) address, country_Cls, POPayment_Day From Trade_Master where trade_cls in ('2','3')"
    Set rsSup = New Recordset
    rsSup.Open sql, Db, adOpenKeyset, adLockOptimistic
    With cboSupp
        .clear
        .columnCount = 4
        .ColumnWidths = "80 pt;280 pt; 0 pt; 0 pt; 0 pt"
        .ListWidth = 360
        .ListRows = 15
        i = 0
        rsSup.Requery
        If Not rsSup.EOF And Not rsSup.BOF Then
            Do Until rsSup.EOF
                .AddItem ""
                .List(i, 0) = Trim(rsSup!supp_code)
                .List(i, 1) = Trim(rsSup!supp_name)
                .List(i, 2) = Trim(rsSup!Address) & " "
                .List(i, 3) = Trim(rsSup!country_cls)
                .List(i, 4) = Val(rsSup!POPayment_Day & "")
                i = i + 1
                rsSup.MoveNext
            Loop
        End If
    End With
    
    
    With cboSearch
        .AddItem "Product Code"
        .AddItem "Product Name"
        .AddItem "DO NO"
        .ListIndex = 0
    End With
End Sub

Sub Header()
  With grid
    .clear

    .Rows = 1
    .ColS = 25

    .ColWidth(ColCls) = 300
    .ColWidth(ColPONo) = 2400
    .ColWidth(colpodate) = 1300
'    If Overseas_Cls = "1" Then
'        .ColWidth(ColDoNo) = 0
'    ElseIf Overseas_Cls = "0" Then
        .ColWidth(ColDoNo) = 1900
        .ColWidth(ColBCNO) = 1200
'    End If
    .ColWidth(colpcode) = 2500
    .ColWidth(colprodname) = 2000
    .ColWidth(ColUnit) = 500
    .ColWidth(colpoqty) = 1000
'    If Overseas_Cls = "1" Then
'        .ColWidth(coldoqty) = 0
'    ElseIf Overseas_Cls = "0" Then
        .ColWidth(coldoqty) = 1000
'    End If
    .ColWidth(colinvqty) = 1000
    .ColWidth(colrem) = 1000
    .ColWidth(ColCurr) = 600
    .ColWidth(ColPrice) = 1200
    .ColWidth(ColAmount) = 1600
    .ColWidth(colPriceMat) = 1500
    .ColWidth(ColAmountMat) = 1600
    .ColWidth(colPriceService) = 1500
    .ColWidth(ColAmountService) = 1600
    
    .ColHidden(ColBantuUnit) = True
    .ColHidden(ColBantuCurr) = True
    .ColHidden(ColCurX) = True
    .ColHidden(ColPoQtyReal) = True
    .ColHidden(colqtyasli) = True
    .ColHidden(ColReceiptSeqNo) = True
    
    .ColHidden(ColCurr) = (bteHakPrice = 0)
    .ColHidden(ColPrice) = (bteHakPrice = 0)
    .ColHidden(ColAmount) = (bteHakPrice = 0)
    .ColHidden(colPriceMat) = (bteHakPrice = 0)
    .ColHidden(ColAmountMat) = (bteHakPrice = 0)
    .ColHidden(colPriceService) = (bteHakPrice = 0)
    .ColHidden(ColAmountService) = (bteHakPrice = 0)
    
    .TextMatrix(0, ColCls) = " "
    .TextMatrix(0, ColPONo) = "PO No"
'    If Overseas_Cls = "1" Then
'        .TextMatrix(0, colpodate) = "PO Date"
'    Else
        .TextMatrix(0, colpodate) = "Receipt Date"
'    End If
    .TextMatrix(0, ColDoNo) = "DO No"
     .TextMatrix(0, ColBCNO) = "BC No"
    .TextMatrix(0, colpcode) = "Product Code"
    .TextMatrix(0, colprodname) = "Product Name"
    .TextMatrix(0, ColUnit) = "Unit"
    .TextMatrix(0, colpoqty) = "PO Qty"
    .TextMatrix(0, coldoqty) = "DO Qty"
    .TextMatrix(0, colinvqty) = "Inv Qty"
    .TextMatrix(0, colrem) = "Remaining"
    .TextMatrix(0, ColCurr) = "Curr"
    .TextMatrix(0, ColPrice) = "Adj. Price"
    .TextMatrix(0, ColAmount) = "Amount"
    .TextMatrix(0, colPriceMat) = "Price Material"
    .TextMatrix(0, ColAmountMat) = "Amount Material"
    .TextMatrix(0, colPriceService) = "Price Service"
    .TextMatrix(0, ColAmountService) = "Amount Service"
    .TextMatrix(0, ColReceiptSeqNo) = "Receipt Date"
    
    .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
    For i = colpoqty To ColAmountService
      .ColAlignment(i) = flexAlignRightCenter
    Next i
    
    .RowHeight(0) = 250
  End With
  
End Sub

Private Sub TxtBlNo_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "'" Then
KeyAscii = 0
End If
End Sub

Private Sub txtFakturPajakNo_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "'" Then
KeyAscii = 0
End If
End Sub

Private Sub txtPPN_GotFocus()
txtPPN.Text = Format(txtPPN.Text, "0")
End Sub

Private Sub txtPPN_LostFocus()
txtPPN.Text = Format(txtPPN.Text, gs_formatAmount)
End Sub

Private Sub ClearData()
    'Db.BeginTrans
    sql = "delete InvoiceSupplier_master where not exists (" & _
        "select invoice_no from InvoiceSupplier_detail " & _
        "where supplier_code = InvoiceSupplier_master.supplier_code " & _
        "and invoice_no = InvoiceSupplier_master.invoice_no) " & _
        "and supplier_code = '" & cboSupp & "' " & _
        "and invoice_no = '" & CboInvNo & "' "
    Db.Execute sql
    'Db.CommitTrans
End Sub
