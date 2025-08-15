VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_invoice_inquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Invoice Inquiry"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frm_invoice_inquiry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   12990
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   390
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.CommandButton cmdSearch 
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
      Index           =   0
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2340
      Width           =   1155
   End
   Begin VB.CommandButton cmd_preview 
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
      Left            =   12495
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9735
      Width           =   1125
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Enabled         =   0   'False
      Height          =   915
      Left            =   180
      TabIndex        =   29
      Top             =   8025
      Width           =   13125
      Begin VB.TextBox txt_invoice_no 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "txt_invoice_no"
         Top             =   473
         Width           =   2430
      End
      Begin VB.TextBox txt_total_amount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5535
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "txt_total_amount"
         Top             =   473
         Width           =   2505
      End
      Begin VB.TextBox txt_grand_total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   10710
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "txt_grand_total"
         Top             =   473
         Width           =   2235
      End
      Begin VB.TextBox txt_ppn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   8100
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "txt_ppn"
         Top             =   473
         Width           =   2550
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   330
         Left            =   1770
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   450
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CustomFormat    =   "MMM yyyy"
         Format          =   151715843
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   150
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   450
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CustomFormat    =   "dd MMM yyyy"
         Format          =   151715843
         CurrentDate     =   37798
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
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
         Left            =   10710
         TabIndex        =   43
         Top             =   90
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPn"
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
         Left            =   8100
         TabIndex        =   42
         Top             =   90
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Left            =   5535
         TabIndex        =   41
         Top             =   90
         Width           =   1140
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
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
         Left            =   3060
         TabIndex        =   40
         Top             =   90
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
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
         Left            =   1770
         TabIndex        =   39
         Top             =   90
         Width           =   1185
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
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
         Left            =   225
         TabIndex        =   34
         Top             =   90
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00A6D2FF&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   90
         Top             =   0
         Width           =   12930
      End
      Begin VB.Shape Shape1 
         Height          =   555
         Left            =   90
         Top             =   360
         Width           =   12930
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   285
      TabIndex        =   23
      Top             =   9015
      Width           =   14550
      Begin VB.Label lbl_pesan 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_pesan"
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
         Left            =   90
         TabIndex        =   24
         Top             =   240
         Width           =   14205
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1245
      Left            =   300
      TabIndex        =   15
      Top             =   1005
      Width           =   14550
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
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "txt_name"
         Top             =   270
         Width           =   2985
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   345
         Left            =   3285
         TabIndex        =   2
         Top             =   705
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   609
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
         Format          =   151715843
         CurrentDate     =   37818
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   345
         Left            =   1395
         TabIndex        =   1
         Top             =   705
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   609
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
         Format          =   151715843
         CurrentDate     =   37818
      End
      Begin VB.Label lbl_name 
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_name"
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
         Left            =   5400
         TabIndex        =   22
         Top             =   765
         Visible         =   0   'False
         Width           =   2985
      End
      Begin MSForms.ComboBox ComboBox1 
         Height          =   315
         Left            =   1395
         TabIndex        =   0
         Top             =   270
         Width           =   1335
         VariousPropertyBits=   746604571
         MaxLength       =   7
         DisplayStyle    =   3
         Size            =   "2355;556"
         ColumnCount     =   2
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         X1              =   3555
         X2              =   6540
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line2 
         X1              =   7650
         X2              =   11475
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   6690
         TabIndex        =   21
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer CD"
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
         Left            =   90
         TabIndex        =   20
         Top             =   285
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
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
         Left            =   90
         TabIndex        =   19
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label11 
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
         Left            =   2970
         TabIndex        =   18
         Top             =   735
         Width           =   210
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   2880
         TabIndex        =   17
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lbl_address 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_address"
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
         Left            =   7665
         TabIndex        =   16
         Top             =   270
         Width           =   3855
      End
   End
   Begin VB.TextBox txt_remarks 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "txt_remarks"
      Top             =   7620
      Width           =   11790
   End
   Begin VB.CommandButton cmd_submit 
      BackColor       =   &H0080FFFF&
      Caption         =   "To Update"
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
      Left            =   13740
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9735
      Width           =   1125
   End
   Begin VB.CommandButton cmd_sub_menu 
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
      TabIndex        =   7
      Top             =   9735
      Width           =   1125
   End
   Begin VB.CommandButton cmd_first 
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
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9735
      Width           =   1125
   End
   Begin VB.CommandButton cmd_previous 
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
      Left            =   5055
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9735
      Width           =   1125
   End
   Begin VB.CommandButton cmd_next 
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
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9735
      Width           =   1125
   End
   Begin VB.CommandButton cmd_last 
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
      Left            =   7545
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9735
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid vsflexGrid1 
      Height          =   4650
      Left            =   315
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2835
      Width           =   14505
      _cx             =   25585
      _cy             =   8202
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
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   275
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm_invoice_inquiry.frx":0E42
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
   Begin MSForms.ComboBox combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   2385
      Width           =   2790
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4921;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lbl_status 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_status"
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
      Left            =   11790
      TabIndex        =   37
      Top             =   2430
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Inquiry"
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
      Left            =   330
      TabIndex        =   28
      Top             =   390
      Width           =   14505
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   345
      TabIndex        =   27
      Top             =   2445
      Width           =   975
   End
   Begin VB.Label lbl_record 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_record"
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
      Height          =   255
      Left            =   1905
      TabIndex        =   26
      Top             =   9795
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   315
      TabIndex        =   25
      Top             =   7650
      Width           =   855
   End
End
Attribute VB_Name = "frm_invoice_inquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim rs_invoice_master As New ADODB.Recordset
Dim rs_invoice_master2 As New ADODB.Recordset
Dim rs_invoice_detail As New ADODB.Recordset
Dim rs_join_delivery_order_model_master As New ADODB.Recordset
Dim rs_trade_master As New ADODB.Recordset

Dim countrycls As String
Dim update As String, kanan_pertama As Integer, l_combo As String
Dim sql_join As String, l_fix As String

Dim bteColItemCode As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColQty As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColService As Byte
Dim bteColAmount As Byte

Dim bteHakPrice As Byte

Private Sub SetCol()
    bteColItemCode = 0
    bteColPartNo = 1
    bteColDesc = 2
    bteColQty = 3
    bteColCurr = 4
    bteColPrice = 5
    bteColService = 6
    bteColAmount = 7
End Sub

Public Sub cmdSearch_Click(Index As Integer)
On Error GoTo ErrorMSg
If Trim(combo1) = "" Then lbl_pesan.Caption = DisplayMsg(4007): Exit Sub
MousePointer = vbDefault


    combo1 = combo1
    'rs_invoice_master.Requery
     '   rs_invoice_master.MoveFirst
      '  rs_invoice_master.Find "invoice_no='" & Trim(combo1.Text) & "'"
    'If Not rs_invoice_master.EOF Then
    If combo1.MatchFound Then
        l_fix = UCase(Trim(combo1.Column(1)))
        rs_invoice_master.Requery
        rs_invoice_master.MoveFirst
        rs_invoice_master.Find "invoice_no='" & Trim(combo1.Text) & "'"
        l_combo = Trim(combo1.Text)
        ComboBox1.Text = Trim(rs_invoice_master!Cust_CodE)
        combo1.Text = l_combo
        Call setting_grid
        lbl_pesan.Caption = ""

        DTPicker1.Value = rs_invoice_master!Invoice_Date
        DTPicker2.Value = Right(Trim(rs_invoice_master!delivery_Date), 2) & " " & Left(Trim(rs_invoice_master!delivery_Date), 4)

        txt_remarks.Text = rs_invoice_master!Remarks

        txt_invoice_no.Text = Trim(rs_invoice_master!Invoice_No)
        txt_total_amount.Text = Format(Trim(rs_invoice_master!Amount), gs_formatAmountIDR)
        txt_ppn.Text = Format(Trim(rs_invoice_master!ppn), gs_formatAmountIDR)
        txt_grand_total.Text = Format(Trim(rs_invoice_master!total_amount), gs_formatAmountIDR)
        If Trim(txt_grand_total.Text) = "" Then txt_grand_total.Text = Format("0", gs_formatAmountIDR)
        If Trim(txt_ppn.Text) = "" Then txt_ppn.Text = Format("0", gs_formatAmountIDR)
            
            'if overseas customer
        If Trim(countrycls) <> "0" Then
            txt_grand_total.Text = Format(Trim(rs_invoice_master!Amount), gs_formatAmountIDR)
            txt_ppn.Text = Format("0", gs_formatAmountIDR)
        End If
            
        lbl_record.Caption = "Page " & rs_invoice_master.AbsolutePosition & " of " & rs_invoice_master.RecordCount
        lbl_status = IIf(Trim(l_fix) = "1", "Status : Fix", "")
        lbl_pesan.Caption = ""
    Else
        lbl_pesan.Caption = DisplayMsg(4006) '"Data not found !"
        Call frame_bawah_clear(False)
        Call set_combo_inv
        'Combobox1.Text = ""
        combo1.SelStart = Len(Trim(combo1.Text))
        lbl_status = ""
    End If

MousePointer = vbDefault
Exit Sub
ErrorMSg:
lbl_pesan = err.number & " " & err.Description
MousePointer = vbDefault
Exit Sub
End Sub
Public Sub dariPjk()

    
            l_combo = Trim(combo1.Text)
            ComboBox1.Text = Trim(rs_invoice_master!Cust_CodE)
            combo1.Text = l_combo
            'DTPicker1.Value = rs_invoice_master!Invoice_Date
            Call setting_grid
            lbl_pesan.Caption = ""
            
            DTPicker1.Value = rs_invoice_master!Invoice_Date
            DTPicker2.Value = Right(Trim(rs_invoice_master!delivery_Date), 2) & " " & Left(Trim(rs_invoice_master!delivery_Date), 4)
            
            txt_remarks.Text = rs_invoice_master!Remarks
            
            txt_invoice_no.Text = Trim(rs_invoice_master!Invoice_No)
            txt_total_amount.Text = Format(Trim(rs_invoice_master!Amount), gs_formatAmountIDR)
            txt_ppn.Text = Format(Trim(rs_invoice_master!ppn), gs_formatAmountIDR)
            txt_grand_total.Text = Format(Trim(rs_invoice_master!total_amount), gs_formatAmountIDR)
                If Trim(txt_grand_total.Text) = "" Then txt_grand_total.Text = Format("0", gs_formatAmountIDR)
                If Trim(txt_ppn.Text) = "" Then txt_ppn.Text = Format("0", gs_formatAmountIDR)
            lbl_record.Caption = "Page " & rs_invoice_master.AbsolutePosition & " of " & rs_invoice_master.RecordCount
            lbl_status = IIf(Trim(l_fix) = "1", "Status : Fix", "")
            lbl_pesan.Caption = ""


End Sub

Private Sub combo1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If Len(Trim(combo1.Text)) = 20 And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0: Exit Sub

End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    lbl_pesan.Caption = ErrMsg
End If
End Sub


Private Sub tombol_navigasi()
Dim l_invoiceNo As String
If rs_invoice_master.EOF Or rs_invoice_master.BOF = False Then
    
    lbl_record.Caption = "Page " & rs_invoice_master.AbsolutePosition & " of " & rs_invoice_master.RecordCount
    ComboBox1.Text = Trim(rs_invoice_master!Cust_CodE)
    combo1.Text = Trim(rs_invoice_master!Invoice_No)
     DTPicker1.Value = rs_invoice_master!Invoice_Date
     DTPicker2.Value = Right(Trim(rs_invoice_master!delivery_Date), 2) & " " & Left(Trim(rs_invoice_master!delivery_Date), 4)
    
     txt_remarks.Text = rs_invoice_master!Remarks
     
     txt_invoice_no.Text = Trim(rs_invoice_master!Invoice_No)
     txt_total_amount.Text = Format(Trim(rs_invoice_master!Amount), gs_formatAmountIDR)
     txt_ppn.Text = Format(Trim(rs_invoice_master!ppn), gs_formatAmountIDR)
     txt_grand_total.Text = Format(Trim(rs_invoice_master!total_amount), gs_formatAmountIDR)
                If Trim(txt_grand_total.Text) = "" Then txt_grand_total.Text = Format("0", gs_formatAmountIDR)
                If Trim(txt_ppn.Text) = "" Then txt_ppn.Text = Format("0", gs_formatAmountIDR)
End If
    
Call set_cust

l_invoiceNo = Trim(combo1.Text)
Call set_invoice_master(False)
combo1.Text = l_invoiceNo
Call setting_grid
'Call combo1_KeyPress(13)

kanan_pertama = 1
    
lbl_record.Caption = "Page " & rs_invoice_master.AbsolutePosition & " of " & rs_invoice_master.RecordCount
    
End Sub
    
    
Private Sub cmd_first_Click()

If rs_invoice_master.EOF = False Or rs_invoice_master.BOF = False Then
    rs_invoice_master.MoveFirst
    Call tombol_navigasi
    lbl_pesan.Caption = DisplayMsg(4020) '"This is the first page !"
End If

End Sub

Private Sub cmd_last_Click()


If rs_invoice_master.EOF = False Or rs_invoice_master.BOF = False Then
    rs_invoice_master.MoveLast
    Call tombol_navigasi
    lbl_pesan.Caption = DisplayMsg(4021) '"This is the last page !"
End If

End Sub

Private Sub cmd_next_Click()

If rs_invoice_master.EOF = False Or rs_invoice_master.BOF = False Then
    
    If kanan_pertama = 0 Then
        rs_invoice_master.MoveFirst: Call tombol_navigasi: lbl_pesan.Caption = DisplayMsg(4020) '"This is the first page !"
        Exit Sub
    End If
    
    rs_invoice_master.MoveNext
    
    If rs_invoice_master.EOF = True Then
        If rs_invoice_master.State <> adStateClosed Then rs_invoice_master.Close
        rs_invoice_master.Open " select * from invoice_master", Db, adOpenKeyset, adLockOptimistic

        rs_invoice_master.MoveLast: Call tombol_navigasi: lbl_pesan.Caption = DisplayMsg(4021) '"This is the last page !"
        Exit Sub
    End If
    Call tombol_navigasi
    
    lbl_pesan.Caption = ""

End If

End Sub

Private Sub cmd_preview_Click()

lbl_pesan.Caption = ""

If Trim(combo1.Text) = "" Then GoTo Kosong
If rs_invoice_master.EOF = False Or rs_invoice_master.BOF = False Then
    rs_invoice_master.MoveFirst
        rs_invoice_master.Find "invoice_no='" & Trim(combo1.Text) & "'"
End If
If rs_invoice_master.EOF = True Then lbl_pesan.Caption = DisplayMsg(4006) '"Please insert a valid invoice number !": Exit Sub


Dim xdo As Recordset
MousePointer = vbHourglass

sql = "select distinct do_no from invoice_detail"
Set xdo = New Recordset
xdo.Open sql, Db, adOpenDynamic, adLockOptimistic
inv_no = "'" & Trim(combo1.Text) & "'"
If countrycls = 0 Then
    Call InvReport
Else
    InvReportExport (bteHakPrice)
End If
MousePointer = vbDefault
Exit Sub

Kosong:
    lbl_pesan.Caption = DisplayMsg(4007) '"Please select an invoice number !"
    combo1.SetFocus

End Sub

Private Sub cmd_previous_Click()

If rs_invoice_master.EOF = False Or rs_invoice_master.BOF = False Then
    rs_invoice_master.MovePrevious
    If rs_invoice_master.BOF = True Then
        rs_invoice_master.MoveFirst: Call tombol_navigasi: lbl_pesan.Caption = DisplayMsg(4020) '"This is the first page !"
        Exit Sub
    End If
    Call tombol_navigasi
    lbl_pesan.Caption = ""
End If

End Sub

Private Sub frame_clear()

ComboBox1.Text = ""
'lbl_name.Caption = ""
txt_name = ""
lbl_address.Caption = ""
'lbl_area.Caption = ""
'lbl_dealer_cls.Caption = ""
DTPicker1.Value = Now
DTPicker2.Value = Now

txt_invoice_no.Text = ""
txt_ppn.Text = ""
txt_grand_total.Text = ""
txt_total_amount.Text = ""

End Sub

Private Sub cmd_sub_menu_Click()

If cmd_sub_menu.Caption = "&Back" Then
    
    
    If cmd_sub_menu.Tag = "" Then
       ' F_DOUpdate.Show
   Else
        FrmPajak_Create_New.Show
        cmd_sub_menu.Tag = ""
   End If
    
   Unload Me
   Exit Sub
End If


frmMainMenu.Show

Unload Me
    
End Sub

Private Sub Cmd_Submit_Click()

If Trim(combo1.Text) = "" Then GoTo Kosong
If rs_invoice_master.EOF = False Or rs_invoice_master.BOF = False Then
    rs_invoice_master.MoveFirst
        rs_invoice_master.Find "invoice_no='" & Trim(combo1.Text) & "'"
End If
If rs_invoice_master.EOF = True Then
    lbl_pesan.Caption = DisplayMsg(4006) '"Please insert a valid invoice number !"
    Exit Sub
End If

If countrycls = 0 Then
    If hakAkses("FrminvoiceExport") = 0 Then lbl_pesan = DisplayMsg(3007):  Exit Sub
    With FrmInvoiceExport
        Call .dr_invInq(ComboBox1, combo1)
        '.nilKosong = True
        .cmdAction(0).Caption = "Back"
        .set_tgl CStr(DTPicker3.Value), CStr(DTPicker4.Value)
        .ShowData
        .formpanggil = "invoiceinquiry"
        .Show
        .lblerror = ""
        '.nilKosong = False
    End With
Else
    If hakAkses("FrmInvoiceExport") = 0 Then lbl_pesan = DisplayMsg(3007):  Exit Sub
    With FrmInvoiceExport
        Call .dr_invInq(ComboBox1, combo1)
        .cmdAction(0).Caption = "Back"
        .set_tgl CStr(DTPicker3.Value), CStr(DTPicker4.Value)
        .ShowData
        .formpanggil = "invoiceinquiry"
        .Show
        .lblerror = ""
    End With
End If

Unload Me
Exit Sub

Kosong:
    lbl_pesan.Caption = DisplayMsg(4007) '"Please insert invoice number !"
End Sub

Public Sub set_tgl(X As String, Y As String)
DTPicker3.Value = X
DTPicker4.Value = Y
'Call set_invoice_master

End Sub

Private Sub Combo1_Click()
'Combo1_KeyPress (13)
If combo1.ListIndex < 0 Then Exit Sub
l_fix = UCase(Trim(combo1.List(combo1.ListIndex, 1)))
End Sub

Function panggil(no_invoice As String)
'Call set_invoice_master
combo1.Text = no_invoice
'combo1_KeyPress (13)
End Function




Private Sub frame_bawah_clear(Optional a As Boolean)

combo1.clear

DTPicker1.Value = Now
DTPicker2.Value = Now

txt_remarks.Text = ""

txt_invoice_no.Text = ""
txt_total_amount.Text = ""
txt_ppn.Text = ""
'txt_pph.Text = ""
txt_grand_total.Text = ""

Call setting_grid
If a = True Then _
Call label_clear

a = True

lbl_record.Caption = "Page 0 of 0"
End Sub

Public Sub ComboBox1_Click()
Call set_cust

Call set_invoice_master

lbl_pesan.Caption = ""
lbl_record.Caption = "Page 0 of 0"
    
End Sub
Public Sub set_cust()
If rs_trade_master.EOF = False Or rs_trade_master.BOF = False Then
    rs_trade_master.MoveFirst
        rs_trade_master.Find " trade_code='" & Trim(ComboBox1.Text) & "'"
    If rs_trade_master.EOF = False Then

        txt_name = rs_trade_master!trade_name
        lbl_address = rs_trade_master!address1
        countrycls = IIf(IsNull(rs_trade_master!country_cls) = True, "0", rs_trade_master!country_cls)
    End If
End If
End Sub

Public Function set_dari_inv_create()
Call set_cust
Call set_combo_inv
End Function

Private Sub set_invoice_master(Optional a As Boolean)

l_combo = Trim(combo1.Text)
combo1.clear
combo1.Text = ""

Call set_combo_inv

kanan_pertama = 0

DTPicker1.Value = Now
DTPicker2.Value = Now

If a = True Then
txt_remarks.Text = ""
txt_invoice_no.Text = ""
txt_total_amount.Text = ""
txt_ppn.Text = ""
txt_grand_total.Text = ""
End If
a = True
Call setting_grid

End Sub

Public Sub set_combo_inv()
If rs_invoice_master2.State <> adStateClosed Then rs_invoice_master2.Close

sql = "select * from invoice_master where cust_code='" & Trim(ComboBox1.Text) & "' " & _
                                " and invoice_date>='" & Format(DTPicker3.Value, "yyyy-MM-dd") & "' " & _
                                " and invoice_date<='" & Format(DTPicker4.Value, "yyyy-MM-dd") & "' "
                                
rs_invoice_master2.Open sql, Db, adOpenDynamic, adLockReadOnly, adCmdText

If rs_invoice_master2.EOF = False Then
rs_invoice_master2.MoveFirst
combo1.clear
combo1.columnCount = 2
combo1.TextColumn = 1
i = 0
While (rs_invoice_master2.EOF = False)
    combo1.AddItem ""
    combo1.List(i, 0) = (rs_invoice_master2!Invoice_No)
    combo1.List(i, 1) = IIf(Trim(rs_invoice_master2!fix_cls) <> "", Trim(rs_invoice_master2!fix_cls), "")
    i = i + 1
    rs_invoice_master2.MoveNext
Wend
 combo1.ColumnWidths = "150 pt; 0 pt"
combo1.ListWidth = 150
End If

'==================================
For i = 0 To combo1.ListCount - 1
    If l_combo = combo1.List(i) Then
        combo1.ListIndex = i
        combo1.Text = combo1.List(combo1.ListIndex)
        Exit For
    End If
Next
'==================================

End Sub

Private Sub Combobox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)


If KeyCode = 13 Then
    If rs_trade_master.EOF = False Or rs_trade_master.BOF = False Then
        rs_trade_master.MoveFirst
            rs_trade_master.Find "trade_code='" & Trim(ComboBox1.Text) & "'"
        If rs_trade_master.EOF = False Then
            lbl_pesan.Caption = ""
            Call ComboBox1_Click
            lbl_record.Caption = "Page 0 of 0"
        Else
            lbl_pesan.Caption = DisplayMsg(4006) '"Data not found !"
            lbl_record.Caption = "Page 0 of 0"
            combo1.Text = ""
            Call setting_grid
            Call frame_bawah_clear(True)
            Call label_clear
        End If
    End If
End If
End Sub



Private Sub Combobox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub DTPicker3_Change()

If DTPicker3.Value > DTPicker4.Value Then
    lbl_pesan.Caption = DisplayMsg(1021) '"The first of delivery date must be equal or lower than " & Format(DTPicker4.value, "dd MMM yyyy")
    DTPicker3.Value = Format(DTPicker4.Value, "dd MMM yyyy")
Else
    lbl_pesan.Caption = ""
    
End If
lbl_record.Caption = "Page 0 of 0"
Call set_invoice_master
End Sub

Private Sub DTPicker4_Change()
If DTPicker3.Value > DTPicker4.Value Then
    lbl_pesan.Caption = DisplayMsg(1021) '"The last of delivery date must be equal or higher than " & Format(DTPicker3.value, "dd MMM yyyy")
    DTPicker4.Value = Format(DTPicker3.Value, "dd MMM yyyy")
Else
    lbl_pesan.Caption = ""
    
End If
lbl_record.Caption = "Page 0 of 0"
Call set_invoice_master
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
Me.Caption = "Invoice Inquiry"
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
bteHakPrice = hakPrice(Me.Name)
Label10(3).Visible = (bteHakPrice = 1)
Label10(4).Visible = (bteHakPrice = 1)
Label10(5).Visible = (bteHakPrice = 1)
txt_total_amount.Visible = (bteHakPrice = 1)
txt_ppn.Visible = (bteHakPrice = 1)
txt_grand_total.Visible = (bteHakPrice = 1)
Call SetCol
Call koneksi
combo1.Text = ""
Call setting
ComboBox1.Text = ""
lbl_status = ""
lbl_pesan.Caption = ""
lbl_record.Caption = "Page 0 of 0"
txt_remarks.Enabled = False
End Sub

Private Sub koneksi()

'=======================================================================================================
rs_invoice_master.Open "select * from invoice_master", Db, adOpenKeyset, adLockOptimistic
rs_invoice_detail.Open "select * from invoice_detail", Db, adOpenKeyset, adLockOptimistic
rs_trade_master.Open "select * from trade_master where  trade_Cls='2'", Db, adOpenKeyset, adLockOptimistic

'=======================================================================================================

End Sub

Private Sub text_clear()

txt_remarks.Text = ""
txt_invoice_no.Text = ""
txt_total_amount.Text = ""
txt_ppn.Text = ""
txt_grand_total.Text = ""

End Sub

Private Sub label_clear()

'lbl_name.Caption = ""
txt_name = ""
'lbl_area.Caption = ""
lbl_address.Caption = ""
'lbl_dealer_cls.Caption = ""

End Sub

Private Sub setting()

'=================setting combobox1=======================
ComboBox1.clear
ComboBox1.columnCount = 3
ComboBox1.TextColumn = 1

i = 0

If rs_trade_master.EOF = False Or rs_trade_master.BOF = False Then
    rs_trade_master.MoveFirst
    While rs_trade_master.EOF = False
        ComboBox1.AddItem ""
        ComboBox1.List(i, 0) = Trim(rs_trade_master!Trade_Code)
        ComboBox1.List(i, 1) = Trim(rs_trade_master!trade_name)
        ComboBox1.List(i, 2) = IIf(IsNull(rs_trade_master!country_cls) = True, "0", rs_trade_master!country_cls)
        rs_trade_master.MoveNext
        i = i + 1
    Wend
    ComboBox1.ColumnWidths = "50 pt; 350 pt;0 pt"
    ComboBox1.ListWidth = 400
End If
'==========================================================

'=================setting label clear============================
Call label_clear
'==========================================================

'=================setting text clear============================
Call text_clear
'==========================================================

'=================setting vsflexgrid1============================
Call setting_grid
'==========================================================

'=================setting dtpicker======================
DTPicker1.Value = Now
DTPicker2.Value = Now
DTPicker3.Value = Now
DTPicker4.Value = Now
'==========================================================

End Sub

Private Sub setting_grid()
Dim l_cur As String, l_price1 As String
Dim l_service1 As String
VSFlexGrid1.Rows = 1
                              
If rs_join_delivery_order_model_master.State <> adStateClosed Then rs_join_delivery_order_model_master.Close

sql_join = "select invoice_no,invoice_detail.item_code,item_name,invoice_detail.makeritem_code, " & _
                "sum(invoice_detail.qty) as qty,invoice_detail.price,isnull(invoice_detail.Service,0) service, " & _
                "invoice_detail.currency_code " & _
                " From " & _
                "invoice_detail Join item_master on " & _
                "invoice_detail.item_code=item_master.item_code where invoice_no='" & Trim(combo1.Text) & "' group by " & _
                "invoice_no , invoice_detail.item_code, item_name, invoice_detail.price,invoice_detail.currency_code,invoice_detail.makeritem_code,invoice_detail.Service "


rs_join_delivery_order_model_master.Open sql_join, Db, adOpenKeyset, adLockOptimistic

If rs_join_delivery_order_model_master.EOF = False Or rs_join_delivery_order_model_master.BOF = False Then

    rs_join_delivery_order_model_master.MoveFirst
        

    While Not rs_join_delivery_order_model_master.EOF
    
        l_cur = uf_GetCurrencyDescription(rs_join_delivery_order_model_master!currency_code)
        If rs_join_delivery_order_model_master!currency_code = "03" Then
            l_price1 = Format(Trim(rs_join_delivery_order_model_master!Price), gs_formatPriceIDR)
            l_service1 = Format(Trim(rs_join_delivery_order_model_master!service), gs_formatPriceIDR)
        Else
            l_price1 = Format(Trim(rs_join_delivery_order_model_master!Price), gs_formatPrice)
            l_service1 = Format(Trim(rs_join_delivery_order_model_master!service), gs_formatPrice)
        End If
        
        VSFlexGrid1.AddItem ""
        VSFlexGrid1.TextMatrix(VSFlexGrid1.Rows - 1, bteColItemCode) = Trim(rs_join_delivery_order_model_master!Item_Code)
        VSFlexGrid1.TextMatrix(VSFlexGrid1.Rows - 1, bteColPartNo) = Trim(rs_join_delivery_order_model_master!MakerItem_Code)
        VSFlexGrid1.TextMatrix(VSFlexGrid1.Rows - 1, bteColDesc) = Trim(uf_GetItemDescription(rs_join_delivery_order_model_master!Item_Code))
        VSFlexGrid1.TextMatrix(VSFlexGrid1.Rows - 1, bteColQty) = Format(Trim(rs_join_delivery_order_model_master!Qty), gs_formatQty)
        VSFlexGrid1.TextMatrix(VSFlexGrid1.Rows - 1, bteColCurr) = l_cur
        VSFlexGrid1.TextMatrix(VSFlexGrid1.Rows - 1, bteColPrice) = l_price1
        VSFlexGrid1.TextMatrix(VSFlexGrid1.Rows - 1, bteColService) = l_service1
        VSFlexGrid1.TextMatrix(VSFlexGrid1.Rows - 1, bteColAmount) = Format((Trim(rs_join_delivery_order_model_master!Qty) * CDbl(l_price1)), gs_formatAmountIDR)
        rs_join_delivery_order_model_master.MoveNext
        
    Wend

Else
    VSFlexGrid1.clear
End If

With VSFlexGrid1
    .TextMatrix(0, bteColItemCode) = "Item Code"
    .TextMatrix(0, bteColPartNo) = "Part Number"
    .TextMatrix(0, bteColDesc) = "Description"
    .TextMatrix(0, bteColQty) = "Qty"
    .TextMatrix(0, bteColCurr) = "Currency"
    .TextMatrix(0, bteColPrice) = "Price"
    .TextMatrix(0, bteColService) = "Service"
    .TextMatrix(0, bteColAmount) = "Amount"
    
    .ColAlignment(bteColItemCode) = flexAlignLeftCenter
    .ColAlignment(bteColPartNo) = flexAlignLeftCenter
    .ColAlignment(bteColDesc) = flexAlignLeftCenter
    .ColAlignment(bteColQty) = flexAlignRightCenter
    .ColAlignment(bteColCurr) = flexAlignLeftCenter
    .ColAlignment(bteColPrice) = flexAlignRightCenter
    .ColAlignment(bteColService) = flexAlignRightCenter
    .ColAlignment(bteColAmount) = flexAlignRightCenter

    .ColWidth(bteColItemCode) = 2000
    .ColWidth(bteColPartNo) = 2000
    .ColWidth(bteColDesc) = 3200
    .ColWidth(bteColQty) = 1100
    .ColWidth(bteColCurr) = 900
    .ColWidth(bteColPrice) = 1500
    .ColWidth(bteColService) = 1500
    .ColWidth(bteColAmount) = 2000
    
    .ColHidden(bteColCurr) = (bteHakPrice = 0)
    .ColHidden(bteColPrice) = (bteHakPrice = 0)
    .ColHidden(bteColService) = (bteHakPrice = 0)
    .ColHidden(bteColAmount) = (bteHakPrice = 0)
    
    .EditMaxLength = 1
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

If rs_invoice_master.State <> adStateClosed Then rs_invoice_master.Close
If rs_invoice_detail.State <> adStateClosed Then rs_invoice_detail.Close
If rs_join_delivery_order_model_master.State <> adStateClosed Then rs_join_delivery_order_model_master.Close
If rs_trade_master.State <> adStateClosed Then rs_trade_master.Close
End Sub


Private Sub vsflexgrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col <> 0 Then Cancel = True
End Sub

Private Sub VSFlexGrid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
    KeyAscii = 0
End If
End Sub

Private Sub vsflexgrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If VSFlexGrid1.Row < VSFlexGrid1.Rows - 1 Then: VSFlexGrid1.Row = VSFlexGrid1.Row + 1: VSFlexGrid1.SetFocus
End If
End Sub

Public Sub xxx()
cmdSearch_Click (0)
End Sub



