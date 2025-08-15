VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmInvoiceExport 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Invoice Create (Export)"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInvoiceExport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Index           =   7
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   9840
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   345
      Index           =   0
      Left            =   1470
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   9870
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtPEBNo 
      Height          =   315
      Left            =   8460
      MaxLength       =   25
      TabIndex        =   8
      Top             =   2610
      Width           =   2265
   End
   Begin VB.TextBox TxtAirCharge 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   7950
      TabIndex        =   13
      Top             =   8580
      Width           =   2265
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Delete"
      Height          =   375
      Index           =   4
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9840
      Width           =   1320
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmInvoiceExport.frx":0E42
      Left            =   300
      List            =   "FrmInvoiceExport.frx":0E4C
      Style           =   2  'Dropdown List
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1155
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "Create"
      Height          =   375
      Index           =   5
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2130
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1215
      Left            =   150
      TabIndex        =   28
      Top             =   810
      Width           =   14715
      Begin MSComCtl2.DTPicker SDate 
         Height          =   315
         Left            =   1650
         TabIndex        =   1
         Top             =   780
         Width           =   1665
         _ExtentX        =   2937
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
         Format          =   335151107
         CurrentDate     =   37802
      End
      Begin MSComCtl2.DTPicker EDate 
         Height          =   315
         Left            =   3720
         TabIndex        =   2
         Top             =   780
         Width           =   1665
         _ExtentX        =   2937
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
         Format          =   335151107
         CurrentDate     =   37802
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing No."
         Height          =   195
         Index           =   4
         Left            =   5730
         TabIndex        =   50
         Top             =   840
         Width           =   1005
      End
      Begin MSForms.ComboBox cboPacking 
         Height          =   315
         Left            =   6870
         TabIndex        =   3
         Top             =   780
         Width           =   2625
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "4630;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbodealer 
         Height          =   315
         Left            =   1620
         TabIndex        =   0
         Top             =   270
         Width           =   1335
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2355;556"
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
         Caption         =   "Customer CD"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   36
         Top             =   330
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   35
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   2
         Left            =   6660
         TabIndex        =   34
         Top             =   330
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Date"
         Height          =   195
         Index           =   5
         Left            =   330
         TabIndex        =   33
         Top             =   810
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   195
         Index           =   6
         Left            =   3420
         TabIndex        =   32
         Top             =   810
         Width           =   165
      End
      Begin VB.Label lbldesc 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxx"
         Height          =   195
         Index           =   0
         Left            =   3720
         TabIndex        =   31
         Top             =   330
         Width           =   2730
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3720
         X2              =   6480
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line2 
         X1              =   7470
         X2              =   14010
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label lbldesc 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
         Height          =   195
         Index           =   1
         Left            =   7470
         TabIndex        =   30
         Top             =   330
         Width           =   6540
      End
      Begin VB.Label lblfix 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   13290
         TabIndex        =   29
         Top             =   870
         Width           =   1185
      End
   End
   Begin VB.TextBox txtremarks 
      Height          =   315
      Left            =   1350
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   7440
      Width           =   10965
   End
   Begin VB.TextBox Txtdisplay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   1
      Left            =   5610
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   8580
      Width           =   2265
   End
   Begin VB.TextBox Txtdisplay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   2
      Left            =   10290
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   8580
      Width           =   1935
   End
   Begin VB.TextBox Txtdisplay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   3
      Left            =   12300
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8580
      Width           =   2565
   End
   Begin VB.TextBox Txtdisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   2970
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8580
      Width           =   2565
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   375
      Index           =   3
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9840
      Width           =   1320
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Index           =   1
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9840
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   2
      Left            =   13830
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9840
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   615
      Left            =   90
      TabIndex        =   22
      Top             =   9120
      Width           =   15000
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
         Height          =   285
         Left            =   150
         TabIndex        =   23
         Top             =   210
         Width           =   12165
      End
   End
   Begin VB.TextBox TXTNo 
      Height          =   315
      Left            =   630
      MaxLength       =   25
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
      Height          =   375
      Index           =   6
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9840
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4020
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1140
      Width           =   2745
   End
   Begin MSComCtl2.DTPicker IDate 
      Height          =   315
      Left            =   8460
      TabIndex        =   6
      Top             =   2160
      Width           =   1665
      _ExtentX        =   2937
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
      Format          =   335216643
      CurrentDate     =   37802
   End
   Begin MSComCtl2.DTPicker DDate 
      Height          =   315
      Left            =   11790
      TabIndex        =   7
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
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
      Format          =   335216643
      UpDown          =   -1  'True
      CurrentDate     =   37802
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4080
      Left            =   90
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3030
      Width           =   15015
      _cx             =   26485
      _cy             =   7197
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
   Begin MSMask.MaskEdBox MEDuedate 
      Height          =   315
      Left            =   13590
      TabIndex        =   12
      Top             =   7440
      Width           =   1245
      _ExtentX        =   2196
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
      Format          =   "dd/MM/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker MyDuedate 
      Height          =   315
      Left            =   13590
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   7440
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
      Format          =   335216643
      CurrentDate     =   37802
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   12990
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   270
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin MSComCtl2.DTPicker dtpPEBDate 
      Height          =   315
      Left            =   11775
      TabIndex        =   9
      Top             =   2610
      Width           =   1665
      _ExtentX        =   2937
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
      CheckBox        =   -1  'True
      CustomFormat    =   "dd MMM yyyy"
      Format          =   335216643
      CurrentDate     =   39346
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PEB Date"
      Height          =   195
      Index           =   8
      Left            =   10905
      TabIndex        =   56
      Top             =   2670
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PEB Number"
      Height          =   195
      Index           =   4
      Left            =   7200
      TabIndex        =   55
      Top             =   2670
      Width           =   1065
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Air Freight Charge"
      Height          =   195
      Index           =   3
      Left            =   8160
      TabIndex        =   53
      Top             =   8160
      Width           =   1575
   End
   Begin MSForms.TextBox txtdesc 
      Height          =   315
      Left            =   3930
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2580
      Width           =   3075
      VariousPropertyBits=   746604563
      BackColor       =   16637923
      Size            =   "5424;556"
      BorderColor     =   16637923
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line3 
      X1              =   3930
      X2              =   7020
      Y1              =   2880
      Y2              =   2880
   End
   Begin MSForms.ComboBox cbocls 
      Height          =   315
      Left            =   3060
      TabIndex        =   5
      Top             =   2580
      Width           =   810
      VariousPropertyBits=   746604571
      MaxLength       =   25
      DisplayStyle    =   3
      Size            =   "1429;556"
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
      Caption         =   "Trade Terms"
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   51
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label pagelbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Page 0 of 0"
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
      Height          =   240
      Left            =   12840
      TabIndex        =   49
      Top             =   10560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Month"
      Height          =   195
      Index           =   2
      Left            =   10440
      TabIndex        =   48
      Top             =   2220
      Width           =   1290
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Date"
      Height          =   195
      Index           =   1
      Left            =   7200
      TabIndex        =   47
      Top             =   2220
      Width           =   1095
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No."
      Height          =   195
      Index           =   0
      Left            =   1530
      TabIndex        =   46
      Top             =   2250
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Index           =   7
      Left            =   150
      TabIndex        =   45
      Top             =   7500
      Width           =   765
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   1
      Left            =   120
      Top             =   8430
      Width           =   14940
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Index           =   1
      Left            =   6060
      TabIndex        =   44
      Top             =   8160
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PPn"
      Height          =   195
      Index           =   2
      Left            =   11040
      TabIndex        =   43
      Top             =   8160
      Width           =   315
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      Height          =   195
      Index           =   4
      Left            =   12990
      TabIndex        =   42
      Top             =   8160
      Width           =   1005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No."
      Height          =   195
      Index           =   0
      Left            =   3750
      TabIndex        =   41
      Top             =   8160
      Width           =   975
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   315
      Left            =   3060
      TabIndex        =   4
      Top             =   2160
      Width           =   3165
      VariousPropertyBits=   746604571
      MaxLength       =   25
      DisplayStyle    =   3
      Size            =   "5583;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FDDFE3&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Create (Export)"
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
      Height          =   435
      Index           =   3
      Left            =   270
      TabIndex        =   40
      Top             =   240
      Width           =   14595
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   195
      Index           =   3
      Left            =   12555
      TabIndex        =   39
      Top             =   7500
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   120
      Top             =   8070
      Width           =   14940
   End
End
Attribute VB_Name = "FrmInvoiceExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Update dudi santosa, Penambahan Field Service Desember 2008
Option Explicit
 
Dim rst As Recordset, rsmaster As Recordset, rsdetail As Recordset, rstpage As Recordset
Dim rssubtot As Recordset, rstcust As Recordset, rssumqty As Recordset, rspajak As Recordset
Dim i As Long, area As String
Dim Model As String, DelDate As Date, Dono As String
Dim currCode As String
Dim totalQty As Double, TotalAmount As Double
Dim blnupdate As Boolean, blndata As Boolean, blndisplay As Boolean
Dim bsave As Boolean, RsDel As Recordset, notfix As Boolean
Dim sqlx As String, sqla As String, sqlB As String, HakU As Integer
Public formpanggil As String, blntotal As Boolean, tempdealer As String
Public Tgl1 As String, Tgl2 As String, tppn As Double, xno As String
Dim tgl_sb As Byte, rstrade As Recordset, Overseas_Cls As String * 1, xpackno As String, abbr As String
Dim dblTemp As Double
Dim listPODate As String
Dim InvDate As Date

Dim bteColSelect As Byte
Dim bteColPackingNo As Byte
Dim bteColPartNumber As Byte
Dim bteColDesc As Byte
Dim bteColPONo As Byte
Dim bteColQty As Byte
Dim bteColRemain As Byte
Dim bteColUnit As Byte
Dim bteColPackingDate As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColDisc As Byte
Dim bteColAmount As Byte
Dim bteColNoCommercial As Byte
Dim bteColQtyTemp As Byte
Dim bteColUnitCls As Byte
Dim bteColCurrCode As Byte
Dim bteColItemCode As Byte
Dim bteColStatus As Byte
Dim bteColDelvDate As Byte
Dim bteColSeqNo As Byte
Dim bteColDOSeqNo As Byte
Dim bteColPriceTemp As Byte
Dim bteHakPrice As Byte
Dim bteColService As Byte 'Dudi Update
Dim BteColServiceTemp As Byte
Dim BteColSerialFrom As Byte
Dim BteColSerialTo As Byte

Private Sub Header()

    bteColSelect = 0
    bteColPackingNo = 1
    bteColPartNumber = 2
    bteColDesc = 3
    bteColPONo = 4
    bteColQty = 5
    
    BteColSerialFrom = 6
    BteColSerialTo = 7
    
    bteColRemain = 6 + 2
    bteColUnit = 7 + 2
    bteColPackingDate = 8 + 2
    bteColCurr = 9 + 2
    bteColPrice = 10 + 2
    bteColService = 11 + 2 'dudi Update
    bteColDisc = 12 + 2
    bteColAmount = 13 + 2
    bteColNoCommercial = 14 + 2
    bteColQtyTemp = 15 + 2
    bteColUnitCls = 16 + 2
    bteColCurrCode = 17 + 2
    bteColItemCode = 18 + 2
    bteColStatus = 19 + 2
    bteColDelvDate = 20 + 2
    bteColSeqNo = 21 + 2
    bteColDOSeqNo = 22 + 2
    bteColPriceTemp = 23 + 2
    BteColServiceTemp = 24 + 2

    With grid
    
        .clear
        .Rows = 1
        .ColS = 25 + 2
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColPackingNo) = "Packing.No (Ref No.)"
        .TextMatrix(0, bteColPartNumber) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColPONo) = "SI/PO No."
        .TextMatrix(0, bteColQty) = "Qty"
        
        .TextMatrix(0, BteColSerialFrom) = "Serial From"
        .TextMatrix(0, BteColSerialTo) = "Serial To"
        
        .TextMatrix(0, bteColRemain) = "Qty Rem"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColPackingDate) = "Packing Date"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColService) = "Service" 'tambahan dudi
        .TextMatrix(0, bteColDisc) = "Disc"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColNoCommercial) = "N/C"
        .TextMatrix(0, bteColQtyTemp) = "Qty Temp"
        .TextMatrix(0, bteColUnitCls) = "Unit Cls"
        .TextMatrix(0, bteColCurrCode) = "Curr Code"
        .TextMatrix(0, bteColItemCode) = "Item Code"
        .TextMatrix(0, bteColStatus) = "Status"
        .TextMatrix(0, bteColDelvDate) = "Delivery Date"
        .TextMatrix(0, bteColSeqNo) = "SeqNo"
        .TextMatrix(0, bteColDOSeqNo) = "DOSeqNo"
        .TextMatrix(0, bteColPriceTemp) = "Price Temp"
        .TextMatrix(0, BteColServiceTemp) = "Service Temp" 'tambahan dudi
        
        .ColWidth(bteColPartNumber) = 2000
        .ColWidth(bteColDesc) = 3700
        .ColWidth(bteColPONo) = 1500
        .ColWidth(bteColQty) = 1000
        
        .ColWidth(BteColSerialFrom) = 1100
        .ColWidth(BteColSerialTo) = 1100
        
        .ColWidth(bteColUnit) = 800
        .ColWidth(bteColPackingDate) = 1300
        .ColWidth(bteColCurr) = 800
        .ColWidth(bteColPrice) = 1350
        .ColWidth(bteColService) = 1350
        .ColWidth(bteColDisc) = 1000
        .ColWidth(bteColAmount) = 1800
        .ColWidth(bteColNoCommercial) = 600
        
        .ColHidden(0) = True
        .ColHidden(bteColPackingNo) = True
        .ColHidden(bteColRemain) = True
        .ColHidden(bteColDisc) = True
        .ColHidden(bteColQtyTemp) = True
        .ColHidden(bteColUnitCls) = True
        .ColHidden(bteColCurrCode) = True
        .ColHidden(bteColItemCode) = True
        .ColHidden(bteColStatus) = True
        .ColHidden(bteColDelvDate) = True
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColDOSeqNo) = True
        .ColHidden(bteColPriceTemp) = True
        .ColHidden(BteColServiceTemp) = True
        .ColHidden(bteColNoCommercial) = True
        
        If bteHakPrice = "0" Then
            .ColHidden(bteColCurr) = True
            .ColHidden(bteColPrice) = True
            .ColHidden(bteColService) = True
            .ColHidden(bteColDisc) = True
            .ColHidden(bteColAmount) = True
            Txtdisplay(1).Visible = False
            Txtdisplay(2).Visible = False
            Txtdisplay(3).Visible = False
            Label5(1).Visible = False
            Label5(2).Visible = False
            Label5(4).Visible = False
        Else
            .ColHidden(bteColCurr) = False
            .ColHidden(bteColPrice) = False
            .ColHidden(bteColService) = False
            .ColHidden(bteColDisc) = True
            .ColHidden(bteColAmount) = False
            Txtdisplay(1).Visible = True
            Txtdisplay(2).Visible = True
            Txtdisplay(3).Visible = True
            Label5(1).Visible = True
            Label5(2).Visible = True
            Label5(4).Visible = True
        End If
        
        .Cell(flexcpAlignment, 0, 1, 0, .ColS - 1) = flexAlignCenterCenter
        .Cell(flexcpBackColor, 0, 0, 0, .ColS - 1) = &HA6D2FF
    
    End With

End Sub

Private Sub cbocls_Change()
cboCls = Trim(cboCls)
If cboCls.MatchFound Then
    txtDesc.Text = cboCls.Column(1)
Else
    txtDesc.Text = ""
End If
End Sub

Private Sub cbodealer_Change()
    
    If cbodealer.MatchFound Then
        cboCls.Text = cbodealer.Column(4)
    Else
        cboCls.Text = ""
    End If
    
End Sub

Private Sub cbodealer_Click()

    MousePointer = vbHourglass
    rstcust.Requery
    rstcust.Find "Cust_code ='" & cbodealer & "'"
    If Not rstcust.EOF Then
        LblDesc(0).Caption = rstcust!Cust_Name
        LblDesc(1).Caption = rstcust!Address
        Overseas_Cls = rstcust!country_cls
        abbr = Left(rstcust!trade_abbr, 3)
        rstcust.Requery
    End If
    
    If combo1.Text = "Create" Then
        'delete
        If Trim(ComboBox1) <> "" Then
            sql = "select * from invoice_detail where invoice_no = '" & ComboBox1 & "'"
            Set RsDel = New Recordset
            RsDel.Open sql, Db, adOpenKeyset, adLockOptimistic
            If RsDel.EOF Then
                Db.Execute ("Delete from  invoice_master where invoice_no = '" & ComboBox1 & "'")
            End If
            Set RsDel = Nothing
        End If
        clear2
    Else
        'delete
        If Trim(ComboBox1) <> "" Then
            sql = "select * from invoice_detail where invoice_no = '" & ComboBox1 & "'"
            Set RsDel = New Recordset
            RsDel.Open sql, Db, adOpenKeyset, adLockOptimistic
            If RsDel.EOF Then
                Db.Execute ("Delete from  invoice_master where invoice_no = '" & ComboBox1 & "'")
            End If
            Set RsDel = Nothing
        End If
        
        nomorinvoice
        
        clear2
    End If
    'mengecek no packing ada atau tidak berdasar cust code dan tgl data-- dudi Januari 2008
    If CekSql("select packing_no, packing_date, etd from packing_master where packing_date >= '" & Format(SDate, "YYYY-MM-DD") & "' and packing_date <= '" & Format(EDate, "YYYY-MM-DD") & "' and cust_Code ='" & cbodealer & "'") Then
        nomorpacking
    Else
        CboPacking.clear
    End If
    
    
    MousePointer = vbDefault
    grid.Editable = False

End Sub
 Function CekSql(SqlData)
Dim Rdata As New ADODB.Recordset
Set Rdata = Nothing
Rdata.Open SqlData, Db, adOpenDynamic, adLockBatchOptimistic
If Rdata.EOF Then
CekSql = False
Else
CekSql = True
End If
Rdata.Close
End Function

Private Sub cbodealer_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    rstcust.Requery
    rstcust.Find "Cust_code ='" & cbodealer & "'"
    If Not rstcust.EOF Then
        cbodealer_Click
    Else
        lblerror = DisplayMsg(4072)
        cbodealer.SetFocus
    End If
End If
End Sub

Private Sub cbodealer_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdAction_Click(Index As Integer)
On Error GoTo ErrMsg
lblerror = ""
MousePointer = vbHourglass
Select Case Index
    Case 0
           If formpanggil = "invoiceinquiry" And cmdAction(0).Caption = "Back" Then
                frm_invoice_inquiry.ComboBox1 = Me.cbodealer
                frm_invoice_inquiry.set_tgl Tgl1, Tgl2
                frm_invoice_inquiry.combo1 = ComboBox1.Text
                Unload Me
                frm_invoice_inquiry.Show
                frm_invoice_inquiry.set_dari_inv_create
                frm_invoice_inquiry.xxx
            Else
                If Trim(ComboBox1.Text) <> "" Then
                    sql = "select * from invoice_detail where invoice_no = '" & ComboBox1 & "'"
                    Set RsDel = New Recordset
                    RsDel.Open sql, Db, adOpenKeyset, adLockOptimistic
                    If RsDel.EOF Then
                        Db.Execute ("Delete from  invoice_master where invoice_no = '" & ComboBox1 & "'")
                    End If
                    Set RsDel = Nothing
                End If
                Unload Me
                frmMainMenu.Show
            End If
            
    Case 1
            MousePointer = vbHourglass
            clear
            lblerror = ""
            CtrlMenu1.MenuText = ""
            cmdAction(5).Caption = "Update"
            combo1.Text = "Update"
            MousePointer = vbDefault
    Case 2
            MousePointer = vbHourglass
            If HakU = 0 Then lblerror = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
            savedetail
            tppn = tax("ppn")
            updateMaster
            listPO (ComboBox1)
            listDO (ComboBox1)
            inquiryupdate
            MousePointer = vbDefault
    Case 3
            Dim xdo As Recordset
            MousePointer = vbHourglass
            If Trim(ComboBox1) <> "" Then
                ComboBox1 = ComboBox1
                If ComboBox1.MatchFound Then
                    sql = "select distinct do_no from invoice_detail where invoice_no ='" & ComboBox1 & "'"
                    Set xdo = New Recordset
                    xdo.Open sql, Db, adOpenDynamic, adLockOptimistic
                    If xdo.EOF Then
                        lblerror = DisplayMsg(4071)
                        MousePointer = 1
                        Exit Sub
                    Else
                        inv_no = "'" & ComboBox1 & "'"
                        Call InvReportExport(bteHakPrice)
                    End If
                Else
                    lblerror = DisplayMsg(4071)
                End If
            End If
            MousePointer = vbDefault
    Case 4
    
            lblerror = up_ValidateDateRange(IDate, True)
            If lblerror <> "" Then Me.MousePointer = vbDefault: Exit Sub
    
            If (MsgBox("Are you sure want to delete this invoice ?", vbQuestion + vbYesNo, "Confirmation") = vbYes) Then
                Dim dbdel As New Connection
                dbdel.ConnectionString = Db.ConnectionString
                dbdel.Open
                dbdel.BeginTrans
                
                sql = "delete from invoice_detail where invoice_no ='" & ComboBox1 & "'"
                dbdel.Execute sql
                sql = "delete from invoice_master where invoice_no ='" & ComboBox1 & "'"
                dbdel.Execute sql
                If err.number = 0 Then
                    dbdel.CommitTrans
                    clear
                    Header
                    lblerror = DisplayMsg(1201)
                Else
                    dbdel.RollbackTrans
                    lblerror = err.Description
                End If
                dbdel.Close
                Set dbdel = Nothing
            End If
    Case 5

            If cbodealer = "" Then
                lblerror = DisplayMsg(1045)
                MousePointer = vbDefault
                cbodealer.SetFocus
                Exit Sub
            End If
            'mengecek packing no 'dudi 5 Januari 2009
            If CboPacking = "" Then
                lblerror = DisplayMsg(4101)
                MousePointer = vbDefault
                CboPacking.SetFocus
                Exit Sub
            ElseIf CboPacking.MatchFound = False Then
                lblerror = DisplayMsg(4010)
                MousePointer = vbDefault
                CboPacking.SetFocus
                Exit Sub
            End If
            
            
            'MENGECEK Invoice no 'dudi 5 Januari 2009
            If ComboBox1.Text = "" Then
                lblerror = DisplayMsg(50)
                MousePointer = vbDefault
                ComboBox1.SetFocus
                Exit Sub
            End If
            
            
            If HakU = 0 Then lblerror = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
            If cmdAction(5).Caption = "Create" Then
                
                lblerror = up_ValidateDateRange(IDate, True)
                If lblerror <> "" Then Me.MousePointer = vbDefault: Exit Sub
                
                cbodealer = cbodealer
                If cbodealer.MatchFound = False Then
                    lblerror = DisplayMsg(4072)
                    cbodealer.SetFocus
                    cbodealer.SetFocus
                    MousePointer = vbDefault
                    Exit Sub
                End If
                cboCls = Trim(cboCls)
                If Not cboCls.MatchFound Then
                    lblerror = DisplayMsg("0045")
                    cboCls.SetFocus
                    MousePointer = vbDefault
                    Exit Sub
                End If
                'update by dudi..mengecek packing,apakah terisi atau tidak 20090101
                If CboPacking = "" Then
                    lblerror = DisplayMsg("4101")
                    CboPacking.SetFocus
                    MousePointer = vbDefault
                Exit Sub
                ElseIf CboPacking.MatchFound = False Then
                    lblerror = DisplayMsg("4010")
                    CboPacking.SetFocus
                    MousePointer = vbDefault
                Exit Sub
                End If
                
                If Trim(ComboBox1.Text) <> "" Then
                    savemaster
                    inquiry
                    inquiryupdate
                    blndisplay = True
                    cmdAction(5).Caption = "Update"
                    bsave = True
                    combo1 = "Update"
                    grid.Editable = True
                    cbodealer.locked = False
                    ComboBox1.locked = False
                    xno = ComboBox1
                    nomorinvoice
                    ComboBox1 = Trim(xno)
                    lblerror.Caption = DisplayMsg(1000)
                    cmdAction(2).Enabled = True
                    blndisplay = False
                End If
            Else
                
                lblerror = up_ValidateDateRange(InvDate, True)
                If lblerror <> "" Then
                    inquiry
                    inquiryupdate
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
                
                If Trim(ComboBox1.Text) <> "" Then
                    ComboBox1 = Trim(ComboBox1)
                    If ComboBox1.MatchFound Then
                        tppn = tax("ppn")
                        If notfix Then updateMaster
                        inquiry
                        inquiryupdate
                        grid.Editable = True
                        If notfix Then lblerror.Caption = DisplayMsg(1101)
                        blndisplay = False
                    Else
                        lblerror = DisplayMsg(4071)
                    End If
                Else
                    lblerror = ""
                End If
            End If
            If lblfix <> "" Then
                grid.Editable = False
            Else
                grid.Editable = True
            End If
            
    Case 6
        If Trim(ComboBox1.Text) <> "" And combo1.Text = "Update" And grid.Rows <> 1 Then
            ComboBox1 = ComboBox1
            If ComboBox1.MatchFound Then
                inquiryupdate
                lblerror = ""
            Else
                lblerror = DisplayMsg(4071)
            End If
        End If
    Case 7
       Unload Me
      frmMainMenu.Show
End Select
MousePointer = vbDefault
Exit Sub
ErrMsg:
lblerror = err.number & " " & err.Description
MousePointer = vbDefault
Exit Sub
End Sub

Private Sub Combo1_Click()
If combo1.Text = "Create" Then
    If Trim(ComboBox1) <> "" Then
        sql = "select * from Invoice_detail where invoice_no = '" & ComboBox1 & "'"
        Set RsDel = New Recordset
        RsDel.Open sql, Db, adOpenKeyset, adLockOptimistic
        If RsDel.EOF Then
            Db.Execute ("Delete from  invoice_master where invoice_no = '" & ComboBox1 & "'")
        End If
        Set RsDel = Nothing
    End If
    rstcust.Requery
    rstcust.Find "Cust_code ='" & cbodealer & "'"
    If Not rstcust.EOF Then
        rstcust.Requery
        Header
    End If
    bsave = False
    blntotal = False
    clear2
    txtRemarks = ""
    txtPEBNo = ""
    dtpPEBDate = Null
    cmdAction(5).Caption = "Create"
    cmdAction(2).Enabled = False
    cbodealer.locked = False
    If GetInvoiceNumber = "" Then
    ComboBox1 = CboPacking
    'GetNoInvoice
    End If
    'ComboBox1.locked = True
    grid.Editable = False
    ComboBox1 = ""
    ComboBox1.locked = False
    ComboBox1.clear
Else
    If bsave Then Exit Sub
    Header
    clear2
    cmdAction(5).Caption = "Update"
    nomorinvoice
    ComboBox1 = ""
    ComboBox1.locked = False
End If
End Sub
Sub GetNoInvoice()
Dim s As String
  Dim RS As Recordset
    sql = "select LEFT((select Description FROM transportation_cls WHERE Transportation_Cls=a.transportation_cls),1) FROM packing_master a where Packing_No='" & CboPacking & "'"
    s = Get_Record(sql)
    
        sql = "select  substring(rtrim(Invoice_no),6,4) Nomor, invoice_no " & _
            "from invoice_master " & _
            "where year(invoice_date) = " & IDate.Year & _
            "AND LEN(Invoice_No)>10 order by invoice_no desc"
    
        Set RS = New Recordset
        RS.Open sql, Db, adOpenDynamic, adLockOptimistic
        If Not RS.EOF Then
            ComboBox1 = Format(IDate, "YYYY") & "-" & Format(Val(RS!nomor + 1), "0000") & "-" & IIf(cbodealer.Column(3) = "", Left(cbodealer.Text, 3), cbodealer.Column(3)) & "-" & UCase(s)
        Else
            ComboBox1 = Format(IDate, "YYYY") & "-0001-" & IIf(cbodealer.Column(3) = "", Left(cbodealer.Text, 3), cbodealer.Column(3)) & "-" & UCase(s)
        End If
    
End Sub

Private Sub ComboBox1_Change()
DoEvents
sql = "select * from invoice_master where invoice_no ='" & ComboBox1 & "'"
Set rst = New Recordset
rst.Open sql, Db, adOpenDynamic, adLockOptimistic
If Not rst.EOF Then
    ComboBox1 = Trim(rst!Invoice_No)
    InvDate = rst!Invoice_Date
    IDate = rst!Invoice_Date
    'DDate = Right(rst!delivery_Date, 2) & "/" & Left(rst!delivery_Date, 4)
    DDate = Month(rst!delivery_Date) & "/" & Year(rst!delivery_Date)
    Txtdisplay(0) = Trim(rst!Invoice_No)
    Txtdisplay(1) = Format(CDbl(rst!Amount), gs_formatAmountIDR)

     If IsNull(rst!AirFreightCharge) Then
        TxtAirCharge = Format(0, gs_formatAmountIDR)
    Else
        TxtAirCharge = Format(CDbl(rst!AirFreightCharge), gs_formatAmountIDR)
    End If

    If Overseas_Cls = "1" Then
        Txtdisplay(2) = Format(0, gs_formatAmountIDR)
    Else
        Txtdisplay(2) = Format((((CDbl(Txtdisplay(1)) + CDbl(TxtAirCharge)) * tax("Ppn")) / 100), gs_formatAmountIDR)
    End If
    Txtdisplay(3) = Format(CDbl(Txtdisplay(1)) + CDbl(Txtdisplay(2)) + CDbl(TxtAirCharge), gs_formatAmountIDR)
    
    If IsNull(rst!due_date) Then
        MEDuedate.Text = "99/99/9999"
    Else
        If Format(MyDuedate.Value, "dd") > 12 Then
            MEDuedate.Format = "dd/MM/yyyy"
        Else
            MEDuedate.Format = "MM/dd/yyyy"
        End If
        MEDuedate.Text = Format(rst!due_date, "DD") & "/" & Format(rst!due_date, "MM") & "/" & Format(rst!due_date, "YYYY")
    End If
    
    txtRemarks = Trim(rst!Remarks)
    txtPEBNo = Trim(rst!PEBNo & "")
    If Not IsNull(rst!PEBDate) Then dtpPEBDate = rst!PEBDate Else dtpPEBDate = Null
    If IsNull(rst!fix_cls) Then
        lblfix = ""
        notfix = True
        cmdAction(2).Enabled = True
        cmdAction(4).Enabled = True
    Else
        lblfix = "Status : Fix"
        notfix = False
        cmdAction(2).Enabled = False
        cmdAction(4).Enabled = False
    End If
    If Trim(ComboBox1) <> "" Then
        If blndisplay = False Or xno <> Trim(ComboBox1) Then grid.Rows = 1
    End If
    lblerror = ""
    PackingBaseInvoice
    If IsNull(rst!tradeterms_cls) Then
    cboCls = ""
    Else
        cboCls = rst!tradeterms_cls
    End If
Else
    lblfix = ""
    notfix = True
    If Trim(ComboBox1) <> "" Then
        If blndisplay = False Or xno <> Trim(ComboBox1) Then grid.Rows = 1
    End If
    lblerror = ""
End If
End Sub

Private Sub Combobox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboPacking_Click()

Dim s  As String
CboPacking = Trim(CboPacking)


If CboPacking.MatchFound Then
    DDate = CboPacking.Column(2)
    isiDueDate CboPacking.Column(2)
    'InvoiceBasePacking

End If
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    lblerror.Caption = ErrMsg
End If
End Sub

Private Sub DDate_Change()
DDate_Click
tgl_sb = DDate.Month
End Sub

Private Sub DDate_Click()
If DDate.Month = 1 And Val(tgl_sb) = 12 Then DDate.Year = DDate.Year + 1
If DDate.Month = 12 And Val(tgl_sb) = 1 Then DDate.Year = DDate.Year - 1
DDate.Value = Format(DDate.Value, "MMM yyyy")
End Sub

Private Sub DDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub edate_Change()

cbodealer = Trim(cbodealer)
If cbodealer.MatchFound Then
    If combo1.Text = "Create" Then
        clear2
        grid.Editable = False
    Else
        If Trim(CboPacking.Text) = "" Then nomorpacking: Exit Sub
        If Trim(ComboBox1.Text) = "" Then nomorinvoice: Exit Sub
        xpackno = CboPacking
        nomorpacking
        CboPacking = xpackno
        xno = ComboBox1.Text
        nomorinvoice
        ComboBox1 = xno
    End If
End If

End Sub

Private Sub edate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
adtocombo
bteHakPrice = hakPrice(Me.Name)
Header
HakU = hakUpdate(Me.Name)

SDate = Format(Date, "dd mmm yyyy")
EDate = Format(Date, "dd mmm yyyy")
IDate = Format(Date, "dd mmm yyyy")
DDate = Format(Date, "dd mmm yyyy")
dtpPEBDate.Value = Format(Date, "dd mmm yyyy")

cmdAction(2).Enabled = False

CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
combo1.ListIndex = 1
clear
cmdAction(5).Caption = "Update"
blndisplay = False
tgl_sb = Month(Date)
MyDuedate.Value = Format(Date, "dd mmm yyyy")


End Sub


Sub adtocombo()
Dim rstcls As New Recordset
'----------
'  Edit SQL Statement 20090202
'---------
sql = "SELECT  rtrim(Trade_Master.trade_Code) cust_code, rtrim(Trade_Master.Trade_Name) cust_name, " & _
    "rtrim(Trade_Master.Address1) address, isnull(country_Cls,'0') country_Cls, rtrim(trade_abbr) trade_abbr, Isnull(Price_Condition, '') price_con " & _
    "From Trade_Master where trade_cls in ('2') " 'and country_cls = '1'"
    
Set rstcust = New Recordset
rstcust.Open sql, Db, adOpenKeyset, adLockOptimistic
With cbodealer
    .clear
    .columnCount = 4
    .ColumnWidths = "50 pt;280 pt; 0 pt; 0 pt; 0 pt"
    .ListWidth = 350
    .ListRows = 15
    
i = 0
Do Until rstcust.EOF
    .AddItem ""
    .List(i, 0) = Trim(rstcust!Cust_CodE)
    .List(i, 1) = Trim(rstcust!Cust_Name)
    .List(i, 2) = IIf(IsNull(Trim(rstcust!Address)), "", Trim(rstcust!Address))
    .List(i, 3) = IIf(IsNull(Trim(rstcust!trade_abbr)), "", Trim(rstcust!trade_abbr))
    .List(i, 4) = IIf(IsNull(Trim(rstcust!price_con)), "", Trim(rstcust!price_con))
    i = i + 1
    rstcust.MoveNext
Loop
End With

sql = "select * from priceCondition_cls order by priceCondition_cls"
If rstcls.State <> adStateClosed Then rstcls.Close
rstcls.Open sql, Db, adOpenStatic, adLockOptimistic
With cboCls
    .clear
    .columnCount = 2
    .ColumnWidths = "50 pt;200 pt"
    .ListWidth = 250
    .ListRows = 15
    i = 0
    Do Until rstcls.EOF
        .AddItem ""
        .List(i, 0) = Trim(rstcls!PriceCondition_Cls)
        .List(i, 1) = Trim(rstcls!Description)
        i = i + 1
        rstcls.MoveNext
    Loop
End With

End Sub

Sub DisplayData(RS As Recordset, i As Long, Model As String, DDate As Date)
Dim L_price, L_Service As Double
With grid
        
    .TextMatrix(i, bteColPackingNo) = Trim(RS.Fields("packing_no").Value)
    If IsNull(RS.Fields("Part_no").Value) Then
        .TextMatrix(i, bteColPartNumber) = ""
    Else
        .TextMatrix(i, bteColPartNumber) = Trim(RS.Fields("Part_no").Value)
    End If
    .TextMatrix(i, bteColDesc) = Trim(RS.Fields("Item_name").Value)
    
    .TextMatrix(i, bteColPONo) = Trim(RS!Order_No)
    .TextMatrix(i, bteColQty) = Format((RS.Fields("Qty").Value), gs_formatQty)
    
    .TextMatrix(i, BteColSerialFrom) = Trim(RS!SerialNoFrom)
    .TextMatrix(i, BteColSerialTo) = Trim(RS!SerialNoTo)
    
    .TextMatrix(i, bteColRemain) = Format(0, gs_formatQty)
    .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(RS!Unit_cls)
    .TextMatrix(i, bteColPackingDate) = Format(RS.Fields("etd").Value, "dd Mmm YYYY")
    If Trim(RS!currency_code) <> "" Then .TextMatrix(i, bteColCurr) = uf_GetCurrencyDescription(RS!currency_code)
    
    
    sql = "select isnull(price,0) price,isnull(service,0)service  from invoice_detail where Invoice_No = '" & ComboBox1 & "' and packing_No ='" & Trim(RS.Fields("packing_no").Value) & "' and Po_no = '" & RS!Order_No & "' and seq_no= '" & RS!order_SeqNo & "' and packingseq_no = '" & RS!PackingSeq_No & "' "
    Set rsdetail = New Recordset
    rsdetail.Open sql, Db, adOpenDynamic, adLockOptimistic
    If Not rsdetail.EOF Then
        If InStr(1, rsdetail!Price, ".") Then
            .TextMatrix(i, bteColPrice) = Format(rsdetail.Fields("Price").Value, gs_formatPrice)
            .TextMatrix(i, bteColService) = Format(rsdetail.Fields("Service").Value, gs_formatPrice)
        Else
            .TextMatrix(i, bteColPrice) = Format(rsdetail.Fields("Price").Value, gs_formatPriceIDR)
            .TextMatrix(i, bteColService) = Format(rsdetail.Fields("Service").Value, gs_formatPriceIDR)
        End If
        If InStr(1, RS!AmountInv, ".") Then
            .TextMatrix(i, bteColAmount) = Format(RS.Fields("Amountinv").Value, gs_formatAmountIDR)
        Else
            .TextMatrix(i, bteColAmount) = Format(RS.Fields("Amountinv").Value, gs_formatAmountIDR)
        End If
    Else
        If InStr(1, RS!Price, ".") Then
            .TextMatrix(i, bteColPrice) = Format(RS.Fields("Price").Value, gs_formatPrice)
            .TextMatrix(i, bteColService) = Format(RS.Fields("Service").Value, gs_formatPrice)
        Else
            .TextMatrix(i, bteColPrice) = Format(RS.Fields("Price").Value, gs_formatPriceIDR)
            .TextMatrix(i, bteColService) = Format(RS.Fields("Service").Value, gs_formatPriceIDR)
        End If
        If InStr(1, RS!Amount, ".") Then
            .TextMatrix(i, bteColAmount) = Format(RS.Fields("Amount").Value, gs_formatAmountIDR)
        Else
            .TextMatrix(i, bteColAmount) = Format(RS.Fields("Amount").Value, gs_formatAmountIDR)
        End If
        
    End If
    
'    If (IsNull(rs!NoCommercial_Cls)) = False Then
'     If rs!NoCommercial_Cls = 0 Then
'      .TextMatrix(i, bteColNoCommercial) = "No"
'     Else
'      .TextMatrix(i, bteColNoCommercial) = "Yes"
'     End If
'    Else
'     .TextMatrix(i, bteColNoCommercial) = ""
'    End If
    
    
    L_price = .TextMatrix(i, bteColPrice)
    L_Service = .TextMatrix(i, bteColService) 'tambahan dudi
    
    .TextMatrix(i, bteColQtyTemp) = Format(RS.Fields("Qty").Value, gs_formatQty)
    .TextMatrix(i, bteColUnitCls) = RS.Fields("unit_Cls").Value
    If Trim(RS!currency_code) <> "" Then .TextMatrix(i, bteColCurrCode) = RS.Fields("currency_code").Value
    .TextMatrix(i, bteColItemCode) = RS.Fields("item_code").Value
    .TextMatrix(i, bteColStatus) = "0"
    .TextMatrix(i, bteColDelvDate) = Format(RS!etd, "dd mmm yyyy")
    .TextMatrix(i, bteColSeqNo) = RS!order_SeqNo
    .TextMatrix(i, bteColDOSeqNo) = RS!PackingSeq_No
    .TextMatrix(i, bteColPriceTemp) = L_price
    .TextMatrix(i, BteColServiceTemp) = L_Service   'tambahan dudi
    
    .Cell(flexcpAlignment, i, bteColPackingNo, i, bteColPONo) = flexAlignLeftCenter
    .Cell(flexcpAlignment, i, bteColQty, i, bteColRemain) = flexAlignRightCenter
    .Cell(flexcpBackColor, i, bteColPackingNo, i, bteColPONo) = &H80000018
    .Cell(flexcpBackColor, i, bteColRemain, i, bteColAmount) = &H80000018
    .Cell(flexcpBackColor, i, bteColPrice) = &H80000018 'vbWhite
    .Cell(flexcpBackColor, i, bteColService) = &H80000018 'vbWhite  'tambahan dudi
    totalQty = totalQty + .TextMatrix(i, bteColQty)
    TotalAmount = TotalAmount + .TextMatrix(i, bteColAmount)
    
End With
End Sub

Sub subtotal(Row As Long)
With grid
.Cell(flexcpBackColor, Row, bteColSelect) = vbWhite
.Cell(flexcpBackColor, Row, bteColPackingNo, Row, bteColNoCommercial) = &HE0E0E0
.Cell(flexcpChecked, Row, bteColSelect) = " "
.TextMatrix(Row, bteColPackingNo) = ""
.TextMatrix(Row, bteColPartNumber) = ""
.TextMatrix(Row, bteColDesc) = ""
.TextMatrix(Row, bteColQty) = Format(totalQty, gs_formatQty)
.TextMatrix(Row, bteColRemain) = ""
.TextMatrix(Row, bteColUnit) = ""
.TextMatrix(Row, bteColPackingDate) = ""
.TextMatrix(Row, bteColAmount) = Format(TotalAmount, gs_formatAmountIDR)
.TextMatrix(Row, bteColQtyTemp) = "subtotal" & Trim(.TextMatrix(Row - 1, bteColItemCode)) & CDbl(.TextMatrix(Row - 1, bteColPrice))
.TextMatrix(Row, bteColItemCode) = "0"
.Cell(flexcpAlignment, Row, bteColQty) = flexAlignRightCenter
.Cell(flexcpAlignment, Row, bteColAmount) = flexAlignRightCenter
End With
totalQty = 0
TotalAmount = 0

End Sub


Sub inquiryupdate()
Dim rupdate As Recordset
Dim Price As String, tempPrice As String

sql = " SELECT A.Packing_No,A.PackingSeq_No,A.Container_No," & vbCrLf
sql = sql & " a.Container_Size , a.Order_No, a.order_SeqNo, a.item_code" & vbCrLf
sql = sql & ",A.MakerItem_Code,A.Qty,A.SerialNoFrom,A.SerialNoTo,A.QtyWeight_Netto" & vbCrLf
sql = sql & ",A.QtyWeight_Gross,A.Qty_Volume,A.Qty_Ctn" & vbCrLf
sql = sql & ",A.Ctn_No,A.Detail_Cls,A.Unit_Cls,A.Currency_Code" & vbCrLf
sql = sql & ",A.DO_No,A.DOSeq_No,A.DO_Date,A.Price" & vbCrLf
sql = sql & ",ISNULL(od.Service,0) service ,A.Amount,A.Length,A.Width,A.Thickness" & vbCrLf
sql = sql & ",A.Last_Update,A.Last_User,A.Register_Date,b.etd," & vbCrLf
sql = sql & "rtrim(b.list_po) list_po, rtrim(b.list_podate) list_Podate," & vbCrLf
sql = sql & "rtrim(c.item_name) item_name, c.makeritem_Code Part_No," & vbCrLf
sql = sql & "isnull((select price from invoice_detail where LTRIM(packing_no) = LTRIM(a.packing_no) and packingseq_no = a.packingseq_no),0) priceinv," & vbCrLf
sql = sql & "isnull((select amount from invoice_detail where LTRIM(packing_no) = LTRIM(a.packing_no) and packingseq_no = a.packingseq_no),0) amountinv" & vbCrLf
sql = sql & ",isnull((select service from invoice_detail where LTRIM(packing_no) = LTRIM(a.packing_no) and packingseq_no = a.packingseq_no),0) serviceInv," & vbCrLf
sql = sql & " om.nocommercial_cls  from packing_detail a, packing_master  b , item_master c,  " & vbCrLf
sql = sql & " orderentry_detail od, orderentry_master om" & vbCrLf
sql = sql & " where LTRIM(a.packing_no) = LTRIM(b.packing_no)  and a.item_code = c.item_code" & vbCrLf & "  and od.PO_No = a.Order_No And od.Seq_No = a.order_SeqNo   and om.PO_No = od.PO_No And om.cust_code = od.cust_code  " & vbCrLf
sql = sql & " and ltrim(a.packing_no) ='" & CboPacking & "'" & vbCrLf
sql = sql & " order by om.nocommercial_cls, c.group_cls, a.makeritem_code" & vbCrLf

Set rupdate = New Recordset
rupdate.CursorLocation = adUseClient
rupdate.Open sql, Db, adOpenKeyset, adLockOptimistic
If rupdate.EOF = False Then
With rupdate
    totalQty = 0
    TotalAmount = 0
    grid.Rows = 1
    currCode = ""
    For i = 1 To .RecordCount
        If Not .EOF Then 'selama ada model yang sama
            If !Priceinv = 0 Then
                tempPrice = !Price
            Else
                tempPrice = !Priceinv
            End If

            If i = 1 Then
                Model = !Item_Code
                DelDate = !etd
                Price = tempPrice
                Dono = !packing_no
                grid.Rows = grid.Rows + 1
                Call DisplayData(rupdate, grid.Rows - 1, Model, DelDate)
            Else
                If !Item_Code = Model Then 'jika code shop sekarang sama dgn code shop pada hal sebelumnya
                    If Price = tempPrice Then
                        DelDate = !etd
                        grid.Rows = grid.Rows + 1
                        Call DisplayData(rupdate, grid.Rows - 1, Model, DelDate)
                    Else
                        grid.Rows = grid.Rows + 1
                        Call subtotal(grid.Rows - 1)
                        Model = !Item_Code
                        DelDate = !etd
                        Price = tempPrice
                        Dono = !packing_no
                        grid.Rows = grid.Rows + 1
                        Call DisplayData(rupdate, grid.Rows - 1, Model, DelDate)
                    End If
                Else
                    grid.Rows = grid.Rows + 1
                    Call subtotal(grid.Rows - 1)
                    Model = !Item_Code
                    DelDate = !etd
                    Dono = !packing_no
                    Price = tempPrice
                    grid.Rows = grid.Rows + 1
                    Call DisplayData(rupdate, grid.Rows - 1, Model, DelDate)
                    grid.MergeCells = flexMergeFree
                End If
            End If
            
        Else 'jika tdk ada data sesudahnya maka total

            Exit For
        End If
        .MoveNext
    Next i
    grid.Rows = grid.Rows + 1
    Call subtotal(grid.Rows - 1)
End With
Else


End If
End Sub

Sub inquiry()
sql = "select * from invoice_master where invoice_no ='" & ComboBox1 & "'"
Set rst = New Recordset
rst.Open sql, Db, adOpenDynamic, adLockOptimistic
If Not rst.EOF Then
    ComboBox1 = Trim(rst!Invoice_No)
    InvDate = rst!Invoice_Date
    IDate = rst!Invoice_Date
    DDate = Right(rst!delivery_Date, 2) & "/" & Left(rst!delivery_Date, 4)
    
    If IsNull(rst!due_date) Then
        MEDuedate.Text = "99/99/9999"
    Else
        If Format(MyDuedate.Value, "dd") > 12 Then
            MEDuedate.Format = "dd/MM/yyyy"
        Else
            MEDuedate.Format = "MM/dd/yyyy"
        End If
        MEDuedate.Text = Format(rst!due_date, "DD") & "/" & Format(rst!due_date, "MM") & "/" & Format(rst!due_date, "YYYY")
    End If
    
    Txtdisplay(0) = Trim(rst!Invoice_No)
    Txtdisplay(1) = Format(CDbl(rst!Amount), gs_formatAmountIDR)
     
     If IsNull(rst!AirFreightCharge) Then
        TxtAirCharge = Format(0, gs_formatAmountIDR)
    Else
        TxtAirCharge = Format(CDbl(rst!AirFreightCharge), gs_formatAmountIDR)
    End If
    
    If Overseas_Cls = "1" Then
        Txtdisplay(2) = Format(0, gs_formatAmountIDR)
    Else
        Txtdisplay(2) = Format((((CDbl(Txtdisplay(1)) + CDbl(TxtAirCharge)) * tax("Ppn")) / 100), gs_formatAmountIDR)
    End If
    Txtdisplay(3) = Format(CDbl(Txtdisplay(1)) + CDbl(Txtdisplay(2)) + CDbl(TxtAirCharge), gs_formatAmountIDR)
    cboCls = Trim(rst!tradeterms_cls)
    txtRemarks = Trim(rst!Remarks)
    txtPEBNo = Trim(rst!PEBNo & "")
    If Not IsNull(rst!PEBDate) Then dtpPEBDate = rst!PEBDate Else dtpPEBDate = Null
    Dim a As String
    a = ComboBox1
    cbodealer = Trim(rst!Cust_CodE)
    ComboBox1 = a
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
lblerror = ""
With grid
If Col = bteColPrice Or Col = bteColService Then
    If Not IsNumeric(.TextMatrix(Row, bteColPrice)) Or Not IsNumeric(.TextMatrix(Row, bteColService)) Then
        If Not IsNumeric(.TextMatrix(Row, bteColPrice)) Then
        .TextMatrix(Row, bteColPrice) = Format(0, gs_formatPrice)
        End If
        If Not IsNumeric(.TextMatrix(Row, bteColService)) Then
        .TextMatrix(Row, bteColService) = Format(0, gs_formatPrice)
        End If
    Else
        If InStr(1, .TextMatrix(Row, bteColPrice), ".") Then
            .TextMatrix(Row, bteColPrice) = Format(.TextMatrix(Row, bteColPrice), gs_formatPrice)
        Else
            .TextMatrix(Row, bteColPrice) = Format(.TextMatrix(Row, bteColPrice), gs_formatPriceIDR)
        End If
        If InStr(1, .TextMatrix(Row, bteColPrice), ".") Then
            .TextMatrix(Row, bteColService) = Format(.TextMatrix(Row, bteColService), gs_formatPrice)
        Else
            .TextMatrix(Row, bteColService) = Format(.TextMatrix(Row, bteColService), gs_formatPriceIDR)
        End If
    End If
    
    If CDbl(.TextMatrix(Row, Col)) > gd_MaxPrice Then
        lblerror = DisplayMsg(4048) & " " & gd_MaxPrice & " !"
        .TextMatrix(Row, bteColPrice) = Format(dblTemp, gs_formatPrice)
    End If
'    .TextMatrix(Row, bteColAmount) = Format(uf_Trunc(.TextMatrix(Row, bteColQty) * .TextMatrix(Row, bteColPrice), gi_decimalDigitAmountIDR), gs_formatAmountIDR)
'    gridsubtotal Trim(.TextMatrix(Row, bteColItemCode)), CDbl(.TextMatrix(Row, bteColPriceTemp))
 '   totalbasechecked
    .TextMatrix(Row, bteColStatus) = 1
End If

If Col = bteColService Then

    If Not IsNumeric(.TextMatrix(Row, bteColService)) Then
        .TextMatrix(Row, bteColService) = Format(0, gs_formatPrice)
    Else
        If InStr(1, .TextMatrix(Row, bteColService), ".") Then
            .TextMatrix(Row, bteColService) = Format(.TextMatrix(Row, bteColService), gs_formatPrice)
        Else
            .TextMatrix(Row, bteColService) = Format(.TextMatrix(Row, bteColService), gs_formatPrice)
        End If
    End If
    
    If CDbl(.TextMatrix(Row, bteColService)) > gd_MaxPrice Then
        lblerror = DisplayMsg(4048) & " " & gd_MaxPrice & " !"
        .TextMatrix(Row, bteColService) = Format(dblTemp, gs_formatPrice)
    End If

End If
If Col = bteColService Or Col = bteColPrice Then
 Dim dd As Double
 dd = (CDbl(.TextMatrix(Row, bteColPrice)) + CDbl(.TextMatrix(Row, bteColService)))
 .TextMatrix(Row, bteColAmount) = Format(uf_Trunc(.TextMatrix(Row, bteColQty) * (dd), gi_decimalDigitAmountIDR), gs_formatAmountIDR)
  gridsubtotal Trim(.TextMatrix(Row, bteColItemCode)), CDbl(.TextMatrix(Row, bteColPriceTemp))
  totalbasechecked
 
End If
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = bteColPrice Or Col = bteColService Then Cancel = True 'Price Dan SErvice tidak boleh di edit

If grid.Col <> bteColPrice And grid.Col <> bteColService Then Cancel = True ' Or Grid.Col <> bteColService Then Cancel = True
If Col = bteColPrice Then
    dblTemp = CDbl(IIf(grid.TextMatrix(Row, bteColPrice) = "", 0, grid.TextMatrix(Row, bteColPrice)))
    grid.EditMaxLength = 12
ElseIf Col = bteColQty Then
    dblTemp = grid.TextMatrix(Row, bteColQty)
    grid.EditMaxLength = 7
End If
End Sub

Private Sub grid_Click()
If grid.Rows <> 1 And grid.Row <> -1 Then
With grid
    If (.Col = bteColPrice) And Trim(.TextMatrix(.Row, bteColPackingNo)) <> "" Then
        .FocusRect = flexFocusInset
    Else
        .FocusRect = flexFocusNone
    End If
End With
End If
End Sub

Private Sub IDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub lbldesc_Change(Index As Integer)
If Index = 0 Then Text1 = LblDesc(0)
End Sub

Private Sub MEDuedate_Change()
If IsDate(MEDuedate.Text) Then
    MyDuedate = Format(MEDuedate, "dd mm yyyy")
End If
End Sub

Private Sub MyDuedate_Change()
MEDuedate.Text = Format(MyDuedate.Value, "dd") & "/" & Format(MyDuedate.Value, "MM") & "/" & Format(MyDuedate.Value, "yyyy")
If Format(MyDuedate.Value, "dd") > 12 Then
    MEDuedate.Format = "dd/MM/yyyy"
Else
    MEDuedate.Format = "MM/dd/yyyy"
End If
If MyDuedate < IDate Then
    lblerror = DisplayMsg(4110)
Else
    lblerror = ""
End If
End Sub

Private Sub sdate_Change()
cbodealer = Trim(cbodealer)
If cbodealer.MatchFound Then
    If combo1.Text = "Create" Then
        clear2
        grid.Editable = False
        nomorpacking
    Else
        If Trim(CboPacking.Text) = "" Then nomorpacking
        If Trim(ComboBox1.Text) = "" Then nomorinvoice: Exit Sub
        xpackno = CboPacking
        nomorpacking
        CboPacking = xpackno
        xno = ComboBox1.Text
        nomorinvoice
        ComboBox1 = xno
    End If
End If
End Sub

Private Sub sdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub Txtdisplay_GotFocus(Index As Integer)
ComboBox1.SetFocus
End Sub

Sub savemaster()
Dim xduedate As String

If MEDuedate.Text = "99/99/9999" Then
    xduedate = Null
Else
    xduedate = Format(MyDuedate, "YYYY-MM-DD")
End If

sql = "select * from invoice_master where invoice_no = '" & ComboBox1.Text & "'"
Set rsmaster = New Recordset
rsmaster.Open sql, Db, adOpenDynamic, adLockOptimistic
With rsmaster
If Not rsmaster.EOF Then
    .Fields("Invoice_No") = ComboBox1
    .Fields("Invoice_date") = Format(IDate.Value, "YYYY-MM-dd")
    '.Fields("Delivery_date") = Format(DDate, "YYYYMM")
    .Fields("Delivery_date") = Format(DDate, "dd-MMM-YYYY")
    .Fields("Cust_code") = cbodealer.Text
    If Val(Txtdisplay(1)) = 0 Then
        .Fields("Amount") = Val(Txtdisplay(1))
    Else
        .Fields("Amount") = CDbl(Txtdisplay(1))
    End If
    If Val(Txtdisplay(2)) = 0 Then
        .Fields("PPN") = Val(Txtdisplay(2))
    Else
        .Fields("PPN") = CDbl(Txtdisplay(2))
    End If
    If IsNumeric(TxtAirCharge) = False Then TxtAirCharge = Format(0, gs_formatAmountIDR)
    .Fields("AirFreightCharge") = CDbl(TxtAirCharge)
    TxtAirCharge = Format(TxtAirCharge, gs_formatAmountIDR)
    Txtdisplay(3) = Format(Val(Txtdisplay(1)) + Val(Txtdisplay(2)) + Val(TxtAirCharge), gs_formatAmountIDR)
    If Val(Txtdisplay(3)) = 0 Then
        .Fields("Total_amount") = Val(Txtdisplay(3))
    Else
        .Fields("Total_Amount") = CDbl(Txtdisplay(3))
    End If
    .Fields("due_date") = xduedate
    .Fields("Remarks") = Trim(txtRemarks)
    .Fields("PEBNo") = Trim(txtPEBNo)
    .Fields("PEBDate") = dtpPEBDate
    .Fields("tradeterms_Cls") = cboCls
    .update
    lblerror.Caption = DisplayMsg(1101)
Else
    .AddNew
    If IsNumeric(TxtAirCharge) = False Then TxtAirCharge = Format(0, gs_formatAmountIDR)
    .Fields("AirFreightCharge") = CDbl(TxtAirCharge)
    TxtAirCharge = Format(TxtAirCharge, gs_formatAmountIDR)
    .Fields("Invoice_No") = ComboBox1
    .Fields("Invoice_date") = Format(IDate, "yyyy-mm-dd")
    .Fields("Delivery_date") = Format(IDate, "yyyy-mm-dd")
    .Fields("Cust_code") = cbodealer.Text
    .Fields("Amount") = Val(Txtdisplay(1))
    .Fields("PPN") = Val(Txtdisplay(2))
    Txtdisplay(3) = Format(Val(Txtdisplay(1)) + Val(Txtdisplay(2)) + Val(TxtAirCharge), gs_formatAmountIDR)
    .Fields("Total_Amount") = Val(Txtdisplay(3))
    .Fields("Remarks") = txtRemarks
    .Fields("PEBNo") = txtPEBNo
    .Fields("PEBDate") = dtpPEBDate
    .Fields("due_date") = xduedate
    .Fields("tradeterms_Cls") = cboCls
    On Error Resume Next
    .update
HandleError:
    If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
        err.clear
        ComboBox1.locked = False
        If IsNumeric(TxtAirCharge) = False Then TxtAirCharge = Format(0, gs_formatAmountIDR)
        .Fields("AirFreightCharge") = CDbl(TxtAirCharge)
        TxtAirCharge = Format(TxtAirCharge, gs_formatAmountIDR)
        .Fields("Invoice_No") = ComboBox1.Text
        .Fields("Invoice_date") = Format(IDate, "yyyy-mm-dd")
        .Fields("Delivery_date") = Format(DDate, "YYYYMM")
        .Fields("Cust_code") = cbodealer.Text
        .Fields("Amount") = Val(Txtdisplay(1))
        .Fields("PPN") = Val(Txtdisplay(2))
        Txtdisplay(3) = Format(Val(Txtdisplay(1)) + Val(Txtdisplay(2)) + Val(TxtAirCharge), gs_formatAmountIDR)
        .Fields("Total_Amount") = Val(Txtdisplay(3))
        .Fields("Remarks") = txtRemarks
        .Fields("PEBNo") = txtPEBNo
        .Fields("PEBDate") = dtpPEBDate
        .Fields("due_date") = xduedate
        .Fields("tradeterms_Cls") = cboCls
        .update
        If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then GoTo HandleError
    End If
    lblerror.Caption = DisplayMsg(1000)
End If
End With
End Sub

Sub savedetail()

    For i = 1 To grid.Rows - 1
        If grid.TextMatrix(i, 1) <> "" Then
                sql = "select * from invoice_detail where Invoice_No = '" & ComboBox1 & "' and packing_No ='" & grid.TextMatrix(i, bteColPackingNo) & "' and Po_no = '" & grid.TextMatrix(i, bteColPONo) & "' and seq_no= '" & grid.TextMatrix(i, bteColSeqNo) & "' and packingseq_no = '" & grid.TextMatrix(i, bteColDOSeqNo) & "' "
                Set rsdetail = New Recordset
                rsdetail.Open sql, Db, adOpenDynamic, adLockOptimistic
                With rsdetail
                Dim lb_Insert As Boolean
                If .EOF Then
                    .AddNew
                End If
                
                    .Fields("invoice_no") = ComboBox1
                    .Fields("packing_no") = CboPacking
                    .Fields("DO_No") = ""
                    .Fields("item_code") = grid.TextMatrix(i, bteColItemCode)
                    .Fields("MakerItem_code") = grid.TextMatrix(i, bteColPartNumber)
                    .Fields("PO_no") = grid.TextMatrix(i, bteColPONo)

                    .Fields("Qty") = CDbl(grid.TextMatrix(i, bteColQty))
                    .Fields("Unit_cls") = grid.TextMatrix(i, bteColUnitCls)
                    .Fields("Delivery_date") = Format(grid.TextMatrix(i, bteColDelvDate), "yyyy-mm-dd")
                    .Fields("currency_code") = grid.TextMatrix(i, bteColCurrCode)
                    .Fields("Price") = CDbl(grid.TextMatrix(i, bteColPrice))
                    .Fields("Service") = CDbl(grid.TextMatrix(i, bteColService))
                    .Fields("Amount") = CDbl(grid.TextMatrix(i, bteColAmount))
                    .Fields("Seq_no") = CDbl(grid.TextMatrix(i, bteColSeqNo))
                    .Fields("packingseq_no") = grid.TextMatrix(i, bteColDOSeqNo)
                    .Fields("doSeq_no") = 0
                    .update
                    

                    lblerror.Caption = DisplayMsg(1000)

                End With

        End If
    Next

End Sub

Sub totalbasechecked()
Dim AmountInv As Double
With grid
AmountInv = 0
For i = 1 To grid.Rows - 1
    If .Cell(flexcpChecked, i, bteColSelect) = 1 Then
        If Val(AmountInv) = 0 Then
            AmountInv = CDbl(.TextMatrix(i, bteColAmount))
        Else
            AmountInv = CDbl(AmountInv) + CDbl(.TextMatrix(i, bteColAmount))
        End If
    End If
Next
Txtdisplay(1).Text = Format(uf_Trunc(AmountInv, gi_decimalDigitAmountIDR), gs_formatAmountIDR)
If Overseas_Cls = "0" Then
    Txtdisplay(2).Text = Format((AmountInv * tax("ppn")) / 100, gs_formatAmountIDR)
Else
    Txtdisplay(2).Text = Format("0", gs_formatAmountIDR)
End If
Txtdisplay(3).Text = Format(CDbl(Txtdisplay(1).Text) + CDbl(Txtdisplay(2).Text) + CDbl(TxtAirCharge.Text), gs_formatAmountIDR)
Txtdisplay(0).Text = ComboBox1
End With
End Sub

Sub clear()
cbodealer.ListIndex = -1
SDate = Format(Date, "dd mmm yyyy")
EDate = Format(Date, "dd mmm yyyy")
IDate = Format(Date, "dd mmm yyyy")
DDate = Format(Date, "dd mmm yyyy")
ComboBox1 = ""
CboPacking = ""
txtRemarks = ""
txtPEBNo = ""
dtpPEBDate = Null
LblDesc(0).Caption = ""
LblDesc(1).Caption = ""
Txtdisplay(0) = ""
Txtdisplay(1) = Format(0, gs_formatAmountIDR)
Txtdisplay(2) = Format(0, gs_formatAmountIDR)
Txtdisplay(3) = Format(0, gs_formatAmountIDR)
TxtAirCharge = Format(0, gs_formatAmountIDR)
'MEDuedate.Text = "99/99/9999"
MEDuedate.Text = Format(Date, "dd/mm/yyyy")
cmdAction(2).Enabled = False
cmdAction(5).Caption = "Update"
cbodealer.locked = False
ComboBox1.locked = False
ComboBox1 = ""
ComboBox1.clear
bsave = False
currCode = ""
End Sub

Sub clear2()
    Txtdisplay(0) = ""
    Txtdisplay(1) = Format(0, gs_formatAmountIDR)
    Txtdisplay(2) = Format(0, gs_formatAmountIDR)
    Txtdisplay(3) = Format(0, gs_formatAmountIDR)
    TxtAirCharge = Format(0, gs_formatAmountIDR)
    txtRemarks = ""
   MEDuedate.Text = Format(Date, "dd/mm/yyyy")
    
End Sub

Sub updateMaster()
Dim rsUpdate As Recordset, xrs As Recordset

sql = "select * from invoice_master where invoice_no = '" & ComboBox1 & "'"
Set rsUpdate = New Recordset
rsUpdate.Open sql, Db, adOpenKeyset, adLockOptimistic
With rsUpdate
If Not .EOF Then
    sql = "select invoice_no ,sum(amount) Amount from invoice_detail where invoice_no ='" & ComboBox1 & "' group by invoice_no"
    Set xrs = New Recordset
    xrs.Open sql, Db, adOpenDynamic, adLockOptimistic
    If xrs.EOF Then
        Txtdisplay(1) = Format(0, gs_formatAmountIDR)
    Else
        Txtdisplay(1) = Format(xrs!Amount, gs_formatAmountIDR)
    End If
    
   If IsNumeric(TxtAirCharge) = False Then TxtAirCharge = Format(0, gs_formatAmountIDR)
    TxtAirCharge = Format(TxtAirCharge, gs_formatAmountIDR)
    
    If Overseas_Cls = "1" Then
        Txtdisplay(2) = Format(0, gs_formatAmountIDR)
    Else
        Txtdisplay(2) = Format((((CDbl(Txtdisplay(1)) + CDbl(TxtAirCharge)) * tax("Ppn")) / 100), gs_formatAmountIDR)
    End If
    Txtdisplay(3) = Format(CDbl(Txtdisplay(1)) + CDbl(Txtdisplay(2)) + CDbl(TxtAirCharge), gs_formatAmountIDR)
    
    !Amount = CDbl(Txtdisplay(1))
    !ppn = CDbl(Txtdisplay(2))  'tppn * CDbl(Txtdisplay(1)) / 100
    !total_amount = Txtdisplay(3)
    !AirFreightCharge = CDbl(TxtAirCharge)
    !Remarks = Trim(txtRemarks.Text)
    !PEBNo = Trim(txtPEBNo.Text)
    !PEBDate = dtpPEBDate
    !Invoice_Date = Format(IDate, "yyyy-mm-dd")
    !delivery_Date = Format(IDate, "yyyy-mm-dd")
    If MEDuedate.Text = "99/99/9999" Then
        !due_date = Null
    Else
        !due_date = Format(MyDuedate, "YYYY-MM-DD")
    End If
    !tradeterms_cls = cboCls
    .update
End If
End With
End Sub

Function tax(kode$) As String
Dim rtax As Recordset, tempidate As String

tempidate = Format(IDate.Value, "yyyymmdd")
sql = "SELECT Tax_Code, Rate, start_Date, End_Date " & _
        "FROM tax_cls " & _
        "" & _
        "WHERE  " & _
        "start_date <= '" & tempidate & "'  and end_date >= '" & tempidate & "' " & _
        "and tax_code= '" & kode & "'"
Set rtax = New Recordset
rtax.Open sql, Db, adOpenKeyset, adLockOptimistic
If Not rtax.EOF Then
    tax = rtax!rate
Else
    tax = 0
End If

End Function

Sub nomorinvoice()
Dim noi As Integer
'sql = "select distinct invoice_master .* from invoice_master, invoice_detail where invoice_master.invoice_no *= invoice_detail.invoice_no and cust_code = '" & cbodealer.Text & "' and year(invoice_date) >= '" & year(sdate) & "' and year(invoice_date) <= '" & year(edate) & "' and packing_no <> '0'"

sql = " SELECT DISTINCT " & vbCrLf & _
            "         invoice_master.* " & vbCrLf & _
            " FROM    invoice_master INNER join " & vbCrLf & _
            "         invoice_detail ON invoice_master.invoice_no = invoice_detail.invoice_no " & vbCrLf & _
            " WHERE   cust_code = '" & cbodealer.Text & "' " & vbCrLf & _
            "         AND YEAR(invoice_date) >= '" & Year(SDate) & "' " & vbCrLf & _
            "         AND YEAR(invoice_date) <= '" & Year(EDate) & "' " & vbCrLf & _
            "         AND packing_no <> '0' "

Set rsmaster = Db.Execute(sql)
If Not rsmaster.EOF Then
    ComboBox1.clear
    ComboBox1.locked = False
    ComboBox1.columnCount = 2
    ComboBox1.ColumnWidths = "110 pt;0 pt; 0 pt;"
    noi = 0
    Do While Not rsmaster.EOF
        Me.ComboBox1.AddItem
           ComboBox1.List(noi, 0) = Trim(rsmaster!Invoice_No)
           ComboBox1.List(noi, 1) = IIf(IsNull(rsmaster!fix_cls), "0", rsmaster!fix_cls)
        rsmaster.MoveNext
        noi = noi + 1
    Loop
Else
    ComboBox1.clear
    ComboBox1.locked = True
End If
ComboBox1 = ""
rsmaster.Close
Set rsmaster = Nothing

End Sub

Public Sub set_tgl(a As String, b As String)
Tgl1 = a
Tgl2 = b
End Sub

Public Sub dr_invInq(p_custCode As String, p_invNo As String)
cbodealer.Text = p_custCode
ComboBox1.Text = p_invNo
End Sub

Sub ShowData()
cmdAction_Click (5)
End Sub

Sub gridsubtotal(Model$, xprice As Double)
Dim qtytemp As Double, bol As Boolean
Dim pricetemp As Double, disctemp As Double, tempamount As Double
Dim tempService As Double
qtytemp = 0
pricetemp = 0
tempamount = 0
With grid
For i = 1 To grid.Rows - 1
    If UCase(Trim(.TextMatrix(i, bteColQtyTemp))) = UCase("subtotal" & Trim(Model) & xprice) Then Exit For
        If Trim(.TextMatrix(i, bteColPrice)) = "" Then
            pricetemp = 0
        Else
            pricetemp = CDbl(.TextMatrix(i, bteColPriceTemp))
        End If
       
    
    
    If Trim(.TextMatrix(i, bteColItemCode)) & pricetemp = Trim(Model) & xprice Then
        qtytemp = qtytemp + .TextMatrix(i, bteColQty)
        tempamount = tempamount + .TextMatrix(i, bteColAmount)
    End If
    
Next
.TextMatrix(i, bteColQty) = Format(qtytemp, gs_formatQty)
.TextMatrix(i, bteColAmount) = Format(tempamount, gs_formatAmountIDR)
End With

End Sub

Sub isiDueDate(dtedate As Date)
Dim rstrade As New ADODB.Recordset
    sql = "Select ISNULL(InvoicePay_days,0) InvoicePayDay From Trade_Master Where Trade_Code = '" & cbodealer & "'"
    Set rstrade = Db.Execute(sql)
    If Not rstrade.EOF Then
        If Trim(rstrade!InvoicePayDay) = "" Then
            MyDuedate = Format(dtedate, "dd MMM yyyy")
        Else
            MyDuedate = Format(DateAdd("d", rstrade!InvoicePayDay, dtedate), "dd MMM yyyy")
        End If
    Else
        MyDuedate = Format(dtedate, "dd MMM yyyy")
    End If
    Call MyDuedate_Change
End Sub

Sub nomorpacking()
Dim rspack As Recordset
sql = "select packing_no, packing_date, etd from packing_master where packing_date >= '" & Format(SDate, "YYYY-MM-DD") & "' and packing_date <= '" & Format(EDate, "YYYY-MM-DD") & "' and cust_Code ='" & cbodealer & "'"
Set rspack = Db.Execute(sql)
If Not rspack.EOF Then
    With CboPacking
        .clear
        .columnCount = 3
        .ColumnWidths = "120pt;0pt;0pt"
        
        Do Until rspack.EOF
            .AddItem ""
            .Column(0, .ListCount - 1) = Trim(rspack!packing_no)
            .Column(1, .ListCount - 1) = Trim(rspack!packing_date)
            .Column(2, .ListCount - 1) = Trim(IIf(IsNull(rspack!etd), "", rspack!etd))
            rspack.MoveNext
        Loop
'        .Text = .Column(0, 0)
    End With
Else
    CboPacking.clear
    CboPacking.Text = ""
End If
rspack.Close
Set rspack = Nothing
End Sub

Sub InvoiceBasePacking()
Dim rsbase As Recordset, Idx As Long

sql = "select distinct invoice_no from invoice_detail where packing_no ='" & Trim(CboPacking.Text) & "'"
Set rsbase = Db.Execute(sql)
If Not rsbase.EOF Then
    combo1.Text = "Update"
    cmdAction(5).Caption = "Update"
    ComboBox1.Text = Trim(rsbase!Invoice_No)
Else
    ComboBox1 = ""
    combo1.Text = "Create"
    cmdAction(5).Caption = "Create"
    ComboBox1 = CboPacking
    If CboPacking.ListCount = 1 Then
    
    End If
    
End If


End Sub

Sub PackingBaseInvoice()
Dim rstbase As Recordset, Idx As Long
sql = "select Packing_no from invoice_detail where invoice_no = '" & Trim(ComboBox1) & "'"
Set rstbase = Db.Execute(sql)
If Not rstbase.EOF Then
    CboPacking.Text = Trim(rstbase!packing_no)
End If
End Sub

Private Function GetInvoiceNumber() As String
    
    Dim adoRs As New ADODB.Recordset
    
    sql = "Select Invoice_No From Invoice_Detail Where DO_No = '" & Trim(CboPacking) & "'"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        GetInvoiceNumber = adoRs.Fields("Invoice_No")
    End If
    adoRs.Close
    
End Function

Function listDO(noinvoice$) As String
Dim rsIsiDO As New ADODB.Recordset
Dim tampungDO As String
    sql = "Select distinct a.Packing_No from invoice_detail  a where a.invoice_no = '" & noinvoice & "'"
    Set rsIsiDO = Db.Execute(sql)
    
    If rsIsiDO.EOF Then
        listDO = ""
    Else
        tampungDO = ""
        Do While Not rsIsiDO.EOF
            tampungDO = tampungDO & Trim(rsIsiDO(0)) & ", "
            rsIsiDO.MoveNext
        Loop
        listDO = Left(Trim(tampungDO), Len(Trim(tampungDO)) - 1)
    End If
sql = "update invoice_master set list_do = '" & listDO & "', Last_Update = getdate(), Last_User = '" & userLogin & "' where invoice_no ='" & noinvoice & "'"
Db.Execute sql
End Function

Function listPO(noinvoice$) As String
Dim rsIsiPO As New ADODB.Recordset
Dim tampungPO As String
Dim tampungPODate As String
    sql = "Select distinct a.PO_No, b.PO_Date from invoice_detail  a inner join orderentry_master b on a.po_no = b.po_no where a.invoice_no = '" & noinvoice & "'"
    Set rsIsiPO = Db.Execute(sql)
    
    If rsIsiPO.EOF Then
        listPO = ""
        listPODate = ""
    Else
        tampungPO = ""
        tampungPODate = ""
        Do While Not rsIsiPO.EOF
            tampungPO = tampungPO & Trim(rsIsiPO(0)) & ", "
            tampungPODate = tampungPODate & Format(rsIsiPO(1), "dd-MMM-yyyy") & ", "
            rsIsiPO.MoveNext
        Loop
        listPO = Left(Trim(tampungPO), Len(Trim(tampungPO)) - 1)
        listPODate = Left(Trim(tampungPODate), Len(Trim(tampungPODate)) - 1)
    End If
sql = "update invoice_master set list_po = '" & listPO & "', list_podate = '" & listPODate & "', Last_Update = getdate(), Last_User = '" & userLogin & "' where invoice_no ='" & noinvoice & "'"
Db.Execute sql
End Function

