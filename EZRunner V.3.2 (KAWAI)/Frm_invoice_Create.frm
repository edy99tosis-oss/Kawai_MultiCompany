VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frm_invoice_Create 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Invoice Create (Export)"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15240
   Icon            =   "Frm_invoice_Create.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPEBNo 
      Height          =   315
      Left            =   8100
      MaxLength       =   25
      TabIndex        =   47
      Top             =   2760
      Width           =   2265
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   4
      Left            =   12522
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   10230
      Width           =   1320
   End
   Begin VB.TextBox Txtdisplay 
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
      Height          =   285
      Index           =   2
      Left            =   9990
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9000
      Width           =   2265
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   6
      Left            =   11238
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10230
      Width           =   1200
   End
   Begin VB.TextBox TXTNo 
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
      Left            =   120
      MaxLength       =   25
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   645
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   615
      Left            =   113
      TabIndex        =   28
      Top             =   9480
      Width           =   15015
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
         Height          =   300
         Left            =   105
         TabIndex        =   29
         Top             =   195
         Width           =   14790
      End
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   2
      Left            =   13928
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10230
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   1
      Left            =   9954
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10230
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
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
      Left            =   113
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   10230
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   3
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10230
      Width           =   1320
   End
   Begin VB.TextBox Txtdisplay 
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
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9000
      Width           =   2565
   End
   Begin VB.TextBox Txtdisplay 
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
      Height          =   285
      Index           =   3
      Left            =   12360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9000
      Width           =   2565
   End
   Begin VB.TextBox Txtdisplay 
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
      Height          =   285
      Index           =   1
      Left            =   7620
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9000
      Width           =   2265
   End
   Begin VB.TextBox txtremarks 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1403
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   7770
      Width           =   10680
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   5
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2670
      Width           =   1200
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
      ItemData        =   "Frm_invoice_Create.frx":0E42
      Left            =   113
      List            =   "Frm_invoice_Create.frx":0E4C
      Style           =   2  'Dropdown List
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1155
   End
   Begin MSComCtl2.DTPicker IDate 
      Height          =   315
      Left            =   8115
      TabIndex        =   4
      Top             =   2280
      Width           =   2235
      _ExtentX        =   3942
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
      Format          =   151650307
      CurrentDate     =   37802
   End
   Begin MSComCtl2.DTPicker DDate 
      Height          =   315
      Left            =   12030
      TabIndex        =   5
      Top             =   2280
      Width           =   1245
      _ExtentX        =   2196
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
      Format          =   151650307
      UpDown          =   -1  'True
      CurrentDate     =   37802
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4560
      Left            =   120
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3120
      Width           =   15015
      _cx             =   26485
      _cy             =   8043
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
   Begin MSMask.MaskEdBox MEDuedate 
      Height          =   315
      Left            =   12210
      TabIndex        =   9
      Top             =   8070
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
      Left            =   12210
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   8070
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
      Format          =   151650307
      CurrentDate     =   37802
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   930
      Width           =   15015
      Begin VB.TextBox Text1 
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
         Height          =   195
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   10320
         Width           =   4035
      End
      Begin MSComCtl2.DTPicker SDate 
         Height          =   315
         Left            =   1650
         TabIndex        =   1
         Top             =   705
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
         Format          =   151650307
         CurrentDate     =   37802
      End
      Begin MSComCtl2.DTPicker EDate 
         Height          =   315
         Left            =   3720
         TabIndex        =   2
         Top             =   705
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
         Format          =   151650307
         CurrentDate     =   37802
      End
      Begin MSForms.ComboBox cboDO 
         Height          =   315
         Left            =   7260
         TabIndex        =   3
         Top             =   705
         Width           =   2565
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "4524;556"
         ColumnCount     =   2
         ListRows        =   7
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN No."
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
         Left            =   6015
         TabIndex        =   45
         Top             =   765
         Width           =   600
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
         Left            =   13710
         TabIndex        =   41
         Top             =   870
         Width           =   1185
      End
      Begin VB.Label lbldesc 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
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
         Left            =   8790
         TabIndex        =   27
         Top             =   345
         Width           =   4425
      End
      Begin VB.Line Line2 
         X1              =   8790
         X2              =   13200
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3225
         X2              =   7245
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label lbldesc 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxx"
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
         Left            =   3225
         TabIndex        =   26
         Top             =   345
         Width           =   4035
      End
      Begin VB.Label Label1 
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
         Index           =   6
         Left            =   3435
         TabIndex        =   25
         Top             =   765
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN Date"
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
         Left            =   330
         TabIndex        =   24
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   7980
         TabIndex        =   23
         Top             =   345
         Width           =   690
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   330
         TabIndex        =   22
         Top             =   345
         Width           =   1170
      End
      Begin MSForms.ComboBox cbodealer 
         Height          =   315
         Left            =   1650
         TabIndex        =   0
         Top             =   285
         Width           =   1545
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2725;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13260
      TabIndex        =   46
      Top             =   330
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin MSComCtl2.DTPicker dtpPEBDate 
      Height          =   315
      Left            =   12000
      TabIndex        =   48
      Top             =   2760
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
      Format          =   151715843
      CurrentDate     =   39346
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   315
      Left            =   2610
      TabIndex        =   54
      Top             =   2280
      Width           =   2565
      VariousPropertyBits=   746604571
      MaxLength       =   25
      DisplayStyle    =   3
      Size            =   "4524;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cbocls 
      Height          =   315
      Left            =   2610
      TabIndex        =   53
      Top             =   2760
      Width           =   870
      VariousPropertyBits=   746604571
      MaxLength       =   25
      DisplayStyle    =   3
      Size            =   "1535;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3540
      X2              =   6540
      Y1              =   3030
      Y2              =   3030
   End
   Begin MSForms.TextBox txtdesc 
      Height          =   315
      Left            =   3600
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2760
      Width           =   3075
      VariousPropertyBits=   746604571
      BackColor       =   16637923
      BorderStyle     =   1
      Size            =   "5424;556"
      BorderColor     =   16637923
      SpecialEffect   =   0
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
      Left            =   1425
      TabIndex        =   51
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PEB Number"
      Height          =   195
      Index           =   4
      Left            =   6750
      TabIndex        =   50
      Top             =   2760
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PEB Date"
      Height          =   195
      Index           =   1
      Left            =   11085
      TabIndex        =   49
      Top             =   2850
      Width           =   795
   End
   Begin VB.Label Label5 
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
      Index           =   2
      Left            =   10965
      TabIndex        =   44
      Top             =   8625
      Width           =   315
   End
   Begin VB.Label Label1 
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
      Index           =   8
      Left            =   12210
      TabIndex        =   42
      Top             =   7770
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FDDFE3&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Create"
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
      Left            =   120
      TabIndex        =   40
      Top             =   240
      Width           =   15015
   End
   Begin VB.Label Label5 
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
      Index           =   0
      Left            =   1155
      TabIndex        =   38
      Top             =   8595
      Width           =   975
   End
   Begin VB.Label Label5 
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
      Index           =   3
      Left            =   13140
      TabIndex        =   37
      Top             =   8595
      Width           =   1005
   End
   Begin VB.Label Label5 
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
      Index           =   1
      Left            =   8175
      TabIndex        =   36
      Top             =   8595
      Width           =   1140
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   1
      Left            =   120
      Top             =   8880
      Width           =   15015
   End
   Begin VB.Label Label1 
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
      Index           =   7
      Left            =   360
      TabIndex        =   35
      Top             =   7950
      Width           =   765
   End
   Begin VB.Label lbl 
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
      Index           =   0
      Left            =   1425
      TabIndex        =   34
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label lbl 
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
      Index           =   1
      Left            =   6750
      TabIndex        =   33
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Month"
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
      Left            =   10590
      TabIndex        =   32
      Top             =   2340
      Width           =   1290
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
      Left            =   13223
      TabIndex        =   31
      Top             =   9630
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   120
      Top             =   8520
      Width           =   15015
   End
End
Attribute VB_Name = "Frm_invoice_Create"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Update Penambahan Field service by Dudi Desember 2008

Option Explicit
 
Dim rst As Recordset, rsmaster As Recordset, rsdetail As Recordset, rstpage As Recordset
Dim rssubtot As Recordset, rstcust As Recordset, rssumqty As Recordset, rspajak As Recordset
Dim i As Long, area As String
Dim Model As String, DelDate As Date, Dono As String
Dim currCode As String
Dim totalQty As Double, TotalAmount As Double
Dim blnupdate As Boolean, blndata As Boolean, blndisplay As Boolean
Dim bsave As Boolean, RsDel As Recordset, notfix As Boolean
Dim HakU As Integer
Public formpanggil As String, blntotal As Boolean, tempdealer As String
Public Tgl1 As String, Tgl2 As String, tppn As Double, xno As String
Dim tgl_sb As Byte, rstrade As Recordset, Overseas_Cls As String * 1
Dim BolExchange As Boolean

Public nilKosong As Boolean

Dim bteColSelect As Byte
Dim bteColSJNo As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColPONo As Byte
Dim bteColQty As Byte
Dim bteColQtyRem As Byte
Dim bteColUnit As Byte
Dim bteColSJDate As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColNoCommercial As Byte
Dim bteColRefQty As Byte
Dim bteColRefUnit As Byte
Dim bteColRefCurr As Byte
Dim bteColItemCode As Byte
Dim bteColStatus As Byte
Dim bteColService As Byte
Dim bteColDelDate As Byte
Dim bteColSeqNo As Byte
Dim bteColDOSeqNo As Byte
Dim bteHakPrice As Byte

Dim strTempPackingNo As String
Dim listPODate As String
Dim InvDate As Date

Private Sub Header()
With Grid

bteColSelect = 0
bteColSJNo = 1
bteColPartNo = 2
bteColDesc = 3
bteColPONo = 4
bteColQty = 5
bteColQtyRem = 6
bteColUnit = 7
bteColSJDate = 8
bteColCurr = 9
bteColPrice = 10
bteColService = 11
bteColAmount = 12
bteColNoCommercial = 13
bteColRefQty = 14
bteColRefUnit = 15
bteColRefCurr = 16
bteColItemCode = 17
bteColStatus = 18
bteColDelDate = 19
bteColSeqNo = 20
bteColDOSeqNo = 21

.Rows = 1
.ColS = 22

.TextMatrix(0, bteColSelect) = ""
.TextMatrix(0, bteColSJNo) = "DN No. (Ref No.)"
.TextMatrix(0, bteColPartNo) = "Part Number"
.TextMatrix(0, bteColDesc) = "Description"
.TextMatrix(0, bteColPONo) = "SI/PO No."
.TextMatrix(0, bteColQty) = "Qty"
.TextMatrix(0, bteColQtyRem) = "Qty Rem"
.TextMatrix(0, bteColUnit) = "Unit"
.TextMatrix(0, bteColSJDate) = "DN Date"
.TextMatrix(0, bteColCurr) = "Curr"
.TextMatrix(0, bteColPrice) = "Price"
.TextMatrix(0, bteColService) = "Service"
.TextMatrix(0, bteColAmount) = "Amount"
.TextMatrix(0, bteColNoCommercial) = "N/C"
.TextMatrix(0, bteColRefQty) = "Qty"
.TextMatrix(0, bteColRefUnit) = "unit"
.TextMatrix(0, bteColRefCurr) = "curr"
.TextMatrix(0, bteColItemCode) = "item_code"
.TextMatrix(0, bteColStatus) = "Status"
.TextMatrix(0, bteColDelDate) = "delivery date"
.TextMatrix(0, bteColSeqNo) = "Seq No"
.TextMatrix(0, bteColDOSeqNo) = "DOSeq No"

.ColWidth(bteColSelect) = 250
.ColWidth(bteColSJNo) = 2000
.ColWidth(bteColPartNo) = 2000
.ColWidth(bteColDesc) = 2500
.ColWidth(bteColPONo) = 1500
.ColWidth(bteColQty) = 1000
.ColWidth(bteColQtyRem) = 1000
.ColWidth(bteColUnit) = 600
.ColWidth(bteColSJDate) = 1300
.ColWidth(bteColCurr) = 600
.ColWidth(bteColPrice) = 1200
.ColWidth(bteColService) = 1200
.ColWidth(bteColAmount) = 1900
.ColWidth(bteColNoCommercial) = 600

.ColHidden(bteColRefQty) = True
.ColHidden(bteColRefUnit) = True
.ColHidden(bteColRefCurr) = True
.ColHidden(bteColItemCode) = True
.ColHidden(bteColStatus) = True
.ColHidden(bteColDelDate) = True
.ColHidden(bteColSeqNo) = True
.ColHidden(bteColDOSeqNo) = True
.ColHidden(bteColNoCommercial) = True
If gb_InvoiceReferToDO_InvoiceCreate And Not gb_AllowMultipleDO_InvoiceCreate Then
    .ColHidden(bteColSelect) = True
End If

.ColHidden(bteColCurr) = (bteHakPrice = 0)
.ColHidden(bteColPrice) = (bteHakPrice = 0)
.ColHidden(bteColService) = (bteHakPrice = 0)
.ColHidden(bteColAmount) = (bteHakPrice = 0)

.Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
.Cell(flexcpBackColor, 0, 0, 0, .ColS - 1) = &HA6D2FF
End With
End Sub

Private Sub cbocls_Change()
cboCls = Trim(cboCls)
If cboCls.MatchFound Then
    TxtDesc.Text = cboCls.Column(1)
Else
    TxtDesc.Text = ""
End If
End Sub

Private Sub cbodealer_Click()
    
    If nilKosong Then Exit Sub
    MousePointer = vbHourglass
    rstcust.Requery
    rstcust.Find "Cust_code ='" & cbodealer & "'"
    If Not rstcust.EOF Then
        LblDesc(0).Caption = rstcust!Cust_Name
        LblDesc(1).Caption = rstcust!Address
        Overseas_Cls = rstcust!country_cls
    rstcust.Requery
    End If
    'Update BY DUDI mengecek apakah ada no DO, untuk mempercepat proses
    Dim s As String
    s = "Select a.DO_No, DO_Date, (Select Max (Delivery_Date)From Delivery_Order Where DO_No = a.DO_No) Delivery_Date From DO_Master a " & _
            "Where a.Cust_Code = '" & Trim(cbodealer.Text) & "' " & _
            "And a.DO_Date >= '" & Format(SDate.Value, "YYYY-MM-DD") & "' " & _
            "And a.DO_Date <= '" & Format(EDate.Value, "YYYY-MM-DD") & "' "
    If CekSql(s) Then
        isiCboDO
        Else
        cboDO.clear
        
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
        Call createnumber
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
        'Update DUDI Januari 2009,mempercepat proses
        'mengambil no invoice jika telah ada
        s = "select * from invoice_master where cust_code = '" & cbodealer.Text & "' and year(invoice_date) >= '" & Year(SDate) & "' and year(invoice_date) <= '" & Year(EDate) & "'"
        If CekSql(s) Then
            nomorinvoice
        Else
            lblFix = ""
            Grid.Rows = 1
            ComboBox1.clear
            ComboBox1.locked = True
            
        End If
        clear2
        
    End If
    
    MousePointer = vbDefault
    Grid.Editable = False

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
End Sub

Private Sub CboDo_Change()
    
    Dim strInvoice As String
    
    If cboDO.MatchFound And gb_InvoiceReferToDO_InvoiceCreate And Not gb_AllowMultipleDO_InvoiceCreate Then
        If Trim(ComboBox1) <> "" Then
            sql = "select * from Invoice_detail where invoice_no = '" & ComboBox1 & "'"
            Set RsDel = New Recordset
            RsDel.Open sql, Db, adOpenKeyset, adLockOptimistic
            If RsDel.EOF Then
                Db.Execute ("Delete from  invoice_master where invoice_no = '" & ComboBox1 & "'")
            End If
            Set RsDel = Nothing
        End If
        nomorinvoice
'        strInvoice = GetInvoiceNumber
'        If strInvoice = "" Then
'            DDate = cboDO.Column(2)
'            isiDueDate cboDO.Column(2)
'            If Combo1.ListIndex = 0 Then
'                ComboBox1 = cboDO
'                IDate = cboDO.Column(1)
'            Else
'                Combo1.ListIndex = 0
'            End If
'        Else
            ComboBox1 = cboDO
'        End If
    End If
    
End Sub

Private Sub cmdAction_Click(Index As Integer)

lblerror.Caption = ""
Dim strDONo As String, ls_sql As String
Select Case Index
    Case 0
           If formpanggil = "invoiceinquiry" And cmdAction(0).Caption = "Back" Then
                
                frm_invoice_inquiry.ComboBox1 = Me.cbodealer
                frm_invoice_inquiry.set_tgl Tgl1, Tgl2
                frm_invoice_inquiry.combo1 = ComboBox1
                
                Unload Me
                frm_invoice_inquiry.Show
                frm_invoice_inquiry.set_dari_inv_create
                frm_invoice_inquiry.xxx
                
            Else
                If Trim(ComboBox1) <> "" Then
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
            If HakU = 0 Then _
                lblerror = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
            
            savedetail
            listPO (ComboBox1)
            listDO (ComboBox1)
            
            If Not BolExchange Then Me.MousePointer = vbDefault: Exit Sub
            
            tppn = tax("ppn")
            updateMaster
            inquiryupdate
            MousePointer = vbDefault
            lblerror = DisplayMsg(1101)
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
                        Call InvReport
                    End If
                Else
                    lblerror = DisplayMsg(4071)
                End If
            End If
            MousePointer = vbDefault
    Case 4
            
            lblerror = up_ValidateDateRange(InvDate, True)
            If lblerror <> "" Then Me.MousePointer = vbDefault: Exit Sub
            
            If (MsgBox("Are you sure want to delete this invoice ?", vbQuestion + vbYesNo, "Confirmation") = vbYes) Then
                Dim dbdel As New Connection
                dbdel.ConnectionString = Db.ConnectionString
                dbdel.Open
                dbdel.BeginTrans
                
                '#Delete Invoice
                sql = "delete from invoice_detail where invoice_no ='" & ComboBox1 & "'"
                dbdel.Execute sql
                sql = "delete from invoice_master where invoice_no ='" & ComboBox1 & "'"
                dbdel.Execute sql
                If err.number = 0 Then
                    dbdel.CommitTrans
                    clear
                    cboDO = ""
                    lblerror = DisplayMsg(1201)
                Else
                    dbdel.RollbackTrans
                    lblerror = err.Description
                End If
                dbdel.Close
                Set dbdel = Nothing
            End If
    Case 5
            MousePointer = vbHourglass
            If HakU = 0 Then _
                lblerror = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
            
            If cmdAction(5).Caption = "Create" Then
                
                lblerror = up_ValidateDateRange(IDate, True)
                If lblerror <> "" Then Me.MousePointer = vbDefault: Exit Sub
                
                cbodealer = cbodealer
                If cbodealer.MatchFound = False Then
                    lblerror = DisplayMsg(4072)
                    cbodealer.SetFocus
                    MousePointer = vbDefault
                    Exit Sub
                End If
                If Not gb_AllowMultipleDO_InvoiceCreate And (cboDO = "" Or Not cboDO.MatchFound) Then
                    lblerror = DisplayMsg("4101")
                    cboDO.SetFocus
                    MousePointer = vbDefault
                    Exit Sub
                End If
                If Trim(ComboBox1) <> "" Then
                    savemaster
                    inquiry
                    inquiryupdate
                    blndisplay = True
                    cmdAction(5).Caption = "Update"
                    bsave = True
                    combo1 = "Update"
                    Grid.Editable = True
                    cbodealer.locked = False
                    ComboBox1.locked = False
                    xno = ComboBox1
                    nomorinvoice
                    ComboBox1 = xno
                    lblerror.Caption = DisplayMsg(1000)
                    cmdAction(2).Enabled = True
                    cmdAction(4).Enabled = True 'BARU
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
                If Trim(ComboBox1) <> "" Then
                    ComboBox1 = ComboBox1
                    If ComboBox1.MatchFound Then
                        tppn = tax("ppn")
                        If notfix Then updateMaster
                        inquiry
                        If notfix Then listDO (ComboBox1)
                        If notfix Then listPO (ComboBox1)
                        inquiryupdate
                        Grid.Editable = True
                        If notfix Then lblerror.Caption = DisplayMsg(1101)
                        blndisplay = False
                    Else
                        lblerror = DisplayMsg(4071)
                    End If
                Else
                    lblerror = ""
                End If
            End If
            If lblFix <> "" Then
                Grid.Editable = False
            Else
                Grid.Editable = True
            End If
            MousePointer = vbDefault
    Case 6
        If Trim(ComboBox1) <> "" And combo1.Text = "Update" And Grid.Rows <> 1 Then
            ComboBox1 = ComboBox1
            If ComboBox1.MatchFound Then
                inquiryupdate
                lblerror = ""
            Else
                lblerror = DisplayMsg(4071)
            End If
        End If
End Select
Exit Sub


End Sub

Private Sub Combo1_Click()
Dim strInvoice As String
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
        If gb_InvoiceReferToDO_InvoiceCreate And Not gb_AllowMultipleDO_InvoiceCreate Then
            strInvoice = GetInvoiceNumber
            If strInvoice = "" Then
                ComboBox1 = cboDO
            Else
                ComboBox1 = strInvoice
            End If
        Else
            'Call createnumber
            If gb_AllowMultipleDO_InvoiceCreate Then cboDO.Text = ""
        End If
        Header
    End If
    bsave = False
    blntotal = False
    clear2
    txtRemarks = ""
    
    cmdAction(5).Caption = "Create"
    cmdAction(2).Enabled = False
    cmdAction(4).Enabled = False '' baru
    cbodealer.locked = False
    'ComboBox1.locked = True
    Grid.Editable = False
    ComboBox1.locked = False
    ComboBox1 = ""
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

Private Sub ComboBox1_Change()
Dim strDONo As String
lblerror.Caption = ""
cmdAction(5).Enabled = True
If nilKosong Then Exit Sub
DoEvents
sql = "select * from invoice_master where invoice_no ='" & ComboBox1 & "'"
Set rst = New Recordset
rst.Open sql, Db, adOpenDynamic, adLockOptimistic
If Not rst.EOF Then


    If Trim(combo1.Text) = "Create" Then
    
        lblerror.Caption = "Invoice No" & " " & Trim(ComboBox1.Text) & " Has been Create"
        cmdAction(5).Enabled = False
        Exit Sub
    
    Else
        cmdAction(5).Enabled = True
 
          ComboBox1 = Trim(rst!Invoice_No)
          If Not gb_AllowMultipleDO_InvoiceCreate Then
              strDONo = GetDONumber
              If strDONo <> "" Then cboDO = GetDONumber
          End If
          InvDate = rst!Invoice_Date
          IDate = rst!Invoice_Date
          DDate = Month(rst!delivery_Date) & "/" & Year(rst!delivery_Date)
          Txtdisplay(0) = Trim(rst!Invoice_No)
          Txtdisplay(1) = Format(CDbl(rst!Amount), gs_formatAmountIDR)
          If Overseas_Cls = "1" Then
              Txtdisplay(2) = Format(0, gs_formatAmountIDR)
              Txtdisplay(3) = Txtdisplay(1)
          Else
              Txtdisplay(2) = Format(((CDbl(Txtdisplay(1)) * tax("Ppn")) / 100), gs_formatAmountIDR)
              Txtdisplay(3) = Format(CDbl(Txtdisplay(1)) + CDbl(Txtdisplay(2)), gs_formatAmountIDR)
          End If
        
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
          cboCls = Trim(rst!tradeterms_cls)
          txtPEBNo = Trim(rst!PEBNo & "")
          If Not IsNull(rst!PEBDate) Then dtpPEBDate = rst!PEBDate Else dtpPEBDate = Null
          
          If IsNull(rst!fix_cls) Then
              lblFix = ""
              notfix = True
              cmdAction(2).Enabled = True
              cmdAction(4).Enabled = True 'baru
          Else
              lblFix = "Status : Fix"
              notfix = False
              cmdAction(2).Enabled = False
              cmdAction(4).Enabled = False 'baru
          End If
          If Trim(ComboBox1) <> "" Then
              If blndisplay = False Or xno <> ComboBox1 Then Grid.Rows = 1
          End If
          lblerror = ""
          
     End If
Else
        lblFix = ""
        notfix = True
        If Trim(ComboBox1) <> "" Then
            If blndisplay = False Or xno <> ComboBox1 Then Grid.Rows = 1
        End If
        lblerror = ""
        

    
End If

End Sub


Private Sub Combobox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
rstcust.Requery
rstcust.Find "cust_code = '" & cbodealer.Text & "'"
If Not rstcust.EOF Then
    rstcust.Requery
    isiCboDO
    If combo1.Text = "Create" Then
        clear2
        Grid.Editable = False
    Else
        If Trim(ComboBox1) = "" Then nomorinvoice: Exit Sub
        xno = ComboBox1
         'Update DUDI Januari 2009,mempercepat proses
        'mengambil no invoice jika telah ada
        Dim s As String
        s = "select * from invoice_master where cust_code = '" & cbodealer.Text & "' and year(invoice_date) >= '" & Year(SDate) & "' and year(invoice_date) <= '" & Year(EDate) & "'"
        If CekSql(s) Then
            nomorinvoice
        Else
            lblFix = ""
            Grid.Rows = 1
            ComboBox1.clear
            ComboBox1.locked = True
            
        End If
        
        ComboBox1 = xno
    End If
End If
End Sub

Private Sub edate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
bteHakPrice = hakPrice(Me.Name)
Label5(1).Visible = (bteHakPrice = 1)
Label5(2).Visible = (bteHakPrice = 1)
Label5(3).Visible = (bteHakPrice = 1)
Txtdisplay(1).Visible = (bteHakPrice = 1)
Txtdisplay(2).Visible = (bteHakPrice = 1)
Txtdisplay(3).Visible = (bteHakPrice = 1)
adtocombo
Header
IDate = Date
HakU = hakUpdate(Me.Name)

SDate = Format(Date, "dd mmm yyyy")
EDate = Format(Date, "dd mmm yyyy")
IDate = Format(Date, "dd mmm yyyy")
DDate = Format(Date, "dd mmm yyyy")
cmdAction(2).Enabled = False
cmdAction(4).Enabled = False 'baru

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
sql = "SELECT  rtrim(Trade_Master.trade_Code) cust_code, rtrim(Trade_Master.Trade_Name) cust_name, " & _
    "rtrim(Trade_Master.Address1) address, isnull(country_Cls,'0') country_Cls " & _
    " From Trade_Master where trade_cls in ('2') --and isnull(country_Cls,'0') = 0"
    
Set rstcust = New Recordset
rstcust.Open sql, Db, adOpenKeyset, adLockOptimistic
With cbodealer
    .clear
    .columnCount = 2
    .ColumnWidths = "60 pt;280 pt; 0 pt"
    .ListWidth = 350
    .ListRows = 15
    i = 0
    Do Until rstcust.EOF
        .AddItem ""
        .List(i, 0) = Trim(rstcust!Cust_CodE)
        .List(i, 1) = Trim(rstcust!Cust_Name)
        .List(i, 2) = IIf(IsNull(Trim(rstcust!Address)), "", Trim(rstcust!Address))
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
With Grid
    
    .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
    .Cell(flexcpBackColor, i, bteColQty) = vbWhite
    .TextMatrix(i, bteColSJNo) = Trim(RS.Fields("DO_no").Value)
    If IsNull(RS.Fields("Part_no").Value) Then
        .TextMatrix(i, bteColPartNo) = ""
    Else
        .TextMatrix(i, bteColPartNo) = Trim(RS.Fields("Part_no").Value)
    End If
    .TextMatrix(i, bteColDesc) = Trim(RS.Fields("Item_name").Value)
    
    '**** lihat detail untuk qty
    sql = "select *,isnull(service,0) as services from invoice_detail where invoice_no = '" & ComboBox1 & "' and Item_code= '" & Model & "' and delivery_date = '" & Format(DDate, "yyyy-mm-dd") & "' and do_no = '" & RS!do_no & "' and po_no = '" & RS!po_no & "' and seq_no = '" & RS!Seq_no & "' and doseq_no = '" & RS!DOSeq_No & "' "
    Set rsdetail = New Recordset
    rsdetail.Open sql, Db, adOpenDynamic, adLockOptimistic
    If Not rsdetail.EOF Then
        If combo1.Text = "Create" Then
            If gb_InvoiceReferToDO_InvoiceCreate And Not gb_AllowMultipleDO_InvoiceCreate Then
                .Cell(flexcpChecked, i, bteColSelect) = flexChecked
            Else
                .Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
            End If
            .TextMatrix(i, bteColPONo) = Trim(RS!po_no)
            .TextMatrix(i, bteColQty) = Format(RS.Fields("qtyresult"), gs_formatQty)
            .TextMatrix(i, bteColQtyRem) = Format(0, gs_formatQty)
            .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(Trim(RS!Unit_cls))
            .TextMatrix(i, bteColSJDate) = Format(RS.Fields("do_date").Value, "dd Mmm YYYY")
            If Trim(RS!currency_code) <> "" Then .TextMatrix(i, bteColCurr) = uf_GetCurrencyDescription(Trim(RS!currency_code))
            If Trim(RS!currency_code) = "03" Then
                .TextMatrix(i, bteColPrice) = Format(RS.Fields("Price").Value, gs_formatPriceIDR)
                .TextMatrix(i, bteColService) = Format(RS.Fields("service").Value, gs_formatPriceIDR)
            Else
                .TextMatrix(i, bteColPrice) = Format(RS.Fields("Price").Value, gs_formatPrice)
                .TextMatrix(i, bteColService) = Format(RS.Fields("Service").Value, gs_formatPrice)
            End If
            .TextMatrix(i, bteColAmount) = Format(RS.Fields("Amount").Value, gs_formatAmountIDR)
            .TextMatrix(i, bteColRefQty) = Format(RS.Fields("Qtyresult").Value, gs_formatQty)
        Else
            .Cell(flexcpChecked, i, bteColSelect) = flexChecked
            .TextMatrix(i, bteColPONo) = Trim(RS!po_no)
            .TextMatrix(i, bteColQty) = Format(Trim(RS.Fields("IQTy").Value), gs_formatQty)
            
            sql = "select DO_No,item_code,sum(qty ) totqty ,delivery_date,PO_NO,seq_no,isnull(service,0)as services  from invoice_detail where do_no = '" & RS!do_no & "' and item_code = '" & Model & "' and delivery_date ='" & Format(DDate, "YYYY-MM-DD") & "' and po_no = '" & RS!po_no & "' and seq_no =  '" & RS!Seq_no & "' and doseq_no = '" & RS!DOSeq_No & "' "
            sql = sql & " group by do_no,item_code,delivery_date,Po_no,seq_no,Service "
            Set rssumqty = New Recordset
            rssumqty.Open sql, Db, adOpenDynamic, adLockOptimistic
            
            If Not rssumqty.EOF Then
                   .TextMatrix(i, bteColQtyRem) = Format(RS!DQty - rssumqty.Fields("totqty").Value, gs_formatQty)
                    If Val(.TextMatrix(i, bteColQtyRem)) <> 0 Then
                        .TextMatrix(i, bteColRefQty) = RS!IQTY + .TextMatrix(i, bteColQtyRem) 'rs!iqty + rssumqty.Fields("totqty").Value
                    Else
                        .TextMatrix(i, bteColRefQty) = RS.Fields("iqty").Value
                    End If
            Else
               .TextMatrix(i, bteColRefQty) = Format(RS.Fields("Qtyresult").Value, gs_formatQty)
               .TextMatrix(i, bteColQtyRem) = Format(Trim(RS.Fields("Qtyresult").Value), gs_formatQty)
            End If
            
            .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(Trim(RS!Unit_cls))
            .TextMatrix(i, bteColSJDate) = Format(RS.Fields("do_date").Value, "dd Mmm YYYY")
            If Trim(RS!currency_code) <> "" Then .TextMatrix(i, bteColCurr) = uf_GetCurrencyDescription(Trim(RS!currency_code))
            If Trim(RS!currency_code) = "03" Then
                .TextMatrix(i, bteColPrice) = Format(RS.Fields("Price").Value, gs_formatPriceIDR)
                .TextMatrix(i, bteColService) = Format(RS.Fields("Service").Value, gs_formatPriceIDR)
            Else
                .TextMatrix(i, bteColPrice) = Format(RS.Fields("Price").Value, gs_formatPrice)
                .TextMatrix(i, bteColService) = Format(RS.Fields("Service").Value, gs_formatPriceIDR)
            End If
            .TextMatrix(i, bteColAmount) = Format(RS.Fields("Amount").Value, gs_formatAmountIDR)
            currCode = Trim(RS!currency_code)
       End If
    Else
        
        If gb_InvoiceReferToDO_InvoiceCreate And Not gb_AllowMultipleDO_InvoiceCreate Then
            .Cell(flexcpChecked, i, bteColSelect) = flexChecked
        Else
                .Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
        End If
        .TextMatrix(i, bteColPONo) = Trim(RS!po_no)
        .TextMatrix(i, bteColQty) = Format(RS.Fields("Qtyresult").Value, gs_formatQty)
        .TextMatrix(i, bteColQtyRem) = Format(0, gs_formatQty)
        .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(Trim(RS!Unit_cls))
        .TextMatrix(i, bteColSJDate) = Format(RS.Fields("do_date").Value, "dd Mmm YYYY")
        If Trim(RS!currency_code) <> "" Then .TextMatrix(i, bteColCurr) = uf_GetCurrencyDescription(Trim(RS!currency_code))
        If Trim(RS!currency_code) = "03" Then
            .TextMatrix(i, bteColPrice) = Format(RS.Fields("Price").Value, gs_formatPriceIDR)
            .TextMatrix(i, bteColService) = Format(RS.Fields("Service").Value, gs_formatPriceIDR)
        Else
            .TextMatrix(i, bteColPrice) = Format(RS.Fields("Price").Value, gs_formatPrice)
            .TextMatrix(i, bteColService) = Format(RS.Fields("service").Value, gs_formatPrice)
        End If
        .TextMatrix(i, bteColAmount) = Format(RS.Fields("Amount").Value, gs_formatAmountIDR)
        .TextMatrix(i, bteColRefQty) = Format(RS.Fields("Qtyresult").Value, gs_formatQty)
    End If
    .TextMatrix(i, bteColRefUnit) = RS.Fields("unit_Cls").Value
    If Trim(RS!currency_code) <> "" Then .TextMatrix(i, bteColRefCurr) = RS.Fields("currency_code").Value
    .TextMatrix(i, bteColItemCode) = RS.Fields("item_code").Value
    If gb_InvoiceReferToDO_InvoiceCreate And Not gb_AllowMultipleDO_InvoiceCreate Then
        .TextMatrix(i, bteColStatus) = "1"
    Else
        .TextMatrix(i, bteColStatus) = "0"
    End If
    .TextMatrix(i, bteColDelDate) = Format(RS!delivery_Date, "dd mmm yyyy")
    .TextMatrix(i, bteColSeqNo) = RS!Seq_no
    .TextMatrix(i, bteColDOSeqNo) = RS!DOSeq_No
    
'    If (IsNull(rs!NoCommercial_Cls)) = False Then
'     If rs!NoCommercial_Cls = 0 Then
'      .TextMatrix(i, bteColNoCommercial) = "No"
'     Else
'      .TextMatrix(i, bteColNoCommercial) = "Yes"
'     End If
'    Else
'     .TextMatrix(i, bteColNoCommercial) = ""
'    End If
    
    .Cell(flexcpAlignment, i, 1, i, bteColPONo) = flexAlignLeftCenter
    .Cell(flexcpAlignment, i, bteColQty, i, bteColQtyRem) = flexAlignRightCenter
    .Cell(flexcpBackColor, i, 1, i, bteColPONo) = &H80000018
    .Cell(flexcpBackColor, i, bteColQtyRem, i, bteColAmount) = &H80000018

    totalQty = totalQty + .TextMatrix(i, bteColQty)
    TotalAmount = TotalAmount + .TextMatrix(i, bteColAmount)
    
End With
End Sub

Sub subtotal(Row As Long)
With Grid
.Cell(flexcpBackColor, Row, bteColSelect) = vbWhite
.Cell(flexcpBackColor, Row, bteColSJNo, Row, bteColNoCommercial) = &HE0E0E0
.Cell(flexcpChecked, Row, bteColSelect) = ""
.TextMatrix(Row, bteColSJNo) = ""
.TextMatrix(Row, bteColPartNo) = ""
.TextMatrix(Row, bteColDesc) = ""
.TextMatrix(Row, bteColQty) = Format(totalQty, gs_formatQty)
.TextMatrix(Row, bteColQtyRem) = ""
.TextMatrix(Row, bteColUnit) = ""
.TextMatrix(Row, bteColSJDate) = ""
.TextMatrix(Row, bteColAmount) = Format(TotalAmount, gs_formatAmountIDR)
.TextMatrix(Row, bteColRefQty) = "subtotal" & Trim(.TextMatrix(Row - 1, bteColItemCode)) & CDbl(.TextMatrix(Row - 1, bteColPrice))
.TextMatrix(Row, bteColItemCode) = "0"
.Cell(flexcpAlignment, Row, bteColQty) = flexAlignRightCenter
.Cell(flexcpAlignment, Row, bteColAmount) = flexAlignRightCenter
End With
totalQty = 0
TotalAmount = 0

End Sub


Sub inquiryupdate()
Dim rupdate As Recordset
Dim Price As String

sql = "Select a.*, om.nocommercial_cls " & _
      "From ( "

sql = sql & _
      "Select Qtyresult, DO_NO, cust_code, Group_Cls, Item_Code, Part_no, Item_name, Po_no, unit_cls, delivery_date, Do_date, currency_code, Price,isnull(sERVICE,0)as service, Amount, AmountKu, Dqty, IQTY, Seq_No, doseq_no " & _
      "From ("

sql = sql & _
        "Select do.Qty - Sum(Isnull(ivd.Qty, 0)) As Qtyresult, do.DO_No, dm.Cust_Code, Group_Cls, do.Item_Code, " & _
        "RTrim(do.makerItem_code) Part_No,im.Item_Name, do.po_no,do.unit_cls,do.Delivery_Date,dm.do_date, do.currency_code,do.Price As Price,   do.Service as service,Sum(do.Amount) As Amount, " & _
        "Abs(do.Qty - Sum(Isnull(ivd.Qty, 0))) * do.Price As amountku, do.Qty As DQty, " & _
        "Sum(IsNull(ivd.Qty, 0)) As IQTY,do.seq_no, do.doseq_no " & _
        "From Delivery_Order do " & _
        "Inner Join Item_Master im On do.Item_Code = im.Item_Code " & _
        "Inner Join DO_Master dm On do.DO_No = dm.DO_No " & _
        "Left Outer Join Invoice_Detail ivd On do.DO_No = ivd.DO_No And do.po_no = ivd.po_no And do.seq_no = ivd.seq_no And do.doseq_no = ivd.doseq_no " & _
        "Inner Join OrderEntry_Detail od On od.PO_No = do.PO_No and od.seq_no = do.seq_no " & _
        "Inner Join OrderEntry_Master om On om.PO_No = od.PO_No and om.Cust_Code = od.cust_code " & _
        "Group By do.Qty, do.DO_No, dm.Cust_Code, Group_Cls, do.Item_Code,do.makerItem_code, do.Delivery_Date, do.currency_code,do.Price, " & _
          "do.Amount, im.Item_Name ,do.po_no ,do.unit_cls,dm.fix_cls,dm.do_date, do.seq_no, do.doseq_no,do.service " & _
          "Having  (dm.Do_Date>= '" & Format(SDate, "YYYY-MM-DD") & "') And (dm.Do_Date <= '" & Format(EDate, "YYYY-MM-DD") & "') And (dm.Cust_Code = '" & Trim(cbodealer.Text) & "') And do.Qty - Sum(Isnull(ivd.Qty, 0))> 0 " & _
          "And (RTrim(do.DO_NO) + RTrim(dm.Cust_code) + RTrim(do.po_no) + RTrim(do.Seq_no) + RTrim(do.doSeq_no)) " & _
          "Not In ( " & _
            "Select RTrim(do.DO_NO) + RTrim(dm.Cust_code) + RTrim(do.po_no) + RTrim(do.Seq_no) + RTrim(do.DoSeq_no) As DoSeq_no " & _
            "From Delivery_Order do " & _
            "Inner Join Item_Master im On do.Item_Code = im.Item_Code " & _
            "Inner Join DO_Master dm On do.DO_No = dm.DO_No " & _
            "Left Outer Join Invoice_Detail ivd On do.DO_No = ivd.DO_No And do.po_no = ivd.po_no And do.seq_no = ivd.seq_no And do.doseq_no = ivd.doseq_no " & _
            "Where (ivd.Invoice_No = '" & Trim(ComboBox1) & "') " & _
            "Group By do.Qty, do.DO_No, dm.Cust_Code, Group_Cls, do.Item_Code, do.makerItem_code,im.Item_Name,do.po_no, " & _
            "do.Delivery_Date,do.po_no,do.Seq_no, do.doSeq_no) "

If cboDO <> strAll Then sql = sql & "And do.DO_No = '" & Trim(cboDO) & "' "
    
sql = sql & _
        "Union " & _
        "Select do.Qty - Sum(Isnull(ivd.Qty, 0)) As Qtyresult, do.DO_No, dm.Cust_Code, Group_Cls, do.Item_Code, " & _
        "RTrim(do.makerItem_code) Part_No,im.Item_Name, do.po_no,do.unit_cls,do.Delivery_Date,dm.do_date,do.currency_code,do.Price As Price,isnull(DO.SERVICE,0)as service " & _
        ",Sum(do.Amount) As Amount, " & _
        "Abs(do.Qty - Sum(Isnull(ivd.Qty, 0))) * do.Price As AmountKu, do.Qty As DQty, " & _
        "Sum(IsNull(ivd.Qty, 0)) As IQTY,do.seq_no, do.DoSeq_no " & _
        "From Delivery_Order do " & _
        "Inner Join Item_Master im On do.Item_Code = im.Item_Code " & _
        "Inner Join DO_Master dm On do.DO_No = dm.DO_No " & _
        "Left Outer Join Invoice_Detail ivd On do.DO_No = ivd.DO_No And do.po_no = ivd.po_no And do.seq_no = ivd.seq_no And do.doseq_no = ivd.doseq_no " & _
        "Inner Join OrderEntry_Detail od On od.PO_No = do.PO_No and od.seq_no = do.seq_no " & _
        "Inner Join OrderEntry_Master om On om.PO_No = od.PO_No and om.Cust_Code = od.cust_code " & _
        "Where (ivd.Invoice_No = '" & Trim(ComboBox1) & "') " & _
        "Group By do.Qty, do.DO_No, dm.Cust_Code, Group_Cls, do.Item_Code, do.makerItem_code,do.Delivery_Date, do.currency_code,do.Price,DO.SERVICE ," & _
        "do.Amount, im.Item_Name, do.po_no,do.unit_cls,dm.do_date,do.seq_no, do.doSeq_no "

sql = sql & _
      ") SQL_Union "
    
sql = sql & _
      ")a " & _
      "Inner Join OrderEntry_Detail od On od.PO_No = a.PO_No and od.seq_no = a.seq_no " & _
      "Inner Join OrderEntry_Master om On om.PO_No = od.PO_No and om.Cust_Code = od.cust_code "
    
sql = sql & _
      "Order By om.nocommercial_cls, a.group_cls, a.Part_no"
    
Set rupdate = New Recordset
rupdate.CursorLocation = adUseClient
rupdate.Open sql, Db, adOpenKeyset, adLockOptimistic

If rupdate.EOF = False Then
With rupdate
    totalQty = 0
    TotalAmount = 0
    Grid.Rows = 1
    currCode = ""
    For i = 1 To .RecordCount
        If Not .EOF Then 'selama ada model yang sama
            If i = 1 Then
                Model = !Item_Code
                DelDate = !delivery_Date
                Price = !Price
                Dono = !do_no
                Grid.Rows = Grid.Rows + 1
                Call DisplayData(rupdate, Grid.Rows - 1, Model, DelDate)
            Else
                If !Item_Code = Model Then 'jika code shop sekarang sama dgn code shop pada hal sebelumnya
                    If Price = !Price Then
                        DelDate = !delivery_Date
                        Grid.Rows = Grid.Rows + 1
                        Call DisplayData(rupdate, Grid.Rows - 1, Model, DelDate)
                    Else
                        Grid.Rows = Grid.Rows + 1
                        Call subtotal(Grid.Rows - 1)
                        Model = !Item_Code
                        DelDate = !delivery_Date
                        Price = !Price
                        Dono = !do_no
                        Grid.Rows = Grid.Rows + 1
                        Call DisplayData(rupdate, Grid.Rows - 1, Model, DelDate)
                    End If
                Else
                    Grid.Rows = Grid.Rows + 1
                    Call subtotal(Grid.Rows - 1)
                    Model = !Item_Code
                    DelDate = !delivery_Date
                    Dono = !do_no
                    Price = !Price
                    Grid.Rows = Grid.Rows + 1
                    Call DisplayData(rupdate, Grid.Rows - 1, Model, DelDate)
                End If
            End If
            
        Else 'jika tdk ada data sesudahnya maka total

            Exit For
        End If
        .MoveNext
    Next i
    Grid.Rows = Grid.Rows + 1
    Call subtotal(Grid.Rows - 1)
End With
Else

Header
End If
End Sub


Sub inquiry()
Dim strDONo As String
sql = "select * from invoice_master where invoice_no ='" & ComboBox1 & "'"
Set rst = New Recordset
rst.Open sql, Db, adOpenDynamic, adLockOptimistic
If Not rst.EOF Then
    ComboBox1 = Trim(rst!Invoice_No)
    If Not gb_AllowMultipleDO_InvoiceCreate Then
        strDONo = GetDONumber
        If strDONo <> "" Then cboDO = strDONo
    End If
    InvDate = rst!Invoice_Date
    IDate = rst!Invoice_Date
    'DDate = Right(rst!delivery_Date, 2) & "/" & Left(rst!delivery_Date, 4)
    DDate = Format(rst!delivery_Date, "MM") & "/" & Format(rst!delivery_Date, "yyyy")
    
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
    
    cbodealer = Trim(rst!Cust_CodE)
    Txtdisplay(0) = Trim(rst!Invoice_No)
    If InStr(1, rst!Amount, ".") > 0 Then
        Txtdisplay(1) = Format(CDbl(rst!Amount), gs_formatAmountIDR)
    Else
        Txtdisplay(1) = Format(CDbl(rst!Amount), gs_formatAmountIDR)
    End If

    If Overseas_Cls = "1" Then
        Txtdisplay(2) = Format(0, gs_formatAmountIDR)
        Txtdisplay(3) = Txtdisplay(1)
    Else
        Txtdisplay(2) = Format(((CDbl(Txtdisplay(1)) * tax("Ppn")) / 100), gs_formatAmountIDR)
        Txtdisplay(3) = Format(CDbl(Txtdisplay(1)) + CDbl(Txtdisplay(2)), gs_formatAmountIDR)
    End If
    
    If InStr(1, (Txtdisplay(1) * tax("Ppn")) / 100, ".") > 0 Then
        Txtdisplay(2) = Format(((CDbl(Txtdisplay(1)) * tax("Ppn")) / 100), gs_formatAmountIDR)
    Else
        Txtdisplay(2) = Format(((CDbl(Txtdisplay(1)) * tax("Ppn")) / 100), gs_formatAmountIDR)
    End If
    
    If InStr(Txtdisplay(1) + Txtdisplay(2), ".") > 0 Then
        Txtdisplay(3) = Format(CDbl(Txtdisplay(1)) + CDbl(Txtdisplay(2)), gs_formatAmountIDR)
    Else
        Txtdisplay(3) = Format(CDbl(Txtdisplay(1)) + CDbl(Txtdisplay(2)), gs_formatAmountIDR)
    End If
    
    txtRemarks = Trim(rst!Remarks)
End If
End Sub

Sub createnumber()

'    Dim rs As Recordset
'
'    rstcust.filter = "cust_code = '" & cbodealer & "'"
'    If Not rstcust.EOF Then
'        sql = "select  RIGHT(rtrim(Invoice_No),4) Nomor, invoice_no " & _
'            "from invoice_master " & _
'            "where year(invoice_date) = " & IDate.Year & _
'            "AND LEN(Invoice_NO)<=9 order by invoice_no desc"
'
'        Set rs = New Recordset
'        rs.Open sql, Db, adOpenDynamic, adLockOptimistic
'        If Not rs.EOF Then
'            ComboBox1 = Format(IDate, "YYYY") & "-" & Format(Val(rs!nomor + 1), "0000")
'        Else
'            ComboBox1 = Format(IDate, "YYYY") & "-0001"
'        End If
'    End If
'    rstcust.filter = ""
'    rstcust.Requery
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim DISC As Double
Dim tmp As Double

blntotal = True
tppn = tax("ppn")
With Grid
    If Col = bteColQty And Trim(.TextMatrix(Row, bteColSJNo)) <> "" Then
        If CDbl(.TextMatrix(Row, bteColQty)) = 0 Then .TextMatrix(Row, bteColQty) = Format(0, gs_formatQty)
        If CDbl(.TextMatrix(Row, bteColQty)) > gd_MaxQty Then
            lblerror.Caption = DisplayMsg(4045) & " " & gd_MaxQty & " !"
            .TextMatrix(Row, bteColQtyRem) = Format(0, gs_formatQty)
            .TextMatrix(Row, bteColQty) = .TextMatrix(Row, bteColRefQty)
        ElseIf CDbl(.TextMatrix(Row, bteColQty)) > CDbl(.TextMatrix(Row, bteColRefQty)) Then
            lblerror.Caption = DisplayMsg(4045) & " " & .TextMatrix(Row, bteColRefQty) & " !"
            .TextMatrix(Row, bteColQtyRem) = Format(0, gs_formatQty)
            .TextMatrix(Row, bteColQty) = .TextMatrix(Row, bteColRefQty)
        Else
            .TextMatrix(Row, bteColQtyRem) = CDbl(.TextMatrix(Row, bteColRefQty)) - CDbl(.TextMatrix(Row, bteColQty))
            lblerror.Caption = ""
        End If
        .TextMatrix(Row, bteColQty) = Format(.TextMatrix(Row, bteColQty), gs_formatQty)
        .TextMatrix(Row, bteColQtyRem) = Format(.TextMatrix(Row, bteColQtyRem), gs_formatQty)
        .TextMatrix(Row, bteColStatus) = 1
        tmp = CDbl(.TextMatrix(Row, bteColPrice)) + CDbl(.TextMatrix(Row, bteColService))
        '.TextMatrix(.Row, bteColAmount) = Format(uf_Trunc((.TextMatrix(Row, bteColQty) * (.TextMatrix(Row, bteColPrice))), gi_decimalDigitAmountIDR), gs_formatAmountIDR)
        .TextMatrix(.Row, bteColAmount) = Format(uf_Trunc((.TextMatrix(Row, bteColQty) * (tmp)), gi_decimalDigitAmountIDR), gs_formatAmountIDR)
        gridsubtotal Trim(.TextMatrix(.Row, bteColItemCode)), CDbl(.TextMatrix(.Row, bteColPrice))
        totalbasechecked
    ElseIf Col = bteColSelect Then
        If Trim(currCode) = "" Then
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
                    currCode = Trim(.TextMatrix(i, bteColRefCurr))
                    lblerror = ""
                    Exit For
                End If
                currCode = ""
            Next
        End If
        If currCode <> Trim(.TextMatrix(Row, bteColRefCurr)) Then lblerror = DisplayMsg(4084): .Cell(flexcpChecked, Row, bteColSelect) = flexUnchecked: Exit Sub
        .TextMatrix(Row, bteColStatus) = 1
        totalbasechecked
        currCode = ""
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
                currCode = Trim(.TextMatrix(i, bteColRefCurr))
                lblerror = ""
                Exit For
            End If
        Next
        lblerror = ""
        Exit Sub
    End If
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col <> bteColSelect And Col <> bteColQty Then Cancel = True
Grid.EditMaxLength = 7
If Not CheckPackingNo(Grid.TextMatrix(Row, bteColSJNo)) Then Cancel = True
End Sub

Private Sub grid_Click()
If Grid.Rows <> 1 And Grid.Row <> -1 Then
With Grid
    If .Col = bteColQty And Trim(.TextMatrix(.Row, bteColSJNo)) <> "" Then
        .FocusRect = flexFocusInset
    Else
        .FocusRect = flexFocusNone
    End If
End With
End If
End Sub

Private Sub Grid_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = vbKeyDelete And Trim(Grid.TextMatrix(Grid.Row, bteColSJNo)) = "" Then KeyCode = 0
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If Trim(Grid.TextMatrix(Row, bteColSJNo)) = "" Then KeyAscii = 0
If Col = bteColSelect Then
    If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
      KeyAscii = 0
    End If
ElseIf Col = bteColQty Then
    If KeyAscii = Asc(".") Then KeyAscii = 0
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyEscape Then
      KeyAscii = 0
    End If
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub IDate_Change()
If UCase(combo1.Text) = "CREATE" Then createnumber
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
rstcust.Requery
rstcust.Find "cust_code = '" & cbodealer.Text & "'"
If Not rstcust.EOF Then
    rstcust.Requery
    isiCboDO
    If combo1.Text = "Create" Then
        clear2
        Grid.Editable = False
    Else
        If Trim(ComboBox1) = "" Then nomorinvoice: Exit Sub
        xno = ComboBox1
         'Update DUDI Januari 2009,mempercepat proses
        'mengambil no invoice jika telah ada
        Dim s As String
        s = "select * from invoice_master where cust_code = '" & cbodealer.Text & "' and year(invoice_date) >= '" & Year(SDate) & "' and year(invoice_date) <= '" & Year(EDate) & "'"
        If CekSql(s) Then
            nomorinvoice
        Else
            lblFix = ""
            Grid.Rows = 1
            ComboBox1.clear
            ComboBox1.locked = True
            
        End If
        
        
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
    xduedate = ""
Else
    xduedate = Format(MyDuedate, "YYYY-MM-DD")
End If
sql = "select * from invoice_master where invoice_no = '" & ComboBox1 & "'"
Set rsmaster = New Recordset
rsmaster.Open sql, Db, adOpenDynamic, adLockOptimistic
With rsmaster

Dim sqltradeterms_cls As String
Dim rstradeterms_cls As New ADODB.Recordset
    
sqltradeterms_cls = "select price_condition from trade_master " & _
                   "where trade_code = '" & cbodealer.Text & "'"
    
Set rstradeterms_cls = Db.Execute(sqltradeterms_cls)

If Not rsmaster.EOF Then
    .Fields("Cust_code") = cbodealer.Text
    .Fields("Invoice_No") = ComboBox1
    .Fields("Invoice_date") = Format(IDate.Value, "YYYY-MM-dd")
    .Fields("Delivery_date") = Format(DDate, "YYYYMM")
    .Fields("Amount") = CDbl(Txtdisplay(1))
    .Fields("PPN") = CDbl(Txtdisplay(2))
    .Fields("Total_Amount") = CDbl(Txtdisplay(3))
    .Fields("Remarks") = Trim(txtRemarks)
    .Fields("Exchange_rate") = 0
    .Fields("Exchange_amount") = 0

    If xduedate = "" Then
        .Fields("Due_date") = Null
    Else
        .Fields("Due_date") = xduedate
    End If
   
    .Fields("TradeTerms_Cls") = Trim(cboCls.Text)
    .Fields("PEBDate") = Format(dtpPEBDate.Value, "YYYY-MM-dd")
    .Fields("PEBNo") = Trim(txtPEBNo)
    
    .Fields("Last_Update") = Date
    .Fields("Last_User") = userLogin
    .Fields("TradeTerms_Cls") = IIf(IsNull(rstradeterms_cls!Price_Condition), "", Trim(rstradeterms_cls!Price_Condition))
    .update
    lblerror.Caption = DisplayMsg(1101)
Else
    .AddNew
    .Fields("Invoice_No") = ComboBox1
    .Fields("Invoice_date") = Format(IDate, "yyyy-mm-dd")
    .Fields("Delivery_date") = Format(DDate, "YYYY-MM-dd")
    .Fields("Cust_code") = cbodealer.Text
    .Fields("Amount") = Val(Txtdisplay(1))
    .Fields("PPN") = Val(Txtdisplay(2))
    .Fields("Total_Amount") = Val(Txtdisplay(3))
    .Fields("Remarks") = txtRemarks
    .Fields("Exchange_rate") = 0
    .Fields("Exchange_amount") = 0
    If xduedate = "" Then
        .Fields("Due_date") = Null
    Else
        .Fields("Due_date") = xduedate
    End If
    'On Error Resume Next
    .Fields("Last_Update") = Date
    .Fields("Last_User") = userLogin
    .Fields("TradeTerms_Cls") = IIf(IsNull(rstradeterms_cls!Price_Condition), "", Trim(rstradeterms_cls!Price_Condition))
    .update
HandleError:
    If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
        err.clear
        ComboBox1.locked = False
        Call createnumber
        .Fields("Invoice_No") = ComboBox1
        .Fields("Invoice_date") = Format(IDate, "yyyy-mm-dd")
        .Fields("Delivery_date") = Format(DDate, "YYYYMM")
        .Fields("Cust_code") = cbodealer.Text
        .Fields("Amount") = Val(Txtdisplay(1))
        .Fields("PPN") = Val(Txtdisplay(2))
        .Fields("Total_Amount") = Val(Txtdisplay(3))
        .Fields("Remarks") = txtRemarks
        .Fields("Exchange_rate") = 0
        .Fields("Exchange_amount") = 0
        
        If xduedate = "" Then
            .Fields("Due_date") = Null
        Else
            .Fields("Due_date") = xduedate
        End If
        .Fields("Last_Update") = Date
        .Fields("Last_User") = userLogin
        .Fields("TradeTerms_Cls") = IIf(IsNull(rstradeterms_cls!Price_Condition), "", Trim(rstradeterms_cls!Price_Condition))
        .update
        Set rstradeterms_cls = Nothing
        If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then GoTo HandleError
    End If
    lblerror.Caption = DisplayMsg(1000)
End If
.filter = ""
.Requery
End With
Set rstradeterms_cls = Nothing
End Sub

Sub savedetail()
    Dim ls_sql As String
    For i = 1 To Grid.Rows - 1
        If Grid.TextMatrix(i, bteColSJNo) <> "" Then
            If Grid.Cell(flexcpChecked, i, bteColSelect, i, bteColSelect) = flexChecked And Grid.TextMatrix(i, bteColStatus) = 1 Then
                sql = "select * from invoice_detail where Invoice_No = '" & ComboBox1 & "' and DO_No ='" & Grid.TextMatrix(i, bteColSJNo) & "' and Po_no = '" & Grid.TextMatrix(i, bteColPONo) & "' and seq_no= '" & Grid.TextMatrix(i, bteColSeqNo) & "' and doSeq_no = '" & Grid.TextMatrix(i, bteColDOSeqNo) & "' "
                Set rsdetail = New Recordset
                rsdetail.Open sql, Db, adOpenDynamic, adLockOptimistic
                With rsdetail
                If .EOF Then
                    .AddNew
                    .Fields("invoice_no") = ComboBox1
                    .Fields("Packing_No") = ""
                    .Fields("PackingSeq_No") = 0
                    .Fields("DO_No") = Grid.TextMatrix(i, bteColSJNo)
                    .Fields("item_code") = Grid.TextMatrix(i, bteColItemCode)
                    .Fields("MakerItem_code") = Grid.TextMatrix(i, bteColPartNo)
                    .Fields("Delivery_date") = Format(Grid.TextMatrix(i, bteColDelDate), "yyyy-mm-dd")
                    .Fields("PO_no") = Grid.TextMatrix(i, bteColPONo)
                    .Fields("Seq_no") = CDbl(Grid.TextMatrix(i, bteColSeqNo))
                    .Fields("doSeq_no") = CDbl(Grid.TextMatrix(i, bteColDOSeqNo))
                    .Fields("Qty") = CDbl(Grid.TextMatrix(i, bteColQty))
                    .Fields("Price") = CDbl(Grid.TextMatrix(i, bteColPrice))
                    .Fields("Service") = CDbl(Grid.TextMatrix(i, bteColService))
                    .Fields("currency_code") = Grid.TextMatrix(i, bteColRefCurr)
                    .Fields("Unit_cls") = Grid.TextMatrix(i, bteColRefUnit)
                    .Fields("Amount") = CDbl(Grid.TextMatrix(i, bteColAmount))
                    .Fields("ExchangeRate_Amount") = Daily_Rate(IDate, ComboBox1, Grid.TextMatrix(i, bteColRefCurr)) * CDbl(Grid.TextMatrix(i, bteColAmount))
                    
                    
                    If Not BolExchange Then Exit For
                    .Fields("Last_Update") = Date
                    .Fields("Last_User") = userLogin
                    .update
                    lblerror.Caption = DisplayMsg(1000)
                    

                Else
                
                    If .Fields("Qty") <> CDbl(Grid.TextMatrix(i, bteColQty)) Or .Fields("Price") <> CDbl(Grid.TextMatrix(i, bteColPrice)) Then
                            .Fields("Packing_No") = ""
                            .Fields("PackingSeq_No") = 0
                            .Fields("DO_No") = Grid.TextMatrix(i, bteColSJNo)
                            .Fields("item_code") = Grid.TextMatrix(i, bteColItemCode)
                            .Fields("MakerItem_code") = Grid.TextMatrix(i, bteColPartNo)
                            .Fields("Delivery_date") = Format(Grid.TextMatrix(i, bteColDelDate), "yyyy-mm-dd")
                            .Fields("PO_no") = Grid.TextMatrix(i, bteColPONo)
                            .Fields("Seq_no") = CDbl(Grid.TextMatrix(i, bteColSeqNo))
                            .Fields("doSeq_no") = CDbl(Grid.TextMatrix(i, bteColDOSeqNo))
                            .Fields("Qty") = CDbl(Grid.TextMatrix(i, bteColQty))
                            .Fields("Price") = CDbl(Grid.TextMatrix(i, bteColPrice))
                            .Fields("currency_code") = Grid.TextMatrix(i, bteColRefCurr)
                            .Fields("Unit_cls") = Grid.TextMatrix(i, bteColRefUnit)
                            .Fields("Amount") = CDbl(Grid.TextMatrix(i, bteColAmount))
                            .Fields("ExchangeRate_Amount") = Daily_Rate(IDate, ComboBox1, Grid.TextMatrix(i, bteColRefCurr)) * CDbl(Grid.TextMatrix(i, bteColAmount))
                            
                            If Not BolExchange Then Exit For
                            .Fields("Last_Update") = Date
                            .Fields("Last_User") = userLogin
                            .update
                            

                    End If

                    lblerror.Caption = DisplayMsg(1101)

                End If
                End With
            Else
                If Trim(Grid.TextMatrix(i, bteColSJNo)) <> "" And Grid.TextMatrix(i, bteColStatus) = 1 Then
                sql = "select * from invoice_detail where Invoice_No = '" & ComboBox1 & "' and DO_No ='" & Grid.TextMatrix(i, bteColSJNo) & "' and PO_no = '" & Grid.TextMatrix(i, bteColPONo) & "' and seq_no = '" & Grid.TextMatrix(i, bteColSeqNo) & "' and doseq_no = '" & Grid.TextMatrix(i, bteColDOSeqNo) & "' "
                Set rsdetail = New Recordset
                rsdetail.Open sql, Db, adOpenDynamic, adLockOptimistic
                    With rsdetail
                    If Not .EOF Then
                        'Delete Invoice
                        sql = "delete invoice_detail where Invoice_No = '" & ComboBox1 & "' and DO_No ='" & Grid.TextMatrix(i, bteColSJNo) & "' and PO_no = '" & Grid.TextMatrix(i, bteColPONo) & "' and seq_no = '" & Grid.TextMatrix(i, bteColSeqNo) & "' and doseq_no ='" & Grid.TextMatrix(i, bteColDOSeqNo) & "' "
                        Db.Execute sql
                        lblerror.Caption = DisplayMsg(1101)
                    End If
                    End With
                End If
            End If
        End If
    Next


End Sub

Sub totalbasechecked()
Dim AmountInv As Double
With Grid
AmountInv = 0
For i = 1 To Grid.Rows - 1
    If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
        If Val(AmountInv) = 0 Then
            AmountInv = CDbl(.TextMatrix(i, bteColAmount))
        Else
            AmountInv = CDbl(AmountInv) + CDbl(.TextMatrix(i, bteColAmount))
        End If
    End If
Next
Txtdisplay(1).Text = Format(AmountInv, gs_formatAmountIDR)
If Overseas_Cls = "0" Then
    Txtdisplay(2).Text = Format((AmountInv * tax("ppn")) / 100, gs_formatAmountIDR)
Else
    Txtdisplay(2).Text = Format("0", gs_formatAmountIDR)
End If
Txtdisplay(3).Text = Format(CDbl(Txtdisplay(1).Text) + CDbl(Txtdisplay(2).Text), gs_formatAmountIDR)
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
txtRemarks = ""
Txtdisplay(0) = ""
Txtdisplay(1) = Format(0, gs_formatAmountIDR)
Txtdisplay(2) = Format(0, gs_formatAmountIDR)
Txtdisplay(3) = Format(0, gs_formatAmountIDR)
LblDesc(0).Caption = ""
LblDesc(1).Caption = ""
MEDuedate.Text = "99/99/9999"
Header
cmdAction(2).Enabled = False
cmdAction(4).Enabled = False 'BARU
cmdAction(5).Caption = "Update"
cbodealer.locked = False
ComboBox1.locked = False
ComboBox1 = ""
ComboBox1.clear
bsave = False
currCode = ""
cboCls.Text = ""
txtPEBNo.Text = ""

End Sub

Sub clear2()
    Txtdisplay(0) = ""
    Txtdisplay(1) = Format(0, gs_formatAmountIDR)
    Txtdisplay(2) = Format(0, gs_formatAmountIDR)
    Txtdisplay(3) = Format(0, gs_formatAmountIDR)
End Sub

Sub updateMaster()
Dim rsUpdate As Recordset, xrs As Recordset
'tomaster
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
    
    If Overseas_Cls = "1" Then
        Txtdisplay(2) = Format(0, gs_formatAmountIDR)
        Txtdisplay(3) = Txtdisplay(1)
    Else
        Txtdisplay(2) = Format(((CDbl(Txtdisplay(1)) * tax("Ppn")) / 100), gs_formatAmountIDR)
        Txtdisplay(3) = Format(CDbl(Txtdisplay(1)) + CDbl(Txtdisplay(2)), gs_formatAmountIDR)
    End If
    
    !Amount = CDbl(Txtdisplay(1))
    !ppn = CDbl(Txtdisplay(2))  'tppn * CDbl(Txtdisplay(1)) / 100
    !total_amount = CDbl(Txtdisplay(1)) + CDbl(Txtdisplay(2))
    !Remarks = Trim(txtRemarks.Text)
    !Invoice_Date = Format(IDate, "yyyy-MM-dd")
    !delivery_Date = Format(DDate, "yyyy-MM-dd")
    If MEDuedate.Text = "99/99/9999" Then
        !due_date = Null
    Else
        !due_date = Format(MyDuedate, "YYYY-MM-DD")
    End If
    !exchange_Rate = Daily_Rate(IDate, Trim(ComboBox1))
    !exchange_amount = !exchange_Rate * !total_amount
    !Last_Update = Date
    !last_user = userLogin
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
sql = "select * from invoice_master where cust_code = '" & cbodealer.Text & "' and year(invoice_date) >= '" & Year(SDate) & "' and year(invoice_date) <= '" & Year(EDate) & "'"
Set rsmaster = Db.Execute(sql)
If Not rsmaster.EOF Then
    ComboBox1.clear
    ComboBox1.locked = False
    ComboBox1.columnCount = 2
    ComboBox1.ColumnWidths = "100 pt;0 pt"
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
Dim prm_invNo As String * 25
prm_invNo = p_invNo
cbodealer.Text = p_custCode
ComboBox1 = prm_invNo
End Sub

Function listDO(noinvoice$) As String
Dim rsIsiDO As New ADODB.Recordset
Dim tampungDO As String
    sql = "Select distinct a.DO_No from invoice_detail  a where a.invoice_no = '" & noinvoice & "'"
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

Sub ShowData()
cmdAction_Click (5)
End Sub


Sub gridsubtotal(Model$, xprice As Double)
Dim qtytemp As Double, bol As Boolean
Dim pricetemp As Double, disctemp As Double, tempamount As Double
qtytemp = 0
pricetemp = 0
tempamount = 0
With Grid
For i = 1 To Grid.Rows - 1
    If UCase(Trim(.TextMatrix(i, bteColRefQty))) = UCase("subtotal" & Trim(Model) & xprice) Then Exit For
        If Trim(.TextMatrix(i, bteColPrice)) = "" Then
            pricetemp = 0
        Else
            pricetemp = CDbl(.TextMatrix(i, bteColPrice))
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

Function Daily_Rate(Tgl As Date, InvoiceNo As String, Optional Curr As String) As Double
Dim rstrate As New Recordset, rstcurr As New Recordset
If Trim(Curr) = "03" Then Daily_Rate = 1: BolExchange = True: Exit Function
If Trim(Curr) = "" Then
    sql = "select distinct currency_code from Invoice_detail where invoice_no = '" & InvoiceNo & "'"
    If rstcurr.State <> adStateClosed Then rstcurr.Close
    rstcurr.Open sql, Db, adOpenStatic, adLockOptimistic
    If Not rstcurr.EOF Then
        sql = "select daily_ExchangeRate from daily_exchangeRate where ExchangeRate_Date = '" & Format(Tgl, "YYYY-MM-DD") & "' and currency_Code = '" & rstcurr!currency_code & "'"
        If rstrate.State <> adStateClosed Then rstrate.Close
        rstrate.Open sql, Db, adOpenStatic, adLockOptimistic
        If Not rstrate.EOF Then
            Daily_Rate = rstrate!daily_ExchangeRate
            BolExchange = True
        Else
            BolExchange = True
            Daily_Rate = 0
        End If
        rstrate.Close
        Set rstrate = Nothing
    Else
        Daily_Rate = 0
    End If
    rstcurr.Close
    Set rstcurr = Nothing
Else
    sql = "select daily_ExchangeRate from daily_exchangeRate where ExchangeRate_Date = '" & Format(Tgl, "YYYY-MM-DD") & "' and currency_Code = '" & Curr & "'"
    If rstrate.State <> adStateClosed Then rstrate.Close
    rstrate.Open sql, Db, adOpenStatic, adLockOptimistic
    If Not rstrate.EOF Then
        Daily_Rate = rstrate!daily_ExchangeRate
        BolExchange = True
    Else
        BolExchange = True
        Daily_Rate = 0
    End If
    rstrate.Close
    Set rstrate = Nothing
End If
End Function

Private Function CheckPackingNo(strTempNo As String) As Boolean
    
    Dim intRow As Integer
    CheckPackingNo = True
    If strTempNo = "" Then Exit Function
    If Not gb_AllowMultipleDO_InvoiceCreate Then
        For intRow = 1 To Grid.Rows - 1
            If Grid.Cell(flexcpChecked, intRow, bteColSelect) = flexChecked And Trim(strTempNo) <> Trim(Grid.TextMatrix(intRow, bteColSJNo)) Then
                lblerror = DisplayMsg("0010")
                CheckPackingNo = False
                Exit For
            End If
        Next
    End If
    
End Function

Sub isiCboDO()

    Dim rscbo As New ADODB.Recordset
    
    With cboDO
        
        .clear
        .columnCount = 3
        .TextColumn = 1
        
        If gb_AllowMultipleDO_InvoiceCreate Then
            .AddItem ""
            .List(0, 0) = strAll
            .List(0, 1) = strAll
            i = 1
        Else
            i = 0
        End If
        
        sql = "Select a.DO_No, DO_Date, (Select Max (Delivery_Date)From Delivery_Order Where DO_No = a.DO_No) Delivery_Date From DO_Master a " & _
            "Where Coalesce(Fix_Cls,'0')='1' and a.Cust_Code = '" & Trim(cbodealer.Text) & "' " & _
            "And a.DO_Date >= '" & Format(SDate.Value, "YYYY-MM-DD") & "' " & _
            "And a.DO_Date <= '" & Format(EDate.Value, "YYYY-MM-DD") & "'  "
        
        Set rscbo = Db.Execute(sql)
        Do While Not (rscbo.EOF)
            .AddItem ""
            .List(i, 0) = Trim(rscbo("DO_No"))
            .List(i, 1) = Trim(rscbo("DO_Date"))
            .List(i, 2) = Trim(IIf(IsNull(rscbo("Delivery_Date")), "1900-01-01", rscbo("Delivery_date")))
            i = i + 1
            rscbo.MoveNext
        Loop
        
        .Text = ""
        .ListWidth = 150
        .ColumnWidths = "150pt;0pt;0pt"
        
        If gb_AllowMultipleDO_InvoiceCreate Then .ListIndex = 0
    
    End With

    Set rscbo = Nothing

End Sub

Private Function GetDONumber() As String
    
    Dim adoRs As New ADODB.Recordset
    
    sql = "Select DO_No From Invoice_Detail Where Invoice_No = '" & Trim(ComboBox1) & "'"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        GetDONumber = Trim(adoRs.Fields("DO_No"))
    End If
    adoRs.Close
    
End Function

Private Function GetInvoiceNumber() As String
    
    Dim adoRs As New ADODB.Recordset
    
    sql = "Select Invoice_No From Invoice_Detail Where DO_No = '" & Trim(cboDO) & "'"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        GetInvoiceNumber = Trim(adoRs.Fields("Invoice_No"))
    End If
    adoRs.Close
    
End Function

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

