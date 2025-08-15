VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOOther 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Scheduled (Others)"
   ClientHeight    =   10815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "POOther.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10815
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRemarks 
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
      Left            =   8760
      MaxLength       =   100
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   8160
      Width           =   6045
   End
   Begin VB.TextBox txtRemarks 
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
      Left            =   8760
      MaxLength       =   100
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   7725
      Width           =   6045
   End
   Begin VB.TextBox txtRemarks 
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
      Left            =   8760
      MaxLength       =   100
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   7320
      Width           =   6045
   End
   Begin VB.TextBox txtpoday 
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
      Left            =   2940
      MaxLength       =   3
      TabIndex        =   14
      Top             =   7725
      Width           =   495
   End
   Begin VB.TextBox txtRev 
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
      Left            =   8520
      MaxLength       =   20
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2280
      Width           =   2415
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13343
      TabIndex        =   50
      Top             =   210
      Width           =   1845
      _extentx        =   3254
      _extenty        =   767
   End
   Begin MSComCtl2.FlatScrollBar hscrollbar 
      Height          =   255
      Left            =   90
      TabIndex        =   45
      Top             =   6870
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Max             =   1
      Orientation     =   1245185
   End
   Begin VB.CommandButton cmdreport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Preview"
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   10200
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   580
      Left            =   83
      TabIndex        =   46
      Top             =   840
      Width           =   15105
      Begin MSComCtl2.DTPicker requestdate1 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   180
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   67436547
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker requestdate2 
         Height          =   315
         Left            =   3840
         TabIndex        =   1
         Top             =   180
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   67436547
         CurrentDate     =   37798
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
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
         Left            =   3360
         TabIndex        =   49
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Request Date From"
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
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1815
      End
      Begin MSForms.ComboBox cborequestno 
         Height          =   315
         Left            =   6720
         TabIndex        =   2
         Top             =   180
         Width           =   1815
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3201;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Request No"
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
         Left            =   5640
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   4170
      Left            =   90
      TabIndex        =   9
      Top             =   2745
      Width           =   15105
      _cx             =   26644
      _cy             =   7355
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
      ScrollBars      =   2
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
      Begin MSComCtl2.DTPicker DelDate 
         Height          =   315
         Left            =   3720
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   67436547
         CurrentDate     =   37798
      End
      Begin MSForms.ComboBox cbocurr 
         Height          =   285
         Left            =   5640
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   855
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "1508;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.TextBox txtppn 
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
      Left            =   9990
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   9045
      Width           =   2355
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
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   9045
      Width           =   2355
   End
   Begin VB.TextBox txtpono2 
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
      Left            =   300
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   9045
      Width           =   2715
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
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
      Left            =   11567
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   10200
      Width           =   1125
   End
   Begin VB.TextBox txtpono 
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
      Left            =   2280
      MaxLength       =   25
      TabIndex        =   5
      Top             =   2280
      Width           =   2470
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Create"
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2235
      Width           =   1125
   End
   Begin VB.TextBox txtgrandtotal 
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
      Left            =   12660
      Locked          =   -1  'True
      MaxLength       =   35
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   9045
      Width           =   2355
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   83
      TabIndex        =   34
      Top             =   9480
      Width           =   15105
      Begin VB.Label lblErrMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   180
         Width           =   14865
      End
   End
   Begin VB.CommandButton command1 
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
      Left            =   14063
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   10200
      Width           =   1125
   End
   Begin VB.CommandButton cmdsubmenu 
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
      Left            =   83
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   10200
      Width           =   1125
   End
   Begin VB.CommandButton command1 
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
      Index           =   1
      Left            =   12814
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   10200
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   720
      Left            =   83
      TabIndex        =   37
      Top             =   1440
      Width           =   15105
      Begin VB.TextBox lblcust 
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
         Height          =   200
         Index           =   0
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   280
         Width           =   4995
      End
      Begin VB.TextBox lblcust 
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
         Index           =   1
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   280
         Width           =   5715
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   9120
         X2              =   14880
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Index           =   4
         Left            =   8280
         TabIndex        =   41
         Top             =   270
         Width           =   840
      End
      Begin MSForms.ComboBox cbocust 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1305
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2302;556"
         ColumnCount     =   12
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   2880
         X2              =   7920
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label LblCode 
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
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   270
         Width           =   1815
      End
   End
   Begin MSComCtl2.DTPicker podate 
      Height          =   315
      Left            =   6120
      TabIndex        =   6
      Top             =   2280
      Width           =   1500
      _ExtentX        =   2646
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
      Format          =   67436547
      CurrentDate     =   37798
   End
   Begin VB.Label Label6 
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
      Index           =   2
      Left            =   7680
      TabIndex        =   57
      Top             =   7365
      Width           =   1020
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "days"
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
      Left            =   3420
      TabIndex        =   56
      Top             =   7755
      Width           =   495
   End
   Begin MSForms.ComboBox cbopoterms 
      Height          =   315
      Left            =   4020
      TabIndex        =   15
      Top             =   7725
      Width           =   2415
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4260;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cbopocode 
      Height          =   315
      Left            =   1860
      TabIndex        =   13
      Top             =   7725
      Width           =   975
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "1720;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Terms "
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
      Left            =   240
      TabIndex        =   55
      Top             =   7755
      Width           =   1575
   End
   Begin MSForms.ComboBox cbodelmode 
      Height          =   315
      Left            =   1860
      TabIndex        =   16
      Top             =   8160
      Width           =   1815
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3201;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transportation"
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
      Left            =   240
      TabIndex        =   54
      Top             =   8205
      Width           =   1620
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Price Condition"
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
      Left            =   240
      TabIndex        =   53
      Top             =   7365
      Width           =   1575
   End
   Begin VB.Label lblpricecondition 
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
      Left            =   2820
      TabIndex        =   52
      Top             =   7380
      Width           =   3525
   End
   Begin VB.Line Line4 
      X1              =   2820
      X2              =   6370
      Y1              =   7605
      Y2              =   7605
   End
   Begin MSForms.ComboBox cbopricecondition 
      Height          =   315
      Left            =   1860
      TabIndex        =   12
      Top             =   7320
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev"
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
      Index           =   0
      Left            =   8040
      TabIndex        =   51
      Top             =   2325
      Width           =   525
   End
   Begin VB.Label lblfix 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Status Fix "
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
      Left            =   14003
      TabIndex        =   42
      Top             =   2325
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PO No"
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
      Index           =   1
      Left            =   1560
      TabIndex        =   40
      Top             =   2325
      Width           =   720
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Date"
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
      Index           =   2
      Left            =   5280
      TabIndex        =   39
      Top             =   2325
      Width           =   825
   End
   Begin MSForms.ComboBox cbopono 
      Height          =   315
      Left            =   2280
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2745
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "4842;556"
      ListWidth       =   5291
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "2143;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   12660
      TabIndex        =   36
      Top             =   8685
      Width           =   2295
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   7320
      TabIndex        =   33
      Top             =   8685
      Width           =   2325
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PO No"
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
      Index           =   0
      Left            =   300
      TabIndex        =   32
      Top             =   8685
      Width           =   2340
   End
   Begin VB.Label Label12 
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
      Height          =   255
      Left            =   9990
      TabIndex        =   31
      Top             =   8685
      Width           =   2325
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Scheduled (Others)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   83
      TabIndex        =   30
      Top             =   240
      Width           =   15105
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   495
      Index           =   2
      Left            =   90
      Top             =   8940
      Width           =   15105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   330
      Index           =   2
      Left            =   90
      Top             =   8625
      Width           =   15105
   End
End
Attribute VB_Name = "frmPOOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sql As String, sqlGrid As String
Dim rs As New ADODB.Recordset, rsGrid As New ADODB.Recordset
Dim i As Integer, orderawal As Double, isippn As Long
Dim ubah As Boolean, ubahgrid As Boolean, ada As Boolean, sampun As Boolean
Dim statusfix As String
Dim actrow As Long, activecurrcd As String, activecurr As String
Dim countrycls As Byte, j As Integer
Const isiPOTerm = "after B/L Date,after Delivery Date,after Invoice Date,prior before Shipment,from Custom Clearance Date,after Receive Invoice,after Receive Goods"

Private Sub ClearData()

    Sql = "Delete from PurchaseOrder_Master Where PO_No = '" & Trim(txtpono.Text) & "' " & _
          "and PO_No not in (select PO_No from PurchaseOrder_Detail) and others_cls = '1' and period is null"
    Db.Execute Sql

End Sub

Private Sub CekPONumber()
    
    Dim adoRs As New ADODB.Recordset
    
    Sql = "Select * From PurchaseOrder_Master Where PO_No = '" & Trim(txtpono.Text) & "'"
    adoRs.Open Sql, Db, adOpenKeyset, adLockOptimistic, adCmdText
    If Not adoRs.EOF Then
        Call poNo(Right(year(podate), 2), Format(month(podate), "0#"))
        adoRs.update "PO_No", Trim(txtpono.Text)
    End If
    adoRs.Close
    Set adoRs = Nothing
    
End Sub

Sub kosong()
    lblErrMsg = ""
    requestdate1.Value = Format(Now, "yyyy-mm-01")
    requestdate2.Value = Format(Now, "dd MMM yyyy")
    Call adtocborequestno
    cborequestno.Text = ""
    
    cbocust.Text = ""
    lblcust(0).Text = "": lblcust(1).Text = ""
    txtpono.Text = "": txtpono2.Text = ""
    podate.Value = Format(Now, "dd MMM yyyy")
    podate.Enabled = True
    Call ppn(podate.Value)
    txtRev.Text = ""
    
    grid.FocusRect = flexFocusNone
    DelDate.Value = Format(Now + 1, "dd MMM yyyy")
    
    cbopricecondition.Text = ""
    cbopocode.ListIndex = -1
    txtpoday.Text = ""
    cbopoterms.ListIndex = -1
    txtRemarks(0).Text = ""
    txtRemarks(1).Text = ""
    txtRemarks(2).Text = ""
    cbodelmode.ListIndex = -1
    
    ubah = False: ada = False
    statusfix = 0
    
    Call Kunci(False)
    Call kosongBwh: Call header
End Sub

Sub kosongBwh()
    txtamount.Text = 0
    txtppn.Text = 0
    txtgrandtotal.Text = 0
End Sub

Function adtocbocust(ByVal Filter As Boolean)
Dim sqlcust As String
Dim RsCust As New Recordset

    If Filter = True Then
        sqlcust = "select tm.trade_code, tm.trade_name, tm.address1, tm.country_cls, tm.po_cls, isnull(tm.Trade_Abbr,'') Trade_Abbr " & _
                  ", tm.popayment_code, tm.popayment_day, tm.popayment_terms, tm.transportation_Cls " & _
                  "from trade_master tm " & _
                  "left outer Join (select distinct pom.supplier_code from PurchaseOrder_Master pom " & _
                  "                 inner join (select PO_No from PurchaseOrder_Detail where porequest_no = '" & Trim(cborequestno.Text) & "') pod on pod.PO_No = pom.PO_No " & _
                  "                 where pom.others_Cls = '1' and pom.period is null) sc " & _
                  "on sc.supplier_code = tm.trade_code " & _
                  "where ((tm.trade_cls='2' And tm.TradeExternal_Cls in ('1','2')) or tm.trade_cls='3') " & _
                  "Order By tm.Trade_Abbr "
    Else
        sqlcust = "select trade_code, trade_name, address1, country_cls, po_cls, isnull(Trade_Abbr,'') Trade_Abbr " & _
                  ", popayment_code, popayment_day, popayment_terms, transportation_Cls " & _
                  "from trade_master where ((trade_cls='2' And TradeExternal_Cls in ('1','2'))  or trade_cls='3') Order By Trade_Abbr"
'         sqlcust = "select trade_code, trade_name, address1, country_cls, po_cls, isnull(Trade_Abbr,'') Trade_Abbr " & _
                  ", popayment_code, popayment_terms, transportation_Cls " & _
                  "from trade_master where (trade_cls='2')"
    End If
    Set RsCust = Db.Execute(sqlcust)

    With cbocust
        .clear
        .ColumnCount = 8
        .ColumnWidths = "40pt;120pt;0pt;0pt;0pt;0pt;0pt;200pt"
        .ListWidth = 360
        .ListRows = 15

        i = 0
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("Trade_code"))
            .List(i, 1) = IIf(IsNull(RsCust("trade_abbr")), "", Trim(RsCust("Trade_abbr")))
            .List(i, 2) = IIf(IsNull(RsCust("address1")), "", Trim(RsCust("Address1")))
            .List(i, 3) = IIf(IsNull(RsCust("country_cls")), 3, Trim(RsCust("country_cls")))
            .List(i, 4) = IIf(IsNull(RsCust("po_cls")), 0, Trim(RsCust("po_cls")))
            .List(i, 5) = IIf(IsNull(RsCust("popayment_code")), "", RsCust("popayment_code")) & "," & IIf(IsNull(RsCust("popayment_day")), "", RsCust("popayment_day")) & "," & IIf(IsNull(RsCust("popayment_terms")), "", RsCust("popayment_terms"))
            .List(i, 6) = IIf(IsNull(RsCust("transportation_Cls")), -1, RsCust("transportation_Cls"))
            .List(i, 7) = IIf(IsNull(RsCust("trade_name")), "", Trim(RsCust("Trade_Name")))
            RsCust.MoveNext
            i = i + 1
        Loop
    End With
    Set RsCust = Nothing
End Function

Sub adtocbo(nmCombo, nmConst, start As Integer, col1 As Integer, Col2 As Integer, kosong As Boolean)
Dim j As Integer, k As Integer

    With nmCombo
        .clear
        .ColumnCount = 2
        .TextColumn = 2
        j = 0: k = start
        
        If kosong Then
            .AddItem
            .List(0, 0) = "Null"
            .List(0, 1) = ""
            j = 1
        End If
        
        For i = 0 To UBound(Split(nmConst, ","))
            .AddItem
            .List(j, 0) = k
            .List(j, 1) = Split(nmConst, ",")(i)
            k = k + 1: j = j + 1
        Next i
        .ListRows = 10
        .ListWidth = col1 + Col2
        .ColumnWidths = col1 & "pt;" & Col2 & "pt"
    End With
End Sub

Sub adtocombo()
Dim sql1 As String
Dim rs1 As New Recordset

    combo1.AddItem "Create"
    combo1.AddItem "Update"
    'Call isiCboUnitCurr(cbocurr, isiCurr, 0, 4)
    Call up_FillCombo(cbocurr, "curr_cls")
    cbocurr.TextColumn = 2
    
    Call adtocbo(cbopocode, isiPOCode, 1, 0, 40, False)
    Call adtocbo(cbopoterms, isiPOTerm, 1, 0, 140, False)
    Call adtocbo(cbodelmode, isiTransport, 0, 0, 80, False)

    'PRICE CONDITION
    sql1 = "select * from PriceCondition_cls"
    If rs1.State <> adStateClosed Then rs1.Close
    rs1.Open sql1, Db, adOpenKeyset, adLockOptimistic
    If Not (rs1.BOF And rs1.EOF) Then
        i = 0
        With cbopricecondition
            .clear
            .ColumnCount = 2
            .ColumnWidths = "25pt;175pt"
            .ListWidth = 200
            .ListRows = 15
            
            Do While Not rs1.EOF
                .AddItem ""
                .List(i, 0) = Trim(rs1!pricecondition_cls)
                .List(i, 1) = Trim(rs1!Description)
                i = i + 1
                rs1.MoveNext
            Loop
        End With
    End If
    Set rs1 = Nothing
End Sub

Sub adtocborequestno()
Dim sqlno As String, tempcust As String
Dim rsno As New Recordset
    
    sqlno = "select porequest_no, porequest_period, isnull(fix_cls,'0') fix_cls " & _
            "from PORequest_Master where isnull(others_cls,'0')='1' and isnull(fix_cls,'0')='1' " & _
            "and porequest_date >= '" & Format(requestdate1.Value, "yyyy-mm-dd") & "' " & _
            "and porequest_date <= '" & Format(requestdate2.Value, "yyyy-mm-dd") & "' "
    Set rsno = Db.Execute(sqlno)
    With cborequestno
        .clear
        .ColumnCount = 3
        
        If Trim(cbocust.Text) <> "" Then tempcust = Trim(cbocust.Text)
        Call adtocbocust(False): cbocust.Text = tempcust: Call cbocust_Click
        
        i = 0
        Do While Not rsno.EOF
            .AddItem
            .List(i, 0) = Trim(rsno("PORequest_No"))
            .List(i, 1) = Trim(rsno("PORequest_Period"))
            .List(i, 2) = Trim(rsno("fix_cls"))
            rsno.MoveNext
            i = i + 1
        Loop
        .ColumnWidths = "90pt;0pt;0pt"
        .ListWidth = 90
        .ListRows = 15
    End With
    Set rsno = Nothing
End Sub

Sub adtocbopono()
Dim sqlno As String
Dim rsno As New Recordset
    
'    sqlno = "select pom.PO_No from PurchaseOrder_Master pom " & _
'            "where pom.others_cls = '1' and pom.period is null " & _
'            "and pom.supplier_Code = '" & Trim(cbocust.Text) & "' " & _
'            "and pom.po_date >= '" & Format(requestdate1, "yyyy-mm-dd") & "' " & _
'            "and pom.po_date <= '" & Format(requestdate2, "yyyy-mm-dd") & "' "

' PO tidak di filter berdasarkan supplier code

    sqlno = "select pom.PO_No from PurchaseOrder_Master pom " & _
            "where pom.others_cls = '1' and pom.period is null " & _
            "and pom.po_date >= '" & Format(requestdate1, "yyyy-mm-dd") & "' " & _
            "and pom.po_date <= '" & Format(requestdate2, "yyyy-mm-dd") & "' "
    Set rsno = Db.Execute(sqlno)
    With cbopono
        .clear
        Do While Not rsno.EOF
            .AddItem Trim(rsno("PO_No"))
            rsno.MoveNext
        Loop
        .ColumnWidths = "150pt"
        .ListWidth = 150
        .ListRows = 15
    End With
    Set rsno = Nothing
End Sub

Sub poNo(ByVal thn As String, ByVal bln As String)
Dim sqlno As String, sqlS As String
Dim rsno As New Recordset, rsS As New Recordset
    
    'POYYMM999
    If Trim(txtpono.Text) = "" Then
'        If Format(podate, "YYYY-MM-01") > "2006-07-30" Then
'            sqlno = "select top 1 rtrim(PO_No) from PurchaseOrder_Master " & _
'                    "where substring(rtrim(PO_No),3,2) = '" & thn & "' and substring(rtrim(PO_No),5,2) > '07' " & _
'                    "order by right(rtrim(PO_No),5) desc"
'        Else
            sqlno = "select top 1 rtrim(PO_No) from PurchaseOrder_Master " & _
                    "where substring(rtrim(PO_No),3,2) = '" & thn & "' " & _
                    "order by right(rtrim(PO_No),5) desc"
'        End If
    Else
'        If Format(podate, "YYYY-MM-01") > "2006-07-30" Then
'            sqlno = "select top 1 rtrim(PO_No) from PurchaseOrder_Master " & _
'                    "where substring(rtrim(PO_No),3,2) = '" & thn & "' and substring(rtrim(PO_No),5,2) > '07' " & _
'                    "and Right(RTrim(PO_No), 9)  < '" & Right(Trim(txtpono.Text), 9) & "' " & _
'                    "order by right(rtrim(PO_No),5) desc"
'        Else
            sqlno = "select top 1 rtrim(PO_No) from PurchaseOrder_Master " & _
                    "where substring(rtrim(PO_No),3,2) = '" & thn & "' " & _
                    "and Right(RTrim(PO_No), 9)  < '" & Right(Trim(txtpono.Text), 9) & "' " & _
                    "order by right(rtrim(PO_No),5) desc"
'        End If
    End If
    Set rsno = Db.Execute(sqlno)
    If Not (rsno.BOF And rsno.EOF) Then
        txtpono.Text = Left(Trim(rsno(0)), 4) & bln & Format(Right(Trim(rsno(0)), 5) + 1, "0000#")
    Else
        sqlS = "select top 1 PO_No from Initial_No"
        Set rsS = Db.Execute(sqlS)
        If Not (rsS.BOF And rsS.EOF) Then
            txtpono.Text = Left(Trim(rsS(0)), 2) & thn & bln & Right(Trim(rsS(0)), 5)
        Else
            txtpono.Text = "PO" & thn & bln & "00001"
        End If
        Set rsS = Nothing
    End If
    txtpono.locked = True
    Set rsno = Nothing
End Sub

Function seqNo() As Long
Dim sqlseqno As String
Dim rsseqno As New Recordset

    sqlseqno = "select Seq_No  from PurchaseOrder_Detail order by PoReq_SeqNo desc"
    If rsseqno.State <> adStateClosed Then rsseqno.Close
    rsseqno.Open sqlseqno, Db, adOpenKeyset, adLockOptimistic
    
    If Not (rsseqno.BOF And rsseqno.EOF) Then
        seqNo = rsseqno!seq_no + 1
    Else
        seqNo = 1
    End If
    Set rsseqno = Nothing
End Function

Sub Kunci(l As Boolean)
    podate.Enabled = Not l
    grid.Editable = Not l
    command1(0).Enabled = Not l
    lblfix.Caption = "Status Fix"
    lblfix.Visible = l
End Sub

Sub ppn(ByVal D As Date)
Dim sqlppn As String
Dim rsppn As New ADODB.Recordset
    
    sqlppn = "select rate from tax_cls where tax_code='PPN' and " & _
             "start_date <= '" & Format(D, "yyyymmdd") & "' and " & _
             "end_date >= '" & Format(D, "yyyymmdd") & "' "
    Set rsppn = Db.Execute(sqlppn)
    If Not (rsppn.BOF And rsppn.EOF) Then
        isippn = IIf(IsNull(rsppn(0)), 0, CDbl(rsppn(0)))
    Else
        isippn = 0
    End If
    Set rsppn = Nothing
End Sub

Sub FormatPrice()
Dim p1 As Byte, p2 As String, p0 As String
Dim jmldigit As Byte, jmldigit0 As Byte, j As Integer

    jmldigit = 0
    With grid
        For i = 1 To .Rows - 1
            If InStr(1, .TextMatrix(i, 17), ".") > 0 Then _
                jmldigit0 = Len(Trim(.TextMatrix(i, 17))) - InStr(1, Trim(.TextMatrix(i, 17)), ".")
            If jmldigit0 > jmldigit Then jmldigit = jmldigit0
        Next i

        For i = 1 To .Rows - 1
            p0 = Trim(.TextMatrix(i, 17))
            If InStr(1, p0, ".") > 0 Then
                p1 = Len(p0) - InStr(1, p0, ".")
                For j = 1 To jmldigit - p1
                    p2 = p0 & " "
                    p0 = p2
                Next j
            End If
            .TextMatrix(i, 17) = p0
        Next i
    End With
End Sub

Sub header()
    With grid
        .clear
        .Rows = 1
        .ColS = 23
        
        .ColWidth(0) = 300
        .ColWidth(1) = 1320
        .ColWidth(2) = 1300
        .ColWidth(3) = 1500
        .ColHidden(4) = True    'Department Cls
        .ColWidth(5) = 1215
        .ColHidden(6) = True    'Class Cls
        .ColWidth(7) = 1500
        .ColWidth(8) = 1770
        .ColWidth(9) = 1170
        .ColWidth(10) = 1680
        .ColWidth(11) = 1170
        .ColWidth(12) = 1125
        .ColHidden(13) = True   'Unit Cls
        .ColWidth(14) = 480
        .ColHidden(15) = True   'Currency Code
        .ColWidth(16) = 800
        .ColWidth(17) = 1500
        .ColWidth(18) = 1800
        .ColHidden(19) = True   'POReq SeqNo
        .ColHidden(20) = True   'Purpose
        .ColHidden(21) = True   'Seq No
        .ColWidth(22) = 1500
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Request No"
        .TextMatrix(0, 2) = "Product Code"
        .TextMatrix(0, 3) = "Product Name"
        .TextMatrix(0, 5) = "Department"
        .TextMatrix(0, 7) = "Class"
        .TextMatrix(0, 8) = "Req Delivery Date"
        .TextMatrix(0, 9) = "Request Qty"
        .TextMatrix(0, 10) = "PO Delivery Date"
        .TextMatrix(0, 11) = "Order Qty"
        .TextMatrix(0, 12) = "Remaining"
        .TextMatrix(0, 14) = "Unit"
        .TextMatrix(0, 16) = "Curr"
        .TextMatrix(0, 17) = "Price"
        .TextMatrix(0, 18) = "Amount"
        .TextMatrix(0, 22) = "Account No."
        
        .Cell(flexcpAlignment, 0, 0, 0, 22) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignLeftCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignCenterCenter
        .ColAlignment(11) = flexAlignRightCenter
        .ColAlignment(12) = flexAlignRightCenter
        .ColAlignment(14) = flexAlignCenterCenter
        .ColAlignment(16) = flexAlignCenterCenter
        .ColAlignment(17) = flexAlignRightCenter
        .ColAlignment(18) = flexAlignRightCenter
        .ColAlignment(22) = flexAlignCenterCenter
        
        .RowHeight(0) = 225
    End With
    sampun = False ' False -> tidak ada data di grid
End Sub

Sub browseitem()
Dim sqlitem As String, RsItem As New ADODB.Recordset
    
    lblErrMsg = ""
    Call header
    
    If ubah = False Then
        txtamount.Text = 0
        txtppn.Text = 0
        txtgrandtotal.Text = 0
    End If
    activecurr = ""
    activecurrcd = ""
    ubahgrid = False
    i = 1
    
    'Detail PO NO with Different POREQUEST NO
    sqlitem = "select a.*, (select description from unit_cls uc where uc.unit_cls= a.unit_cls ) unit_desc,  (select description from curr_cls where curr_cls.Curr_cls= a.Currency_Code) Curr_desc  From ( " & _
              "select distinct '1' No, pod.item_code, pod.unit_cls, pod.item_name, prd.accountno, " & _
              "isnull(prd.qty,0) RequestQty, " & _
              "isnull( (select sum(qty) qty from PurchaseOrder_Detail pod2 " & _
              "         inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "         where pod2.item_name = pod.item_name and pod2.porequest_no = pod.PORequest_No and pod2.POReq_SeqNo = pod.POReq_SeqNo " & _
              "         and isnull(pom.others_cls,'0')='1') ,0) totalPOQty, " & _
              "isnull(prd.qty,0) - isnull( (select sum(qty) qty from PurchaseOrder_Detail pod2 " & _
              "                             inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "                             where pod2.item_name = pod.item_name and pod2.porequest_no = pod.PORequest_No " & _
              "                             and pod2.POReq_SeqNo = pod.POReq_SeqNo and isnull(pom.others_cls,'0')='1') " & _
              ",0) RemainingQty, " & _
              "prm.Department_Cls, dc.Description Department, prd.ReqDelivery_Date, pod.PORequest_No, pod.POReq_SeqNo, pod.Purpose, " & _
              "pod.Currency_Code, isnull(pod.Price,0) Price, prd.Class " & _
              "from PurchaseOrder_Detail pod " & _
              "left outer join PORequest_Detail prd on prd.PORequest_No = pod.PORequest_No and prd.Seq_No = pod.POReq_SeqNo " & _
              "left outer join (select * from PORequest_Master where isnull(others_cls,'0') = '1') prm " & _
              "     on prm.porequest_no = prd.porequest_no " & _
              "left outer join Department_Cls dc on dc.Department_Cls = prm.Department_Cls " & _
              "where pod.PO_No = '" & Trim(txtpono.Text) & "' and pod.PORequest_No <> (case isnull(prm.Complete_Cls,'0') when '1' then '' else '" & Trim(cborequestno.Text) & "' end) "
    
    'From selected POREQUEST NO
    sqlitem = sqlitem & _
              "UNION " & _
              "select distinct '3' No, prd.item_code, prd.unit_cls, prd.item_name, prd.accountno, " & _
              "isnull(prd.qty,0) RequestQty, " & _
              "isnull( (select sum(qty) qty from PurchaseOrder_Detail pod " & _
              "         inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "         where pod.item_code = prd.item_code and pod.porequest_no = '" & Trim(cborequestno.Text) & "' " & _
              "         and pod.POReq_Seqno = prd.Seq_no and isnull(pom.others_cls,'0')='1') ,0) totalPOQty, " & _
              "isnull(prd.qty,0) - isnull( (select sum(qty) qty from PurchaseOrder_Detail pod " & _
              "                             inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "                             where pod.item_code = prd.item_code and pod.porequest_no = '" & Trim(cborequestno.Text) & "' " & _
              "                             and pod.POReq_Seqno = prd.Seq_no and isnull(pom.others_cls,'0')='1') " & _
              ",0) RemainingQty, " & _
              "prm.Department_Cls, dc.Description Department, prd.ReqDelivery_Date, prd.PORequest_No, prd.Seq_No, prd.Purpose, " & _
              "'' Currency_code, Null Price, prd.Class " & _
              "from PORequest_Detail prd " & _
              "inner join (select * from PORequest_Master where isnull(others_cls,'0') = '1') prm " & _
              "     on prm.porequest_no = prd.porequest_no " & _
              "left outer join Department_Cls dc on dc.Department_Cls = prm.Department_Cls " & _
              "where prd.porequest_no = '" & Trim(cborequestno.Text) & "' and isnull(prm.Complete_Cls,'0')='0' " & _
              ") a "

     Set RsItem = Db.Execute(sqlitem)
    If Not (RsItem.BOF And RsItem.EOF) Then
        With grid
        Do While Not RsItem.EOF
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, i, 0) = &HFFFFFF
            .Cell(flexcpChecked, i, 0) = flexUnchecked
            .TextMatrix(i, 1) = IIf(IsNull(RsItem("PORequest_No")), "", Trim(RsItem("PORequest_No")))
            .TextMatrix(i, 2) = IIf(IsNull(RsItem("item_code")), "", Trim(RsItem("item_code")))
            .TextMatrix(i, 3) = IIf(IsNull(RsItem("item_name")), "", Trim(RsItem("item_name")))
            .TextMatrix(i, 4) = IIf(IsNull(RsItem("Department_cls")), "", Trim(RsItem("department_Cls")))
            .TextMatrix(i, 5) = IIf(IsNull(RsItem("Department")), "", Trim(RsItem("Department")))
            If IsNull(RsItem("Class")) Then
                .TextMatrix(i, 6) = ""
                .TextMatrix(i, 7) = ""
            Else
                For j = 0 To UBound(Split(POClass1, ","))
                    If Trim(RsItem("Class")) = Split(POClass1, ",")(j) Then
                        .TextMatrix(i, 6) = Trim(RsItem("class"))
                        .TextMatrix(i, 7) = Split(POClass2, ",")(j)
                        Exit For
                    End If
                Next j
            End If
            .TextMatrix(i, 8) = IIf(IsNull(RsItem("ReqDelivery_Date")), "", Format(RsItem("ReqDelivery_Date"), "dd MMM yyyy"))
            .TextMatrix(i, 9) = IIf(IsNull(RsItem("RequestQty")), 0, Format(RsItem("RequestQty"), "#,##0.#0"))
            .TextMatrix(i, 10) = ""  'Delivery Date
            .Cell(flexcpBackColor, i, 10) = vbWhite
            .TextMatrix(i, 11) = 0   'Order
            .Cell(flexcpBackColor, i, 11) = &HFFFFFF
            .TextMatrix(i, 12) = IIf(IsNull(RsItem("RemainingQty")), 0, Format(RsItem("RemainingQty"), "#,##0.#0"))
            If IsNull(RsItem("unit_cls")) Then
              .TextMatrix(i, 13) = " "
              .TextMatrix(i, 14) = " "
            Else
              .TextMatrix(i, 13) = Trim(RsItem("Unit_cls"))
              '.TextMatrix(i, 14) = Split(isiunit, ",")(Val(Trim(RsItem("Unit_Cls"))) - 1)
              .TextMatrix(i, 14) = Trim(RsItem("Unit_desc"))
            End If
            If RsItem("No") = "3" Then
                .TextMatrix(i, 15) = ""
                .TextMatrix(i, 16) = ""
                .TextMatrix(i, 17) = Format(0, "#,##0.00###")
            Else
                If IsNull(RsItem("currency_code")) Then
                   .TextMatrix(i, 15) = ""
                   .TextMatrix(i, 16) = ""
                Else
                  .TextMatrix(i, 15) = Trim(RsItem("currency_code"))
                  '.TextMatrix(i, 16) = Split(isiCurr, ",")(Val(Trim(RsItem("Currency_code"))) - 1)
                  .TextMatrix(i, 16) = Trim(RsItem("Curr_desc"))
                End If
                .TextMatrix(i, 17) = IIf(IsNull(RsItem("Price")), 0, Format(Trim(RsItem("Price")), "##,##0.00###"))
            End If
            .Cell(flexcpBackColor, i, 16) = &HFFFFFF
            .Cell(flexcpBackColor, i, 17) = &HFFFFFF
            .Cell(flexcpFontName, i, 17) = "Courier New"
            .TextMatrix(i, 18) = 0  'Amount
            .TextMatrix(i, 19) = IIf(IsNull(RsItem("POReq_SeqNo")), "", RsItem("POReq_SeqNo"))
            .TextMatrix(i, 20) = IIf(IsNull(RsItem("Purpose")), "", Trim(RsItem("Purpose")))
            .TextMatrix(i, 21) = ""
            .TextMatrix(i, 22) = IIf(IsNull(RsItem("AccountNo")), "", Trim(RsItem("AccountNo")))
            
            RsItem.MoveNext
            i = i + 1
        Loop
        End With
    Else
        lblErrMsg = DisplayMsg(4006)
    End If
    Set RsItem = Nothing
End Sub

Sub Browse()
Dim a As Double

    lblErrMsg = ""
    
    Sql = "select * from PurchaseOrder_Master where PO_No = '" & txtpono.Text & "' and others_cls = '1' and period is null"
    If rs.State <> adStateClosed Then rs.Close
    rs.Open Sql, Db, adOpenKeyset, adLockOptimistic
    
    If Not (rs.BOF And rs.EOF) Then
        ada = True: ubah = True
        statusfix = IIf(IsNull(rs("fix_cls")), 0, rs("fix_cls"))
        Call browseitem
        Call BrowseGrid
        Call FormatPrice
        
        'Count TOTAL AMOUNT
        For i = 1 To grid.Rows - 1
            If grid.Cell(flexcpChecked, i, 0) = flexChecked Then _
                a = a + grid.TextMatrix(i, 18)
        Next i
        txtamount.Text = a
        If (txtamount.Text <> 0) Then txtamount.Text = Format(txtamount.Text, "##,##0.#0")
        If ((cbocust.Column(3) = 1) Or (cbocust.Column(3) = 2) Or (cbocust.Column(3) = 3) Or (cbocust.Column(3) = 5)) Then
            txtppn = 0
        Else
            txtppn.Text = CDbl(isippn / 100) * CDbl(txtamount.Text)
        End If
        If (txtppn.Text <> 0) Then txtppn.Text = Format(txtppn.Text, "##,##0.#0")
        txtgrandtotal = CDbl(txtppn.Text) + CDbl(txtamount.Text)
        If (txtgrandtotal.Text <> 0) Then txtgrandtotal.Text = Format(txtgrandtotal.Text, "##,##0.#0")
        
        If statusfix = 1 Then Call Kunci(True) Else Call Kunci(False)
    Else
        ada = False
    End If
End Sub

Sub BrowseGrid()
Dim g As Integer
    
    sqlGrid = "select *, (select description from unit_cls uc where uc.unit_cls= PurchaseOrder_Detail.unit_cls ) unit_desc,  (select description from curr_cls where curr_cls.Curr_cls= PurchaseOrder_Detail .Currency_Code) Curr_desc from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "' order by item_name"
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic

    With grid
    Do While Not rsGrid.EOF
        For g = 1 To .Rows - 1
            If Trim(.TextMatrix(g, 3)) = Trim(rsGrid("Item_Name")) And Trim(.TextMatrix(g, 1)) = Trim(rsGrid("PORequest_No")) And Trim(.TextMatrix(g, 19)) = CStr(rsGrid("POReq_SeqNo")) Then
                ubahgrid = True
                .Cell(flexcpChecked, g, 0) = flexChecked
                .TextMatrix(g, 10) = IIf(IsNull(rsGrid("Delivery_Date")), "", Format(rsGrid("Delivery_Date"), "dd MMM yyyy"))
                .TextMatrix(g, 11) = IIf(IsNull(rsGrid("qty")), 0, Format(Trim(rsGrid("qty")), "##,##0.#0"))
                .TextMatrix(g, 13) = Trim(rsGrid("Unit_cls"))
                '.TextMatrix(g, 14) = Split(isiunit, ",")(Val(Trim(rsGrid("Unit_Cls"))) - 1)
                .TextMatrix(g, 14) = Trim(rsGrid("Unit_desc"))
                If IsNull(rsGrid("currency_code")) Then
                    .TextMatrix(g, 15) = ""
                    .TextMatrix(g, 16) = ""
                Else
                    .TextMatrix(g, 15) = Trim(rsGrid("currency_code"))
                    '.TextMatrix(g, 16) = Split(isiCurr, ",")(Val(Trim(rsGrid("Currency_code"))) - 1)
                    .TextMatrix(g, 16) = Trim(rsGrid("Curr_desc"))
                End If
                activecurrcd = .TextMatrix(g, 15)
                activecurr = .TextMatrix(g, 16)
                .TextMatrix(g, 17) = IIf(IsNull(rsGrid("Price")), 0, Format(Trim(rsGrid("Price")), "##,##0.00###"))
                .TextMatrix(g, 18) = IIf(IsNull(rsGrid("Amount")), 0, Format(rsGrid("Amount"), "##,##0.#0"))
                .TextMatrix(g, 21) = rsGrid("Seq_no")
            End If
        Next g
        rsGrid.MoveNext
    Loop
    End With
End Sub

Sub BrowseAtas()
    Sql = "select * from PurchaseOrder_Master where PO_No = '" & Trim(txtpono.Text) & "' and isnull(others_cls,'0') = '1' and period is null"
    If rs.State <> adStateClosed Then rs.Close
    rs.Open Sql, Db, adOpenKeyset, adLockOptimistic

    If Not (rs.BOF And rs.EOF) Then
        podate.Value = IIf(IsNull(rs("po_date")), "", Format(Trim(rs("po_date")), "dd MMM yyyy"))
        cbocust.Text = Trim(rs("Supplier_code"))
        txtRev.Text = IIf(IsNull(rs("revisi")), "", Trim(rs("Revisi")))
        cbopricecondition.Text = IIf(IsNull(rs("PriceCondition_Cls")), "", Trim(rs("PriceCondition_Cls")))
        cbopocode.ListIndex = IIf(IsNull(rs("POPayment_Code")), -1, rs("POPayment_Code") - 1)
        txtpoday.Text = IIf(IsNull(rs("popayment_days")), "", rs("popayment_days"))
        cbopoterms.ListIndex = IIf(IsNull(rs("POPayment_Terms")), -1, rs("POPayment_Terms") - 1)
        txtRemarks(0).Text = IIf(IsNull(rs("remarks")), "", Trim(rs("remarks")))
        txtRemarks(1).Text = IIf(IsNull(rs("remarks2")), "", Trim(rs("remarks2")))
       txtRemarks(2).Text = IIf(IsNull(rs("remarks3")), "", Trim(rs("remarks3")))
        cbodelmode.ListIndex = IIf(IsNull(rs("transportation_Cls")), -1, rs("transportation_Cls"))
        statusfix = IIf(IsNull(rs("fix_cls")), 0, rs("fix_cls"))
        If statusfix = 1 Then Call Kunci(True) Else Call Kunci(False)
    End If
End Sub

Function cekrecqty(ByVal seqNo As String, ByVal poNo As String) As Double
Dim sqlcek As String, rsCek As New Recordset
    
    cekrecqty = 0
    sqlcek = "select dailyseq_no, sum(qty) recqty from Part_Receipt " & _
             "where PO_No = '" & Trim(poNo) & "' and dailyseq_no = '" & IIf(Trim(seqNo) = "", 0, Trim(seqNo)) & "' " & _
             "group by dailyseq_no"
    If rsCek.State <> adStateClosed Then rsCek.Close
    rsCek.Open sqlcek, Db, adOpenKeyset, adLockOptimistic
    If Not (rsCek.BOF And rsCek.EOF) Then _
        cekrecqty = CDbl(rsCek("recqty"))
    Set rsCek = Nothing
End Function

Private Sub Form_Load()
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

    Call adtocbocust(False)
    Call adtocombo
    Call kosong
    combo1.ListIndex = 1
End Sub

Private Sub requestdate1_Change()
Dim ketemu As Boolean

    lblErrMsg.Caption = ""
    If Format(requestdate1, "yyyy-mm-dd") > Format(requestdate2, "yyyy-mm-dd") Then
       lblErrMsg.Caption = DisplayMsg(4025) & " " & Format(requestdate2, "MMM yyyy") '"Start Date must be lower than "
       Exit Sub
    End If
    
    Call adtocborequestno
        
    If combo1.ListIndex = 1 Then    'UPDATE
        If cbocust.Text <> "" Then Call adtocbopono
        For i = 0 To cbopono.ListCount - 1
            If txtpono.Text = cbopono.List(i) Then
                ketemu = True
                cbopono.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtpono.Text = ""
    End If
    
'    txtpono.Text = ""
'    cbopono.clear
    Call header
    Call kosongBwh
End Sub

Private Sub requestdate2_Change()
Dim ketemu As Boolean

    lblErrMsg.Caption = ""
    If Format(requestdate2, "yyyy-mm-01") < Format(requestdate1, "yyyy-mm-01") Then
       lblErrMsg.Caption = DisplayMsg(4024) & " " & Format(requestdate1, "MMM yyyy") '"End Date must be higher than "
       Exit Sub
    End If

    Call adtocborequestno
    
    If combo1.ListIndex = 1 Then    'UPDATE
        If cbocust.Text <> "" Then Call adtocbopono
        For i = 0 To cbopono.ListCount - 1
            If txtpono.Text = cbopono.List(i) Then
                ketemu = True
                cbopono.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtpono.Text = ""
    End If

'    txtpono.Text = ""
'    cbopono.clear
    Call header
    Call kosongBwh
End Sub

Private Sub cborequestno_Click()
Dim ketemu As Boolean, tempcust As String

    lblErrMsg = ""
    If cborequestno.ListIndex <> -1 Then
        If Trim(cbocust.Text) <> "" Then tempcust = Trim(cbocust.Text)
        If cborequestno.Text <> "" Then
            Call adtocbocust(True)
            Call adtocbopono
        Else
            Call adtocbocust(False)
        End If
        cbocust.Text = tempcust: Call cbocust_Click
        
        If combo1.ListIndex = 1 Then    'UPDATE
'            If cbocust.Text <> "" Then Call adtocbopono
'            For i = 0 To cbopono.ListCount - 1
'                If txtpono.Text = cbopono.List(i) Then
'                    ketemu = True
'                    cbopono.ListIndex = i
'                    Exit For
'                End If
'            Next i
'            If ketemu = False Then txtpono.Text = ""
            Call kosongBwh: Call header
        End If
    Else
'        cbopono.clear
        If Trim(cbocust.Text) <> "" Then tempcust = Trim(cbocust.Text)
        If cborequestno.Text <> "" Then Call adtocbocust(True) Else Call adtocbocust(False)
        cbocust.Text = tempcust: Call cbocust_Click
        
        If combo1.ListIndex = 1 Then 'UPDATE
            Call kosongBwh
            Call header
'            txtpono.Text = ""
        End If
        If cborequestno.Text <> "" Then lblErrMsg.Caption = DisplayMsg(4144) '"Record with this Request No not found"
    End If
End Sub

Private Sub cborequestno_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cborequestno_Click
End Sub

Private Sub cborequestno_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbocust_Click()
Dim ketemu As Boolean

    lblErrMsg = ""
    ketemu = False
    Call Kunci(False)

    If cbocust.ListIndex <> -1 Then
        lblcust(0).Text = cbocust.Column(1)
        lblcust(1).Text = cbocust.Column(2)
        countrycls = cbocust.Column(3)
        If Split(cbocust.Column(5), ",")(0) = "" Then cbopocode.ListIndex = -1 Else cbopocode.ListIndex = Split(cbocust.Column(5), ",")(0) - 1
        txtpoday.Text = Split(cbocust.Column(5), ",")(1)
        If Split(cbocust.Column(5), ",")(2) = "" Then cbopoterms.ListIndex = -1 Else cbopoterms.ListIndex = Split(cbocust.Column(5), ",")(2) - 1
        If cbocust.Column(6) = "" Then cbodelmode.ListIndex = -1 Else cbodelmode.ListIndex = cbocust.Column(6)
        
        If combo1.ListIndex = 1 Then    'UPDATE
            Call ClearData
            'Call adtocbopono
            For i = 0 To cbopono.ListCount - 1
                If txtpono.Text = cbopono.List(i) Then
                    ketemu = True
                    'cbopono.ListIndex = i
                    Exit For
                End If
            Next i
            If ketemu = False Then txtpono.Text = ""
            Call kosongBwh
            Call header
        End If
    Else
        lblcust(0).Text = ""
        lblcust(1).Text = ""
        countrycls = 3
        cbopocode.ListIndex = -1
        txtpoday.Text = ""
        cbopoterms.ListIndex = -1
        cbodelmode.ListIndex = -1
        'cbopono.clear
        If combo1.ListIndex = 1 Then 'UPDATE
            Call kosongBwh
            Call header
            txtpono.Text = ""
        End If
        If cbocust.Text <> "" Then lblErrMsg.Caption = DisplayMsg(4050) '"Record with this Supplier Code not Exist"
        Exit Sub
    End If
        
    If (countrycls = 1 Or countrycls = 2) Then  'OVERSEAS
        isippn = 0
        txtppn.Text = 0
        txtgrandtotal.Text = txtamount.Text
        If (txtgrandtotal.Text <> 0) Then txtgrandtotal.Text = Format(txtgrandtotal.Text, "##,##0.#0")
    Else 'DOMESTIC
        Call ppn(podate.Value)
        txtppn.Text = CDbl(isippn / 100) * CDbl(txtamount.Text)
        If (txtppn.Text <> 0) Then txtppn.Text = Format(txtppn.Text, "##,##0.#0")
        txtgrandtotal = CDbl(txtppn.Text) + CDbl(txtamount.Text)
        If (txtgrandtotal.Text <> 0) Then txtgrandtotal.Text = Format(txtgrandtotal.Text, "##,##0.#0")
    End If
End Sub

Private Sub cbocust_LostFocus()
    Call cbocust_Click
End Sub

Private Sub cbocust_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbocust_Click
End Sub

Private Sub cbocust_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Combo1_Click()
Dim ketemu As Boolean

    lblErrMsg = ""
    ketemu = False
    Call Kunci(False)
    Call kosongBwh
    Call header

    If combo1.ListIndex = 0 Then    'CREATE
        Call ClearData
        command1(2).Caption = "&Create"
        ubah = False
        cbopono.locked = True
        txtpono.Text = ""
        podate.Value = Format(Now, "dd MMM yyyy")
        podate.Enabled = False
        Call poNo(Right(year(podate), 2), Format(month(podate), "0#"))
        cbopricecondition.ListIndex = -1
        txtRemarks(0).Text = "": txtRemarks(1).Text = "": txtRemarks(2).Text = ""
    Else    'UPDATE
        If cbocust.Text = "" Then   'Or cborequestno.Text = ""
            cbopono.clear
            txtpono.Text = ""
        Else
            Call adtocbopono
        End If
        ubah = True
        command1(2).Caption = "&Update"
        cbopono.locked = False
        txtpono.locked = False
        podate.Enabled = True

        For i = 0 To cbopono.ListCount - 1
            If txtpono.Text = cbopono.List(i) Then
                ketemu = True
                cbopono.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtpono.Text = ""
    End If
End Sub

Private Sub combo1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call Combo1_Click
End Sub

Private Sub cbopono_Click()
    lblErrMsg = ""
    txtpono.Text = cbopono.Text
    Call header
    Call kosongBwh
    Call BrowseAtas
End Sub

Private Sub cbopono_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbopono_Click
End Sub

Private Sub txtpono_Change()
Dim ketemu As Boolean

    txtpono2.Text = txtpono.Text
    If combo1.ListIndex = 1 Then
        For i = 0 To cbopono.ListCount - 1
            If txtpono.Text = cbopono.List(i) Then
                ketemu = True
                cbopono.ListIndex = i
                Exit For
            End If
        Next
        If ketemu = False Then cbopono.ListIndex = -1
    End If
End Sub

Private Sub txtpono_KeyPress(KeyAscii As Integer)
    lblErrMsg = ""
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then
        If combo1.ListIndex = 0 Then
            SendKeys vbTab
        Else
            Call header
            Call kosongBwh
            Call BrowseAtas
        End If
    End If
End Sub

Private Sub podate_Change()
    lblErrMsg = ""
    'CREATE
    If combo1.ListIndex = 0 Then _
        Call poNo(Right(year(podate), 2), Format(month(podate), "0#"))
    If (countrycls = 1 Or countrycls = 2) Then isippn = 0 Else Call ppn(podate.Value)
End Sub

Private Sub deldate_Change()
    lblErrMsg = ""
    grid.TextMatrix(actrow, 10) = Format(DelDate, "dd mmm yyyy")
End Sub

Private Sub deldate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub deldate_LostFocus()
    DelDate.Visible = False
End Sub

Private Sub CboCurr_Click()
    If cbocurr.ListIndex <> -1 Then
        grid.TextMatrix(actrow, 15) = cbocurr.Column(0)
        grid.TextMatrix(actrow, 16) = cbocurr.Column(1)
    End If
End Sub

Private Sub cbocurr_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call CboCurr_Click
End Sub

Private Sub cbocurr_LostFocus()
    cbocurr.Visible = False
End Sub

Private Sub grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim a As Double

    a = 0
    With grid
        If Col = 11 Then 'ORDER QTY
            If .TextMatrix(Row, 11) = "" Then .TextMatrix(Row, 11) = 0
            If IsNumeric(.TextMatrix(Row, 11)) = False Then .TextMatrix(Row, 11) = 0
            If CDbl(.TextMatrix(Row, 11)) > 9999999.99 Then lblErrMsg = DisplayMsg(4045) & " 9,999,999.99": .SetFocus: Exit Sub  '"Quantity must be lower or equal than 9,999,999.99"
            
            .TextMatrix(Row, 11) = Format(.TextMatrix(Row, 11), "#,##0.#0")
            .TextMatrix(Row, 12) = Format((CDbl(.TextMatrix(Row, 12)) - CDbl(.TextMatrix(Row, 11)) + orderawal), "##,##0.#0")
        End If
        
        If Col = 17 Then 'PRICE
            If .TextMatrix(Row, 17) = "" Then .TextMatrix(Row, 17) = 0
            If IsNumeric(.TextMatrix(Row, 17)) = False Then .TextMatrix(Row, 17) = 0
            If CDbl(.TextMatrix(Row, 17)) > 9999999999.99999 Then lblErrMsg = DisplayMsg(4048) & " 9,999,999,999.999999": .SetFocus: Exit Sub  '"Price must be lower or equal than 9,999,999,999.99999"
            .TextMatrix(Row, 17) = Format(.TextMatrix(Row, 17), "#,##0.00###")
        End If
        
        If Col = 0 Or Col = 11 Or Col = 17 Then
            Call FormatPrice
            .TextMatrix(Row, 18) = CDbl(.TextMatrix(Row, 11)) * CDbl(.TextMatrix(Row, 17))
            If .TextMatrix(Row, 18) <> 0 Then .TextMatrix(Row, 18) = Format(.TextMatrix(Row, 18), "##,##0.#0")
            
            'Count TOTAL
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, 0) = flexChecked Then
                    a = a + .TextMatrix(i, 18)
                End If
            Next i
            txtamount.Text = a
            If (txtamount.Text <> 0) Then txtamount.Text = Format(txtamount.Text, "##,##0.#0")
            If ((cbocust.Column(3) = 1) Or (cbocust.Column(3) = 2) Or (cbocust.Column(3) = 3) Or (cbocust.Column(3) = 5)) Then
                txtppn = 0
            Else
                txtppn.Text = CDbl(isippn / 100) * CDbl(txtamount.Text)
            End If
            If (txtppn.Text <> 0) Then txtppn.Text = Format(txtppn.Text, "##,##0.#0")
            txtgrandtotal = CDbl(txtppn.Text) + CDbl(txtamount.Text)
            If (txtgrandtotal.Text <> 0) Then txtgrandtotal.Text = Format(txtgrandtotal.Text, "##,##0.#0")
        End If
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    actrow = Row
    If grid.Cell(flexcpChecked, Row, 0) <> flexChecked Then
        If Col <> 0 Then Cancel = True
    Else
        If Col <> 0 And Col <> 11 And Col <> 10 And Col <> 16 And Col <> 17 Then _
            Cancel = True
'        If cborequestno.ListIndex <> -1 Then _
'            If Col = 11 And cborequestno.Column(2) = "1" Then lblErrMsg = DisplayMsg(4089): Cancel = True
        If Col = 11 Then orderawal = CDbl(grid.TextMatrix(Row, 11))
    End If
End Sub

Private Sub Grid_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 11 Or Col = 17 Then _
        If InStr(1, grid.TextMatrix(Row, Col), ",") = 1 Then grid.TextMatrix(Row, Col) = Right(grid.TextMatrix(Row, Col), Len(grid.TextMatrix(Row, Col)) - 1)
End Sub

Private Sub Grid_Click()
    With grid
    If statusfix = 0 Then
        If .Row > 0 Then
            If .Cell(flexcpChecked, .Row, 0) = flexChecked Then
                If .Col = 11 Or .Col = 10 Or .Col = 16 Or .Col = 17 Then
                    .SelectionMode = flexSelectionFree
                Else
                    .SelectionMode = flexSelectionByRow
                End If
                
                If .Col = 10 Then   'DELIVERY DATE
                    DelDate.top = .Cell(flexcpTop, .Row, 10)
                    DelDate.Left = .Cell(flexcpLeft, .Row, 10)
                    DelDate.Width = .CellWidth + 30
                    DelDate.Height = .CellHeight + 30
                    If .TextMatrix(.Row, 10) <> "" Then
                        DelDate.Value = Format(.TextMatrix(.Row, 10), "yyyy-mm-dd")
                    Else
                        If Trim(.TextMatrix(.Row, 8)) <> "" Then
                            DelDate.Value = Format(.TextMatrix(.Row, 8), "yyyy-mm-dd")
                        End If
                        .TextMatrix(.Row, 10) = Format(DelDate.Value, "dd MMM yyyy")
                    End If
                    DelDate.Visible = True
                    DelDate.SetFocus
                    cbocurr.Visible = False
                ElseIf .Col = 16 Then   'CURRENCY
                    cbocurr.top = .Cell(flexcpTop, .Row, 16)
                    cbocurr.Left = .Cell(flexcpLeft, .Row, 16)
                    cbocurr.Width = .CellWidth + 30
                    Call up_FillCombo(cbocurr, "curr_cls")
                    'Call isiCboUnitCurr(cbocurr, isiCurr, 0, 4)
                    cbocurr.TextColumn = 2
                    If grid.TextMatrix(.Row, 16) <> "" Then
                        cbocurr.Text = grid.TextMatrix(.Row, 16)
                        For i = 0 To cbocurr.ListCount - 1
                            If Trim(grid.TextMatrix(.Row, 15)) = Trim(cbocurr.List(i)) Then
                                cbocurr.ListIndex = i
                                Exit For
                            End If
                        Next i
                    End If
                    cbocurr.Visible = True
                    cbocurr.SetFocus
                    DelDate.Visible = False
                Else
                    cbocurr.Visible = False
                    DelDate.Visible = False
                End If
                
                If .Col = 11 Then .FocusRect = flexFocusInset Else .FocusRect = flexFocusNone
            Else
                .SelectionMode = flexSelectionByRow
            End If
        End If
    End If
    End With
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    lblErrMsg = ""
    If Col = 11 Then
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
            KeyAscii = 0
    ElseIf Col = 17 Then
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
              KeyAscii = 0
'        If InStr(1, grid.TextMatrix(Row, 17), ".") > 1 Then _
'            If KeyAscii = Asc(".") Then KeyAscii = 0
    End If
End Sub

Private Sub grid_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    cbocurr.Visible = False
    DelDate.Visible = False
End Sub

Private Sub grid_AfterSort(ByVal Col As Long, Order As Integer)
    cbocurr.Visible = False
    DelDate.Visible = False
End Sub

Private Sub Grid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call Grid_Click
End Sub

Private Sub cbopricecondition_Click()
    If cbopricecondition.ListIndex <> -1 Then
        lblpricecondition.Caption = cbopricecondition.Column(1)
    Else
        lblpricecondition.Caption = ""
    End If
End Sub

Private Sub cbopricecondition_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbopricecondition_Click
End Sub

Private Sub cboPriceCondition_Change()
    If cbopricecondition.Text = "" Then lblpricecondition.Caption = ""
End Sub

Private Sub txtpoday_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then _
        KeyAscii = 0
    If KeyAscii = Asc("'") Or KeyAscii = Asc(".") Then KeyAscii = 0
End Sub

Private Sub txtpoday_LostFocus()
    txtpoday.Text = Val(txtpoday.Text)
    If txtpoday.Text = 0 Then txtpoday.Text = ""
End Sub

Private Sub hscrollbar_Change()
    Dim k As Integer
    For k = 4 To grid.ColS - 1
       grid.ColHidden(k) = False
    Next k
    If hscrollbar.Value = 1 Then
        For k = 4 To 14
            grid.ColHidden(k) = True
        Next k
    End If
    cbocurr.Visible = False
    DelDate.Visible = False
    grid.ColHidden(4) = True    'Department Cls
    grid.ColHidden(6) = True    'Class Cls
    grid.ColHidden(13) = True   'Unit Cls
    grid.ColHidden(15) = True   'Currency Code
    grid.ColHidden(19) = True   'POReq SeqNo
    grid.ColHidden(20) = True   'Purpose
    grid.ColHidden(21) = True   'Seq No
End Sub

Private Sub hscrollbar_Scroll()
    Call hscrollbar_Change
End Sub

Private Sub Command1_Click(Index As Integer)
Dim sql4 As String
Dim rs4 As New Recordset

lblErrMsg = ""

Select Case Index
    Case 0: 'SUBMIT
            If hakUpdate(Me.Name) = 0 Then _
                lblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
            
            'HEADER VALIDATION
            If cborequestno.Text = "" Then
'                cborequestno.SetFocus
'                lblErrMsg = DisplayMsg(1067) '"Please Input Request No"
'                Exit Sub
            ElseIf cborequestno.Text <> "" Then
                If cborequestno.MatchFound = False Then
                    lblErrMsg = DisplayMsg(4144)    'Record with This Request No Not found !
                    cborequestno.SetFocus
                    Exit Sub
                End If
            End If
            If cbocust.Text = "" Then
                cbocust.SetFocus
                lblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
                Exit Sub
            ElseIf cbocust.Text <> "" Then
                If cbocust.MatchFound = False Then
                    lblErrMsg = DisplayMsg(4050)    'Record with This Supplier Code Not found !
                    cbocust.SetFocus
                    Exit Sub
                End If
            End If
            If txtpono.Text = "" Then
                txtpono.SetFocus
                lblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                Exit Sub
            End If
                            
            '-------------------------------------------------------------
            
            'FOOTER VALIDATION
            If cbopricecondition.Text <> "" Then
                If cbopricecondition.MatchFound = False Then
                    lblErrMsg = DisplayMsg(4147)    'Record with This Price Condition not found !
                    cbopricecondition.SetFocus
                    Exit Sub
                End If
            End If
            If cbopocode.Text <> "" Or txtpoday.Text <> "" Or cbopoterms.Text <> "" Then
                If cbopocode.Text = "" Then
                    cbopocode.SetFocus
                    lblErrMsg = DisplayMsg(1062)
                    Exit Sub
                ElseIf txtpoday.Text = "" Or Val(txtpoday.Text) = 0 Then
                    txtpoday.SetFocus
                    lblErrMsg = DisplayMsg(1062)
                    Exit Sub
                ElseIf cbopoterms.Text = "" Then
                    cbopoterms.SetFocus
                    lblErrMsg = DisplayMsg(1062)
                    Exit Sub
                End If
                If cbopocode.Text <> "" Then
                    If cbopocode.MatchFound = False Then
                        lblErrMsg = DisplayMsg(4097)
                        cbopocode.SetFocus
                        Exit Sub
                    End If
                End If
                If cbopoterms.Text <> "" Then
                    If cbopoterms.MatchFound = False Then
                        lblErrMsg = DisplayMsg(4097)
                        cbopoterms.SetFocus
                        Exit Sub
                    End If
                End If
            End If
            If cbodelmode.Text <> "" Then
                If cbodelmode.MatchFound = False Then
                    lblErrMsg = DisplayMsg(4099)    '"Record with this Transportation not found"
                    cbodelmode.SetFocus
                    Exit Sub
                End If
            End If
            '-------------------------------------------------------------
            
'            Sql = "select * from PurchaseOrder_Master where PO_No = '" & txtpono.Text & "' and others_cls = '1' and period is null"
'            If rs.State <> adStateClosed Then rs.Close
'            rs.Open Sql, Db, adOpenKeyset, adLockOptimistic
'            If rs.BOF And rs.EOF Then
'                LblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
'                txtpono.SetFocus
'                Exit Sub
'            End If
            
            If grid.Rows = 1 Then lblErrMsg = DisplayMsg(5012): Exit Sub  'There is no data to submit !
            
            If ubah = True Then
'                rs("po_date") = Format(podate.Value, "yyyy-mm-dd")
'                rs("amount") = CDbl(txtAmount.Text)
'                rs("ppn") = CDbl(txtppn.Text)
'                rs("total_amount") = CDbl(txtgrandtotal.Text)
'                rs("remarks") = Trim(txtRemarks(0).Text)
'                rs("remarks2") = Trim(txtRemarks(1).Text)
'                rs("remarks3") = Trim(txtRemarks(2).Text)
'                rs("revisi") = Trim(txtRev.Text)
'                If cbopocode.Text = "" Then rs("popayment_code") = Null Else rs("popayment_code") = cbopocode.Column(0)
'                If txtpoday.Text = "" Then rs("popayment_days") = Null Else rs("popayment_days") = txtpoday.Text
'                If cbopoterms.Text = "" Then rs("popayment_terms") = Null Else rs("popayment_terms") = cbopoterms.Column(0)
'                If cbodelmode.Text = "" Then rs("transportation_cls") = Null Else rs("transportation_cls") = cbodelmode.Column(0)
'                If cbopricecondition.Text = "" Then rs("pricecondition_cls") = Null Else rs("pricecondition_cls") = cbopricecondition.Text
'                rs.update

                Dim recqty As Double
                                
                'DETAIL VALIDATION
                With grid
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = flexChecked Then
                            If .TextMatrix(i, 11) = 0 Then
                                hscrollbar.Value = 0
                                .SelectionMode = flexSelectionFree
                                .Col = 11: .Row = i: actrow = i
                                .SetFocus
                                lblErrMsg = DisplayMsg(1012) '"Please Input Quantity"
                                Exit Sub
                            ElseIf CDbl(.TextMatrix(i, 11)) > 9999999.99 Then
                                hscrollbar.Value = 0
                                .SelectionMode = flexSelectionFree
                                .Col = 11: .Row = i: actrow = i
                                .SetFocus
                                lblErrMsg = DisplayMsg(4045) & " 9,999,999.99" '"Quantity must be lower or equal than 9,999,999.99"
                                Exit Sub
                            ElseIf CDbl(.TextMatrix(i, 12)) < 0 Then
                                hscrollbar.Value = 0
                                .SelectionMode = flexSelectionFree
                                .Col = 11: .Row = i: actrow = i
                                .SetFocus
                                lblErrMsg = DisplayMsg(4045) & " Qty Remaining" '"Quantity must be lower or equal than Qty Remaining"
                                Exit Sub
                            ElseIf .TextMatrix(i, 10) = "" Then
                                hscrollbar.Value = 0
                                .SelectionMode = flexSelectionFree
                                .Col = 10: .Row = i: actrow = i
                                .SetFocus
                                Call Grid_Click
                                lblErrMsg = DisplayMsg(8124)    'Please Input Delivery Date
                                Exit Sub
                            ElseIf .TextMatrix(i, 16) = "" Then
                                hscrollbar.Value = 1
                                .SelectionMode = flexSelectionFree
                                .Col = 16: .Row = i: actrow = i
                                .SetFocus
                                Call Grid_Click
                                lblErrMsg = DisplayMsg(1028)    'Please Select Currency
                                Exit Sub
                            ElseIf .TextMatrix(i, 17) = 0 Then
                                hscrollbar.Value = 1
                                .SelectionMode = flexSelectionFree
                                .Col = 17: .Row = i: actrow = i
                                .SetFocus
                                Call Grid_Click
                                lblErrMsg = DisplayMsg(1029) '"Please Input Price"
                                Exit Sub
                            End If
                            
                            recqty = cekrecqty(.TextMatrix(i, 21), txtpono.Text)
                            If CDbl(.TextMatrix(i, 11)) < recqty Then
                                hscrollbar.Value = 0
                                .SelectionMode = flexSelectionFree
                                .Col = 11: .Row = i: actrow = i
                                .SetFocus
                                lblErrMsg = DisplayMsg(4036) & " " & recqty     '"Quantity must be higher or equal than "
                                Exit Sub
                            End If
                        Else
                            sql4 = "select * from Part_Receipt where po_no = '" & txtpono.Text & "' and dailyseq_no = '" & IIf(.TextMatrix(i, 21) = "", 0, .TextMatrix(i, 21)) & "'"
                            Set rs4 = Db.Execute(sql4)
                            If Not (rs4.BOF And rs4.EOF) Then
                                .Row = i: actrow = i
                                .SetFocus
                                lblErrMsg = DisplayMsg(1204)
                                Exit Sub
                            End If
                            Set rs4 = Nothing
                        End If
                    Next i
                                                                                                    
                    ' Cek currency harus sama
                    Dim a As Integer, C As Integer
                    If ubahgrid = False Or activecurr = "" Then  '*****
                        For a = 1 To .Rows - 1
                            If .Cell(flexcpChecked, a, 0) = flexChecked Then
                                activecurr = .TextMatrix(a, 16): activecurrcd = .TextMatrix(a, 15): Exit For
                            End If
                        Next a
                    End If
                    
                    Dim barisan As Integer, indexawal As Integer, barisawal As Integer, jumlahBarisan As Integer
                    indexawal = 0
                    jumlahBarisan = 0
                    For barisan = 1 To .Rows - 1
                        If .Cell(flexcpChecked, barisan, 0) = flexChecked Then
                                If indexawal = 0 Then barisawal = barisan: indexawal = 1
                                jumlahBarisan = jumlahBarisan + 1
                        End If
                    Next barisan
                    
                    For barisan = 1 To .Rows - 1
                        If .Cell(flexcpChecked, barisan, 0) = flexChecked Then
                        Else
                            .TextMatrix(barisan, 15) = ""
                        End If
                    Next barisan
                    
            If jumlahBarisan > 1 Then
                    For C = a + 1 To .Rows - 1
                        If .Cell(flexcpChecked, C, 0) = flexChecked Then
                            If .TextMatrix(C, 15) <> .TextMatrix(barisawal, 15) Then  'tadinya <> activecurr
                                hscrollbar.Value = 1
                                .Col = 15: .Row = C: actrow = C
                                .SetFocus
                                Call Grid_Click
                                lblErrMsg = DisplayMsg(4084) 'Cannot Select Different Currency !!
                                Exit Sub
                            End If
                        End If
                    Next C
            End If

                                                            
                    Dim rscekC As New Recordset
                    
                    'UPDATE DETAIL
                   'Set rsGrid = Nothing
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = flexChecked Then
                            If Trim(.TextMatrix(i, 21)) = "" Then
                                sqlGrid = "select * From PurchaseOrder_Detail"
                                If rsGrid.State <> adStateClosed Then rsGrid.Close
                                rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
                                rsGrid.AddNew
                                rsGrid("Seq_No") = seqNo
                            Else
                                sqlGrid = "select * From PurchaseOrder_Detail where seq_no = '" & .TextMatrix(i, 21) & "' "
                                If rsGrid.State <> adStateClosed Then rsGrid.Close
                                rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
                            End If
                            rsGrid("po_no") = Trim(txtpono.Text)
                            rsGrid("PORequest_No") = Trim(.TextMatrix(i, 1))
                            rsGrid("item_Code") = Trim(.TextMatrix(i, 2))
                            rsGrid("item_name") = Trim(.TextMatrix(i, 3))
                            rsGrid("class") = Trim(.TextMatrix(i, 6))
                            rsGrid("POReq_SeqNo") = .TextMatrix(i, 19)
                            rsGrid("Delivery_Date") = Format(.TextMatrix(i, 10), "yyyy-mm-dd")
                            rsGrid("price") = CDbl(.TextMatrix(i, 17))
                            rsGrid("currency_code") = .TextMatrix(i, 15)
                            rsGrid("unit_cls") = .TextMatrix(i, 13)
                            rsGrid("qty") = CDbl(.TextMatrix(i, 11))
                            ' TAMBAH VAR UTK AMO ============================================================================
                            rsGrid("amount") = CDbl(.TextMatrix(i, 18))
                            rsGrid("purpose") = Trim(.TextMatrix(i, 20))
                            rsGrid("Department_Cls") = Trim(.TextMatrix(i, 4))
                            rsGrid("AccountNo") = Trim(.TextMatrix(i, 22))
                            rsGrid.update
                        Else
                            If Trim(.TextMatrix(i, 21)) <> "" Then
                                Sql = "Delete from PurchaseOrder_Detail where seq_no = '" & .TextMatrix(i, 21) & "'"
                                Db.Execute Sql
                            End If
                        End If
                        
                        'UPDATE COMPLETE_CLS (POREQUEST_MASTER)
                        Sql = "select PORequest_No, avg(DComplete) Complete " & _
                              "from ( select prd.PORequest_No, prd.Seq_No, isnull(prd.Qty,0) Qty, isnull(sum(pod.Qty),0) POQty, " & _
                              "        (case when isnull(prd.Qty,0) = isnull(sum(pod.Qty),0) then 1 else 0 end) DComplete " & _
                              "        from PORequest_Detail prd " & _
                              "        left outer join PurchaseOrder_Detail pod on pod.PORequest_No = prd.PORequest_No and pod.POReq_SeqNo = prd.Seq_No " & _
                              "        group by prd.PORequest_No, prd.Seq_No, prd.Qty ) a " & _
                              "where PORequest_No = '" & Trim(.TextMatrix(i, 1)) & "' " & _
                              "group by PORequest_No "
                        If rscekC.State <> adStateClosed Then rscekC.Close
                        rscekC.Open Sql, Db, adOpenKeyset, adLockOptimistic
                        If Not (rscekC.BOF And rscekC.EOF) Then
                            If rscekC("Complete") = "1" Then
                                Sql = "Update PORequest_Master set Complete_Cls = '1' where PORequest_No = '" & Trim(rscekC("PORequest_No")) & "' "
                            ElseIf rscekC("Complete") = "0" Then
                                Sql = "Update PORequest_Master set Complete_Cls = '0' where PORequest_No = '" & Trim(rscekC("PORequest_No")) & "' "
                            End If
                            Db.Execute Sql
                        End If
                    Next i
                    Call updatemaster(True)
                    Call CekPONumber
                    Call Browse
                    lblErrMsg = DisplayMsg(1101)
                    ubahgrid = True
                End With
          End If

    Case 1: 'CLEAR
            Call kosong
            combo1.ListIndex = 1
            Call Combo1_Click
            cborequestno.SetFocus
            
    Case 2: 'CREATE / UPDATE
            If combo1.ListIndex = 0 Then    'CREATE
                If hakUpdate(Me.Name) = 0 Then _
                    lblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
                
                'HEADER VALIDATION
                If cborequestno.Text = "" Then
                    cborequestno.SetFocus
                    lblErrMsg = DisplayMsg(1067) '"Please Input Request No"
                    Exit Sub
                ElseIf cborequestno.Text <> "" Then
                    If cborequestno.MatchFound = False Then
                        lblErrMsg = DisplayMsg(4144)    'Record with This Request No Not found !
                        cborequestno.SetFocus
                        Exit Sub
                    End If
                End If
'                If cborequestno.Column(2) = "1" Then
'                    cborequestno.SetFocus
'                    lblErrMsg = DisplayMsg(4088) '"Can't Create PO, PO Request No already Fix"
'                    Exit Sub
'                End If
                If cbocust.Text = "" Then
                    cbocust.SetFocus
                    lblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
                    Exit Sub
                ElseIf cbocust.Text <> "" Then
                    If cbocust.MatchFound = False Then
                        lblErrMsg = DisplayMsg(4050)    'Record with This Supplier Code Not found !
                        cbocust.SetFocus
                        Exit Sub
                    End If
                End If
                If txtpono.Text = "" Then
                    txtpono.SetFocus
                    lblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                    Exit Sub
                End If
                                
                '-------------------------------------------------------------
                
                'FOOTER VALIDATION
                If cbopricecondition.Text <> "" Then
                    If cbopricecondition.MatchFound = False Then
                        lblErrMsg = DisplayMsg(4147)    'Record with This Price Condition not found !
                        cbopricecondition.SetFocus
                        Exit Sub
                    End If
                End If
                If cbopocode.Text <> "" Or txtpoday.Text <> "" Or cbopoterms.Text <> "" Then
                    If cbopocode.Text = "" Then
                        cbopocode.SetFocus
                        lblErrMsg = DisplayMsg(8123)
                        Exit Sub
                    ElseIf txtpoday.Text = "" Or Val(txtpoday.Text) = 0 Then
                        txtpoday.SetFocus
                        lblErrMsg = DisplayMsg(8123)
                        Exit Sub
                    ElseIf cbopoterms.Text = "" Then
                        cbopoterms.SetFocus
                        lblErrMsg = DisplayMsg(1062)
                        Exit Sub
                    End If
                    If cbopocode.Text <> "" Then
                        If cbopocode.MatchFound = False Then
                            lblErrMsg = DisplayMsg(4097)
                            cbopocode.SetFocus
                            Exit Sub
                        End If
                    End If
                    If cbopoterms.Text <> "" Then
                        If cbopoterms.MatchFound = False Then
                            lblErrMsg = DisplayMsg(4097)
                            cbopoterms.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
                If cbodelmode.Text <> "" Then
                    If cbodelmode.MatchFound = False Then
                        lblErrMsg = DisplayMsg(4099)    '"Record with this Transportation not found"
                        cbodelmode.SetFocus
                        Exit Sub
                    End If
                End If
                '-------------------------------------------------------------
                
                
                If ubah = False Then
                    Sql = "select * from PurchaseOrder_Master where PO_No = '" & txtpono.Text & "'"
                    If rs.State <> adStateClosed Then rs.Close
                    rs.Open Sql, Db, adOpenKeyset, adLockOptimistic
                    If Not (rs.BOF And rs.EOF) Then
                        lblErrMsg.Caption = DisplayMsg(1023)
                        txtpono.SetFocus
                        Exit Sub
                    Else
                        rs.AddNew
                        rs("po_no") = txtpono.Text
                        rs("supplier_code") = cbocust.Text
                    End If
                End If
                    
                rs("po_date") = Format(podate.Value, "yyyy-mm-dd")
                rs("amount") = CDbl(txtamount.Text)
                rs("ppn") = CDbl(txtppn.Text)
                rs("total_amount") = CDbl(txtgrandtotal.Text)
                rs("remarks") = Trim(txtRemarks(0).Text)
                rs("remarks2") = Trim(txtRemarks(1).Text)
                rs("remarks3") = Trim(txtRemarks(2).Text)
                rs("others_cls") = "1"
                rs("Revisi") = CDbl(Trim(txtRev.Text))
                If cbopocode.Text = "" Then rs("popayment_code") = Null Else rs("popayment_code") = cbopocode.Column(0)
                If txtpoday.Text = "" Then rs("popayment_days") = Null Else rs("popayment_days") = txtpoday.Text
                If cbopoterms.Text = "" Then rs("popayment_terms") = Null Else rs("popayment_terms") = cbopoterms.Column(0)
                If cbodelmode.Text = "" Then rs("transportation_cls") = Null Else rs("transportation_cls") = cbodelmode.Column(0)
                If cbopricecondition.Text = "" Then rs("pricecondition_cls") = Null Else rs("pricecondition_cls") = cbopricecondition.Text
On Error Resume Next
                rs.update
ErrHandler:
                If InStr(1, Err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                    Call poNo(Right(year(podate), 2), Format(month(podate), "0#"))
                    txtpono2.Text = txtpono.Text
                    rs("po_No") = txtpono.Text
                    rs.update
                    If InStr(1, Err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                        GoTo ErrHandler
                    Else
                        If Trim$(Err.Description) <> "" Then
                            lblErrMsg = Trim$(Err.number) + " : " + Trim$(Err.Description)
                            Exit Sub
                        End If
                    End If
                Else
                    If Trim$(Err.Description) <> "" Then
                        lblErrMsg = Trim$(Err.number) + " : " + Trim$(Err.Description)
                        Exit Sub
                    End If
                End If
    
                If CDate(podate.Value) > CDate(requestdate1.Value) Then
                    If CDate(podate.Value) > CDate(requestdate2.Value) Then _
                        requestdate2.Value = Format(podate.Value, "dd MMM yyyy")
                Else
                    requestdate1.Value = Format(podate.Value, "dd MMM yyyy")
                End If
                
'                Call updateMaster(True)
                
                combo1.Text = "Update"
                If cbocust.Text <> "" And cborequestno.Text <> "" Then Call browseitem: Call FormatPrice
                lblErrMsg.Caption = DisplayMsg(1000)
                ubah = True
    
            Else    'UPDATE
             If sampun Then                     'sampun = true->ada data di grid; false->tidak ada data di grid
                Call updatemaster(False)
             Else
                Dim ketemu As Boolean
                
                If cborequestno.Text = "" Then
'                    cborequestno.SetFocus
                    'LblErrMsg = DisplayMsg(1067) '"Please Input Request No"
                    'Exit Sub
                ElseIf cborequestno.Text <> "" Then
                    If cborequestno.MatchFound = False Then
                        lblErrMsg = DisplayMsg(4144)    'Record with This Request No Not found !
                        cborequestno.SetFocus
                        Exit Sub
                    End If
                End If
                If cbocust.Text = "" Then
                    cbocust.SetFocus
                    lblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
                    Exit Sub
                ElseIf cbocust.Text <> "" Then
                    If cbocust.MatchFound = False Then
                        lblErrMsg = DisplayMsg(4050)    'Record with This Supplier Code Not found !
                        cbocust.SetFocus
                        Exit Sub
                    End If
                End If
                
                If txtpono.Text = "" Then
                    txtpono.SetFocus
                    lblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                    Exit Sub
                End If
                
                If CDate(podate.Value) > CDate(requestdate1.Value) Then
                    If CDate(podate.Value) > CDate(requestdate2.Value) Then _
                        requestdate2.Value = Format(podate.Value, "dd MMM yyyy")
                Else
                    requestdate1.Value = Format(podate.Value, "dd MMM yyyy")
                End If
    
                If cbocust.Text = "" Then   'Or cborequestno.Text = ""
                    cbopono.clear: txtpono.Text = ""
                Else
                    Call adtocbopono
                End If
                For i = 0 To cbopono.ListCount - 1
                    If txtpono.Text = cbopono.List(i) Then
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False Then GoTo here
                
                Call Browse
                sampun = True   'true -> ada data di grid
                Call updatemaster(False)
            End If
            
                If ada = False Then
here:
                    Call kosongBwh
                    txtRemarks(0).Text = "": txtRemarks(1).Text = "": txtRemarks(2).Text = ""
                    Call header
                    lblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
                    txtpono.SetFocus
                    Exit Sub
                End If
            End If
    
    Case 3: 'CANCEL
            If txtpono.Text <> "" And cbocust.Text <> "" Then   'And cboRequestNo.Text <> ""
                For i = 0 To cbopono.ListCount - 1
                    If txtpono.Text = cbopono.List(i) Then
                        ketemu = True
                        Exit For
                    End If
                Next i
                If ketemu = False Then
                    Call kosongBwh
                    txtRemarks(0).Text = "": txtRemarks(1).Text = "": txtRemarks(2).Text = ""
                    Call header
                    lblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
                    txtpono.SetFocus
                    Exit Sub
                End If
                Call BrowseAtas
                Call Browse
            End If
End Select
End Sub

Private Sub cmdReport_Click()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3
Dim sqlcekdet As String, SqlRpt As String
Dim rscekdet As New Recordset
  
    If combo1.ListIndex = 1 And txtpono.Text <> "" And cbocust.Text <> "" Then
        sqlcekdet = "select pom.PO_No from PurchaseOrder_Master pom " & _
                    "inner join PurchaseOrder_Detail pod on pod.PO_No = pom.PO_No " & _
                    "where pom.others_cls = '1' and pom.period is null " & _
                    "and pom.PO_No = '" & Trim(txtpono.Text) & "' and pom.supplier_Code = '" & Trim(cbocust.Text) & "'"
        Set rscekdet = Db.Execute(sqlcekdet)
        If rscekdet.EOF Then lblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
        Set rscekdet = Nothing
        
        Me.MousePointer = vbHourglass

        If cbocust.Column(4) = 1 Then   'PO CLS=YES
            If cborequestno.Text = "" Then
'                cborequestno.SetFocus
'                lblErrMsg = DisplayMsg(1067) '"Please Input Request No"
'                Me.MousePointer = vbDefault
'                Exit Sub
            ElseIf cborequestno.Text <> "" Then
                If cborequestno.MatchFound = False Then
                    lblErrMsg = DisplayMsg(4144)    'Record with This Request No Not found !
                    cborequestno.SetFocus
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        
            'PURCHASE ORDER DETAIL
            SqlRpt = "select '1' Sort, rtrim(pom.po_no) po_no, pom.po_date, rtrim(pom.supplier_Code) Supplier_Code, " & _
                    " trade_name = case when CHARINDEX ( ',' , rtrim(tm.trade_name)) <> 0 then " & _
                    " ltrim(substring(rtrim(tm.trade_name), CHARINDEX ( ',' , rtrim(tm.trade_name)) + 1, 5)) + " & _
                    " '.' + left(rtrim(tm.trade_name),CHARINDEX ( ',' , rtrim(tm.trade_name)) - 1) " & _
                    " Else " & _
                    " RTrim (tm.trade_name) " & _
                    " End, "
            SqlRpt = SqlRpt + " " & _
                     "rtrim(tm.address1) taddress1, rtrim(tm.address2) taddress2, rtrim(tm.city) tcity, rtrim(tm.postal_code) tpostal_code, " & _
                     "rtrim(tm.contact_person) contact_person, isnull(tm.telephone,'') Supplierphone, isnull(tm.Fax,'') SupplierFax, pom.POPayment_code, pom.POPayment_Days, pom.POPayment_Terms, " & _
                     "rtrim(pod.PORequest_No) PORequest_No, pod.POReq_SeqNo, rtrim(pod.item_code) item_code, rtrim(prd.item_name) item_name, " & _
                     "pod.unit_cls,(select description from unit_cls uc where uc.unit_cls= pod.unit_cls ) unit_desc,  isnull(pod.qty,0) Qty, pod.currency_code, (select description from curr_cls where curr_cls.Curr_cls= pod.Currency_Code) Curr_desc ,isnull(pod.price,0) Price, isnull(pod.amount,0) Amount, " & _
                     "pod.Delivery_Date, pom.PriceCondition_Cls, rtrim(pc.Description) PriceCondition, pom.Transportation_Cls, " & _
                     "rtrim(pom.remarks) Remarks, rtrim(pom.remarks2) Remarks2, rtrim(pom.remarks3) Remarks3, isnull(pom.amount,0) as TAmount, isnull(pom.ppn,0) PPN, isnull(pom.total_amount,0) Total_Amount, " & _
                     "rtrim(cp.company_name) company_name, rtrim(cp.address1) caddress1, rtrim(cp.address2) caddress2, " & _
                     "rtrim(cp.Province) cprovince, rtrim(cp.City) ccity, rtrim(cp.postal_code) cpostal_code, rtrim(cp.phone1) cphone1, " & _
                     "rtrim(cp.phone2) cphone2, rtrim(cp.fax) cfax, rtrim(cp.PO_position) po_position, rtrim(cp.PO_person) po_person, " & _
                     "rtrim(cp.POAcknowledge_Person) POAcknowledge_Person, rtrim(cp.POAcknowledge_Position) POAcknowledge_Position, " & _
                     "rtrim(cp.POApproved_Person) POApproved_Person, rtrim(cp.POApproved_Position) POApproved_Position, tm.Trade_Cls, tm.Country_Cls, rtrim(pod.Department_cls) Department_Cls " & _
                     "from PurchaseOrder_Master pom " & _
                     "inner join PurchaseOrder_Detail pod on pod.PO_No = pom.PO_No " & _
                     "left outer join PORequest_Detail prd on prd.item_code = pod.Item_code and prd.Seq_No = pod.POReq_SeqNo " & _
                     "left outer join Trade_Master tm on tm.trade_code = pom.supplier_code " & _
                     "left outer join PriceCondition_Cls pc on pc.PriceCondition_Cls = pom.PriceCondition_Cls " & _
                     "cross join Company_Profile cp " & _
                     "where pom.po_no = '" & Trim(txtpono.Text) & "' and pom.others_cls = '1' and pom.period is null "
                     
            'from selected POREQUEST No
            SqlRpt = SqlRpt & _
                     "UNION " & _
                     "select '3' Sort, rtrim(pom.po_no) po_no, pom.po_date, rtrim(pom.supplier_Code) Supplier_Code, " & _
                    " trade_name = case when CHARINDEX ( ',' , rtrim(tm.trade_name)) <> 0 then " & _
                    " ltrim(substring(rtrim(tm.trade_name), CHARINDEX ( ',' , rtrim(tm.trade_name)) + 1, 5)) + " & _
                    " '.' + left(rtrim(tm.trade_name),CHARINDEX ( ',' , rtrim(tm.trade_name)) - 1) " & _
                    " Else " & _
                    " RTrim (tm.trade_name) " & _
                    " End, "
            SqlRpt = SqlRpt + " " & _
                     "rtrim(tm.address1) taddress1, rtrim(tm.address2) taddress2, rtrim(tm.city) tcity, rtrim(tm.postal_code) tpostal_code, " & _
                     "rtrim(tm.contact_person) contact_person, isnull(tm.telephone,'') Supplierphone, isnull(tm.Fax,'') SupplierFax, pom.POPayment_code, pom.POPayment_Days, pom.POPayment_Terms, " & _
                     "prd.PORequest_No, prd.Seq_No, rtrim(prd.item_code) item_code, rtrim(prd.item_name) item_name, " & _
                     "prd.unit_cls, (select description from unit_cls uc where uc.unit_cls= prd.unit_cls ) unit_desc, 0 Qty, " & _
                     "(select top 1 currency_code from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "' and Currency_Code is not null) currency_code, (select description from curr_cls where curr_cls.Curr_cls= (select top 1 currency_code from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "' and Currency_Code is not null)) Curr_desc, 0 Price, 0 Amount, " & _
                     "Null Delivery_Date, pom.PriceCondition_Cls, '' PriceCondition, pom.Transportation_Cls, " & _
                     "rtrim(pom.remarks) Remarks, rtrim(pom.remarks2) Remarks2, rtrim(pom.remarks3) Remarks3, isnull(pom.amount,0) as TAmount, isnull(pom.ppn,0) PPN, isnull(pom.total_amount,0) Total_Amount, " & _
                     "rtrim(cp.company_name) company_name, rtrim(cp.address1) caddress1, rtrim(cp.address2) caddress2, " & _
                     "rtrim(cp.Province) cprovince, rtrim(cp.City) ccity, rtrim(cp.postal_code) cpostal_code, rtrim(cp.phone1) cphone1, " & _
                     "rtrim(cp.phone2) cphone2, rtrim(cp.fax) cfax, rtrim(cp.PO_position) po_position, rtrim(cp.PO_person) po_person, " & _
                     "rtrim(cp.POAcknowledge_Person) POAcknowledge_Person, rtrim(cp.POAcknowledge_Position) POAcknowledge_Position, " & _
                     "rtrim(cp.POApproved_Person) POApproved_Person, rtrim(cp.POApproved_Position) POApproved_Position, tm.Trade_Cls, tm.Country_Cls, rtrim(prm.Department_cls) Department_Cls " & _
                     "from PurchaseOrder_Master pom, Trade_Master tm, " & _
                     "PORequest_Detail prd, (select * from PORequest_Master where isnull(others_cls,'0') = '1') prm, " & _
                     "Company_Profile cp "
            SqlRpt = SqlRpt & _
                     "Where prm.porequest_no = prd.porequest_no " & _
                     "And pom.Supplier_Code = tm.Trade_Code " & _
                     "and pom.others_Cls = '1' and pom.period is null and pom.PO_No = '" & Trim(txtpono.Text) & "' And tm.PO_Cls = '1' " & _
                     "and prd.porequest_no = '" & Trim(cborequestno.Text) & "' and isnull(prm.Complete_Cls,'0') = '0' " & _
                     "and prd.seq_no not in (select POReq_SeqNo from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "') " & _
                     "Order by Sort "

        Else    'PO CLS=NO
            SqlRpt = "select '1' Sort, rtrim(pom.po_no) po_no, pom.po_date, rtrim(pom.supplier_Code) Supplier_Code, " & _
                    " trade_name = case when CHARINDEX ( ',' , rtrim(tm.trade_name)) <> 0 then " & _
                    " ltrim(substring(rtrim(tm.trade_name), CHARINDEX ( ',' , rtrim(tm.trade_name)) + 1, 5)) + " & _
                    " '.' + left(rtrim(tm.trade_name),CHARINDEX ( ',' , rtrim(tm.trade_name)) - 1) " & _
                    " Else " & _
                    " RTrim (tm.trade_name) " & _
                    " End, "
            SqlRpt = SqlRpt + " " & _
                     "rtrim(tm.address1) taddress1, rtrim(tm.address2) taddress2, rtrim(tm.city) tcity, rtrim(tm.postal_code) tpostal_code, " & _
                     "rtrim(tm.contact_person) contact_person, isnull(tm.telephone,'') Supplierphone, isnull(tm.Fax,'') SupplierFax, pom.POPayment_code, pom.POPayment_Days, pom.POPayment_Terms, " & _
                     "rtrim(pod.PORequest_No) PORequest_No, pod.POReq_SeqNo, rtrim(pod.item_code) item_code, rtrim(prd.item_name) item_name, " & _
                     "pod.unit_cls, (select description from unit_cls uc where uc.unit_cls= pod.unit_cls ) unit_desc, isnull(pod.qty,0) Qty, pod.currency_code, (select description from curr_cls where curr_cls.Curr_cls= pod.Currency_Code) Curr_desc,  isnull(pod.price,0) Price, isnull(pod.amount,0) Amount, " & _
                     "pod.Delivery_Date, pom.PriceCondition_Cls, rtrim(pc.Description) PriceCondition, pom.Transportation_Cls, " & _
                     "rtrim(pom.remarks) Remarks, rtrim(pom.remarks2) Remarks2, rtrim(pom.remarks3) Remarks3, isnull(pom.amount,0) as TAmount, isnull(pom.ppn,0) PPN, isnull(pom.total_amount,0) Total_Amount, " & _
                     "rtrim(cp.company_name) company_name, rtrim(cp.address1) caddress1, rtrim(cp.address2) caddress2, " & _
                     "rtrim(cp.Province) cprovince, rtrim(cp.City) ccity, rtrim(cp.postal_code) cpostal_code, rtrim(cp.phone1) cphone1, " & _
                     "rtrim(cp.phone2) cphone2, rtrim(cp.fax) cfax, rtrim(cp.PO_position) po_position, rtrim(cp.PO_person) po_person, " & _
                     "rtrim(cp.POAcknowledge_Person) POAcknowledge_Person, rtrim(cp.POAcknowledge_Position) POAcknowledge_Position, " & _
                     "rtrim(cp.POApproved_Person) POApproved_Person, rtrim(cp.POApproved_Position) POApproved_Position, tm.Trade_Cls, tm.Country_Cls, rtrim(pod.Department_cls) Department_Cls " & _
                     "from PurchaseOrder_Master pom " & _
                     "inner join PurchaseOrder_Detail pod on pod.PO_No = pom.PO_No " & _
                     "left outer join PORequest_Detail prd on prd.PORequest_No = pod.PORequest_No and prd.Seq_no = pod.POReq_SeqNo " & _
                     "left outer join Trade_Master tm on tm.trade_code = pom.supplier_code " & _
                     "left outer join PriceCondition_Cls pc on pc.PriceCondition_Cls = pom.PriceCondition_Cls " & _
                     "cross join Company_Profile cp " & _
                     "where pom.po_no = '" & Trim(txtpono.Text) & "' and pom.others_cls = '1' and pom.period is null " & _
                     "order by pod.PORequest_No, pod.Item_Name, pod.POReq_SeqNo "
        End If
        
        If rsRpt.State <> adStateClosed Then rsRpt.Close
        rsRpt.Open SqlRpt, Db, adOpenDynamic, adLockOptimistic
        
        sqlprint = SqlRpt
        reportcode = "poparts"
        Fbulan = txtpono.Text
        printorient = 1
        
        If rsRpt.EOF Then lblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
        Set report = application.OpenReport(App.path & "\Reports\rptPOParts.rpt")
        report.Database.Tables(1).SetDataSource rsRpt
        
        Rpt.CRViewer1.ReportSource = report
        Rpt.CRViewer1.ViewReport
        Rpt.CRViewer1.Zoom 1
        Rpt.WindowState = 2
        Rpt.Show 1
        
        Set rsRpt = Nothing
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdSubMenu_Click()
    ClearData
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        lblErrMsg.Caption = ErrMsg
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    Set rsGrid = Nothing
End Sub

Private Sub updatemaster(Flag As Boolean)
    Dim sQl_Master As String
    Dim rs_Master As New ADODB.Recordset
        sQl_Master = "select * from PurchaseOrder_Master where PO_No = '" & txtpono.Text & "' and others_cls = '1' and period is null"
        If rs_Master.State <> adStateClosed Then rs_Master.Close
        rs_Master.Open sQl_Master, Db, adOpenKeyset, adLockOptimistic
        If rs_Master.BOF And rs_Master.EOF Then
            lblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
            txtpono.SetFocus
            rs_Master.Close
            Set rs_Master = Nothing
            Exit Sub
        End If
        rs_Master("po_date") = Format(podate.Value, "yyyy-mm-dd")
        rs_Master("revisi") = Trim(txtRev.Text)
        rs_Master("supplier_Code") = Trim(cbocust.Text)
        If Flag = True Then
            rs_Master("amount") = CDbl(txtamount.Text)
            rs_Master("ppn") = CDbl(txtppn.Text)
            rs_Master("total_amount") = CDbl(txtgrandtotal.Text)
            rs_Master("remarks") = Trim(txtRemarks(0).Text)
            rs_Master("remarks2") = Trim(txtRemarks(1).Text)
            rs_Master("remarks3") = Trim(txtRemarks(2).Text)
            If cbopocode.Text = "" Then rs_Master("popayment_code") = Null Else rs_Master("popayment_code") = cbopocode.Column(0)
            If txtpoday.Text = "" Then rs_Master("popayment_days") = Null Else rs_Master("popayment_days") = txtpoday.Text
            If cbopoterms.Text = "" Then rs_Master("popayment_terms") = Null Else rs_Master("popayment_terms") = cbopoterms.Column(0)
            If cbodelmode.Text = "" Then rs_Master("transportation_cls") = Null Else rs_Master("transportation_cls") = cbodelmode.Column(0)
            If cbopricecondition.Text = "" Then rs_Master("pricecondition_cls") = Null Else rs_Master("pricecondition_cls") = cbopricecondition.Text
        End If
        rs_Master.update
        rs_Master.Close
        Set rs_Master = Nothing
End Sub
