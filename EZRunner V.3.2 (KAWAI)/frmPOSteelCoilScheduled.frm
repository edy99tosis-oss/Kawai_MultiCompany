VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOSteelCoilScheduled 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Scheduled (Steel/Coil)"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "frmPOSteelCoilScheduled.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10260
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtSubAmount 
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
      Left            =   3150
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   8685
      Width           =   2355
   End
   Begin VB.TextBox TxtDisc 
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
      Left            =   5550
      MaxLength       =   25
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   8685
      Width           =   2355
   End
   Begin VB.TextBox txtMarking 
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
      Index           =   5
      Left            =   12900
      MaxLength       =   25
      TabIndex        =   20
      Top             =   7110
      Width           =   2085
   End
   Begin VB.CommandButton command2 
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
      Index           =   4
      Left            =   6390
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   9765
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton command2 
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
      Index           =   3
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   9765
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton command2 
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
      Index           =   2
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   9765
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton command2 
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
      Index           =   1
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   9765
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtMarking 
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
      Left            =   7275
      MaxLength       =   25
      TabIndex        =   16
      Top             =   7110
      Width           =   2085
   End
   Begin VB.TextBox txtMarking 
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
      Index           =   3
      Left            =   10050
      MaxLength       =   25
      TabIndex        =   18
      Top             =   7125
      Width           =   2085
   End
   Begin VB.TextBox txtPacking 
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
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   7125
      Width           =   3165
   End
   Begin VB.TextBox txtInsurance 
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
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   7485
      Width           =   3165
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
      Left            =   7590
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   7695
      Width           =   7515
   End
   Begin VB.TextBox txtMarking 
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
      Index           =   4
      Left            =   12915
      MaxLength       =   25
      TabIndex        =   19
      Top             =   6735
      Width           =   2085
   End
   Begin VB.TextBox txtTransport 
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
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   7890
      Width           =   3165
   End
   Begin VB.TextBox txtPaymentTerm 
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
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   6735
      Width           =   3165
   End
   Begin VB.TextBox txtPriceCondition 
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
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   6345
      Width           =   3165
   End
   Begin VB.TextBox txtMarking 
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
      Left            =   7275
      MaxLength       =   25
      TabIndex        =   15
      Top             =   6735
      Width           =   2085
   End
   Begin VB.TextBox txtMarking 
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
      Left            =   10050
      MaxLength       =   25
      TabIndex        =   17
      Top             =   6735
      Width           =   2085
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
      MaxLength       =   1
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2145
      Width           =   615
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13343
      TabIndex        =   55
      Top             =   210
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
   End
   Begin MSComCtl2.FlatScrollBar hscrollbar 
      Height          =   255
      Left            =   90
      TabIndex        =   50
      Top             =   5850
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Max             =   1
      Orientation     =   1638401
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
      Left            =   10395
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9765
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   580
      Left            =   83
      TabIndex        =   51
      Top             =   795
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
         Format          =   141230083
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
         Format          =   141230083
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   240
         Width           =   1095
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   3300
      Left            =   90
      TabIndex        =   9
      Top             =   2565
      Width           =   15105
      _cx             =   26644
      _cy             =   5821
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
         TabIndex        =   26
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
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin MSForms.ComboBox cbocurr 
         Height          =   285
         Left            =   5640
         TabIndex        =   28
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
      Begin MSForms.ComboBox cboprice 
         Height          =   285
         Left            =   6840
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         VariousPropertyBits=   746604571
         MaxLength       =   19
         DisplayStyle    =   3
         Size            =   "3625;503"
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
      Left            =   10290
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   8685
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
      Left            =   7920
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   8685
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
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   8685
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
      Left            =   12825
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9765
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
      Top             =   2145
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
      Top             =   2115
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
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   8685
      Width           =   2355
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   83
      TabIndex        =   39
      Top             =   9120
      Width           =   15105
      Begin VB.Label lblErrMsg 
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
         Height          =   255
         Left            =   120
         TabIndex        =   40
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
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9765
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
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9765
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
      Left            =   11610
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9765
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   720
      Left            =   83
      TabIndex        =   42
      Top             =   1395
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   46
         Top             =   270
         Width           =   840
      End
      Begin MSForms.ComboBox cbocust 
         Height          =   315
         Left            =   900
         TabIndex        =   3
         Top             =   240
         Width           =   1890
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3334;556"
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
         Caption         =   "Supplier"
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
         TabIndex        =   43
         Top             =   270
         Width           =   825
      End
   End
   Begin MSComCtl2.DTPicker podate 
      Height          =   315
      Left            =   6120
      TabIndex        =   6
      Top             =   2145
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
      Format          =   141230083
      CurrentDate     =   37798
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
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
      Index           =   29
      Left            =   3885
      TabIndex        =   83
      Top             =   8340
      Width           =   810
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Index           =   28
      Left            =   6285
      TabIndex        =   82
      Top             =   8340
      Width           =   735
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO Marking"
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
      Left            =   6705
      TabIndex        =   79
      Top             =   6345
      Width           =   975
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line6"
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
      Left            =   12330
      TabIndex        =   78
      Top             =   7170
      Width           =   450
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   6630
      Top             =   6300
      Width           =   8475
   End
   Begin MSForms.ComboBox cboPacking 
      Height          =   315
      Left            =   1905
      TabIndex        =   12
      Top             =   7065
      Width           =   1305
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "2302;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line2 "
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
      Index           =   19
      Left            =   6750
      TabIndex        =   73
      Top             =   7170
      Width           =   510
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line4"
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
      Index           =   21
      Left            =   9510
      TabIndex        =   72
      Top             =   7185
      Width           =   450
   End
   Begin VB.Line Line8 
      X1              =   3315
      X2              =   6465
      Y1              =   7740
      Y2              =   7740
   End
   Begin VB.Label lblCaption 
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
      Index           =   8
      Left            =   6750
      TabIndex        =   71
      Top             =   7755
      Width           =   765
   End
   Begin VB.Line Line7 
      X1              =   3315
      X2              =   6450
      Y1              =   8130
      Y2              =   8130
   End
   Begin VB.Line Line6 
      X1              =   3315
      X2              =   6465
      Y1              =   7365
      Y2              =   7365
   End
   Begin VB.Line Line5 
      X1              =   3315
      X2              =   6465
      Y1              =   6990
      Y2              =   6990
   End
   Begin VB.Line Line4 
      X1              =   3315
      X2              =   6465
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line5"
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
      Index           =   23
      Left            =   12345
      TabIndex        =   70
      Top             =   6795
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Covered"
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
      Index           =   16
      Left            =   90
      TabIndex        =   69
      Top             =   7500
      Width           =   1650
   End
   Begin MSForms.ComboBox cboInsuranceCls 
      Height          =   315
      Left            =   1890
      TabIndex        =   13
      Top             =   7470
      Width           =   1305
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "2302;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
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
      Index           =   17
      Left            =   90
      TabIndex        =   68
      Top             =   7890
      Width           =   1245
   End
   Begin MSForms.ComboBox cboTransport 
      Height          =   315
      Left            =   1905
      TabIndex        =   14
      Top             =   7830
      Width           =   1305
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "2302;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line1"
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
      Index           =   18
      Left            =   6750
      TabIndex        =   67
      Top             =   6795
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line3"
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
      Index           =   20
      Left            =   9510
      TabIndex        =   66
      Top             =   6795
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Packing"
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
      Index           =   15
      Left            =   90
      TabIndex        =   65
      Top             =   7125
      Width           =   660
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Term"
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
      Index           =   14
      Left            =   90
      TabIndex        =   64
      Top             =   6750
      Width           =   1260
   End
   Begin MSForms.ComboBox cboPaymentTerm 
      Height          =   315
      Left            =   1905
      TabIndex        =   11
      Top             =   6690
      Width           =   1305
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "2302;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
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
      Index           =   13
      Left            =   90
      TabIndex        =   63
      Top             =   6360
      Width           =   1290
   End
   Begin MSForms.ComboBox cboPriceCondition 
      Height          =   315
      Left            =   1905
      TabIndex        =   10
      Top             =   6300
      Width           =   1305
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "2302;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00A6D2FF&
      Height          =   975
      Left            =   6630
      Top             =   6600
      Width           =   8475
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO Marking"
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
      Index           =   22
      Left            =   6720
      TabIndex        =   62
      Top             =   6330
      Width           =   975
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
      TabIndex        =   56
      Top             =   2190
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
      Left            =   14010
      TabIndex        =   47
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
      TabIndex        =   45
      Top             =   2190
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
      TabIndex        =   44
      Top             =   2190
      Width           =   825
   End
   Begin MSForms.ComboBox cbopono 
      Height          =   315
      Left            =   2280
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2145
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
      Top             =   2145
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
      TabIndex        =   41
      Top             =   8370
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
      Left            =   7935
      TabIndex        =   38
      Top             =   8325
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
      TabIndex        =   37
      Top             =   8325
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
      Left            =   10290
      TabIndex        =   36
      Top             =   8325
      Width           =   2325
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Scheduled (Steel/Coil)"
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
      Left            =   83
      TabIndex        =   35
      Top             =   240
      Width           =   15105
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   495
      Index           =   2
      Left            =   90
      Top             =   8580
      Width           =   15105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   330
      Index           =   2
      Left            =   90
      Top             =   8265
      Width           =   15105
   End
End
Attribute VB_Name = "frmPOSteelCoilScheduled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String, sqlGrid As String
Dim RS As New ADODB.Recordset, rsGrid As New ADODB.Recordset
Dim i As Long, orderawal As Double, isippn As Long
Dim ubah As Boolean, ubahgrid As Boolean, ada As Boolean, sampun As Boolean
Dim statusfix As String
Dim actrow As Long, activecurrcd As String, activecurr As String
Dim countrycls As Byte
'Const isiPOTerm = "after B/L Date,after Delivery Date,after Invoice Date,prior before Shipment,from Custom Clearance Date,after Receive Invoice,after Receive Goods"

Private Sub ClearData()

    sql = "Delete from PurchaseOrder_Master Where PO_No = '" & Trim(txtPoNo.Text) & "' " & _
          "and PO_No not in (select PO_No from PurchaseOrder_Detail) and others_cls = '0' and period is null"
    Db.Execute sql

End Sub

Private Sub CekPONumber()
    
    Dim adoRs As New ADODB.Recordset
    
    sql = "Select * From PurchaseOrder_Master Where PO_No = '" & Trim(txtPoNo.Text) & "'"
    adoRs.Open sql, Db, adOpenKeyset, adLockOptimistic, adCmdText
    If Not adoRs.EOF Then
        Call PONO(Right(Year(PODate), 2), Format(Month(PODate), "0#"))
        adoRs.update "PO_No", Trim(txtPoNo.Text)
    End If
    adoRs.Close
    Set adoRs = Nothing
    
End Sub

Sub Kosong()
    Dim X As Integer
    
    LblErrMsg = ""
    requestdate1.Value = Format(Now, "yyyy-mm-01")
    requestdate2.Value = Format(Now, "dd MMM yyyy")
    Call adtocborequestno
    cborequestno.Text = ""
    
    cboCust.Text = ""
    lblcust(0).Text = "": lblcust(1).Text = ""
    txtPoNo.Text = "": txtPONo2.Text = ""
    PODate.Value = Format(Now, "dd MMM yyyy")
    PODate.Enabled = True
    Call ppn(PODate.Value)
    txtRev.Text = ""
    
    grid.FocusRect = flexFocusNone
    DelDate.Value = Format(Now + 1, "dd MMM yyyy")
    
    cboPriceCondition.Text = ""
    cboPaymentTerm.ListIndex = -1
    CboPacking.Text = ""
    cboInsuranceCls = ""
    For X = 1 To 4
        txtMarking(X).Text = ""
    Next X
    txtremarks.Text = ""
    cboTransport.ListIndex = -1
    
    ubah = False: ada = False
    statusfix = 0
    
    Call kunci(False)
    Call kosongBwh: Call Header
End Sub

Sub kosongBwh()
    ' Add 20090112
    TxtSubAmount.Text = 0
    TxtDisc.Text = 0
    txtamount.Text = 0
    ' ---
    txtPPN.Text = 0
      txtGrandTotal.Text = 0
    cboPriceCondition.Text = ""
    cboPaymentTerm.Text = ""
    CboPacking.Text = ""
    cboInsuranceCls.Text = ""
    cboTransport.Text = ""
    txtMarking(0).Text = "": txtMarking(1).Text = "": txtMarking(2).Text = "": txtMarking(3).Text = ""
    txtMarking(4).Text = "": txtMarking(5).Text = "": txtremarks = ""
    TxtTransport = "": txtInsurance = "": TxtPacking = "": txtPaymentTerm = "": txtPriceCondition = ""
    
End Sub

Function adtocboCust(ByVal filter As Boolean)
Dim sqlcust As String
Dim RsCust As New Recordset
    ' Supplier External ( Non Subcon )
    If filter = True Then
        sqlcust = "select tm.trade_code, tm.trade_name, tm.address1, tm.country_cls, tm.po_cls, " & _
                  "tm.popayment_code, tm.popayment_day, tm.popayment_terms, tm.transportation_Cls, isnull(tm.Trade_Abbr,'') Trade_Abbr " & _
                  "from trade_master tm " & _
                  "Inner Join " & _
                  "(select distinct pm.trade_code as supplier_code From price_master pm " & _
                  "    inner join (Select item_code from porequest_detail where porequest_no = '" & Trim(cborequestno.Text) & "') prd on prd.item_code = pm.item_code " & _
                  "    where pm.price_cls = '01' " & _
                  "    Union " & _
                  "    select distinct im.supplier_code from item_master im " & _
                  "    inner join (Select item_code from porequest_detail where porequest_no = '" & Trim(cborequestno.Text) & "') prd on prd.item_code = im.item_code " & _
                  "    Union " & _
                  "    select distinct pom.supplier_code from PurchaseOrder_Master pom " & _
                  "    inner join (select PO_No from PurchaseOrder_Detail where porequest_no = '" & Trim(cborequestno.Text) & "') pod on pod.PO_No = pom.PO_No " & _
                  "    where isnull(pom.others_Cls,'0') = '0' and pom.period is null " & _
                  ") sc on sc.supplier_code = tm.trade_code " & _
                  "where (tm.trade_cls='2') " & _
                  "Order By tm.Trade_Abbr "
    Else
        sqlcust = "select trade_code, trade_name, address1, country_cls, po_cls, popayment_code, popayment_day, popayment_terms, transportation_cls, isnull(Trade_Abbr,'') Trade_Abbr " & _
                  "from trade_master where (trade_cls='2') Order By Trade_Abbr"
    End If
    Set RsCust = Db.Execute(sqlcust)

    With cboCust
        .clear
        .columnCount = 8
        .ColumnWidths = "80pt;0pt;0pt;0pt;0pt;0pt;0pt;270pt"
        .ListWidth = 350
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

Sub adtocbo(nmCombo, nmConst, start As Integer, col1 As Integer, Col2 As Integer, Kosong As Boolean)
Dim j As Integer, k As Integer

    With nmCombo
        .clear
        .columnCount = 2
        .TextColumn = 2
        j = 0: k = start
        
        If Kosong Then
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
    Call up_FillCombo(cbocurr, "Curr_Cls")
    cbocurr.TextColumn = 2
    
    'Call adtocbo(cboPaymentTerm, isiPOTerm, 1, 0, 140, False)
    'Call adtocbo(cboTransport, isiTransport, 0, 0, 80, False)

    'PRICE CONDITION
    sql1 = "select * from PriceCondition_cls"
    If rs1.State <> adStateClosed Then rs1.Close
    rs1.Open sql1, Db, adOpenKeyset, adLockOptimistic
    If Not (rs1.BOF And rs1.EOF) Then
        i = 0
        With cboPriceCondition
            .clear
            .columnCount = 2
            .ColumnWidths = "25pt;175pt"
            .ListWidth = 200
            .ListRows = 15
            
            Do While Not rs1.EOF
                .AddItem ""
                .List(i, 0) = Trim(rs1!PriceCondition_Cls)
                .List(i, 1) = Trim(rs1!Description)
                i = i + 1
                rs1.MoveNext
            Loop
        End With
    End If
    Set rs1 = Nothing
        
    'PAYMENT TERM
    sql1 = "select * from PaymentTerm_cls"
    If rs1.State <> adStateClosed Then rs1.Close
    rs1.Open sql1, Db, adOpenKeyset, adLockOptimistic
    If Not (rs1.BOF And rs1.EOF) Then
        i = 0
        With cboPaymentTerm
            .clear
            .columnCount = 2
            .ColumnWidths = "25pt;175pt"
            .ListWidth = 200
            .ListRows = 15
            
            Do While Not rs1.EOF
                .AddItem ""
                .List(i, 0) = Trim(rs1!PaymentTerm_Cls)
                .List(i, 1) = Trim(rs1!Description)
                i = i + 1
                rs1.MoveNext
            Loop
        End With
    End If
    Set rs1 = Nothing
    
    'PACKING
    sql1 = "select * from PackingStyle_cls"
    If rs1.State <> adStateClosed Then rs1.Close
    rs1.Open sql1, Db, adOpenKeyset, adLockOptimistic
    If Not (rs1.BOF And rs1.EOF) Then
        i = 0
        With CboPacking
            .clear
            .columnCount = 2
            .ColumnWidths = "25pt;175pt"
            .ListWidth = 200
            .ListRows = 15
            
            Do While Not rs1.EOF
                .AddItem ""
                .List(i, 0) = Trim(rs1!PackingStyle_Cls)
                .List(i, 1) = Trim(rs1!Description)
                i = i + 1
                rs1.MoveNext
            Loop
        End With
    End If
    Set rs1 = Nothing

    'INSURANCE CLS
    sql1 = "select * from Insurance_cls"
    If rs1.State <> adStateClosed Then rs1.Close
    rs1.Open sql1, Db, adOpenKeyset, adLockOptimistic
    If Not (rs1.BOF And rs1.EOF) Then
        i = 0
        With cboInsuranceCls
            .clear
            .columnCount = 2
            .ColumnWidths = "25pt;175pt"
            .ListWidth = 200
            .ListRows = 15
            
            Do While Not rs1.EOF
                .AddItem ""
                .List(i, 0) = Trim(rs1!Insurance_Cls)
                .List(i, 1) = Trim(rs1!Description)
                i = i + 1
                rs1.MoveNext
            Loop
        End With
    End If
    Set rs1 = Nothing

    'Transportation
    sql1 = "select * from Transportation_cls"
    If rs1.State <> adStateClosed Then rs1.Close
    rs1.Open sql1, Db, adOpenKeyset, adLockOptimistic
    If Not (rs1.BOF And rs1.EOF) Then
        i = 0
        With cboTransport
            .clear
            .columnCount = 2
            .ColumnWidths = "25pt;175pt"
            .ListWidth = 200
            .ListRows = 15
            
            Do While Not rs1.EOF
                .AddItem ""
                .List(i, 0) = Trim(rs1!Transportation_Cls)
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
            "from PORequest_Master where isnull(others_cls,'0')='0' and isnull(fix_cls,'0')='1' " & _
            "and porequest_date >= '" & Format(requestdate1.Value, "yyyy-mm-dd") & "' " & _
            "and porequest_date <= '" & Format(requestdate2.Value, "yyyy-mm-dd") & "' " & _
            "and sheetcoil_cls = '1' "
            
    Set rsno = Db.Execute(sqlno)
    With cborequestno
        .clear
        .columnCount = 3
        
        If Trim(cboCust.Text) <> "" Then tempcust = Trim(cboCust.Text)
        Call adtocboCust(False): cboCust.Text = tempcust: Call cboCust_Click
        
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
'            "where pom.others_cls = '0' and pom.period is null " & _
'            "and pom.supplier_Code = '" & Trim(CboCust.Text) & "' " & _
'            "and pom.po_date >= '" & Format(RequestDate1, "yyyy-mm-dd") & "' " & _
'            "and pom.po_date <= '" & Format(RequestDate2, "yyyy-mm-dd") & "' "
    
    'NO PO tidak di filter berdasarkan supplier [W0-008 12 Juni 2007]
    sqlno = "select pom.PO_No from PurchaseOrder_Master pom " & _
            "where pom.others_cls = '0' and pom.period is null and sheetcoil_cls = '1' " & _
            "and pom.po_date >= '" & Format(requestdate1, "yyyy-mm-dd") & "' " & _
            "and pom.po_date <= '" & Format(requestdate2, "yyyy-mm-dd") & "' "
    
    'tambahan dudi Januari 2009, di filter berdasar no Request
    'hanya menampilkan PO yang berdasar request saja
    If cborequestno <> "" Then
    'sqlno = sqlno & " AND PO_NO IN (SELECT PO_NO FROM PurchaseOrder_Detail WHERE PORequest_NO='" & cborequestno & "')"
    End If
            
    Set rsno = Db.Execute(sqlno)
    With CboPOnO
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

Sub PONO(ByVal thn As String, ByVal bln As String)
Dim sqlno As String, SqlS As String
Dim rsno As New Recordset, rsS As New Recordset
    
    'POYYMM999
    If Trim(txtPoNo.Text) = "" Then
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
                    "and Right(RTrim(PO_No), 9)  < '" & Right(Trim(txtPoNo.Text), 9) & "' " & _
                    "order by right(rtrim(PO_No),5) desc"
'        End If
    End If
    Set rsno = Db.Execute(sqlno)
    If Not (rsno.BOF And rsno.EOF) Then
        txtPoNo.Text = Left(Trim(rsno(0)), 4) & bln & Format(Right(Trim(rsno(0)), 5) + 1, "0000#")
    Else
' Sementara dinonaktifkan !!!
'        SqlS = "select top 1 PO_No from Initial_No"
'        Set rsS = Db.Execute(SqlS)
'        If Not (rsS.BOF And rsS.EOF) Then
'            txtpono.Text = Left(Trim(rsS(0)), 2) & thn & bln & Right(Trim(rsS(0)), 5)
'        Else
            txtPoNo.Text = "PO" & thn & bln & "00001"
'        End If
'        Set rsS = Nothing
    End If
    txtPoNo.locked = True
    Set rsno = Nothing
End Sub

Function seqNo() As Long
Dim sqlseqno As String
Dim rsseqno As New Recordset

    sqlseqno = "select Seq_No from PurchaseORder_Detail order by Seq_No desc"
    If rsseqno.State <> adStateClosed Then rsseqno.Close
    rsseqno.Open sqlseqno, Db, adOpenKeyset, adLockOptimistic
    
    If Not (rsseqno.BOF And rsseqno.EOF) Then
        seqNo = rsseqno!Seq_no + 1
    Else
        seqNo = 1
    End If
    Set rsseqno = Nothing
End Function

Sub kunci(l As Boolean)
    PODate.Enabled = Not l
    grid.Editable = Not l
    Command1(0).Enabled = Not l
    lblFix.Caption = "Status Fix"
    lblFix.Visible = l
End Sub

Sub ppn(ByVal d As Date)
Dim sqlppn As String
Dim rsppn As New ADODB.Recordset
    
    sqlppn = "select rate from tax_cls where tax_code='PPN' and " & _
             "start_date <= '" & Format(d, "yyyymmdd") & "' and " & _
             "end_date >= '" & Format(d, "yyyymmdd") & "' "
    Set rsppn = Db.Execute(sqlppn)
    If Not (rsppn.BOF And rsppn.EOF) Then
        isippn = IIf(IsNull(rsppn(0)), 0, CDbl(rsppn(0)))
    Else
        isippn = 0
    End If
    Set rsppn = Nothing
End Sub

Function cekprice(ByVal Baris As Integer) As Boolean
Dim sqlcp As String
Dim rsCP As New Recordset
    
    cekprice = False
    sqlcp = "select price from price_master " & _
            "where item_code = '" & grid.TextMatrix(Baris, 2) & "' and price_cls = '01' " & _
            "and (trade_code = '" & cboCust.Text & "' or trade_code = '000000') " & _
            "and start_date <= '" & Format(grid.TextMatrix(Baris, 11), "yyyymmdd") & "' " & _
            "and end_date >= '" & Format(grid.TextMatrix(Baris, 11), "yyyymmdd") & "' "
    Set rsCP = Db.Execute(sqlcp)
    If Not (rsCP.BOF And rsCP.EOF) Then
        Do While Not rsCP.EOF
            If rsCP(0) = 0 Then cekprice = True: Exit Function
            rsCP.MoveNext
        Loop
    End If
    Set rsCP = Nothing
End Function

Sub browseprice()
Dim sql2 As String, rs2 As New Recordset
Dim tgldel As String
    
    If Trim(grid.TextMatrix(actrow, 11)) = "" Then
        tgldel = Trim(grid.TextMatrix(actrow, 18))
    Else
        tgldel = Format(grid.TextMatrix(actrow, 11), "yyyymmdd")
    End If
    
    sql2 = "select trade_code, priority_cls, isnull(currency_code,'') currency_code, price, unit_cls " & _
           "from price_master " & _
           "where item_code = '" & grid.TextMatrix(actrow, 2) & "' and price_cls = '01' " & _
           "and (trade_code = '" & cboCust.Text & "' or trade_code = '000000') " & _
           "and start_date <= '" & tgldel & "' " & _
           "and end_date >= '" & tgldel & "' " & _
           "order by trade_code desc, priority_cls desc"
    Set rs2 = Db.Execute(sql2)

    With cboprice
        .clear
        .columnCount = 4
        .ColumnWidths = "70pt;70pt;0pt;0pt"
        .ListWidth = 140
        .ListRows = 4
        
        i = 0
        Do While Not rs2.EOF
            .AddItem
            .List(i, 0) = Format(Trim(rs2("price")), "##,##0.00###")
            If rs2("trade_code") = "000000" Then
              .List(i, 1) = "Common " & Trim(rs2("priority_cls"))
            Else
              .List(i, 1) = "Priority " & Trim(rs2("priority_cls"))
            End If
            .List(i, 2) = Trim(rs2("Currency_Code"))
            .List(i, 3) = Trim(rs2("unit_cls"))

            rs2.MoveNext
            i = i + 1
        Loop
    End With
    Set rs2 = Nothing
End Sub

Sub formatprice()
Dim p1 As Byte, p2 As String, p0 As String
Dim jmldigit As Byte, jmldigit0 As Byte, j As Integer

    jmldigit = 0
    With grid
        For i = 1 To .Rows - 1
            If InStr(1, .TextMatrix(i, 14), ".") > 0 Then _
                jmldigit0 = Len(Trim(.TextMatrix(i, 14))) - InStr(1, Trim(.TextMatrix(i, 14)), ".")
            If jmldigit0 > jmldigit Then jmldigit = jmldigit0
        Next i

        For i = 1 To .Rows - 1
            p0 = Trim(.TextMatrix(i, 14))
            If InStr(1, p0, ".") > 0 Then
                p1 = Len(p0) - InStr(1, p0, ".")
                For j = 1 To jmldigit - p1
                    p2 = p0 & " "
                    p0 = p2
                Next j
            End If
            .TextMatrix(i, 14) = p0
        Next i
    End With
End Sub

Sub Header()
    With grid
        .clear
        .Rows = 2
        .ColS = 25
        
        .ColWidth(0) = 300
        .ColWidth(1) = 1320
        .ColWidth(2) = 2000
        .ColWidth(3) = 2500
        
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        
        .ColHidden(7) = True
        .ColWidth(8) = 480
        .ColWidth(9) = 930
        .ColWidth(10) = 750
        .ColWidth(11) = 1170
        .ColWidth(12) = 1170
        .ColWidth(13) = 1125
        .ColWidth(14) = 1440
        .ColHidden(15) = True
        .ColWidth(16) = 800
        .ColWidth(17) = 1500
        .ColWidth(18) = 1800
        .ColHidden(19) = True 'POReq SeqNo
        .ColHidden(20) = True 'Purpose
        .ColHidden(21) = True 'POReq Delivery Date (yyyymmdd)
        .ColHidden(22) = True 'Seq No
        .ColHidden(23) = True 'Department Cls
        .ColHidden(24) = True 'Account No
        
        .ColWidth(24) = 1500
        
        .TextMatrix(0, 0) = " "
        .TextMatrix(0, 1) = "Request No"
        .TextMatrix(0, 2) = "Product Code"
        .TextMatrix(0, 3) = "Steel Kind"
        .TextMatrix(0, 4) = "Material Measure"
        .TextMatrix(0, 5) = "Material Measure"
        .TextMatrix(0, 6) = "Material Measure"
        .TextMatrix(0, 8) = "Unit"
        .TextMatrix(0, 9) = "Qty / Box"
        .TextMatrix(0, 10) = "Lot Qty"
        .TextMatrix(0, 11) = "Request Qty"
        .TextMatrix(0, 12) = "Order"
        .TextMatrix(0, 13) = "Remaining"
        .TextMatrix(0, 14) = "Delivery Date"
        .TextMatrix(0, 16) = "Curr"
        .TextMatrix(0, 17) = "Price"
        .TextMatrix(0, 18) = "Amount"
        .TextMatrix(0, 24) = "Account No."
        
        .TextMatrix(1, 0) = " "
        .TextMatrix(1, 1) = "Request No"
        .TextMatrix(1, 2) = "Product Code"
        .TextMatrix(1, 3) = "Steel Kind"
        .TextMatrix(1, 4) = "T"
        .TextMatrix(1, 5) = "W"
        .TextMatrix(1, 6) = "M"
        .TextMatrix(1, 8) = "Unit"
        .TextMatrix(1, 9) = "Qty / Box"
        .TextMatrix(1, 10) = "Lot Qty"
        .TextMatrix(1, 11) = "Request Qty"
        .TextMatrix(1, 12) = "Order"
        .TextMatrix(1, 13) = "Remaining"
        .TextMatrix(1, 14) = "Delivery Date"
        .TextMatrix(1, 16) = "Curr"
        .TextMatrix(1, 17) = "Price"
        .TextMatrix(1, 18) = "Amount"
        .TextMatrix(1, 24) = "Account No."
        
        
'        .ColHidden(bteColCurr) = (bteHakPrice = 0)
'        .ColHidden(bteColPrice) = (bteHakPrice = 0)
'        .ColHidden(bteColAmount) = (bteHakPrice = 0)
        
        .MergeRow(0) = True
        .MergeRow(1) = True
        
        For i = 0 To .ColS - 1
            .MergeCol(i) = True
        Next i
        .MergeCells = flexMergeFixedOnly
        
        
        .Cell(flexcpAlignment, 0, 0, 1, 24) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(8) = flexAlignCenterCenter
        
        For i = 9 To 13
          .ColAlignment(i) = flexAlignRightCenter
        Next i
        
        .ColAlignment(14) = flexAlignCenterCenter
        .ColAlignment(16) = flexAlignCenterCenter
        .ColAlignment(17) = flexAlignRightCenter
        .ColAlignment(18) = flexAlignRightCenter
        .ColAlignment(24) = flexAlignCenterCenter
        
        
        .ColHidden(7) = True
        .ColHidden(15) = True
        
        .RowHeight(0) = 225
        .RowHeight(1) = 225
        
        .FrozenCols = 7

        .ColHidden(19) = True
        .ColHidden(20) = True
        .ColHidden(21) = True
        .ColHidden(22) = True
        .ColHidden(23) = True
        .ColHidden(24) = True
        
        .RowHeight(0) = 225
        .RowHeight(1) = 225
        
    End With
    sampun = False
End Sub

Sub browseitem()
Dim sqlitem As String, RsItem As New ADODB.Recordset
Dim nextperiod As Date, endtgl As Integer, endperiod As Date, tglperiod As Date
    
    LblErrMsg = ""
    
'    Call header
    
    If ubah = False Then
        ' Add 20090112
        TxtSubAmount.Text = 0
        TxtDisc.Text = 0
        ' ---
        txtamount.Text = 0
        txtPPN.Text = 0
        txtGrandTotal.Text = 0
    End If
    activecurr = ""
    activecurrcd = ""
    ubahgrid = False
    i = 2
    
    If cborequestno.Text <> "" And cborequestno.MatchFound Then
    tglperiod = Left(cborequestno.Column(1), 4) & "-" & Right(cborequestno.Column(1), 2) & "-01"
    nextperiod = DateAdd("m", 1, tglperiod)
    endtgl = DateDiff("d", Format(tglperiod, "yyyy-mm-01"), Format(nextperiod, "yyyy-mm-01"))
    endperiod = Year(tglperiod) & "-" & Month(tglperiod) & "-" & Format(endtgl, "0#")
    End If
    
    'Detail PO NO with Different POREQUEST NO
    sqlitem = "select a.*,iim.*,sccm.*, " & _
              "(select description from unit_cls uc where uc.unit_cls= a.unit_cls ) unit_desc , " & _
              "(select description from curr_cls where curr_cls.Curr_cls= a.Currency_Code) Curr_desc " & _
              "From ( " & _
              "select distinct '1' No, pod.item_code, pod.unit_cls, im.item_name, " & _
              "im.finishgoodpart_cls, im.number_entering, im.number_box, im.lot_qty, im.orderpoint_qty, " & _
              "isnull(prd.qty,0) RequestQty, " & _
              "isnull( (select sum(qty) qty from PurchaseOrder_Detail pod2 " & _
              "         inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "         where pod2.item_code = pod.item_code and pod2.porequest_no = pod.PORequest_No and pod2.POReq_SeqNo = pod.POReq_SeqNo " & _
              "         and isnull(pom.others_cls,'0')='0') ,0) totalPOQty, " & _
              "isnull(prd.qty,0) - isnull( (select sum(qty) qty from PurchaseOrder_Detail pod2 " & _
              "                             inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "                             where pod2.item_code = pod.item_code and pod2.porequest_no = pod.PORequest_No " & _
              "                             and pod2.POReq_SeqNo = pod.POReq_SeqNo and isnull(pom.others_cls,'0')='0') " & _
              ",0) RemainingQty, " & _
              "prd.ReqDelivery_Date1, pod.PORequest_No, pod.POReq_SeqNo, prd.Purpose, prd.accountno, pod.Currency_Code, isnull(pod.Price,0) Price, Department_Cls "
              
   sqlitem = sqlitem & " from PurchaseOrder_Detail pod " & _
              "inner join item_master im on im.item_code = pod.item_Code " & _
              "left outer join (select prd.*, Department_Cls,isnull(prm.complete_cls,'0') Complete_Cls, " & _
              "                 cast(year(reqdelivery_date) as char(4)) + " & _
              "                 cast((case when month(reqdelivery_date) < 10 then '0' else '' end) + cast(month(reqdelivery_date) as char) as char(2)) + " & _
              "                 cast((case when day(reqdelivery_date) < 10 then '0' else '' end) + cast(day(reqdelivery_date) as char) as char(2)) as ReqDelivery_Date1 " & _
              "                 from PORequest_Detail prd " & _
              "                 inner join (select PORequest_No, Department_Cls,Complete_Cls from PORequest_Master where isnull(others_cls,'0') = '0') prm " & _
              "                     on prm.porequest_no = prd.porequest_no) prd " & _
              "on prd.PORequest_No = pod.PORequest_No and prd.POReq_SeqNo = pod.POReq_SeqNo " & _
              "where pod.PO_No = '" & Trim(txtPoNo.Text) & "' and pod.PORequest_No <> (case prd.Complete_Cls when '1' then '' else '" & Trim(cborequestno.Text) & "' end) "
    
    'from PRICE MASTER and selected POREQUEST No
    sqlitem = sqlitem & _
              "UNION " & _
              "select distinct '2' No, pm.item_code, im.unit_cls, im.item_name, " & _
              "im.finishgoodpart_cls, im.number_entering, im.number_box, im.lot_qty, im.orderpoint_qty, " & _
              "isnull(prd.qty,0) RequestQty, " & _
              "isnull( (select sum(qty) qty from PurchaseOrder_Detail pod " & _
              "         inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "         where pod.item_code = pm.item_code and pod.porequest_no = '" & Trim(cborequestno.Text) & "' and pod.POReq_Seqno = prd.POReq_SeqNo and isnull(pom.others_cls,'0')='0') " & _
              ",0) totalPOQty, " & _
              "isnull(prd.qty,0) - isnull( (select sum(qty) qty from PurchaseOrder_Detail pod " & _
              "                             inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "                             where pod.item_code = pm.item_code and pod.porequest_no = '" & Trim(cborequestno.Text) & "' and pod.POReq_Seqno = prd.POReq_SeqNo and isnull(pom.others_cls,'0')='0') " & _
              ",0) RemainingQty, " & _
              "prd.ReqDelivery_Date1, prd.PORequest_No, prd.POReq_SeqNo, prd.Purpose, prd.accountno, " & _
              "(select top 1 currency_code from Price_Master p " & _
              " where p.price_cls = '01' and p.item_code = pm.item_code and p.start_date <= prd.ReqDelivery_Date1 and p.end_date >= prd.ReqDelivery_Date1 " & _
              " and p.trade_code in ('" & Trim(cboCust.Text) & "','000000') order by p.trade_Code desc, p.priority_Cls desc) Currency_Code, " & _
              "(select top 1 price from Price_Master p " & _
              " where p.price_cls = '01' and p.item_code = pm.item_code and p.start_date <= prd.ReqDelivery_Date1 and p.end_date >= prd.ReqDelivery_Date1 " & _
              " and p.trade_code in ('" & Trim(cboCust.Text) & "','000000') order by p.trade_Code desc, p.priority_Cls desc) Price, Department_cls " & _
              "From Price_Master pm " & _
              "inner join Item_Master im on pm.item_code = im.item_code "
    sqlitem = sqlitem & _
              "inner join (select prd.*, isnull(prm.complete_cls,'0') Complete_Cls, " & _
              "            cast(year(reqdelivery_date) as char(4)) + " & _
              "            cast((case when month(reqdelivery_date) < 10 then '0' else '' end) + cast(month(reqdelivery_date) as char) as char(2)) + " & _
              "            cast((case when day(reqdelivery_date) < 10 then '0' else '' end) + cast(day(reqdelivery_date) as char) as char(2)) as ReqDelivery_Date1, Department_cls " & _
              "            from PORequest_Detail prd " & _
              "            inner join (select PORequest_No, Department_Cls, Complete_cls from PORequest_Master where isnull(others_cls,'0') = '0') prm " & _
              "                 on prm.porequest_no = prd.porequest_no) prd " & _
              "on prd.item_code = pm.item_code " & _
              "where pm.price_cls = '01' and prd.Complete_Cls = '0' " & _
              "and pm.start_date <= prd.ReqDelivery_Date1 and pm.end_date >= prd.ReqDelivery_Date1 " & _
              "and pm.trade_code in ('" & Trim(cboCust.Text) & "','000000') and prd.porequest_no = '" & Trim(cborequestno.Text) & "' "

    'From ITEM MASTER and selected POREQUEST NO
    sqlitem = sqlitem & _
              "UNION " & _
              "select distinct '3' No, im.item_code, im.unit_cls, im.item_name, " & _
              "im.finishgoodpart_cls, im.number_entering, im.number_box, im.lot_qty, im.orderpoint_qty, isnull(prd.qty,0) RequestQty, " & _
              "isnull( (select sum(qty) qty from PurchaseOrder_Detail pod inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "         where pod.item_code = im.item_code and pod.porequest_no = '" & Trim(cborequestno.Text) & "' and pod.POReq_Seqno = prd.POReq_SeqNo and isnull(pom.others_cls,'0')='0') ,0) totalPOQty, " & _
              "isnull(prd.qty,0) - isnull( (select sum(qty) qty from PurchaseOrder_Detail pod " & _
              "                             inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "                             where pod.item_code = im.item_code and pod.porequest_no = '" & Trim(cborequestno.Text) & "' and pod.POReq_Seqno = prd.POReq_SeqNo and isnull(pom.others_cls,'0')='0') ,0) RemainingQty, " & _
              "prd.ReqDelivery_Date1, prd.PORequest_No, prd.poreq_SeqNo, prd.Purpose, prd.accountno, '' Currency_Code, Null Price, Department_Cls " & _
              "from item_master im " & _
              "inner join (select prd.*, isnull(prm.complete_cls,'0') Complete_Cls, " & _
              "            cast(year(reqdelivery_date) as char(4)) + " & _
              "            cast((case when month(reqdelivery_date) < 10 then '0' else '' end) + cast(month(reqdelivery_date) as char) as char(2)) + " & _
              "            cast((case when day(reqdelivery_date) < 10 then '0' else '' end) + cast(day(reqdelivery_date) as char) as char(2)) as ReqDelivery_Date1, Department_Cls " & _
              "            from PORequest_Detail prd " & _
              "            inner join (select PORequest_No, Department_Cls, Complete_Cls from PORequest_Master where isnull(others_cls,'0') = '0') prm " & _
              "                 on prm.porequest_no = prd.porequest_no) prd " & _
              "on prd.item_code = im.item_code " & _
              "where im.supplier_code = '" & Trim(cboCust.Text) & "' and prd.porequest_no = '" & Trim(cborequestno.Text) & "' and im.use_endday >= '" & Format(endperiod, "yyyymmdd") & "' " & _
              "and im.item_code not in " & _
              "    (select distinct pm2.item_Code From Price_Master pm2 " & _
              "     where pm2.price_cls = '01' and pm2.start_date <= prd.ReqDelivery_Date1 and pm2.end_date >= prd.ReqDelivery_Date1 " & _
              "     and pm2.trade_code in ('" & Trim(cboCust.Text) & "','000000') ) " & _
              "and prd.Complete_Cls = '0' "
    
    sqlitem = sqlitem & ") a "
' ------ Tambahan untuk menampilkan T,W dan L
        sqlitem = sqlitem & " inner join item_master iim on a.item_code=iim.item_code "
        sqlitem = sqlitem & " inner join Sheetcoil_cls Sccm on iim.sheetcoil_cls=Sccm.Sheetcoil_cls "
' -----
    
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
            
            '----
            .TextMatrix(i, 4) = IIf(IsNull(RsItem("Thickness")), "", Trim(RsItem("Thickness")))
            .TextMatrix(i, 5) = IIf(IsNull(RsItem("Width")), "", Trim(RsItem("Width")))
            .TextMatrix(i, 6) = IIf(IsNull(RsItem("Length")), "", Trim(RsItem("Length")))
            '---
            
            If IsNull(RsItem("unit_cls")) Then
              .TextMatrix(i, 7) = " "
              .TextMatrix(i, 8) = " "
            Else
              .TextMatrix(i, 7) = Trim(RsItem("Unit_cls"))
              '.TextMatrix(i, 5) = Split(isiunit, ",")(Val(Trim(RsItem("Unit_Cls"))) - 1)
              .TextMatrix(i, 8) = Trim(RsItem("Unit_desc"))
            End If
            If RsItem("finishgoodpart_cls") = "01" Then
                .TextMatrix(i, 9) = IIf(IsNull(RsItem("number_entering")), 0, Format(RsItem("number_entering"), "##,##0"))
            Else
                .TextMatrix(i, 9) = IIf(IsNull(RsItem("number_box")), 0, Format(RsItem("number_box"), "##,##0"))
            End If
            .TextMatrix(i, 10) = IIf(IsNull(RsItem("lot_qty")), 0, Format(RsItem("lot_qty"), "#,##0"))
            .TextMatrix(i, 11) = IIf(IsNull(RsItem("RequestQty")), 0, Format(RsItem("RequestQty"), "#,##0.#0"))
            .TextMatrix(i, 12) = 0 'RsItem("TotalPOQty") 'Order
            .Cell(flexcpBackColor, i, 12) = &HFFFFFF
            .TextMatrix(i, 13) = IIf(IsNull(RsItem("RemainingQty")), 0, Format(RsItem("RemainingQty"), "#,##0.#0"))
            .TextMatrix(i, 14) = ""  'Delivery Date
            .Cell(flexcpBackColor, i, 14) = vbWhite
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
                  '.TextMatrix(i, 13) = Split(isiCurr, ",")(Val(Trim(RsItem("Currency_code"))) - 1)
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
            .TextMatrix(i, 21) = IIf(IsNull(RsItem("ReqDelivery_Date1")), "", Trim(RsItem("ReqDelivery_Date1")))
            .TextMatrix(i, 22) = ""
            .TextMatrix(i, 23) = IIf(IsNull(RsItem("Department_Cls")), "", Trim(RsItem("Department_Cls")))
            .TextMatrix(i, 24) = IIf(IsNull(RsItem("AccountNo")), "", Trim(RsItem("AccountNo")))
            
            RsItem.MoveNext
            i = i + 1
        Loop
        End With
    Else
        LblErrMsg = DisplayMsg(4006)
    End If
    Set RsItem = Nothing
End Sub

Sub Browse()
Dim a As Double

    LblErrMsg = ""
    
    sql = "select * from PurchaseOrder_Master where PO_No = '" & txtPoNo.Text & "' and others_cls = '0' and period is null"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If Not (RS.BOF And RS.EOF) Then
        ada = True: ubah = True
        statusfix = IIf(IsNull(RS("fix_cls")), 0, RS("fix_cls"))
        Call browseitem
        Call BrowseGrid
        Call formatprice
        
        'Count TOTAL AMOUNT
        For i = 2 To grid.Rows - 1
            If grid.Cell(flexcpChecked, i, 0) = flexChecked Then _
                a = a + grid.TextMatrix(i, 18)
        Next i
        'Add 20090112
        TxtSubAmount.Text = a
        If (TxtSubAmount.Text <> 0) Then TxtSubAmount.Text = Format(TxtSubAmount.Text, "##,##0.#0")
        txtamount.Text = CDbl(TxtSubAmount) - CDbl(TxtDisc)
        If (txtamount.Text <> 0) Then txtamount.Text = Format(txtamount.Text, "##,##0.#0")
        ' ---
        
        If ((cboCust.Column(3) = 1) Or (cboCust.Column(3) = 2) Or (cboCust.Column(3) = 3) Or (cboCust.Column(3) = 5)) Then
            txtPPN = 0
        Else
            txtPPN.Text = CDbl(isippn / 100) * CDbl(txtamount.Text)
        End If
        If (txtPPN.Text <> 0) Then txtPPN.Text = Format(txtPPN.Text, "##,##0.#0")
        txtGrandTotal = CDbl(txtPPN.Text) + CDbl(txtamount.Text)
        If (txtGrandTotal.Text <> 0) Then txtGrandTotal.Text = Format(txtGrandTotal.Text, "##,##0.#0")
        
        If statusfix = 1 Then Call kunci(True) Else Call kunci(False)
    Else
        ada = False
    End If
End Sub

Sub BrowseGrid()
Dim g As Integer
    
    sqlGrid = " select (select description from unit_cls uc where uc.unit_cls= PurchaseOrder_Detail.unit_cls ) unit_desc , " & _
              " (select description from curr_cls where curr_cls.Curr_cls= PurchaseOrder_Detail.Currency_Code) Curr_desc, " & _
              " * from PurchaseOrder_Detail where PO_No = '" & Trim(txtPoNo.Text) & "' order by item_code"
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic

    With grid
    Do While Not rsGrid.EOF
        For g = 2 To .Rows - 1
            If Trim(.TextMatrix(g, 2)) = Trim(rsGrid("Item_Code")) And Trim(.TextMatrix(g, 1)) = Trim(rsGrid("PORequest_No")) And Trim(.TextMatrix(g, 19)) = CStr(rsGrid("POReq_SeqNo")) Then
                ubahgrid = True
                .Cell(flexcpChecked, g, 0) = flexChecked
                .TextMatrix(g, 7) = Trim(rsGrid("Unit_cls"))
                '.TextMatrix(g, 5) = Split(isiunit, ",")(Val(Trim(rsGrid("Unit_Cls"))) - 1)
                .TextMatrix(g, 8) = Trim(rsGrid("Unit_desc"))
                .TextMatrix(g, 12) = IIf(IsNull(rsGrid("qty")), 0, Format(Trim(rsGrid("qty")), "##,##0.#0"))
                .TextMatrix(g, 14) = IIf(IsNull(rsGrid("Delivery_Date")), "", Format(rsGrid("Delivery_Date"), "dd MMM yyyy"))
                If IsNull(rsGrid("currency_code")) Then
                    .TextMatrix(g, 15) = ""
                    .TextMatrix(g, 16) = ""
                Else
                    .TextMatrix(g, 15) = Trim(rsGrid("currency_code"))
                    '.TextMatrix(g, 13) = Split(isiCurr, ",")(Val(Trim(rsGrid("Currency_code"))) - 1)
                    .TextMatrix(g, 16) = Trim(rsGrid("Curr_Desc"))
                End If
                activecurrcd = .TextMatrix(g, 15)
                activecurr = .TextMatrix(g, 16)
                .TextMatrix(g, 17) = IIf(IsNull(rsGrid("Price")), 0, Format(Trim(rsGrid("Price")), "##,##0.00###"))
                .TextMatrix(g, 18) = IIf(IsNull(rsGrid("Amount")), 0, Format(rsGrid("Amount"), "##,##0.#0"))
                .TextMatrix(g, 22) = rsGrid("Seq_no")
'                .TextMatrix(g, 24) = Trim(rsGrid("AccountNo") & "")
            End If
        Next g
        rsGrid.MoveNext
    Loop
    End With
End Sub

Sub BrowseAtas()
    sql = "select * from PurchaseOrder_Master where PO_No = '" & Trim(txtPoNo.Text) & "' and isnull(others_cls,'0') = '0' and period is null"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic

    If Not (RS.BOF And RS.EOF) Then
        PODate.Value = IIf(IsNull(RS("po_date")), "", Format(Trim(RS("po_date")), "dd MMM yyyy"))
        cboCust.Text = Trim(RS("Supplier_code"))
        txtRev.Text = IIf(IsNull(RS("revise_No")), "", Trim(RS("revise_No")))
        cboPriceCondition.Text = IIf(IsNull(RS("PriceCondition_Cls")), "", Trim(RS("PriceCondition_Cls")))
        cboPaymentTerm.Text = IIf(IsNull(RS("PaymentTerm_cls")), "", RS("PaymentTerm_cls"))
        CboPacking.Text = IIf(IsNull(RS("POPacking_Cls")), "", RS("POPacking_Cls"))
        cboInsuranceCls.Text = IIf(IsNull(RS("Insurance_Cls")), "", RS("Insurance_Cls"))
        cboTransport.Text = IIf(IsNull(RS("Transportation_Cls")), "", RS("Transportation_Cls"))
        txtMarking(0).Text = IIf(IsNull(RS("POMarking1")), "", Trim(RS("PoMarking1")))
        txtMarking(1).Text = IIf(IsNull(RS("POMarking2")), "", Trim(RS("PoMarking2")))
        txtMarking(2).Text = IIf(IsNull(RS("POMarking3")), "", Trim(RS("PoMarking3")))
        txtMarking(3).Text = IIf(IsNull(RS("POMarking4")), "", Trim(RS("PoMarking4")))
        txtMarking(4).Text = IIf(IsNull(RS("POMarking5")), "", Trim(RS("PoMarking5")))
        txtMarking(5).Text = IIf(IsNull(RS("POMarking6")), "", Trim(RS("PoMarking6")))
        txtremarks.Text = IIf(IsNull(RS("remarks")), "", Trim(RS("remarks")))
        TxtDisc.Text = Format(RS("Discount"), gs_formatAmount)
        
'        cbopocode.ListIndex = IIf(IsNull(rs("POPayment_Code")), -1, rs("POPayment_Code") - 1)
'        txtpoday.Text = IIf(IsNull(rs("popayment_day")), "", rs("popayment_day"))
'        txtremarks(0).Text = IIf(IsNull(rs("remarks")), "", Trim(rs("remarks")))
'        txtremarks(1).Text = IIf(IsNull(rs("remarks2")), "", Trim(rs("remarks2")))
'        txtremarks(2).Text = IIf(IsNull(rs("remarks3")), "", Trim(rs("remarks3")))
        statusfix = IIf(IsNull(RS("fix_cls")), 0, RS("fix_cls"))
        If statusfix = 1 Then Call kunci(True) Else Call kunci(False)
    End If
End Sub

Function cekrecqty(ItemCode As String, PONO As String) As Double
Dim sqlcek As String, rsCek As New Recordset
    
    cekrecqty = 0
    sqlcek = "select item_code, sum(qty) recqty from Part_Receipt " & _
             "where PO_No = '" & Trim(PONO) & "' and item_code = '" & Trim(ItemCode) & "' " & _
             "group by item_code "
    If rsCek.State <> adStateClosed Then rsCek.Close
    rsCek.Open sqlcek, Db, adOpenKeyset, adLockOptimistic
    If Not (rsCek.BOF And rsCek.EOF) Then _
        cekrecqty = CDbl(rsCek("recqty"))
    Set rsCek = Nothing
End Function

Private Sub cboInsuranceCls_Change()
    If cboInsuranceCls.Text = "" Then cboInsuranceCls.Text = ""
End Sub

Private Sub cboInsuranceCls_Click()
    If cboInsuranceCls.ListIndex <> -1 Then
        txtInsurance.Text = cboInsuranceCls.Column(1)
    Else
        txtInsurance.Text = ""
    End If
End Sub

Private Sub cboInsuranceCls_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboInsuranceCls_Click
End Sub

Private Sub cboPacking_Change()
    If CboPacking.Text = "" Then CboPacking.Text = ""
End Sub

Private Sub cboPacking_Click()
    If CboPacking.ListIndex <> -1 Then
        TxtPacking.Text = CboPacking.Column(1)
    Else
        TxtPacking.Text = ""
    End If
End Sub

Private Sub cboPacking_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboPacking_Click
End Sub

Private Sub cboPaymentTerm_Change()
    If cboPaymentTerm.Text = "" Then cboPaymentTerm.Text = ""
End Sub

Private Sub cboPaymentTerm_Click()
    If cboPaymentTerm.ListIndex <> -1 Then
        txtPaymentTerm.Text = cboPaymentTerm.Column(1)
    Else
        txtPaymentTerm.Text = ""
    End If

End Sub

Private Sub cboPaymentTerm_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboPaymentTerm_Click
End Sub

Private Sub cboTransport_Change()
    If cboTransport.Text = "" Then cboTransport.Text = ""
End Sub

Private Sub cbotransport_Click()
    If cboTransport.ListIndex <> -1 Then
        TxtTransport.Text = cboTransport.Column(1)
    Else
        TxtTransport.Text = ""
    End If
End Sub

Private Sub cbotransport_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbotransport_Click
End Sub

Private Sub Form_Load()
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

    Call adtocboCust(False)
    Call adtocombo
    Call Kosong
    combo1.ListIndex = 1
End Sub


Private Sub requestdate1_Change()
Dim ketemu As Boolean

    LblErrMsg.Caption = ""
    If Format(requestdate1, "yyyy-mm-dd") > Format(requestdate2, "yyyy-mm-dd") Then
       LblErrMsg.Caption = DisplayMsg(4025) & " " & Format(requestdate2, "MMM yyyy") '"Start Date must be lower than "
       Exit Sub
    End If
    
    Call adtocborequestno
        
    If combo1.ListIndex = 1 Then    'UPDATE
        If cboCust.Text <> "" Then Call adtocbopono
        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtPoNo.Text = ""
    End If
    
'    txtpono.Text = ""
'    cbopono.clear
    Call Header
    Call kosongBwh
End Sub

Private Sub requestdate2_Change()
Dim ketemu As Boolean

    LblErrMsg.Caption = ""
    If Format(requestdate2, "yyyy-mm-01") < Format(requestdate1, "yyyy-mm-01") Then
       LblErrMsg.Caption = DisplayMsg(4024) & " " & Format(requestdate1, "MMM yyyy") '"End Date must be higher than "
       Exit Sub
    End If

    Call adtocborequestno
    
    If combo1.ListIndex = 1 Then    'UPDATE
        If cboCust.Text <> "" Then Call adtocbopono
        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtPoNo.Text = ""
    End If

'    txtpono.Text = ""
'    cbopono.clear
    'Call header
    Call kosongBwh
End Sub

Private Sub cborequestno_Click()
Dim ketemu As Boolean, tempcust As String

    LblErrMsg = ""
    If cborequestno.ListIndex <> -1 Then
        If Trim(cboCust.Text) <> "" Then tempcust = Trim(cboCust.Text)
        If cborequestno.Text <> "" Then
            Call adtocboCust(True)
            Call adtocbopono
        Else
            'Call adtocboCust(False)
        End If
        cboCust.Text = tempcust: Call cboCust_Click
        
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
            Call kosongBwh: Call Header
        End If
    Else
'        cbopono.clear
        If Trim(cboCust.Text) <> "" Then tempcust = Trim(cboCust.Text)
        If cborequestno.Text <> "" Then Call adtocboCust(True) Else Call adtocboCust(False)
        cboCust.Text = tempcust: Call cboCust_Click
        
        If combo1.ListIndex = 1 Then 'UPDATE
            Call kosongBwh
            'Call header
'            txtpono.Text = ""
        End If
        If cborequestno.Text <> "" Then LblErrMsg.Caption = DisplayMsg(4144) '"Record with this Request No not found"
    End If
End Sub

Private Sub cborequestno_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cborequestno_Click
End Sub

Private Sub cborequestno_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboCust_Click()
Dim ketemu As Boolean

    LblErrMsg = ""
    ketemu = False
    Call kunci(False)

    If cboCust.ListIndex <> -1 Then
        lblcust(0).Text = cboCust.Column(7)
        lblcust(1).Text = cboCust.Column(2)
        countrycls = cboCust.Column(3)
        If Split(cboCust.Column(5), ",")(2) = "" Then cboPaymentTerm.ListIndex = -1 Else cboPaymentTerm.ListIndex = Split(cboCust.Column(5), ",")(2) - 1
        If cboCust.Column(6) = "" Then cboTransport.ListIndex = -1 Else cboTransport.ListIndex = cboCust.Column(6)
        
        If combo1.ListIndex = 1 Then    'UPDATE
            Call ClearData
            'Call adtocbopono
            For i = 0 To CboPOnO.ListCount - 1
                If txtPoNo.Text = CboPOnO.List(i) Then
                    ketemu = True
                    'cbopono.ListIndex = i
                    Exit For
                End If
            Next i
            If ketemu = False Then txtPoNo.Text = ""
            Call kosongBwh
            'Call header
        End If
    Else
        lblcust(0).Text = ""
        lblcust(1).Text = ""
        countrycls = 3
        cboPaymentTerm.ListIndex = -1
        cboTransport.ListIndex = -1
        'cbopono.clear
        If combo1.ListIndex = 1 Then 'UPDATE
            Call kosongBwh
            'Call header
            txtPoNo.Text = ""
        End If
        If cboCust.Text <> "" Then LblErrMsg.Caption = DisplayMsg(4050) '"Record with this Supplier Code not Exist"
        Exit Sub
    End If
        
    If (countrycls = 1 Or countrycls = 2) Then  'OVERSEAS
        isippn = 0
        txtPPN.Text = 0
        txtGrandTotal.Text = txtamount.Text
        If (txtGrandTotal.Text <> 0) Then txtGrandTotal.Text = Format(txtGrandTotal.Text, "##,##0.#0")
    Else 'DOMESTIC
        Call ppn(PODate.Value)
        txtPPN.Text = CDbl(isippn / 100) * CDbl(txtamount.Text)
        If (txtPPN.Text <> 0) Then txtPPN.Text = Format(txtPPN.Text, "##,##0.#0")
        txtGrandTotal = CDbl(txtPPN.Text) + CDbl(txtamount.Text)
        If (txtGrandTotal.Text <> 0) Then txtGrandTotal.Text = Format(txtGrandTotal.Text, "##,##0.#0")
    End If
End Sub

Private Sub cbocust_LostFocus()
    If sampun = False Then Call cboCust_Click   'sampun->false=tidak ada data di grid
End Sub

Private Sub cboCust_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboCust_Click
End Sub

Private Sub cbocust_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Combo1_Click()
Dim ketemu As Boolean

    LblErrMsg = ""
    ketemu = False
    Call kunci(False)
    Call kosongBwh
    Call Header

   If combo1.ListIndex = 0 Then    'CREATE
        Call ClearData
        Command1(2).Caption = "&Create"
        ubah = False
        CboPOnO.locked = True
        txtPoNo.Text = ""
        PODate.Value = Format(Now, "dd MMM yyyy")
        PODate.Enabled = False
        Call PONO(Right(Year(PODate), 2), Format(Month(PODate), "0#"))
        cboPriceCondition.ListIndex = -1
'        txtRemarks(0).Text = "": txtRemarks(1).Text = "": txtRemarks(2).Text = ""
    Else    'UPDATE
        If cboCust.Text = "" Then   'Or cborequestno.Text = ""
            CboPOnO.clear
            txtPoNo.Text = ""
        Else
            Call adtocbopono
        End If
        ubah = True
        Command1(2).Caption = "&Update"
        CboPOnO.locked = False
        txtPoNo.locked = False
        PODate.Enabled = True

        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtPoNo.Text = ""
    End If
End Sub

Private Sub combo1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call Combo1_Click
End Sub

Private Sub cbopono_Click()
    LblErrMsg = ""
    txtPoNo.Text = CboPOnO.Text
    Call Header
    Call kosongBwh
    Call BrowseAtas
End Sub

Private Sub cbopono_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbopono_Click
End Sub

Private Sub TxtDisc_Change()
    If InStr(1, TxtDisc.Text, ",") = 1 Then TxtDisc.Text = Right(TxtDisc, Len(TxtDisc) - 1)
    If TxtDisc <> "" And TxtSubAmount <> "" And IsNumeric(TxtDisc) And IsNumeric(TxtSubAmount) Then txtamount.Text = Format(CDbl(TxtSubAmount.Text) - CDbl(TxtDisc.Text), "##,##0.#0")
End Sub

Private Sub TxtDisc_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    If InStr(1, TxtDisc, ".") > 0 Then If KeyAscii = Asc(".") Then KeyAscii = 0
End Sub

Private Sub TxtDisc_LostFocus()
txtamount = CDbl(TxtSubAmount) - CDbl(TxtDisc)
TxtDisc = Format(TxtDisc, gs_formatAmount)
End Sub

Private Sub txtpono_Change()
Dim ketemu As Boolean

    txtPONo2.Text = txtPoNo.Text
    If combo1.ListIndex = 1 Then
        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next
        If ketemu = False Then CboPOnO.ListIndex = -1
    End If
End Sub

Private Sub txtpono_KeyPress(KeyAscii As Integer)
    LblErrMsg = ""
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then
        If combo1.ListIndex = 0 Then
            SendKeys vbTab
        Else
            'Call header
            Call kosongBwh
            Call BrowseAtas
        End If
    End If
End Sub

Private Sub PODate_Change()
    LblErrMsg = ""
    'CREATE
    If combo1.ListIndex = 0 Then _
        Call PONO(Right(Year(PODate), 2), Format(Month(PODate), "0#"))
    If (countrycls = 1 Or countrycls = 2) Then isippn = 0 Else Call ppn(PODate.Value)
End Sub

Private Sub deldate_Change()
    LblErrMsg = ""
    grid.TextMatrix(actrow, 14) = Format(DelDate, "dd mmm yyyy")
End Sub

Private Sub deldate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub deldate_LostFocus()
    grid.TextMatrix(grid.RowSel, 14) = Format(DelDate, "dd mmm yyyy")
    DelDate.Visible = False
End Sub

Private Sub cbocurr_Click()
    If cbocurr.ListIndex <> -1 Then
        grid.TextMatrix(actrow, 12) = cbocurr.Column(0)
        grid.TextMatrix(actrow, 13) = cbocurr.Column(1)
    End If
End Sub

Private Sub cbocurr_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbocurr_Click
End Sub

Private Sub cbocurr_LostFocus()
    cbocurr.Visible = False
End Sub

Private Sub cboprice_Change()
    If InStr(1, cboprice.Text, ",") = 1 Then cboprice.Text = Right(cboprice, Len(cboprice) - 1)
End Sub

Private Sub cboprice_Click()
    If cboprice.ListIndex <> -1 Then
        grid.TextMatrix(actrow, 12) = cboprice.Column(2)
        For i = 0 To cbocurr.ListCount - 1
            If Trim(grid.TextMatrix(actrow, 12)) = Trim(cbocurr.List(i)) Then
                cbocurr.ListIndex = i
                Exit For
            End If
        Next i
        If Trim(cboprice.Column(2)) <> "" Then grid.TextMatrix(actrow, 13) = uf_GetCurrencyDescription(cboprice.Column(2))
        'Grid.TextMatrix(actrow, 4) = cboprice.Column(3)
        'Grid.TextMatrix(actrow, 5) = uf_GetUnitDescription(cboprice.Column(3))
    End If
End Sub

Private Sub cboprice_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboprice_Click
End Sub

Private Sub CboPrice_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    If InStr(1, cboprice.Text, ".") > 1 Then _
        If KeyAscii = Asc(".") Then KeyAscii = 0
End Sub

Private Sub cboPrice_LostFocus()
Dim ketemu As Boolean
Dim z
    
    If cboprice.Text = "" Then cboprice.Text = 0
    z = CDec(cboprice.Text)
    If z > 9999999999.99 Then cboprice.Text = Left(z, 10)
        
    grid.TextMatrix(actrow, 14) = Format(cboprice.Text, "#,##0.00###")
    Call Grid_AfterEdit(actrow, 14)
    
    cboprice.Text = Format(cboprice.Text, "#,##0.00###")
    For i = 0 To cboprice.ListCount - 1
        If Trim(cboprice.Text) = Trim(cboprice.List(i)) Then
            ketemu = True
            cboprice.ListIndex = i
            Exit For
        End If
    Next i
    cboprice.Visible = False
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim a As Double

    a = 0
    With grid
        If Col = 12 Then 'ORDER QTY
            If .TextMatrix(Row, 12) = "" Then .TextMatrix(Row, 12) = 0
            If IsNumeric(.TextMatrix(Row, 12)) = False Then .TextMatrix(Row, 12) = 0
            If CDbl(.TextMatrix(Row, 12)) > 9999999.99 Then LblErrMsg = DisplayMsg(4045) & " 9,999,999.99": .SetFocus: Exit Sub  '"Quantity must be lower or equal than 9,999,999.99"
            
            .TextMatrix(Row, 12) = Format(.TextMatrix(Row, 12), "#,##0.#0")
            .TextMatrix(Row, 13) = Format((CDbl(.TextMatrix(Row, 13)) - CDbl(.TextMatrix(Row, 12)) + orderawal), "##,##0.#0")
        End If
        
        If Col = 0 Or Col = 12 Or Col = 17 Then
            Call formatprice
            .TextMatrix(Row, 18) = CDbl(.TextMatrix(Row, 12)) * CDbl(.TextMatrix(Row, 17))
            If .TextMatrix(Row, 18) <> 0 Then .TextMatrix(Row, 18) = Format(.TextMatrix(Row, 18), "##,##0.#0")
            
            'Count TOTAL
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, 0) = flexChecked Then
                    a = a + .TextMatrix(i, 18)
                End If
            Next i
            ' Add 20090112
            TxtSubAmount.Text = a
            If (TxtSubAmount.Text <> 0) Then TxtSubAmount.Text = Format(TxtSubAmount.Text, "##,##0.#0")

            txtamount.Text = CDbl(TxtSubAmount) - CDbl(TxtDisc)
            If (txtamount.Text <> 0) Then txtamount.Text = Format(txtamount.Text, "##,##0.#0")
            '---
             
             If ((cboCust.Column(3) = 1) Or (cboCust.Column(3) = 2) Or (cboCust.Column(3) = 3) Or (cboCust.Column(3) = 5)) Then
                  txtPPN = 0
             Else
                  txtPPN.Text = CDbl(isippn / 100) * CDbl(txtamount.Text)
             End If
            If (txtPPN.Text <> 0) Then txtPPN.Text = Format(txtPPN.Text, "##,##0.#0")
            txtGrandTotal = CDbl(txtPPN.Text) + CDbl(txtamount.Text)
            If (txtGrandTotal.Text <> 0) Then txtGrandTotal.Text = Format(txtGrandTotal.Text, "##,##0.#0")
        End If
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    actrow = Row
    If grid.Cell(flexcpChecked, Row, 0) <> flexChecked Then
        If Col <> 0 Then Cancel = True
    Else
        If Col <> 0 And Col <> 12 And Col <> 14 Then _
            Cancel = True
            'And Col <> 17 And Col <> 18
'        If cborequestno.ListIndex <> -1 Then _
'            If Col = 9 And cborequestno.Column(2) = "1" Then lblErrMsg = DisplayMsg(4089): Cancel = True
        If Col = 12 Then orderawal = CDbl(grid.TextMatrix(Row, 12))
    End If
End Sub

Private Sub Grid_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 12 Then _
        If InStr(1, grid.TextMatrix(Row, Col), ",") = 1 Then grid.TextMatrix(Row, Col) = Right(grid.TextMatrix(Row, Col), Len(grid.TextMatrix(Row, Col)) - 1)
End Sub

Private Sub grid_Click()
Dim reqdel As Date

    With grid
    If statusfix = 0 Then
        If .Row > 0 Then
            If .Cell(flexcpChecked, .Row, 0) = flexChecked Then
                If .Col = 12 Or .Col = 14 Or .Col = 16 Or .Col = 17 Then
                    .SelectionMode = flexSelectionFree
                Else
                    .SelectionMode = flexSelectionByRow
                End If
                
                If .Col = 14 Then   'DELIVERY DATE
                    DelDate.top = .Cell(flexcpTop, .Row, 14)
                    DelDate.Left = .Cell(flexcpLeft, .Row, 14)
                    DelDate.Width = .CellWidth + 30
                    DelDate.Height = .CellHeight + 30
                    If .TextMatrix(.Row, 14) <> "" Then
                        DelDate.Value = Format(.TextMatrix(.Row, 14), "yyyy-mm-dd")
                    Else
                        If Trim(.TextMatrix(.Row, 21)) <> "" Then
                            reqdel = (Left(Trim(.TextMatrix(.Row, 21)), 4) & "-" & Mid(Trim(.TextMatrix(.Row, 21)), 5, 2)) & "-" & Right(Trim(.TextMatrix(.Row, 21)), 2)
                            DelDate.Value = Format(reqdel, "yyyy-mm-dd")
                        End If
                        .TextMatrix(.Row, 21) = Format(DelDate.Value, "dd MMM yyyy")
                    End If
                    DelDate.Visible = True
                    DelDate.SetFocus
                    cbocurr.Visible = False
                    cboprice.Visible = False
                ElseIf .Col = 16 Then   'CURRENCY
'                    cbocurr.top = .Cell(flexcpTop, .Row, 16)
'                    cbocurr.Left = .Cell(flexcpLeft, .Row, 16)
'                    cbocurr.Width = .CellWidth + 30
'                    Call up_FillCombo(cbocurr, "curr_cls")
'                    'Call isiCboUnitCurr(cbocurr, isiCurr, 0, 4)
'                    cbocurr.TextColumn = 2
'                    If grid.TextMatrix(.Row, 16) <> "" Then
'                        cbocurr.Text = grid.TextMatrix(.Row, 16)
'                        For i = 0 To cbocurr.ListCount - 1
'                            If Trim(grid.TextMatrix(.Row, 15)) = Trim(cbocurr.List(i)) Then
'                                cbocurr.ListIndex = i
'                                Exit For
'                            End If
'                        Next i
'                    End If
'                    cbocurr.Visible = True
'                    cbocurr.SetFocus
                    cboprice.Visible = False
                    DelDate.Visible = False
                ElseIf .Col = 17 Then   'PRICE
'                    cboprice.top = .Cell(flexcpTop, .Row, 17)
'                    cboprice.Left = .Cell(flexcpLeft, .Row, 17)
'                    cboprice.Width = .CellWidth + 30
'                    cboprice.Text = ""
'                    Call browseprice
'                    If grid.TextMatrix(.Row, 17) <> "0" Then
'                        cboprice.Text = grid.TextMatrix(.Row, 17)
'                        For i = 0 To cboprice.ListCount - 1
'                            If Trim(grid.TextMatrix(.Row, 17)) = Trim(cboprice.List(i)) Then
'                                cboprice.ListIndex = i
'                                Exit For
'                            End If
'                        Next i
'                    End If
'                    cboprice.Visible = True
'                    cboprice.SetFocus
                    cbocurr.Visible = False
                    DelDate.Visible = False
                Else
                    cbocurr.Visible = False
                    cboprice.Visible = False
                    DelDate.Visible = False
                End If
                
                If .Col = 12 Then .FocusRect = flexFocusInset Else .FocusRect = flexFocusNone
            Else
                .SelectionMode = flexSelectionByRow
            End If
        End If
    End If
    End With
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    LblErrMsg = ""
    If Col = 12 Then
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
            KeyAscii = 0
    End If
End Sub

Private Sub Grid_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    cbocurr.Visible = False
    cboprice.Visible = False
    DelDate.Visible = False
End Sub

Private Sub grid_AfterSort(ByVal Col As Long, Order As Integer)
    cbocurr.Visible = False
    cboprice.Visible = False
    DelDate.Visible = False
End Sub

Private Sub Grid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call grid_Click
End Sub

Private Sub cbopricecondition_Click()
    If cboPriceCondition.ListIndex <> -1 Then
        txtPriceCondition.Text = cboPriceCondition.Column(1)
    Else
        txtPriceCondition.Text = ""
    End If
End Sub

Private Sub cbopricecondition_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbopricecondition_Click
End Sub

Private Sub cbopricecondition_Change()
    If cboPriceCondition.Text = "" Then txtPriceCondition.Text = ""
End Sub

Private Sub hscrollbar_Change()
Dim k As Integer

    For k = 7 To grid.ColS - 1
       grid.ColHidden(k) = False
    Next k
   
    If hscrollbar.Value = 1 Then
        For k = 7 To 13
            grid.ColHidden(k) = True
        Next k
    End If
    cboprice.Visible = False
    cbocurr.Visible = False
    DelDate.Visible = False
    grid.ColHidden(7) = True
    grid.ColHidden(15) = True
    grid.ColHidden(19) = True 'POReq SeqNo
    grid.ColHidden(20) = True 'Purpose
    grid.ColHidden(21) = True 'POReq Delivery Date (yyyymmdd)
    grid.ColHidden(22) = True 'Seq No
    grid.ColHidden(23) = True 'Department Cls
    grid.ColHidden(24) = True 'Account No
End Sub

Private Sub hscrollbar_Scroll()
    Call hscrollbar_Change
End Sub

Private Sub Command1_Click(Index As Integer)


Dim sql4 As String
Dim rs4 As New Recordset

LblErrMsg = ""
MousePointer = vbHourglass
Select Case Index
    Case 0: 'SUBMIT
            If hakUpdate(Me.Name) = 0 Then _
                LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
            
            'HEADER VALIDATION
            If cborequestno.Text = "" Then
'                cboRequestNo.SetFocus
'                LblErrMsg = DisplayMsg(1067) '"Please Input Request No"
'                Exit Sub
            ElseIf cborequestno.Text <> "" Then
                If cborequestno.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4144)    'Record with This Request No Not found !
                    cborequestno.SetFocus
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            If cboCust.Text = "" Then
                cboCust.SetFocus
                LblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
                MousePointer = vbDefault
                Exit Sub
            ElseIf cboCust.Text <> "" Then
                If cboCust.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4050)    'Record with This Supplier Code Not found !
                    cboCust.SetFocus
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            If txtPoNo.Text = "" Then
                txtPoNo.SetFocus
                MousePointer = vbDefault
                LblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                Exit Sub
            End If
                            
            'FOOTER VALIDATION
            If cboPriceCondition.Text <> "" Then
                If cboPriceCondition.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4147)    'Record with This Price Condition not found !
                    cboPriceCondition.SetFocus
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            If cboPaymentTerm.Text <> "" Then
                    If cboPaymentTerm.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4097)
                        cboPaymentTerm.SetFocus
                        MousePointer = vbDefault
                        Exit Sub
                    End If
            End If
            If cboTransport.Text <> "" Then
                If cboTransport.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4099)    '"Record with this Transportation not found"
                    cboTransport.SetFocus
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            '-------------------------------------------------------------
            
'            sQl = "select * from PurchaseOrder_Master where PO_No = '" & txtpono.Text & "' and others_cls = '0' and period is null"
'            If rs.State <> adStateClosed Then rs.Close
'            rs.Open sQl, Db, adOpenKeyset, adLockOptimistic
'            If rs.BOF And rs.EOF Then
'                lblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
'                txtpono.SetFocus
'                Exit Sub
'            End If
            
            If grid.Rows = 2 Then LblErrMsg = DisplayMsg(4047): Exit Sub  'There is no data to submit !
            
            If ubah = True Then
                If Not ValidDataSupplier(txtPoNo.Text) Then
                    LblErrMsg = "Can't change supplier! System found item[s] which no have price for this supplier. Please Input Price Master First!"
                    MousePointer = vbDefault
                    
                    Exit Sub
                End If
                RS("po_date") = Format(PODate.Value, "yyyy-mm-dd")
                RS("discount") = CDbl(TxtDisc.Text) 'Add 20090112
                RS("amount") = CDbl(txtamount.Text)
                RS("ppn") = CDbl(txtPPN.Text)
                RS("total_amount") = CDbl(txtGrandTotal.Text)
                
                RS("PoMarking1") = Trim(txtMarking(0).Text)
                RS("PoMarking2") = Trim(txtMarking(1).Text)
                RS("PoMarking3") = Trim(txtMarking(2).Text)
                RS("PoMarking4") = Trim(txtMarking(3).Text)
                RS("PoMarking5") = Trim(txtMarking(4).Text)
                RS("PoMarking6") = Trim(txtMarking(5).Text)

                RS("remarks") = Trim(txtremarks.Text)
                RS("revise_No") = Trim(txtRev.Text)
                
                If cboPriceCondition.Text = "" Then RS("pricecondition_cls") = Null Else RS("pricecondition_cls") = cboPriceCondition.Text
                If cboPaymentTerm.Text = "" Then RS("paymentterm_cls") = Null Else RS("paymentterm_cls") = cboPaymentTerm.Column(0)
                If CboPacking.Text = "" Then RS("popacking_Cls") = Null Else RS("popacking_Cls") = CboPacking.Column(0)
                If cboInsuranceCls.Text = "" Then RS("insurance_cls") = Null Else RS("insurance_cls") = cboInsuranceCls.Column(0)
                If cboTransport.Text = "" Then RS("transportation_cls") = Null Else RS("transportation_cls") = cboTransport.Column(0)
                RS.update

                Dim recqty As Double
                                
                'DETAIL VALIDATION
                With grid
                    For i = 2 To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = flexChecked Then
                            If .TextMatrix(i, 12) = 0 Then
                                hscrollbar.Value = 0
                                .SelectionMode = flexSelectionFree
                                .Col = 12: .Row = i: actrow = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(1012) '"Please Input Quantity"
                                MousePointer = vbDefault
                                Exit Sub
                            ElseIf CDbl(.TextMatrix(i, 12)) > 9999999.99 Then
                                hscrollbar.Value = 0
                                .SelectionMode = flexSelectionFree
                                .Col = 12: .Row = i: actrow = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(4045) & " 9,999,999.99" '"Quantity must be lower or equal than 9,999,999.99"
                                MousePointer = vbDefault
                                Exit Sub
                            ElseIf CDbl(.TextMatrix(i, 12)) < 0 Then
                                hscrollbar.Value = 0
                                .SelectionMode = flexSelectionFree
                                .Col = 12: .Row = i: actrow = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(4045) & " Qty Remaining" '"Quantity must be lower or equal than Qty Remaining"
                                MousePointer = vbDefault
                                Exit Sub
                            ElseIf .TextMatrix(i, 14) = "" Then
                                hscrollbar.Value = 0
                                .SelectionMode = flexSelectionFree
                                .Col = 14: .Row = i: actrow = i
                                .SetFocus
                                Call grid_Click
                                LblErrMsg = DisplayMsg(8124)    'Please Input Delivery Date
                                MousePointer = vbDefault
                                Exit Sub
                            ElseIf .TextMatrix(i, 16) = "" Then
                                hscrollbar.Value = 1
                                .SelectionMode = flexSelectionFree
                                .Col = 16: .Row = i: actrow = i
                                .SetFocus
                                Call grid_Click
                                LblErrMsg = DisplayMsg(1028)    'Please Select Currency
                                MousePointer = vbDefault
                                Exit Sub
                            ElseIf .TextMatrix(i, 17) = 0 Then
                                If cekprice(i) = False Then
                                    hscrollbar.Value = 1
                                    .SelectionMode = flexSelectionFree
                                    .Col = 17: .Row = i: actrow = i
                                    .SetFocus
                                    Call grid_Click
                                    LblErrMsg = DisplayMsg(1029) '"Please Input Price"
                                    MousePointer = vbDefault
                                    Exit Sub
                                End If
                            End If
                            
                            recqty = cekrecqty(.TextMatrix(i, 2), txtPoNo.Text)
                            If CDbl(.TextMatrix(i, 12)) < recqty Then
                                hscrollbar.Value = 0
                                .SelectionMode = flexSelectionFree
                                .Col = 12: .Row = i: actrow = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(4036) & " " & recqty     '"Quantity must be higher or equal than "
                                MousePointer = vbDefault
                                Exit Sub
                            End If
                        Else
                            sql4 = "select * from Part_Receipt where po_no = '" & txtPoNo.Text & "' and item_code = '" & .TextMatrix(i, 2) & "'"
                            Set rs4 = Db.Execute(sql4)
                            If Not (rs4.BOF And rs4.EOF) Then
                                .Row = i: actrow = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(1204)
                                MousePointer = vbDefault
                                Exit Sub
                            End If
                            Set rs4 = Nothing
                        End If
                    Next i
                                                                                                    
                    ' Cek currency harus sama
'                    Dim a As Integer, c As Integer
'                    If ubahgrid = False Or activecurr = "" Then  '*****
'                        For a = 1 To .Rows - 1
'                            If .Cell(flexcpChecked, a, 0) = flexChecked Then
'                                activecurr = .TextMatrix(a, 13): activecurrcd = .TextMatrix(a, 12): Exit For
'                            End If
'                        Next a
'                    End If
'                    For c = a + 1 To .Rows - 1
'                        If .Cell(flexcpChecked, c, 0) = flexChecked Then
'                            If .TextMatrix(c, 13) <> .TextMatrix(1, 13) Then
'                                hscrollbar.Value = 1
'                                .Col = 13: .Row = c: actrow = c
'                                .SetFocus
'                                Call Grid_Click
'                                LblErrMsg = DisplayMsg(4084)  'Cannot Select Different Currency !!
'                                Exit Sub
'                            End If
'                        End If
'                    Next c

                    Dim a As Integer, C As Integer
                    If ubahgrid = False Or activecurr = "" Then  '*****
                        For a = 2 To .Rows - 1
                            If .Cell(flexcpChecked, a, 0) = flexChecked Then
                                activecurr = .TextMatrix(a, 16): activecurrcd = .TextMatrix(a, 15): Exit For
                            End If
                        Next a
                    End If
                    
                    Dim barisan As Integer, indexawal As Integer, barisawal As Integer, jumlahBarisan As Integer
                    indexawal = 0
                    jumlahBarisan = 0
                    For barisan = 2 To .Rows - 1
                        If .Cell(flexcpChecked, barisan, 0) = flexChecked Then
                                If indexawal = 0 Then barisawal = barisan: indexawal = 1
                                jumlahBarisan = jumlahBarisan + 1
                        End If
                    Next barisan
                    
                    For barisan = 2 To .Rows - 1
                        If .Cell(flexcpChecked, barisan, 0) = flexChecked Then
                        Else
                            .TextMatrix(barisan, 16) = ""
                        End If
                    Next barisan
                    
            If jumlahBarisan > 2 Then
                    For C = a + 1 To .Rows - 1
                        If .Cell(flexcpChecked, C, 0) = flexChecked Then
                            If .TextMatrix(C, 16) <> .TextMatrix(barisawal, 16) Then
                                hscrollbar.Value = 1
                                .Col = 16: .Row = C: actrow = C
                                .SetFocus
                                Call grid_Click
                                LblErrMsg = DisplayMsg(4084)  'Cannot Select Different Currency !!
                                MousePointer = vbDefault
                                Exit Sub
                            End If
                        End If
                    Next C
            End If
                    
                    Dim rscekC As New Recordset
                    
                    'UPDATE DETAIL
                    For i = 2 To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = flexChecked Then
                            If Trim(.TextMatrix(i, 22)) = "" Then
                                sqlGrid = "select * From PurchaseOrder_Detail"
                                If rsGrid.State <> adStateClosed Then rsGrid.Close
                                rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
                                rsGrid.AddNew
                                rsGrid("Seq_no") = seqNo
                            Else
                                sqlGrid = "select * From PurchaseOrder_Detail where seq_No = '" & .TextMatrix(i, 22) & "' "
                                If rsGrid.State <> adStateClosed Then rsGrid.Close
                                rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
                            End If
                            rsGrid("po_no") = Trim(txtPoNo.Text)
                            rsGrid("PORequest_No") = Trim(.TextMatrix(i, 1))
                            rsGrid("item_Code") = Trim(.TextMatrix(i, 2))
                            rsGrid("POReq_SeqNo") = .TextMatrix(i, 19)
                            rsGrid("Delivery_Date") = Format(.TextMatrix(i, 14), "yyyy-mm-dd")
                            rsGrid("price") = CDbl(.TextMatrix(i, 17))
                            rsGrid("currency_code") = .TextMatrix(i, 15)
                            rsGrid("unit_cls") = .TextMatrix(i, 7)
                            rsGrid("qty") = CDbl(.TextMatrix(i, 12))
                            rsGrid("amount") = CDbl(.TextMatrix(i, 18))
                            'rsGrid("purpose") = Trim(.TextMatrix(i, 20)) Sudah tidak di gunakan -- dudi januari 2009
                           ' rsGrid("department_cls") = Trim(.TextMatrix(i, 23))
                           'rsGrid("accountno") = Trim(.TextMatrix(i, 24))Sudah tidak di gunakan -- dudi januari 2009
                            rsGrid.update
                        Else
                            If Trim(.TextMatrix(i, 22)) <> "" Then
                                sql = "Delete from PurchaseOrder_Detail where seq_no = '" & .TextMatrix(i, 22) & "'"
                                Db.Execute sql
                            End If
                        End If
                        
                        'UPDATE COMPLETE_CLS (POREQUEST_MASTER)
                        sql = "select PORequest_No, avg(DComplete) Complete " & _
                              "from ( select prd.PORequest_No, prd.POReq_SeqNo, isnull(prd.Qty,0) Qty, isnull(sum(pod.Qty),0) POQty, " & _
                              "        (case when isnull(prd.Qty,0) = isnull(sum(pod.Qty),0) then 1 else 0 end) DComplete " & _
                              "        from PORequest_Detail prd " & _
                              "        left outer join PurchaseOrder_Detail pod on pod.PORequest_No = prd.PORequest_No and pod.POReq_SeqNo = prd.POReq_SeqNo " & _
                              "        group by prd.PORequest_No, prd.POReq_SeqNo, prd.Qty ) a " & _
                              "where PORequest_No = '" & Trim(.TextMatrix(i, 1)) & "' " & _
                              "group by PORequest_No "
                        If rscekC.State <> adStateClosed Then rscekC.Close
                        rscekC.Open sql, Db, adOpenKeyset, adLockOptimistic
                        If Not (rscekC.BOF And rscekC.EOF) Then
                            If rscekC("Complete") = "1" Then
                                sql = "Update PORequest_Master set Complete_Cls = '1' where PORequest_No = '" & Trim(rscekC("PORequest_No")) & "' "
                            ElseIf rscekC("Complete") = "0" Then
                                sql = "Update PORequest_Master set Complete_Cls = '0' where PORequest_No = '" & Trim(rscekC("PORequest_No")) & "' "
                            End If
                            Db.Execute sql
                        End If
                    Next i
                    Call updateMaster(True)
                    Call CekPONumber
                    Call Browse
                    LblErrMsg = DisplayMsg(1101)
                    ubahgrid = True
                End With
          End If

    Case 1: 'CLEAR
            Call Kosong
            combo1.ListIndex = 1
            Call Combo1_Click
            cborequestno.SetFocus
            
    Case 2: 'CREATE / UPDATE
            If combo1.ListIndex = 0 Then    'CREATE
                If hakUpdate(Me.Name) = 0 Then _
                    LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
                
                'HEADER VALIDATION
                If cborequestno.Text = "" Then
                    cborequestno.SetFocus
                    LblErrMsg = DisplayMsg(1067) '"Please Input Request No"
                    MousePointer = vbDefault
                    Exit Sub
                ElseIf cborequestno.Text <> "" Then
                    If cborequestno.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4144)    'Record with This Request No Not found !
                        cborequestno.SetFocus
                        MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
'                If cborequestno.Column(2) = "1" Then
'                    cborequestno.SetFocus
'                    lblErrMsg = DisplayMsg(4088) '"Can't Create PO, PO Request No already Fix"
'                    Exit Sub
'                End If
                If cboCust.Text = "" Then
                    cboCust.SetFocus
                    LblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
                    MousePointer = vbDefault
                    Exit Sub
                ElseIf cboCust.Text <> "" Then
                    If cboCust.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4050)    'Record with This Supplier Code Not found !
                        cboCust.SetFocus
                        MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                If txtPoNo.Text = "" Then
                    txtPoNo.SetFocus
                    LblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                    MousePointer = vbDefault
                    Exit Sub
                End If
                                
                'FOOTER VALIDATION
                If cboPriceCondition.Text <> "" Then
                    If cboPriceCondition.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4147)    'Record with This Price Condition not found !
                        cboPriceCondition.SetFocus
                        MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
'                If cbopocode.Text <> "" Or txtpoday.Text <> "" Or cboPaymentTerm.Text <> "" Then
'                    If cbopocode.Text = "" Then
'                        cbopocode.SetFocus
'                        lblErrMsg = DisplayMsg(1062)
'                        Exit Sub
'                    ElseIf txtpoday.Text = "" Or Val(txtpoday.Text) = 0 Then
'                        txtpoday.SetFocus
'                        lblErrMsg = DisplayMsg(1062)
'                        Exit Sub
'                    ElseIf cboPaymentTerm.Text = "" Then
'                        cboPaymentTerm.SetFocus
'                        lblErrMsg = DisplayMsg(1062)
'                        Exit Sub
'                    End If
'                    If cbopocode.Text <> "" Then
'                        If cbopocode.MatchFound = False Then
'                            lblErrMsg = DisplayMsg(4097)
'                            cbopocode.SetFocus
'                            Exit Sub
'                        End If
'                    End If
                    If cboPaymentTerm.Text <> "" Then
                        If cboPaymentTerm.MatchFound = False Then
                            LblErrMsg = DisplayMsg(4097)
                            cboPaymentTerm.SetFocus
                            MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
'                End If
                If cboTransport.Text <> "" Then
                    If cboTransport.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4099)    '"Record with this Transportation not found"
                        cboTransport.SetFocus
                        MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                '-------------------------------------------------------------
                
                If ubah = False Then
                    sql = "select * from PurchaseOrder_Master where PO_No = '" & txtPoNo.Text & "'"
                    If RS.State <> adStateClosed Then RS.Close
                    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
                    If Not (RS.BOF And RS.EOF) Then
                        LblErrMsg.Caption = DisplayMsg(1023)
                        txtPoNo.SetFocus
                        MousePointer = vbDefault
                        Exit Sub
                    Else
                        RS.AddNew
                        RS("po_no") = txtPoNo.Text
                        RS("supplier_code") = cboCust.Text
                    End If
                End If
'  sampai sini
                RS("po_date") = Format(PODate.Value, "yyyy-mm-dd")
                RS("discount") = CDbl(TxtDisc.Text) 'Add 20090112
                RS("amount") = CDbl(txtamount.Text)
                RS("ppn") = CDbl(txtPPN.Text)
                RS("total_amount") = CDbl(txtGrandTotal.Text)
                
                RS("PoMarking1") = Trim(txtMarking(0).Text)
                RS("PoMarking2") = Trim(txtMarking(1).Text)
                RS("PoMarking3") = Trim(txtMarking(2).Text)
                RS("PoMarking4") = Trim(txtMarking(3).Text)
                RS("PoMarking5") = Trim(txtMarking(4).Text)
                RS("PoMarking6") = Trim(txtMarking(5).Text)
                
                RS("remarks") = Trim(txtremarks.Text)
                RS("others_cls") = "0"
                RS("sheetcoil_cls") = "1"
                RS("revise_No") = Trim(txtRev.Text)
                
                If cboPriceCondition.Text = "" Then RS("pricecondition_cls") = Null Else RS("pricecondition_cls") = cboPriceCondition.Text
                If cboPaymentTerm.Text = "" Then RS("paymentterm_cls") = Null Else RS("paymentterm_cls") = cboPaymentTerm.Column(0)
                If CboPacking.Text = "" Then RS("popacking_Cls") = Null Else RS("popacking_Cls") = CboPacking.Column(0)
                If cboInsuranceCls.Text = "" Then RS("insurance_cls") = Null Else RS("insurance_cls") = cboInsuranceCls.Column(0)
                If cboTransport.Text = "" Then RS("transportation_cls") = Null Else RS("transportation_cls") = cboTransport.Column(0)
                
'On Error Resume Next
                RS.update
errHandler:
                If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                    Call PONO(Right(Year(PODate), 2), Format(Month(PODate), "0#"))
                    txtPONo2.Text = txtPoNo.Text
                    RS("po_No") = txtPoNo.Text
                    RS.update
                    If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                        GoTo errHandler
                    Else
                        If Trim$(err.Description) <> "" Then
                            LblErrMsg = Trim$(err.number) + " : " + Trim$(err.Description)
                            MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                Else
                    If Trim$(err.Description) <> "" Then
                        LblErrMsg = Trim$(err.number) + " : " + Trim$(err.Description)
                        MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
    
                If CDate(PODate.Value) > CDate(requestdate1.Value) Then
                    If CDate(PODate.Value) > CDate(requestdate2.Value) Then _
                        requestdate2.Value = Format(PODate.Value, "dd MMM yyyy")
                Else
                    requestdate1.Value = Format(PODate.Value, "dd MMM yyyy")
                End If
                
                combo1.Text = "Update"
                If cboCust.Text <> "" And cborequestno.Text <> "" Then Call browseitem: Call formatprice
                LblErrMsg.Caption = DisplayMsg(1000)
                ubah = True
                MousePointer = vbDefault
    
            Else    'UPDATE
            If sampun Then                      'sampun=true->ada data di grid; false->tidak ada data di grid
                Call updateMaster(False)
            Else
                Dim ketemu As Boolean
                
                If cborequestno.Text = "" Then
'                    cboRequestNo.SetFocus
'                    LblErrMsg = DisplayMsg(1067) '"Please Input Request No"
'                    Exit Sub
                ElseIf cborequestno.Text <> "" Then
                    If cborequestno.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4144)    'Record with This Request No Not found !
                        cborequestno.SetFocus
                        MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                If cboCust.Text = "" Then
                    cboCust.SetFocus
                    LblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
                    MousePointer = vbDefault
                    Exit Sub
                ElseIf cboCust.Text <> "" Then
                    If cboCust.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4050)    'Record with This Supplier Code Not found !
                        cboCust.SetFocus
                        MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                
                If txtPoNo.Text = "" Then
                    txtPoNo.SetFocus
                    LblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                       MousePointer = vbDefault
                    Exit Sub
                End If
                
                If CDate(PODate.Value) > CDate(requestdate1.Value) Then
                    If CDate(PODate.Value) > CDate(requestdate2.Value) Then _
                        requestdate2.Value = Format(PODate.Value, "dd MMM yyyy")
                Else
                    requestdate1.Value = Format(PODate.Value, "dd MMM yyyy")
                End If
    
                If cboCust.Text = "" Then   'Or cborequestno.Text = ""
                    CboPOnO.clear: txtPoNo.Text = ""
                Else
                    Call adtocbopono
                End If
                For i = 0 To CboPOnO.ListCount - 1
                    If txtPoNo.Text = CboPOnO.List(i) Then
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False Then GoTo here
                If Not ValidDataSupplier(txtPoNo.Text) Then
                    LblErrMsg = "Can't change supplier! System found item[s] which no have price for this supplier. Please Input Price Master First!"
                    MousePointer = vbDefault
                    Exit Sub
                End If
                
                Call Browse
                sampun = True  ' true = Ada data di grid
                Call updateMaster(False)
            End If
                If ada = False Then
here:
                   Call kosongBwh
'                    txtRemarks(0).Text = "": txtRemarks(1).Text = "": txtRemarks(2).Text = ""
                    
                    'Call header
                    LblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
                    txtPoNo.SetFocus
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
    
    Case 3: 'CANCEL
            If txtPoNo.Text <> "" And cboCust.Text <> "" Then   'And cboRequestNo.Text <> ""
                For i = 0 To CboPOnO.ListCount - 1
                    If txtPoNo.Text = CboPOnO.List(i) Then
                        ketemu = True
                        Exit For
                    End If
                Next i
                If ketemu = False Then
                    Call kosongBwh
'                    txtRemarks(0).Text = "": txtRemarks(1).Text = "": txtRemarks(2).Text = ""
                    'Call header
                    LblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
                    txtPoNo.SetFocus
                    MousePointer = vbDefault
                    Exit Sub
                End If
                Call BrowseAtas
                Call Browse
            End If
End Select
MousePointer = vbDefault
Exit Sub

End Sub

Private Sub cmdReport_Click()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3
Dim sqlcekdet As String, SqlRpt As String
Dim rscekdet As New Recordset
  
    If combo1.ListIndex = 1 And txtPoNo.Text <> "" And cboCust.Text <> "" Then
        sqlcekdet = "select pom.PO_No from PurchaseOrder_Master pom " & _
                    "inner join PurchaseOrder_Detail pod on pod.PO_No = pom.PO_No " & _
                    "where pom.others_cls = '0' and pom.period is null " & _
                    "and pom.PO_No = '" & Trim(txtPoNo.Text) & "' and pom.supplier_Code = '" & Trim(cboCust.Text) & "'"
        Set rscekdet = Db.Execute(sqlcekdet)
        If rscekdet.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
        Set rscekdet = Nothing
        
        Me.MousePointer = vbHourglass

'        If cbocust.Column(4) = 1 Then   'PO CLS=YES
'            If cborequestno.Text = "" Then
''                cboRequestNo.SetFocus
''                LblErrMsg = DisplayMsg(1067) '"Please Input Request No"
''                Me.MousePointer = vbDefault
''                Exit Sub
'            ElseIf cborequestno.Text <> "" Then
'                If cborequestno.MatchFound = False Then
'                    lblErrMsg = DisplayMsg(4144)    'Record with This Request No Not found !
'                    cborequestno.SetFocus
'                    Me.MousePointer = vbDefault
'                    Exit Sub
'                End If
'            End If
'
''            Dim nextperiod As Date, endtgl As Integer, endperiod As Date, tglperiod As Date
''            'tglperiod = Left(cboRequestNo.Column(1), 4) & "-" & Right(cboRequestNo.Column(1), 2) & "-01"
''            'nextperiod = DateAdd("m", 1, tglperiod)
''            'endtgl = DateDiff("d", Format(tglperiod, "yyyy-mm-01"), Format(nextperiod, "yyyy-mm-01"))
''            'endperiod = year(tglperiod) & "-" & month(tglperiod) & "-" & Format(endtgl, "0#")
''
'            'PURCHASE ORDER DETAIL
'            SqlRpt = "select '1' Sort, rtrim(pom.po_no) po_no, pom.po_date, rtrim(pom.supplier_Code) Supplier_Code, " & _
'                    " trade_name = case when CHARINDEX ( ',' , rtrim(tm.trade_name)) <> 0 then " & _
'                    " ltrim(substring(rtrim(tm.trade_name), CHARINDEX ( ',' , rtrim(tm.trade_name)) + 1, 5)) + " & _
'                    " '.' + left(rtrim(tm.trade_name),CHARINDEX ( ',' , rtrim(tm.trade_name)) - 1) " & _
'                    " Else " & _
'                    " RTrim (tm.trade_name) " & _
'                    " End, "
'            SqlRpt = SqlRpt + " " & _
'                     "rtrim(tm.address1) taddress1, rtrim(tm.address2) taddress2, rtrim(tm.city) tcity, rtrim(tm.postal_code) tpostal_code, " & _
'                     "rtrim(tm.contact_person) contact_person, isnull(rtrim(tm.telephone),'') Supplierphone, isnull(rtrim(tm.Fax),'') SupplierFax, pom.paymentterm_cls, " & _
'                     "rtrim(pod.PORequest_No) PORequest_No, pod.POReq_SeqNo, rtrim(pod.item_code) item_code, rtrim(im.item_name) item_name, " & _
'                     "pod.unit_cls, (select description from unit_cls uc where uc.unit_cls= pod.unit_cls ) unit_desc, isnull(pod.qty,0) Qty, pod.currency_code, (select description from curr_cls where curr_cls.Curr_cls= pod.Currency_Code) Curr_desc, isnull(pod.price,0) Price, isnull(pod.amount,0) Amount, " & _
'                     "pod.Delivery_Date, pom.PriceCondition_Cls, rtrim(pc.Description) PriceCondition, pom.Transportation_Cls, " & _
'                     "rtrim(pom.remarks) Remarks,  isnull(pom.amount,0) as TAmount, isnull(pom.ppn,0) PPN, isnull(pom.total_amount,0) Total_Amount, " & _
'                     "rtrim(cp.company_name) company_name, rtrim(cp.address1) caddress1, rtrim(cp.address2) caddress2, " & _
'                     "rtrim(cp.Province) cprovince, rtrim(cp.City) ccity, rtrim(cp.postal_code) cpostal_code, rtrim(cp.phone1) cphone1, " & _
'                     "rtrim(cp.phone2) cphone2, rtrim(cp.fax) cfax, rtrim(cp.PO_position) po_position, rtrim(cp.PO_person) po_person, " '& _
'                     "rtrim(cp.POAcknowledge_Person) POAcknowledge_Person, rtrim(cp.POAcknowledge_Position) POAcknowledge_Position, " & _
'                     "rtrim(cp.POApproved_Person) POApproved_Person, rtrim(cp.POApproved_Position) POApproved_Position,
'            SqlRpt = SqlRpt + " " & _
'                     "tm.Trade_Cls, tm.Country_Cls " & _
'                     "from PurchaseOrder_Master pom " & _
'                     "inner join PurchaseOrder_Detail pod on pod.PO_No = pom.PO_No " & _
'                     "left outer join Item_Master im on im.item_code = pod.Item_code " & _
'                     "left outer join Trade_Master tm on tm.trade_code = pom.supplier_code " & _
'                     "left outer join PriceCondition_Cls pc on pc.PriceCondition_Cls = pom.PriceCondition_Cls " & _
'                     "cross join Company_Profile cp " & _
'                     "where pom.po_no = '" & Trim(txtpono.Text) & "' and pom.others_cls = '0' and pom.period is null "
'
'            'from PRICE MASTER and selected POREQUEST No
'            SqlRpt = SqlRpt & _
'                     "UNION " & _
'                     "select '2' Sort, rtrim(pom.po_no) po_no, pom.po_date, rtrim(pom.supplier_Code) Supplier_Code, " & _
'                    " trade_name = case when CHARINDEX ( ',' , rtrim(tm.trade_name)) <> 0 then " & _
'                    " ltrim(substring(rtrim(tm.trade_name), CHARINDEX ( ',' , rtrim(tm.trade_name)) + 1, 5)) + " & _
'                    " '.' + left(rtrim(tm.trade_name),CHARINDEX ( ',' , rtrim(tm.trade_name)) - 1) " & _
'                    " Else " & _
'                    " RTrim (tm.trade_name) " & _
'                    " End, "
'            SqlRpt = SqlRpt + " " & _
'                     "rtrim(tm.address1) taddress1, rtrim(tm.address2) taddress2, rtrim(tm.city) tcity, rtrim(tm.postal_code) tpostal_code, " & _
'                     "rtrim(tm.contact_person) contact_person, isnull(rtrim(tm.telephone),'') Supplierphone, isnull(rtrim(tm.Fax),'') SupplierFax,  pom.paymentterm_cls, " & _
'                     "prd.PORequest_No, prd.Seq_No, rtrim(pm.item_code) item_code, rtrim(im.item_name) item_name, " & _
'                     "im.unit_cls, (select description from unit_cls uc where uc.unit_cls= im.unit_cls ) unit_desc , 0 Qty, pm.currency_code, (select description from curr_cls where curr_cls.Curr_cls= pm.Currency_Code) Curr_desc, isnull(pm.price,0) Price, 0 Amount, " & _
'                     "Null Delivery_Date, pom.PriceCondition_Cls, (select rtrim(pc.Description) from PriceCondition_Cls pc where pc.PriceCondition_cls=pom.PriceCondition_Cls) PriceCondition, pom.Transportation_Cls, " & _
'                     "rtrim(pom.remarks) Remarks,  isnull(pom.amount,0) as TAmount, isnull(pom.ppn,0) PPN, isnull(pom.total_amount,0) Total_Amount, " & _
'                     "rtrim(cp.company_name) company_name, rtrim(cp.address1) caddress1, rtrim(cp.address2) caddress2, " & _
'                     "rtrim(cp.Province) cprovince, rtrim(cp.City) ccity, rtrim(cp.postal_code) cpostal_code, rtrim(cp.phone1) cphone1, " & _
'                     "rtrim(cp.phone2) cphone2, rtrim(cp.fax) cfax, rtrim(cp.PO_position) po_position, rtrim(cp.PO_person) po_person, " '& _
'                     "rtrim(cp.POAcknowledge_Person) POAcknowledge_Person, rtrim(cp.POAcknowledge_Position) POAcknowledge_Position, " & _
'                     "rtrim(cp.POApproved_Person) POApproved_Person, rtrim(cp.POApproved_Position) POApproved_Position,
'            SqlRpt = SqlRpt + " " & _
'                     "tm.Trade_Cls, tm.Country_Cls, rtrim(Department_cls) Department_Cls " & _
'                     "from PurchaseOrder_Master pom, Price_Master pm, Trade_Master tm, Item_Master im, " & _
'                     "(select prd.*, " & _
'                     "        cast(year(reqdelivery_date) as char(4)) + " & _
'                     "        cast((case when month(reqdelivery_date) < 10 then '0' else '' end) + cast(month(reqdelivery_date) as char) as char(2)) + " & _
'                     "        cast((case when day(reqdelivery_date) < 10 then '0' else '' end) + cast(day(reqdelivery_date) as char) as char(2)) " & _
'                     "        as ReqDelivery_Date1, Department_Cls, isnull(prm.Complete_Cls,'0') Complete_Cls " & _
'                     "        from PORequest_Detail prd " & _
'                     "        left outer join (select PORequest_No, Department_Cls, Complete_Cls from PORequest_Master where isnull(others_cls,'0') = '0') prm " & _
'                     "        on prm.porequest_no = prd.porequest_no) prd, " & _
'                     "Company_Profile cp "
'            SqlRpt = SqlRpt & _
'                     "Where pm.item_code = im.item_code And prd.item_code = pm.item_code " & _
'                     "And pom.Supplier_Code = tm.Trade_Code " & _
'                     "and pom.others_Cls = '0' and pom.period is null and pom.PO_No = '" & Trim(txtpono.Text) & "' And tm.PO_Cls = '1' " & _
'                     "and pm.trade_code in ('" & Trim(cbocust.Text) & "','000000') and pm.price_cls = '01' " & _
'                     "and start_date <= prd.ReqDelivery_Date1 and end_date >= prd.ReqDelivery_Date1 " & _
'                     "and prd.porequest_no = '" & Trim(cborequestno.Text) & "' and prd.Complete_Cls = '0' " & _
'                     "and prd.seq_no not in (select POReq_SeqNo from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "') " & _
'                     "and pm.currency_code = (select top 1 currency_code from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "' and Currency_Code is not null) "
'
'            'from ITEM MASTER and selected POREQUEST No
'            SqlRpt = SqlRpt & _
'                     "UNION " & _
'                     "select '3' Sort, rtrim(pom.po_no) po_no, pom.po_date, rtrim(pom.supplier_Code) Supplier_Code, " & _
'                    " trade_name = case when CHARINDEX ( ',' , rtrim(tm.trade_name)) <> 0 then " & _
'                    " ltrim(substring(rtrim(tm.trade_name), CHARINDEX ( ',' , rtrim(tm.trade_name)) + 1, 5)) + " & _
'                    " '.' + left(rtrim(tm.trade_name),CHARINDEX ( ',' , rtrim(tm.trade_name)) - 1) " & _
'                    " Else " & _
'                    " RTrim (tm.trade_name) " & _
'                    " End, "
'            SqlRpt = SqlRpt + " " & _
'                     "rtrim(tm.address1) taddress1, rtrim(tm.address2) taddress2, rtrim(tm.city) tcity, rtrim(tm.postal_code) tpostal_code, " & _
'                     "rtrim(tm.contact_person) contact_person, isnull(rtrim(tm.telephone),'') Supplierphone, isnull(rtrim(tm.Fax),'') SupplierFax, pom.paymentterm_cls, " & _
'                     "prd.PORequest_No, prd.Seq_No, rtrim(im.item_code) item_code, rtrim(im.item_name) item_name, " & _
'                     "im.unit_cls, (select description from unit_cls uc where uc.unit_cls= im.unit_cls ) unit_desc, 0 Qty, " & _
'                     "(select top 1 currency_code from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "' and Currency_Code is not null) currency_code, (select description from curr_cls where curr_cls.Curr_cls= (select top 1 currency_code from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "' and Currency_Code is not null)) Curr_desc, 0 Price, 0 Amount, " & _
'                     "Null Delivery_Date, pom.PriceCondition_Cls, (select rtrim(pc.Description) from PriceCondition_Cls pc where pc.PriceCondition_cls=pom.PriceCondition_Cls) PriceCondition, pom.Transportation_Cls, " & _
'                     "rtrim(pom.remarks) Remarks, isnull(pom.amount,0) as TAmount, isnull(pom.ppn,0) PPN, isnull(pom.total_amount,0) Total_Amount, " & _
'                     "rtrim(cp.company_name) company_name, rtrim(cp.address1) caddress1, rtrim(cp.address2) caddress2, " & _
'                     "rtrim(cp.Province) cprovince, rtrim(cp.City) ccity, rtrim(cp.postal_code) cpostal_code, rtrim(cp.phone1) cphone1, " & _
'                     "rtrim(cp.phone2) cphone2, rtrim(cp.fax) cfax, rtrim(cp.PO_position) po_position, rtrim(cp.PO_person) po_person, " '& _
'                     "rtrim(cp.POAcknowledge_Person) POAcknowledge_Person, rtrim(cp.POAcknowledge_Position) POAcknowledge_Position, " & _
'                     "rtrim(cp.POApproved_Person) POApproved_Person, rtrim(cp.POApproved_Position) POApproved_Position,
'
'            SqlRpt = SqlRpt + " " & _
'                     "tm.Trade_Cls, tm.Country_Cls, rtrim(Department_cls) Department_Cls " & _
'                     "from PurchaseOrder_Master pom, Trade_Master tm, Item_Master im, " & _
'                     "(select prd.*, " & _
'                     "        cast(year(reqdelivery_date) as char(4)) + " & _
'                     "        cast((case when month(reqdelivery_date) < 10 then '0' else '' end) + cast(month(reqdelivery_date) as char) as char(2)) + " & _
'                     "        cast((case when day(reqdelivery_date) < 10 then '0' else '' end) + cast(day(reqdelivery_date) as char) as char(2)) " & _
'                     "        as ReqDelivery_Date1, Department_cls, isnull(prm.Complete_Cls,'0') Complete_Cls " & _
'                     "        from PORequest_Detail prd " & _
'                     "        inner join (select PORequest_No, Department_Cls, Complete_Cls from PORequest_Master where isnull(others_cls,'0') = '0') prm " & _
'                     "        on prm.porequest_no = prd.porequest_no) prd, " & _
'                     "Company_Profile cp "
'                     'dicabut setelah request no and im.use_endday >= '" & Format(endperiod, "yyyymmdd") & "'
'            SqlRpt = SqlRpt & _
'                     "Where prd.item_code = im.item_code And pom.Supplier_Code = tm.Trade_Code " & _
'                     "and pom.others_Cls = '0' and pom.period is null and pom.PO_No = '" & Trim(txtpono.Text) & "' And tm.PO_Cls = '1' " & _
'                     "and im.supplier_code = '" & Trim(cbocust.Text) & "' and prd.porequest_no = '" & Trim(cborequestno.Text) & "' " & _
'                     "and prd.seq_no not in (select POReq_SeqNo from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "') and prd.Complete_Cls = '0' " & _
'                     "and im.item_code not in " & _
'                     "    (select distinct pm2.item_Code From Price_Master pm2 " & _
'                     "     where pm2.price_cls = '01' and pm2.start_date <= prd.ReqDelivery_Date1 and pm2.end_date >= prd.ReqDelivery_Date1 " & _
'                     "     and pm2.trade_code in ('" & Trim(cbocust.Text) & "','000000') ) " & _
'                     "Order by Sort "
'
'        Else    'PO CLS=NO rtrim(tm.trade_name) trade_name,
'            SqlRpt = "select '1' Sort, rtrim(pom.po_no) po_no, pom.po_date, rtrim(pom.supplier_Code) Supplier_Code, " & _
'                    " trade_name = case when CHARINDEX ( ',' , rtrim(tm.trade_name)) <> 0 then " & _
'                    " ltrim(substring(rtrim(tm.trade_name), CHARINDEX ( ',' , rtrim(tm.trade_name)) + 1, 5)) + " & _
'                    " '.' + left(rtrim(tm.trade_name),CHARINDEX ( ',' , rtrim(tm.trade_name)) - 1) " & _
'                    " Else " & _
'                    " RTrim (tm.trade_name) " & _
'                    " End, "
'            SqlRpt = SqlRpt + " " & _
'                     "rtrim(tm.address1) taddress1, rtrim(tm.address2) taddress2, rtrim(tm.city) tcity, rtrim(tm.postal_code) tpostal_code, " & _
'                     "rtrim(tm.contact_person) contact_person, isnull(rtrim(tm.telephone),'') Supplierphone, isnull(rtrim(tm.Fax),'') SupplierFax, pom.paymentterm_cls, " & _
'                     "rtrim(pod.PORequest_No) PORequest_No, pod.POReq_SeqNo, rtrim(pod.item_code) item_code, rtrim(im.item_name) item_name, " & _
'                     "pod.unit_cls, (select description from unit_cls uc where uc.unit_cls= pod.unit_cls ) unit_desc, isnull(pod.qty,0) Qty, pod.currency_code, (select description from curr_cls where curr_cls.Curr_cls= pod.Currency_Code) Curr_desc ,isnull(pod.price,0) Price, isnull(pod.amount,0) Amount, " & _
'                     "pod.Delivery_Date, pom.PriceCondition_Cls, rtrim(pc.Description) PriceCondition, pom.Transportation_Cls, " & _
'                     "rtrim(pom.remarks) Remarks,  isnull(pom.amount,0) as TAmount, isnull(pom.ppn,0) PPN, isnull(pom.total_amount,0) Total_Amount, " & _
'                     "rtrim(cp.company_name) company_name, rtrim(cp.address1) caddress1, rtrim(cp.address2) caddress2, " & _
'                     "rtrim(cp.Province) cprovince, rtrim(cp.City) ccity, rtrim(cp.postal_code) cpostal_code, rtrim(cp.phone1) cphone1, " & _
'                     "rtrim(cp.phone2) cphone2, rtrim(cp.fax) cfax, rtrim(cp.PO_position) po_position, rtrim(cp.PO_person) po_person, " '& _
'                     "rtrim(cp.POAcknowledge_Person) POAcknowledge_Person, rtrim(cp.POAcknowledge_Position) POAcknowledge_Position, " & _
'                     "rtrim(cp.POApproved_Person) POApproved_Person, rtrim(cp.POApproved_Position) POApproved_Position,
'            SqlRpt = SqlRpt + " " & _
'                     "tm.Trade_Cls, tm.Country_Cls " & _
'                     "from PurchaseOrder_Master pom " & _
'                     "inner join PurchaseOrder_Detail pod on pod.PO_No = pom.PO_No " & _
'                     "left outer join Item_Master im on im.item_code = pod.Item_code " & _
'                     "left outer join Trade_Master tm on tm.trade_code = pom.supplier_code " & _
'                     "left outer join PriceCondition_Cls pc on pc.PriceCondition_Cls = pom.PriceCondition_Cls "
'
'            ' ------ Tambahan untuk menampilkan T,W dan L
'                    SqlRpt = SqlRpt & " inner join item_master iim on pod.item_code=iim.item_code "
'                    SqlRpt = SqlRpt & " inner join Sheetcoil_cls Sccm on iim.sheetcoil_cls=Sccm.Sheetcoil_cls "
'            ' -----
'
'                     SqlRpt = SqlRpt + "cross join Company_Profile cp " & _
'                     "where pom.po_no = '" & Trim(txtpono.Text) & "' and pom.others_cls = '0' and pom.period is null " & _
'                     "order by pod.PORequest_No, pod.Item_Code, pod.POReq_SeqNo "
'        End If
        
' -----
' Perubahan Format PO sesuai Musashi ( Lokal=Import )
' -----
  
SqlRpt = " Select POM.Po_No, POM.Po_Date,POM.delivery_Date,PRD.PoRequest_No,PRM.PersonInCharge_Cls,PIC.Description, " & _
            vbLf & " POM.Supplier_Code,TM.Trade_Name,TM.Contact_Person,TM.Address1,TM.Address2,TM.City,TM.Country, " & _
            vbLf & " TM.Telephone,Tm.Fax,POM.PaymentTerm_Cls, " & _
            vbLf & " POD.Item_code,POD.Price,POD.Qty,POD.Amount,IM.Item_Name, " & _
            vbLf & " POD.Unit_Cls,U.Description Unit,POD.Currency_Code,C.Description Currency," & _
            vbLf & "' ' Ref,' ' ShipVia,POM.Remarks comments, " & _
            vbLf & " isnull((Select ChildRequirement_Qty from requirement_Master Where ChildRequirement_Month=month(POM.delivery_Date)+1 and ChildRequirement_Year=year(POM.delivery_Date) and ChildItem_Code=POD.Item_code),0) F1, " & _
            vbLf & " isnull((Select ChildRequirement_Qty from requirement_Master Where ChildRequirement_Month=month(POM.delivery_Date)+2 and ChildRequirement_Year=year(POM.delivery_Date) and ChildItem_Code=POD.Item_code),0) F2 " & _
            vbLf & " From PurchaseOrder_Master POM inner join PurchaseOrder_Detail POD " & _
            vbLf & " On POM.Po_No=POd.Po_no " & _
            vbLf & " Inner Join Trade_Master TM on POM.Supplier_Code=TM.Trade_Code " & _
            vbLf & " inner Join Item_Master IM on POD.Item_Code=IM.Item_Code " & _
            vbLf & " inner Join Unit_Cls U on POD.Unit_Cls=U.Unit_Cls " & _
            vbLf & " inner Join PORequest_Detail PRD on POD.PORequest_No=PRD.PoRequest_No and POD.PoReq_SeqNo=PRD.PoReq_SeqNo " & _
            vbLf & " inner Join PoREquest_Master PRM on POD.PORequest_No=PRM.PoRequest_No " & _
            vbLf & " inner join PersonInCharge_Cls PIC on PRM.PersonInCharge_Cls=PIC.PersonInCharge_Cls " & _
            vbLf & " inner join curr_cls C on POD.Currency_Code=C.Curr_Cls "
            ' ------ Tambahan untuk menampilkan T,W dan L
                    SqlRpt = SqlRpt & " inner join Sheetcoil_cls Sccm on im.sheetcoil_cls=Sccm.Sheetcoil_cls "
            ' -----
SqlRpt = SqlRpt & _
            vbLf & " where pom.po_no = '" & Trim(txtPoNo.Text) & "' and pom.others_cls = '0' and pom.period is null " & _
            vbLf & " order by pod.PORequest_No, pod.Item_Code, pod.POReq_SeqNo "
  
' ------
        
        
        If rsRpt.State <> adStateClosed Then rsRpt.Close
        rsRpt.Open SqlRpt, Db, adOpenDynamic, adLockOptimistic
        
        sqlprint = SqlRpt
        reportcode = "poparts"
        Fbulan = txtPoNo.Text
        printorient = 1
        
        If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
        Set report = application.OpenReport(App.path & "\Reports\rptPONew.rpt")
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

Private Sub CmdSubMenu_Click()
    ClearData
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblErrMsg.Caption = ErrMsg
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RS = Nothing
    Set rsGrid = Nothing
End Sub

Private Sub updateMaster(Flag As Boolean)
    Dim sQl_Master As String
    Dim rs_Master As New ADODB.Recordset
        sQl_Master = "select * from PurchaseOrder_Master where PO_No = '" & Trim$(txtPoNo.Text) & "' and others_cls = '0' and period is null"
        If rs_Master.State <> adStateClosed Then rs_Master.Close
        rs_Master.Open sQl_Master, Db, adOpenKeyset, adLockOptimistic
        If rs_Master.BOF And rs_Master.EOF Then
            LblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
            txtPoNo.SetFocus
            rs_Master.Close
            Set rs_Master = Nothing
            Exit Sub
        End If
        rs_Master("po_date") = Format(PODate.Value, "yyyy-mm-dd")
        rs_Master("revise_No") = Trim(txtRev.Text)
        rs_Master("supplier_code") = Trim(cboCust.Text)
        
        If Flag = True Then
            rs_Master("discount") = CDbl(TxtDisc.Text) ' Addd 20090112
            rs_Master("amount") = CDbl(txtamount.Text)
            rs_Master("ppn") = CDbl(txtPPN.Text)
            rs_Master("total_amount") = CDbl(txtGrandTotal.Text)
            rs_Master("PoMarking1") = Trim(txtMarking(0).Text)
            rs_Master("PoMarking2") = Trim(txtMarking(1).Text)
            rs_Master("PoMarking3") = Trim(txtMarking(2).Text)
            rs_Master("PoMarking4") = Trim(txtMarking(3).Text)
            rs_Master("PoMarking5") = Trim(txtMarking(4).Text)
            rs_Master("PoMarking6") = Trim(txtMarking(5).Text)
                
            rs_Master("remarks") = Trim(txtremarks.Text)
            
            If CboPacking.Text = "" Then rs_Master("popacking_Cls") = Null Else rs_Master("popacking_Cls") = CboPacking.Column(0)
            If cboInsuranceCls.Text = "" Then rs_Master("insurance_cls") = Null Else rs_Master("insurance_cls") = cboInsuranceCls.Column(0)
            If cboPaymentTerm.Text = "" Then rs_Master("paymentterm_cls") = Null Else rs_Master("paymentterm_cls") = cboPaymentTerm.Column(0)
            If cboTransport.Text = "" Then rs_Master("transportation_cls") = Null Else rs_Master("transportation_cls") = cboTransport.Column(0)
            If cboPriceCondition.Text = "" Then rs_Master("pricecondition_cls") = Null Else rs_Master("pricecondition_cls") = cboPriceCondition.Text
        
        End If
        rs_Master.update
        rs_Master.Close
        Set rs_Master = Nothing
End Sub


Function ValidDataSupplier(pPONO As String) As Boolean
Dim ls_sql  As String
Dim rsCek As New Recordset, rsCek2 As New Recordset, lint_recordcount As Integer

ValidDataSupplier = True
ls_sql = "Select distinct item_code, poreq_seqno  from purchaseOrder_detail where po_no ='" & pPONO & "' order by item_code"
If rsCek.State <> adStateClosed Then rsCek.Close
rsCek.CursorLocation = adUseClient
rsCek.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
If rsCek.EOF = False Then


'    lint_recordcount = rsCek.RecordCount
'    ls_sql = " select * from (select distinct item_Code from  price_master where price_Cls = '01' and trade_code = '" & CboCust.Text & "' " & _
'             "   and exists " & _
'             "       ( " & _
'             "       (select distinct a.item_code, reqdelivery_date " & _
'             "       from purchaseOrder_detail a, PORequest_Detail b " & _
'             "       Where a.porequest_no = b.porequest_no " & _
'             "       and a.poreq_seqno = b.seq_no and a.item_code = price_master.item_Code " & _
'             "       and (convert(char(8), Reqdelivery_date,112) between start_date and end_date)  " & _
'             "       and a.PO_no ='" & pPONO & "'))" & _
'             "  union " & _
'             "   select b.item_Code from item_master a,porequest_detail b " & _
'             "   where supplier_code = '" & CboCust.Text & "' and a.item_code = b.item_code " & _
'             "   and POrequest_no = '" & cboRequestNo & "') a order by item_code "
    Do While Not rsCek.EOF
        ls_sql = "select distinct item_code from price_master where trade_code in ('" & cboCust.Text & "','000000') and item_code= '" & rsCek!Item_Code & "' " & _
                    " and convert(char(8), (select reqdelivery_date from porequest_detail where poreq_seqno='" & rsCek!POReq_seqno & "' ),112) between start_date and end_date " & _
                    " Union " & _
                    " select item_code from item_master where supplier_code = '" & cboCust.Text & "' and item_code ='" & rsCek!Item_Code & "' "
    
        If rsCek2.State <> adStateClosed Then rsCek2.Close
        rsCek2.CursorLocation = adUseClient
        rsCek2.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
        If rsCek2.EOF Then
            ValidDataSupplier = False
            Set rsCek2 = Nothing
            Set rsCek = Nothing
            Exit Function
        End If
    rsCek.MoveNext
    Loop
    Set rsCek2 = Nothing
End If
rsCek.Close
Set rsCek = Nothing
End Function
