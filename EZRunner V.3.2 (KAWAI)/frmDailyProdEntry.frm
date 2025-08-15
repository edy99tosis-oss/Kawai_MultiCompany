VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDailyProdEntry 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Production Schedule Entry (Update)"
   ClientHeight    =   10515
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   15210
   Icon            =   "frmDailyProdEntry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtSerialTo 
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
      Left            =   13320
      MaxLength       =   7
      TabIndex        =   9
      Top             =   7725
      Width           =   1635
   End
   Begin VB.TextBox TxtSerialFrom 
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
      Left            =   11610
      MaxLength       =   7
      TabIndex        =   8
      Top             =   7725
      Width           =   1635
   End
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "..."
      Height          =   300
      Left            =   2250
      TabIndex        =   44
      Top             =   7725
      Width           =   300
   End
   Begin VB.TextBox lblPart 
      Appearance      =   0  'Flat
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
      Left            =   2625
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   7755
      Width           =   2055
   End
   Begin VB.TextBox txtremark1 
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
      Left            =   2100
      MaxLength       =   35
      TabIndex        =   11
      Top             =   8700
      Width           =   6105
   End
   Begin VB.TextBox lbldesc 
      Appearance      =   0  'Flat
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
      Left            =   4860
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   7755
      Width           =   3825
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "To Forecast Inquiry"
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
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   9960
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "To Formula"
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
      Left            =   2925
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   9960
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "To Order Inquiry"
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
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   9960
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.TextBox txtseqno 
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
      Left            =   4470
      MaxLength       =   18
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   10005
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton command1 
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9960
      Width           =   1125
   End
   Begin VB.CommandButton command1 
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
      Index           =   3
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9960
      Width           =   1125
   End
   Begin VB.TextBox txtlotno 
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
      Left            =   8820
      MaxLength       =   7
      TabIndex        =   6
      Top             =   7725
      Width           =   885
   End
   Begin VB.CommandButton command1 
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
      Index           =   2
      Left            =   10290
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9960
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   615
      Left            =   105
      TabIndex        =   26
      Top             =   9210
      Width           =   15000
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
         Height          =   300
         Left            =   120
         TabIndex        =   27
         Top             =   195
         Width           =   14760
      End
   End
   Begin VB.TextBox txtqty 
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
      Left            =   9750
      MaxLength       =   12
      TabIndex        =   7
      Top             =   7725
      Width           =   1125
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
      Left            =   13980
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9960
      Width           =   1125
   End
   Begin VB.CommandButton command3 
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9960
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker scheduledate 
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   8700
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
      Format          =   128909315
      CurrentDate     =   37859
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   4905
      Left            =   105
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2250
      Width           =   15000
      _cx             =   26458
      _cy             =   8652
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
      HighLight       =   2
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
      RowHeightMax    =   0
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
      ExplorerBar     =   1
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1170
      Left            =   105
      TabIndex        =   28
      Top             =   1005
      Width           =   15000
      Begin MSComCtl2.DTPicker scheduledate1 
         Height          =   315
         Left            =   9390
         TabIndex        =   2
         Top             =   690
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
         Format          =   128909315
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker scheduledate2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   11490
         TabIndex        =   3
         Top             =   690
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
         Format          =   128909315
         CurrentDate     =   37798
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Cls"
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
         Left            =   9390
         TabIndex        =   47
         Top             =   300
         Width           =   855
      End
      Begin MSForms.ComboBox CboGroup 
         Height          =   315
         Left            =   10365
         TabIndex        =   46
         Top             =   240
         Width           =   1125
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "1984;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label11 
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
         Left            =   11625
         TabIndex        =   45
         Top             =   300
         Width           =   2955
      End
      Begin VB.Line Line5 
         X1              =   11625
         X2              =   14610
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Production Schedule Date"
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
         Left            =   6975
         TabIndex        =   35
         Top             =   750
         Width           =   2205
      End
      Begin VB.Label Label4 
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
         Left            =   11190
         TabIndex        =   34
         Top             =   750
         Width           =   165
      End
      Begin VB.Line Line1 
         X1              =   3105
         X2              =   6345
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label lbllinecd 
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
         Left            =   3105
         TabIndex        =   33
         Top             =   750
         Width           =   3255
      End
      Begin MSForms.ComboBox cbolinecd 
         Height          =   315
         Left            =   1695
         TabIndex        =   1
         Top             =   690
         Width           =   1305
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "2302;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machine No"
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
         Left            =   360
         TabIndex        =   32
         Top             =   750
         Width           =   975
      End
      Begin MSForms.ComboBox cbocust 
         Height          =   315
         Left            =   1695
         TabIndex        =   0
         Top             =   255
         Width           =   1305
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "2302;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3105
         X2              =   9105
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label LblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory Code"
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
         Left            =   360
         TabIndex        =   30
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label lblcust 
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
         Left            =   3105
         TabIndex        =   29
         Top             =   315
         Width           =   6015
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13260
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   255
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No To"
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
      Left            =   13320
      TabIndex        =   49
      Top             =   7365
      Width           =   1050
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No From"
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
      Left            =   11640
      TabIndex        =   48
      Top             =   7365
      Width           =   1275
   End
   Begin VB.Line Line4 
      X1              =   2625
      X2              =   4650
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label Label7 
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
      Left            =   2685
      TabIndex        =   42
      Top             =   7365
      Width           =   1080
   End
   Begin VB.Line Line3 
      X1              =   4890
      X2              =   8700
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   525
      Index           =   2
      Left            =   90
      Top             =   7635
      Width           =   15000
   End
   Begin VB.Label lblunit1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
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
      Left            =   2490
      TabIndex        =   39
      Top             =   10050
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label5 
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
      Left            =   8820
      TabIndex        =   38
      Top             =   7365
      Width           =   540
   End
   Begin VB.Label lblunit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "pcs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   10965
      TabIndex        =   37
      Top             =   7725
      Width           =   570
   End
   Begin VB.Label Label2 
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
      Left            =   2160
      TabIndex        =   36
      Top             =   8355
      Width           =   675
   End
   Begin VB.Label Label1 
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
      Left            =   10980
      TabIndex        =   31
      Top             =   7365
      Width           =   330
   End
   Begin MSForms.ComboBox cboitemcode 
      Height          =   315
      Left            =   225
      TabIndex        =   5
      Top             =   7725
      Width           =   1950
      VariousPropertyBits=   746604571
      MaxLength       =   25
      DisplayStyle    =   3
      Size            =   "3440;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label8 
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
      Left            =   4860
      TabIndex        =   25
      Top             =   7365
      Width           =   960
   End
   Begin VB.Label Label6 
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
      Left            =   255
      TabIndex        =   24
      Top             =   7365
      Width           =   1155
   End
   Begin VB.Label Label9 
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
      Left            =   10590
      TabIndex        =   23
      Top             =   7365
      Width           =   300
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule Date"
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
      Top             =   8355
      Width           =   1245
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Production Schedule Entry (Update)"
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
      Left            =   105
      TabIndex        =   21
      Top             =   345
      Width           =   15000
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   90
      Top             =   7275
      Width           =   15000
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   90
      Top             =   8250
      Width           =   15000
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   525
      Index           =   0
      Left            =   60
      Top             =   8610
      Width           =   15000
   End
End
Attribute VB_Name = "frmDailyProdEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim sqlGrid As String
Dim ubah As Boolean, ubahBulan As Boolean, ubahBulanQty As Boolean, Status As String
Dim PONO As String, poSEqNo As String, lblFNo As String, startDaily As String
Dim lbllotno As String, lblschdate As Date, lblQty As Double, changeQty As Double
Public dailypanggil As String
Dim gabung As Boolean, gabungState As Boolean
Dim OrderEntry As Double, ForecastAwal As Double, TotalDaily As Double
Dim tOrderEntry As Double, tForecastAwal As Double, tTotalDaily As Double, tUpdateQty As Double
Dim insertActForecast As Boolean, insertActOrder As Boolean, insertActDaily As Boolean, insertActWIP As Boolean
Dim tinsertActForecast As Boolean, tinsertActOrder As Boolean, tinsertActDaily As Boolean, tinsertActWIP As Boolean

Dim notFinishGood As Boolean, tnotFinishGood As Boolean
Dim WIPLimit As Double, tWIPLimit As Double
Dim tmpParentLotNo As String, ttmpParentLotNo As String

Dim bteColSelect As Byte
Dim bteColProdCode As Byte
Dim bteColPart As Byte
Dim bteColDesc As Byte
Dim bteColLotNo As Byte
Dim bteColQty As Byte
Dim bteColUnitCls As Byte
Dim bteColUnit As Byte
Dim BteColSerialFrom As Byte   ' Add 20090207
Dim BteColSerialTo As Byte      ' Add 20090207
Dim bteColDate As Byte
Dim bteColRemark As Byte
Dim bteColAuto As Byte
Dim bteColCustCode As Byte
Dim bteColCustName As Byte
Dim bteColPONo As Byte
Dim bteColSeqNo As Byte

Dim Auto_Cls As Byte
Dim StrStartSerial As String

Sub Header()
    Dim i As Long
    
    bteColSelect = 0
    bteColProdCode = 1
    bteColPart = 2
    bteColDesc = 3
    bteColLotNo = 4
    bteColQty = 5
    bteColUnitCls = 6
    bteColUnit = 7
    ' Add and Update 20090207
    BteColSerialFrom = 8
    BteColSerialTo = 9
    bteColDate = 8 + 2
    bteColRemark = 9 + 2
    bteColAuto = 10 + 2
    bteColCustCode = 11 + 2
    bteColCustName = 12 + 2
    bteColPONo = 13 + 2
    bteColSeqNo = 14 + 2
    
    With grid
      .clear
      .Rows = 1
      .ColS = 15 + 2
      
      .ColWidth(bteColSelect) = 300
      .ColWidth(bteColProdCode) = 1400
      .ColWidth(bteColPart) = 1400
      .ColWidth(bteColDesc) = 3000
      .ColWidth(bteColLotNo) = 1000
      .ColWidth(bteColQty) = 1230
      .ColWidth(bteColUnit) = 650
      .ColWidth(BteColSerialFrom) = 1100        'Add 20090207
      .ColWidth(BteColSerialTo) = 1100          'Add 20090207
      .ColWidth(bteColDate) = 1450
      .ColWidth(bteColRemark) = 3250
      .ColWidth(bteColAuto) = 1000
      .ColWidth(bteColCustCode) = 1200
      .ColWidth(bteColCustName) = 2800
      .ColWidth(bteColPONo) = 2000
      
      .TextMatrix(0, bteColSelect) = ""
      .TextMatrix(0, bteColPart) = "Part Number"
      .TextMatrix(0, bteColProdCode) = "Product Code"
      .TextMatrix(0, bteColDesc) = "Description"
      .TextMatrix(0, bteColLotNo) = "Lot No"
      .TextMatrix(0, bteColQty) = "Qty"
      .TextMatrix(0, bteColUnitCls) = "UnitCls"
      .TextMatrix(0, bteColUnit) = "Unit"
      .TextMatrix(0, BteColSerialFrom) = "Serial From"      'Add 20090207
      .TextMatrix(0, BteColSerialTo) = "Serial To"            'Add 20090207
      .TextMatrix(0, bteColDate) = "Schedule Date"
      .TextMatrix(0, bteColRemark) = "Remark"
      .TextMatrix(0, bteColAuto) = "Auto"
      .TextMatrix(0, bteColCustCode) = "Cust. Code"
      .TextMatrix(0, bteColCustName) = "Cust. Name"
      .TextMatrix(0, bteColPONo) = "PO No."
      .TextMatrix(0, bteColSeqNo) = "SeqNo"
      
      .ColHidden(bteColUnitCls) = True
      .ColHidden(bteColSeqNo) = True
      .ColHidden(bteColCustCode) = True
      .ColHidden(bteColCustName) = True
      .ColHidden(bteColPONo) = True
      
      .ColDataType(bteColDate) = flexDTDate
      
      .Cell(flexcpAlignment, 0, 0, 0, bteColSeqNo) = flexAlignCenterCenter
      .ColAlignment(bteColSelect) = flexAlignCenterCenter
      .ColAlignment(bteColPart) = flexAlignLeftCenter
      .ColAlignment(bteColProdCode) = flexAlignLeftCenter
      .ColAlignment(bteColDesc) = flexAlignLeftCenter
      .ColAlignment(bteColLotNo) = flexAlignCenterCenter
      .ColAlignment(bteColQty) = flexAlignRightCenter
      .ColAlignment(bteColUnit) = flexAlignCenterCenter
      .ColAlignment(BteColSerialFrom) = flexAlignCenterCenter
      .ColAlignment(BteColSerialTo) = flexAlignCenterCenter
      .ColAlignment(bteColDate) = flexAlignCenterCenter
      .ColAlignment(bteColRemark) = flexAlignLeftCenter
      .ColAlignment(bteColAuto) = flexAlignCenterCenter
      .ColAlignment(bteColCustCode) = flexAlignLeftCenter
      .ColAlignment(bteColCustName) = flexAlignLeftCenter
      .ColAlignment(bteColPONo) = flexAlignLeftCenter
      
      .EditMaxLength = 1
    End With
 
End Sub

Sub Kosong()
    scheduledate1.Value = Format(Now, "dd MMM yyyy")
    scheduledate2.Value = Format(Now, "dd MMM yyyy")
    lblcust.Caption = ""
    cbocust.Text = ""
    cbolinecd.clear
    cbolinecd.Text = ""
    lbllinecd.Caption = ""
    CboGroup.ListIndex = 0   'Add 20090207
    
    lblErrMsg = ""
    scheduledate.Value = Format(Now, "dd MMM yyyy")
    lblschdate = scheduledate.Value
    dailypanggil = ""
    
    kosongBwh
    Header
End Sub

Sub kosongBwh()
    cboitemcode.Enabled = True
    cboitemcode.Text = ""
    lblPart.Text = ""
    lbldesc.Text = ""
    txtqty.Text = Format(0, gs_formatQty)
    lblQty = 0
    lblunit.Caption = ""
    txtlotno.Text = ""
' Add Serial No For KAWAI 20090207
    TxtSerialFrom = ""
    TxtSerialTo = ""
' ---
    lbllotno = ""
    txtremark1.Text = ""
    lblFNo = ""
    ubah = False
    txtseqno.Text = ""
    lblunit1.Caption = ""
    PONO = ""
    poSEqNo = 0
    Status = ""
End Sub

Sub adtocboCust()
Dim sqlcust As String
Dim RsCust As New Recordset
Dim i As Integer

    sqlcust = "select distinct(manufacture_code), trade_name from manufacture_line inner join trade_master " & _
              "on trade_master.trade_code = manufacture_line.manufacture_code order by manufacture_code"

    Set RsCust = Db.Execute(sqlcust)
    
    With cbocust
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;300pt"
        .ListWidth = 350
        .ListRows = 15
        
        i = 0
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("manufacture_code"))
            .List(i, 1) = IIf(IsNull(RsCust("trade_name")), "", Trim(RsCust("Trade_Name")))
            RsCust.MoveNext
            i = i + 1
        Loop
    End With
End Sub

Sub adtocboGroup()
Dim SqlGroup As String
Dim RsGroup As New Recordset
Dim i As Integer

    SqlGroup = "Select * From Group_Cls"

    Set RsGroup = Db.Execute(SqlGroup)
    
    With CboGroup
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;100pt"
        .ListWidth = 150
        .ListRows = 5
        
       .AddItem
        .List(0, 0) = "All"
        .List(0, 1) = "All"
        
        i = 1
        Do While Not RsGroup.EOF
            .AddItem
            .List(i, 0) = Trim(RsGroup("Group_Cls"))
            .List(i, 1) = IIf(IsNull(RsGroup("Description")), "", Trim(RsGroup("Description")))
            RsGroup.MoveNext
            i = i + 1
        Loop
    End With
End Sub

Sub addToCboItemCode()
Dim sqlLine As String
Dim RsLine As New Recordset
Dim i As Long

    sqlLine = "select item_code, makeritem_code, item_name, unit_cls, (select description from unit_cls uc where uc.unit_cls=item_master.unit_cls)  Unit_Desc " & _
        "from item_master where production_cls = 01 and manufacture_code = '" & Trim(cbocust.Text) & "' " & _
        "and use_endday > convert(char(8), getdate(), 112) order by item_name"
    Set RsLine = Db.Execute(sqlLine)
    
    With cboitemcode
        .clear
        .columnCount = 5
        .ColumnWidths = "80pt;80pt;240pt;0pt;0pt"
        .ListWidth = 400
        .ListRows = 10
        
        lblPart.Text = ""
        lbldesc.Text = ""
        i = 0
        Do While Not RsLine.EOF
            .AddItem
            .List(i, 0) = Trim(RsLine("item_code"))
            .List(i, 1) = Trim(RsLine("makeritem_code"))
            .List(i, 2) = Trim(RsLine("item_Name"))
            .List(i, 3) = Trim(RsLine("unit_cls"))
            .List(i, 4) = Trim(RsLine("unit_desc"))
            RsLine.MoveNext
            i = i + 1
        Loop
    End With
End Sub

Sub adtocbolinecd()
Dim sqlLine As String
Dim RsLine As New Recordset
Dim i As Integer

    sqlLine = "select * from manufacture_line where manufacture_code='" & cbocust.Text & "'"
    Set RsLine = Db.Execute(sqlLine)
    
    With cbolinecd
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;140pt"
        .ListWidth = 190
        .ListRows = 15
        
        lbllinecd.Caption = ""
        i = 0
        Do While Not RsLine.EOF
            .AddItem
            .List(i, 0) = Trim(RsLine("Line_code"))
            .List(i, 1) = Trim(RsLine("Line_Name"))
            RsLine.MoveNext
            i = i + 1
        Loop
    End With
End Sub

Function seqNo() As Long
Dim sqlseqno As String
Dim rsseqno As New Recordset

sqlseqno = "select * from daily_production order by seq_no"
If rsseqno.State <> adStateClosed Then rsseqno.Close
rsseqno.Open sqlseqno, Db, adOpenKeyset, adLockOptimistic

If Not (rsseqno.BOF And rsseqno.EOF) Then
    rsseqno.MoveLast
    seqNo = rsseqno!Seq_no + 1
Else
    seqNo = 1
End If
End Function

Sub Browse()
    
    Dim i As Long
    Dim strSQL As String
    Dim rsGrid As New ADODB.Recordset

    Header
    Auto_Cls = 0
    If Status = "update" Then kosongBwh
    
    strSQL = "select a.*, b.makeritem_code, b.item_name, b.group_Cls,pcd.cust_code, pcd.po_no, tm.trade_name " & _
        "from daily_production a " & _
        "left join item_master b on b.item_code=a.item_code " & _
        "left join productioncalculate_detail pcd on a.plancust_code = pcd.cust_code and a.planpo_seqno = pcd.seq_no and a.plan_seqno = pcd.plan_seqno and a.item_code = pcd.planitem_code  and a.planpo_no = pcd.po_no " & _
        "left join trade_master tm on pcd.cust_code = tm.trade_code " & _
        "where a.factory_code='" & Trim(cbocust.Text) & "' " & _
        "and a.line_code='" & Trim(cbolinecd.Text) & "' "
    
' Add Group Filter for Browsing Grid 20090207
    If CboGroup.Text <> "All" Then strSQL = strSQL & " And b.group_cls='" & Trim(CboGroup.Text) & "' "
'--------
    strSQL = strSQL & _
        "and a.schedule_date>='" & Format(scheduledate1.Value, "yyyymmdd") & "' " & _
        "and a.schedule_date<='" & Format(scheduledate2.Value, "yyyymmdd") & "' " & _
        "order by a.schedule_date, a.item_code, a.lot_no, a.seq_no "
              
'    strSql = "select a.*, b.makeritem_code, b.item_name, ' ' cust_code, ' ' po_no, ' ' trade_name " & _
'        "from daily_production a " & _
'        "left join item_master b on b.item_code=a.item_code " & _
'        "where a.factory_code='" & Trim(cbocust.Text) & "' " & _
'        "and a.line_code='" & Trim(cbolinecd.Text) & "' " & _
'        "and a.schedule_date>='" & Format(scheduledate1.Value, "yyyymmdd") & "' " & _
'        "and a.schedule_date<='" & Format(scheduledate2.Value, "yyyymmdd") & "' " & _
'        "order by a.schedule_date, a.item_code, a.lot_no, a.seq_no "
              
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    Set rsGrid = Db.Execute(strSQL)
    
    'rsGrid.Open strSql, Db, adOpenKeyset, adLockOptimistic
      
    i = 1
    If Not (rsGrid.BOF And rsGrid.EOF) Then
        With grid
            Do While Not rsGrid.EOF
                .Rows = .Rows + 1
                .TextMatrix(i, bteColProdCode) = Trim(rsGrid("Item_Code"))
                .TextMatrix(i, bteColPart) = Trim(rsGrid("MakerItem_Code"))
                .TextMatrix(i, bteColDesc) = IIf(IsNull(rsGrid("item_name")), "", Trim(rsGrid("item_name")))
                
                .TextMatrix(i, bteColLotNo) = IIf(IsNull(rsGrid("lot_no")), "", Trim(rsGrid("lot_no")))
                
                .TextMatrix(i, bteColQty) = IIf(IsNull(rsGrid("Qty")), 0, Trim(rsGrid("Qty")))
                If InStr(1, .TextMatrix(i, bteColQty), ".") > 0 Then
                    .TextMatrix(i, bteColQty) = Format(.TextMatrix(i, bteColQty), gs_formatQty)
                Else
                    .TextMatrix(i, bteColQty) = Format(.TextMatrix(i, bteColQty), gs_formatQty)
                End If
        
                If IsNull(rsGrid("unit_cls")) Then
                  .TextMatrix(i, bteColUnitCls) = ""
                  .TextMatrix(i, bteColUnit) = ""
                Else
                  .TextMatrix(i, bteColUnitCls) = Trim(rsGrid("Unit_cls"))
                  .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(Trim(rsGrid("Unit_Cls")))
                End If
                
                .TextMatrix(i, BteColSerialFrom) = IIf(IsNull(rsGrid("SerialNoFrom")), "", Trim(rsGrid("SerialNoFrom"))) 'Add 20090207
                .TextMatrix(i, BteColSerialTo) = IIf(IsNull(rsGrid("SerialNoTo")), "", Trim(rsGrid("SerialNoTo"))) 'Add 20090207
                
                .TextMatrix(i, bteColDate) = Format(Trim(rsGrid("schedule_date")), "dd MMM yyyy")
                .TextMatrix(i, bteColRemark) = IIf(IsNull(rsGrid("remark")), "", Trim(rsGrid("remark")))
                .TextMatrix(i, bteColSeqNo) = Val(rsGrid("seq_no"))
                
                If Val(rsGrid("auto_cls") & "") = 0 Then .TextMatrix(i, bteColAuto) = "No" Else .TextMatrix(i, bteColAuto) = "Yes"
                
                .TextMatrix(i, bteColCustCode) = rsGrid("cust_code") & ""
                .TextMatrix(i, bteColCustName) = rsGrid("trade_name") & ""
                .TextMatrix(i, bteColPONo) = rsGrid("po_no") & ""
                
                .Cell(flexcpBackColor, i, bteColSelect) = &HFFFFFF
                rsGrid.MoveNext
                i = i + 1
            Loop
        End With
    End If
    rsGrid.Close
    Set rsGrid = Nothing
    
End Sub

Function cekQty(ByVal i As String, ByVal pn As String, ByVal psn As String, ByVal Sn As String) As Double
    Dim sql1 As String, sql2 As String
    Dim rs1 As New Recordset, rs2 As New Recordset
    Dim OrderQty As Double, DailyQty As Double
    
    If pn = "Forecast" Then
        sql1 = "select isnull(qty,0) from production_planning " & _
               "where item_code='" & Trim(i) & "' and cust_code='000000' and lot_no='" & psn & "' "
    Else
        sql1 = "select isnull(qty,0) as orderqty From orderentry_detail " & _
               "where item_code='" & Trim(i) & "' and po_no='" & Trim(pn) & "' and seq_no=" & psn
    End If
    Set rs1 = Db.Execute(sql1)
    If Not (rs1.BOF And rs1.EOF) Then
        OrderQty = rs1(0)
    Else
        OrderQty = 0
    End If
    
    If Sn = "" Then Sn = 0
    
    sql2 = "select isnull(sum(qty),0) as dailyqty From daily_production " & _
           "where item_code='" & Trim(i) & "' and po_no='" & Trim(pn) & "' and po_seqno=" & psn & _
           " and seq_no<>" & CDbl(Sn)
    Set rs2 = Db.Execute(sql2)
    If Not (rs2.BOF And rs2.EOF) Then
        DailyQty = rs2(0)
    Else
        DailyQty = 0
    End If
    
    cekQty = OrderQty - DailyQty
End Function

Private Sub cboGroup_Change()
    If CboGroup.ListIndex <> -1 Then
        Label11.Caption = CboGroup.Column(1)
        Else
        Label11.Caption = "All"
    End If
    Browse
End Sub

Private Sub CboGroup_LostFocus()
If CboGroup.ListIndex = -1 Then CboGroup.ListIndex = 0
End Sub

Private Sub CboItemCode_Change()
    Call StartSerial
End Sub

Private Sub cboitemcode_Click()
    If cboitemcode.ListIndex <> -1 Then
        lblPart.Text = cboitemcode.Column(1)
        lbldesc.Text = cboitemcode.Column(2)
        lblunit1.Caption = cboitemcode.Column(3)
        lblunit.Caption = cboitemcode.Column(4)
    End If
    Call StartSerial
End Sub

Private Sub cboitemcode_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  Dim i As Integer
  If KeyCode = 13 Then
    For i = 0 To cboitemcode.ListCount - 1
        If cboitemcode.Text = cboitemcode.List(i) Then
            If cboitemcode.Column(1) = cboitemcode.List(i, 1) Then
                cboitemcode.ListIndex = i
                Exit For
            End If
        End If
    Next
  
    cboitemcode_Click
  End If
End Sub

Private Sub cboItemCode_LostFocus()
    Dim i As Integer
    For i = 0 To cboitemcode.ListCount - 1
        If cboitemcode.Text = cboitemcode.List(i) Then
            If cboitemcode.Column(1) = cboitemcode.List(i, 1) Then
                cboitemcode.ListIndex = i
                cboitemcode_Click
                Exit For
            End If
        End If
    Next
End Sub

Private Sub cmdBrowser_Click()
 If cboitemcode.Enabled = True Then
  Me.MousePointer = vbHourglass
  frm_BrowseItem.getItemCode = cboitemcode.Text
  frm_BrowseItem.Show 1
  cboitemcode.Text = frm_BrowseItem.getItemCode
  Me.MousePointer = vbDefault
 End If
End Sub

Private Sub Form_Load()
    
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
        
    adtocboCust
    adtocboGroup
    
    Call Kosong
    
    'GET START DAILY
    Dim Ret As String, NC As Long, TempPWD As String
    Ret = String(255, 0)
    NC = GetPrivateProfileString("StartDaily", "Date", "", Ret, 255, IniFile)
    If NC <> 0 Then Ret = Left$(Ret, NC)
    startDaily = Ret
End Sub

Private Sub cboCust_Click()
    If cbocust.ListIndex <> -1 Then
        lblcust.Caption = cbocust.Column(1)
        adtocbolinecd
    End If
    If cbocust.Text <> "" And cbolinecd.Text <> "" Then
        Browse
    Else
        Header
    End If
    addToCboItemCode
End Sub

Private Sub cboCust_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cboCust_Click
End Sub

Private Sub cbolinecd_Click()
    If cbolinecd.ListIndex <> -1 Then
        lbllinecd.Caption = cbolinecd.Column(1)
    End If
    If cbocust.Text <> "" And cbolinecd.Text <> "" Then
        Browse
    Else
        Header
    End If
End Sub

Private Sub cbolinecd_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cbolinecd_Click
End Sub

Private Sub scheduledate1_Change()
   If CDate(scheduledate1) > CDate(scheduledate2) Then
      lblErrMsg.Caption = DisplayMsg(4068)
      Exit Sub
   Else
      lblErrMsg.Caption = ""
   End If
    
    If cbocust.Text <> "" And cbolinecd.Text <> "" Then
        Browse
    End If
End Sub

Private Sub scheduledate2_Change()
   If CDate(scheduledate2) < CDate(scheduledate1) Then
      lblErrMsg.Caption = DisplayMsg(4066)
      Exit Sub
   Else
      lblErrMsg.Caption = ""
   End If
    If cbocust.Text <> "" And cbolinecd.Text <> "" Then
        Browse
    End If
End Sub

Private Sub txtqty_Change()
If txtqty <> "" Then
    TxtSerialFrom = StrStartSerial
    If TxtSerialFrom <> "" Then TxtSerialTo = GetSerialTo(Trim(TxtSerialFrom), txtqty)
End If
End Sub

Private Sub txtqty_GotFocus()
'SendKeys "{home}"
'SendKeys "+{End}"
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
          KeyAscii = 0
    End If
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim TextGrid As String

With grid
    TextGrid = grid.Text

    If TextGrid = "S" Then
        Status = "update"
        cboitemcode.Text = .TextMatrix(Row, bteColProdCode)
        cboitemcode.Enabled = False
        lblPart.Text = .TextMatrix(Row, bteColPart)
        lbldesc.Text = .TextMatrix(Row, bteColDesc)
        txtlotno.Text = .TextMatrix(Row, bteColLotNo)
        lbllotno = .TextMatrix(Row, bteColLotNo)
        
        If InStr(1, .TextMatrix(Row, bteColQty), ".") > 0 Then
            txtqty.Text = Format(.TextMatrix(Row, bteColQty), gs_formatQty)
        Else
            txtqty.Text = Format(.TextMatrix(Row, bteColQty), gs_formatQty)
        End If
        lblQty = CDbl(.TextMatrix(Row, bteColQty))
                        
        lblunit1.Caption = .TextMatrix(Row, bteColUnitCls)
        lblunit.Caption = .TextMatrix(Row, bteColUnit)
        scheduledate.Value = .TextMatrix(Row, bteColDate)
        lblschdate = .TextMatrix(Row, bteColDate)
        txtremark1.Text = .TextMatrix(Row, bteColRemark)
        txtseqno.Text = .TextMatrix(Row, bteColSeqNo)
        TxtSerialFrom = .TextMatrix(Row, BteColSerialFrom)
        TxtSerialTo = .TextMatrix(Row, BteColSerialTo)
       Call kosongColGrid
       lblErrMsg = ""
    ElseIf TextGrid = "D" Then
       Call kosongColGrid("S")
       lblErrMsg = ""
    End If

    .TextMatrix(Row, Col) = TextGrid
    If grid.TextMatrix(Row, bteColAuto) = "Yes" Then Auto_Cls = 1 Else Auto_Cls = 0
End With

End Sub

Private Sub kosongColGrid(Optional Kolom As String)
    Dim i As Integer
    
    With grid
        .Col = 0
    
        If Kolom <> "" Then
           For i = 1 To .Rows - 1

              If .Text = Kolom Then .Text = ""
              If .TextMatrix(i, bteColSelect) <> "D" Then .TextMatrix(i, bteColSelect) = ""
           Next i
           kosongBwh
        Else
           For i = 1 To .Rows - 1

              If .TextMatrix(i, bteColSelect) <> "" Then .TextMatrix(i, bteColSelect) = ""

           Next i
           
        End If
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Col <> bteColSelect Then Cancel = True
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
  If grid.Col = bteColSelect Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> Asc("S") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
  End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim i As Integer
Dim qty1 As Double
Dim sqlcekf As String, rscekf As New Recordset
Dim tanya
Dim sql1 As String
Dim rs1 As New ADODB.Recordset
Dim hapus As Boolean
Dim bedaQty As Boolean, bedaBulan As Boolean, bedaLot As Boolean
Dim dbw As New Connection
Dim rsUpdate As New ADODB.Recordset
Dim req_cls As String
Dim sql As String

Dim Cekserial As String

Dim X As Long
Dim awal As Long, akhir As Long, Panjang As Integer
Dim TempSerial As String, Depan As String
Dim rsCek As New ADODB.Recordset

dbw.ConnectionString = Db.ConnectionString

lblErrMsg = ""
ubah = False

On Error Resume Next
Select Case Index
  Case 0:   If hakUpdate(Me.Name) = 0 Then _
            lblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
            
            If cbocust.Text = "" Then
              cbocust.SetFocus
              lblErrMsg = DisplayMsg(1040)  '"Please Select Factory Code"
              Exit Sub
            ElseIf cbolinecd.Text = "" Then
              cbolinecd.SetFocus
              lblErrMsg = DisplayMsg(1041)  '"Please Select Machine No"
              Exit Sub
            End If
                        
            If cbocust.Text <> "" Then
                cbocust.MatchEntry = 1
                cbocust.Text = cbocust.Text
                If cbocust.MatchFound = False Then
                    lblErrMsg = DisplayMsg(4016)    'Record with this Factory code not found
                    cbocust.SetFocus
                    cbocust.MatchEntry = 2
                    Exit Sub
                End If
                cbocust.MatchEntry = 2
            End If
            
                        
            If cbolinecd.Text <> "" Then
                cbolinecd.MatchEntry = 1
                cbolinecd.Text = cbolinecd.Text
                If cbolinecd.MatchFound = False Then
                    lblErrMsg = DisplayMsg(4017)    'Record With this Machine No. not found
                    cbolinecd.SetFocus
                    cbolinecd.MatchEntry = 2
                    Exit Sub
                End If
                cbolinecd.MatchEntry = 2
            End If
                
            Cekserial = ""
            Cekserial = uf_ValidasiSerialNo
                                        
            If Cekserial <> "" Then
                
               lblErrMsg.Caption = Cekserial
               Exit Sub
            End If
            
                                       
                           
                
            If cboitemcode.Text <> "" Then
                cboitemcode.MatchEntry = 1
                cboitemcode.Text = cboitemcode.Text
                If cboitemcode.MatchFound = False Then
                    '**** cek Data cocok / tdk dgn Database
                    Dim rsDB As New ADODB.Recordset
                    rsDB.Open "select Item_Code from Item_Master where Item_Code='" & Trim(cboitemcode.Text) & "'", Db, adOpenKeyset, adLockOptimistic
                    If rsDB.EOF = True Then
                        lblErrMsg = DisplayMsg(4003)    'Record With this Product No. not found
                        cboitemcode.SetFocus
                        cboitemcode.MatchEntry = 2
                        Exit Sub
                    End If
                    rsDB.Close
                End If
            End If
            
            If Auto_Cls = 1 Then
                lblErrMsg = DisplayMsg(8111)    'Cannot modify data that generated automatically
                Exit Sub
            End If
            
                With grid
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, bteColSelect) = "D" Then
                            If IsEmpty(tanya) Then tanya = MsgBox("Do You Really Want To Delete This Data ?", vbQuestion & vbYesNo, "Confirmation")
                            If tanya = vbYes Then
                                sql1 = "select * from part_receipt where dailyseq_no='" & .TextMatrix(i, bteColSeqNo) & "' and receipt_cls = 'P1'"
                                Set rs1 = Db.Execute(sql1)
                                If Not (rs1.BOF And rs1.EOF) Then
                                    lblErrMsg.Caption = DisplayMsg(1204)
                                    .Row = i
                                    .SetFocus
                                    Exit Sub
                                End If

                                sql1 = "select complete_cls from daily_production where seq_no='" & .TextMatrix(i, bteColSeqNo) & "' and complete_cls='1' "
                                Set rs1 = Db.Execute(sql1)
                                If Not (rs1.BOF And rs1.EOF) Then
                                    lblErrMsg.Caption = DisplayMsg("0058")
                                    .Row = i
                                    .SetFocus
                                    Exit Sub
                                End If
                                
                                sql1 = "select Request_Cls from daily_production where seq_no='" & .TextMatrix(i, bteColSeqNo) & "' and coalesce(Request_Cls,'')<>''"
                                Set rs1 = Db.Execute(sql1)
                                If Not (rs1.BOF And rs1.EOF) Then
                                    req_cls = Trim(rs1.Fields("request_cls"))
                                End If

                                sql1 = "select * from PartSupplyRequest_Master where Request_Cls='" & req_cls & "'"
                                Set rs1 = Db.Execute(sql1)
                                If Not (rs1.BOF And rs1.EOF) Then
                                    lblErrMsg.Caption = DisplayMsg(1204)
                                    .Row = i
                                    .SetFocus
                                    Exit Sub
                                End If
                                
                                
                                dbw.Open
                                dbw.BeginTrans
                                
                        ' Add Validation for Serial Number on Process before Delete
                        ' Update 20090205
                        If .TextMatrix(i, BteColSerialFrom) <> "" And .TextMatrix(i, BteColSerialTo) <> "" Then
                            Panjang = Len(Trim(.TextMatrix(i, BteColSerialFrom)))

                            awal = Val(Mid(.TextMatrix(i, BteColSerialFrom), 2, (Panjang - 1)))
                            akhir = Val(Mid(.TextMatrix(i, BteColSerialTo), 2, (Panjang - 1)))

                            For X = awal To akhir
                                TempSerial = Left(.TextMatrix(i, BteColSerialFrom), 1) & Format(X, String(Panjang - 1, "0"))
                                sql1 = "Select * from Serial_Detail where item_code='" & .TextMatrix(i, bteColProdCode) & "' and " & _
                                    vbLf & " Serial_No='" & TempSerial & "' And Serial_Status<>'2'"
                                Set rsCek = dbw.Execute(sql1)
                                If Not rsCek.EOF Then
                                    lblErrMsg = "[000]- Serial Number " & Trim(TempSerial) & " is in processed and can't delete !!! "
                                    rsCek.Close
                                    Me.MousePointer = vbDefault
                                    Exit Sub
                                End If
                            Next X
                        End If
                        
                        ' ----------------------------
                                'Hapus Bom Production 202520507
                                sql = "EXEC dbo.sp_BomMasterProduction_InsDel " & _
                                    "@Seqno = '" & .TextMatrix(i, bteColSeqNo) & "', " & _
                                    "@UserID = '" & Replace(userLogin, "'", "''") & "', " & _
                                    "@Type = '2'"
                        
                                ' Eksekusi SQL
                                dbw.Execute sql
                        
                                sql1 = "delete from daily_production where seq_no='" & .TextMatrix(i, bteColSeqNo) & "' "
                                dbw.Execute sql1
                                
                                sql1 = "Delete EZRunnerV3_KawaiLive.dbo.Serial_Detail" & vbCrLf & _
                                        "where Serial_No  between '" & .TextMatrix(i, BteColSerialFrom) & "' and '" & .TextMatrix(i, BteColSerialTo) & "'"
                                
                                dbw.Execute sql1
                                
                        
                        ' Update Serial Number Back to Order Data
                        ' Update 20090204
                        If .TextMatrix(i, BteColSerialFrom) <> "" And .TextMatrix(i, BteColSerialTo) <> "" Then
                            Panjang = Len(Trim(.TextMatrix(i, BteColSerialFrom)))
                            
                            awal = Val(Mid(.TextMatrix(i, BteColSerialFrom), 2, (Panjang - 1)))
                            akhir = Val(Mid(.TextMatrix(i, BteColSerialTo), 2, (Panjang - 1)))
                            
                            For X = awal To akhir
                                TempSerial = Left(.TextMatrix(i, BteColSerialFrom), 1) & Format(X, String(Panjang - 1, "0"))
                                sql1 = "Update Serial_Detail " & _
                                    " Set Product_No=Null, Serial_Status='1' where Product_No='" & .TextMatrix(i, bteColSeqNo) & _
                                    "'  and item_code='" & .TextMatrix(i, bteColProdCode) & "' and " & _
                                    " Serial_No='" & TempSerial & "'"
                                dbw.Execute sql1
                            Next X
                        End If
                        ' -----------------------------
                        
                                dbw.CommitTrans
                                dbw.Close
                                
                                hapus = True
                            
                            Else
                                Exit For
                            End If
                        ElseIf .TextMatrix(i, bteColSelect) = "S" Then
                            sql1 = "select complete_cls from daily_production where seq_no='" & .TextMatrix(i, bteColSeqNo) & "' and complete_cls='1' "
                            Set rs1 = Db.Execute(sql1)
                            If Not (rs1.BOF And rs1.EOF) Then
                                lblErrMsg.Caption = DisplayMsg("0059")
                                .Row = i
                                .SetFocus
                                Exit Sub
                            End If
                            
                           ubah = True
                           
                            sql = "select Request_Cls from daily_production where seq_no='" & grid.TextMatrix(i, bteColSeqNo) & "' "
                            Set rs1 = Db.Execute(sql)
                            If Not (rs1.BOF And rs1.EOF) Then
                                req_cls = Trim(rs1.Fields("request_cls"))
                            End If
                                    
                            'Validasi cek ke tabel PartSupplyRequest_Master
                            sql1 = "select * from PartSupplyRequest_Master where Request_Cls='" & req_cls & "'"
                            Set rs1 = Db.Execute(sql1)
                            If Not (rs1.BOF And rs1.EOF) Then
                                lblErrMsg.Caption = DisplayMsg("9010")
                                grid.Row = i
                                Exit Sub
                            End If
                        
                        End If
                                
                    Next i
                    
                    If (hapus) Then kosongBwh: Browse: lblErrMsg = DisplayMsg(1201): Exit Sub
                End With
'
                
                'Schedule Date harus lebih besar dari Start Daily
                If Format(scheduledate.Value, "yyyy-mm-dd") <= Format(startDaily, "yyyy-mm-dd") Then
                    lblErrMsg.Caption = DisplayMsg("0060") & "" & Format(startDaily, "dd MMM yyyy") & " !"
                    If scheduledate.Enabled Then scheduledate.SetFocus
                    Exit Sub
                End If
                
                'Validasi Schedule Date
                sql1 = "select * from calendar_master where factory_code = '" & cbocust.Text & "' and cal_date = '" & CDate(scheduledate.Value) & "'"
                Set rs1 = Db.Execute(sql1)
                If Not (rs1.BOF And rs1.EOF) Then
                  lblErrMsg.Caption = DisplayMsg(1022)
                  command1(3).SetFocus
                  Exit Sub
                'Bila ada perubahan Schedule Date maka perubahan harus tetap dalam Bulan atau Tahun yg sama
                ElseIf ubah And Not (Month(scheduledate.Value) = Month(lblschdate) And Year(scheduledate.Value) = Year(lblschdate)) Then
                   sql1 = "select isnull(sum(qty),0) Qtyresult from part_receipt " & _
                          "where receipt_cls='P1' and dailySeq_no = " & grid.TextMatrix(i, bteColSeqNo)
                   Set rs1 = Db.Execute(sql1)
                   If Not rs1.EOF Then
                        If rs1("Qtyresult") > 0 Then
                            lblErrMsg.Caption = DisplayMsg(1022)
                            command1(3).SetFocus
                            Exit Sub
                        End If
                   End If
                End If
                
                If cboitemcode.Text <> "" Or txtqty.Text <> "" Or txtremark1.Text <> "" Or txtlotno.Text <> "" Then
                
                    If cboitemcode.Text = "" Then
                       cboitemcode.SetFocus
                       lblErrMsg = DisplayMsg(1024)  '"Please Select Product code"
                       Exit Sub
                    ElseIf txtlotno = "" Then
                       txtlotno.SetFocus
                       lblErrMsg = " [2008] Please Input LOT Number "
                       Exit Sub
                    ElseIf CDbl(txtqty.Text) = 0 Then
                       txtqty.SetFocus
                       SendKeys "{home}"
                       SendKeys "+{end}"
                       lblErrMsg = DisplayMsg(1012) '"Please Input Quantity"
                       Exit Sub
                    ElseIf CDbl(txtqty.Text) > gd_MaxQty Then
                       txtqty.SetFocus
                       lblErrMsg = DisplayMsg(4045) & " " & gd_MaxQty & " !"
                       Exit Sub
                    Else
                    
                        ' Validation of Serial No
                        ' Update 20090207
                    If Trim(TxtSerialFrom) <> "" And Trim(TxtSerialTo) <> "" Then
                        Panjang = Len(Trim(TxtSerialFrom))
                        Depan = Left(TxtSerialFrom, 1)
                                            
                        awal = Val(Mid(TxtSerialFrom, 2, Panjang - 1))
                        akhir = Val(Mid(Trim(TxtSerialTo), 2, Panjang - 1))
                        
                        ' Check For Valid Serial No
                        If awal > akhir Then
                            lblErrMsg = "[000] - Invalid Serial Number ! "
                            Me.MousePointer = vbDefault
                            Exit Sub
                        End If
                        
                        If Depan <> Left(TxtSerialTo, 1) Then
                            lblErrMsg = "[000] - Invalid Serial Number ! "
                            Me.MousePointer = vbDefault
                            Exit Sub
                        End If
                        
                        If (akhir - awal) + 1 <> CDbl(txtqty) Then
                            lblErrMsg = "[000] - Data doesn't match between Qty and Serial No ! "
                            Me.MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                    
                    ' ----------------------------
                    
                       Dim NO As Long
                       Dim Test As String
                       Dim SerialStatus As Boolean
                       SerialStatus = False
                       
                       dbw.Open
                       dbw.BeginTrans
                       
                       If ubah = False Then
                       
                            ' Check For Exist Serial No
                            ' When New Record Check Serial No in All Record
                            If Trim(TxtSerialFrom) <> "" And Trim(TxtSerialTo) <> "" Then
                                For X = awal To akhir
                                    TempSerial = Depan & Format(X, String(Panjang - 1, "0"))
                                ' sql1 = "Select * from Serial_Detail where item_code='" & cboitemcode.Text & "' and "
                                   sql1 = "Select * from Serial_Detail where --item_code='" & cboitemcode.Text & "' and " & vbCrLf & _
                                    " Serial_NO='" & TempSerial & "' "
                                    Set rsCek = Db.Execute(sql1)
                                    If Not rsCek.EOF Then
                                        If Not IsNull(rsCek("Product_No")) Then
                                            lblErrMsg = "[000] - Serial Number Already Process at Daily Number : " & Trim(rsCek("Product_No"))
                                            Me.MousePointer = vbDefault
                                            Exit Sub
                                        End If
                                    Else
'                                        lblErrMsg = "[000] - Serial Number " & TempSerial & " doesn't Exist at Order Data !!! "
'                                        Me.MousePointer = vbDefault
'                                        Exit Sub
                                        sql1 = "Insert Into Serial_Detail (Item_Code,Serial_No,Po_No,PO_SeqNo,Serial_Status) Values ('" & Trim(cboitemcode) & "'," & _
                                        vbLf & "'" & Trim(TempSerial) & "','',0,'1')"
                                        dbw.Execute (sql1)
                                        SerialStatus = True
                                    End If
                                Next X
                            End If
                                ' -------------------------
                       
                            sql1 = "select a.*, b.item_name from daily_production a left join item_master b on b.item_code=a.item_code " & _
                                      "where a.factory_code='" & cbocust.Text & "' and " & _
                                      "a.line_code='" & cbolinecd.Text & "' and a.schedule_date>='" & Format(scheduledate1.Value, "yyyymmdd") & _
                                      "' and a.schedule_date<='" & Format(scheduledate2.Value, "yyyymmdd") & "' order by a.schedule_date, a.item_code, a.lot_no, a.seq_no"

                            Test = "Select * From daily_Production"
                            If rsUpdate.State <> adStateClosed Then rsUpdate.Close
                            rsUpdate.Open Test, dbw, adOpenKeyset, adLockOptimistic

                            'TAMBAHKAN RECORD BARU KE DAILY_PRODUCTION
                            rsUpdate.AddNew
                            rsUpdate("factory_Code") = cbocust.Text
                            rsUpdate("line_code") = cbolinecd.Text
                            rsUpdate("prod_barcode") = Trim(cbocust.Text) & Trim(cbolinecd.Text) & Format(scheduledate.Value, "YYYYMMDD") & seqNo
                            
                            NO = seqNo
                            rsUpdate("seq_no") = NO ' Udah mo nyimpen Niy
                            
                            
                       Else
                       
                        ' Check For Exist Serial No
                        ' When Update Record Check Serial No in All Record ( Not Include This Record )
                        X = 0
                        
                        If Trim(TxtSerialFrom) <> "" And Trim(TxtSerialTo) <> "" Then
                            For X = awal To akhir
                                TempSerial = Depan & Format(X, String(Panjang - 1, "0"))
                                
                                sql1 = " select * From ( " & vbLf & _
                                    " Select * From Serial_Detail Where Po_SeqNo is not null ) n " & vbLf & _
                                    " Where Item_Code='" & cboitemcode.Text & "' and Serial_No='" & TempSerial & "' and product_No <>'" & txtseqno & "'"
                                
'                                sql1 = " Select * From " & _
'                                    vbLf & " (Select * from Serial_Detail where item_code='" & cboItemCode.Text & "' and " & _
'                                    vbLf & "  Serial_No not in (Select Serial_No From Serial_Detail Where Product_No='" & txtseqno & "' " & _
'                                    vbLf & " )) chk Where Serial_No='" & TempSerial & "' "
                
                                Set rsCek = Db.Execute(sql1)
                                If Not rsCek.EOF Then
                                    If Not IsNull(rsCek("Product_No")) Then
                                        lblErrMsg = "[000] - Serial Number Already Process at Daily Number : " & Trim(rsCek("Product_No"))
                                        Me.MousePointer = vbDefault
                                        Exit Sub
                                    End If
                                Else
                                    sql1 = "Select * From Serial_Detail Where Item_Code='" & cboitemcode & "' And " & _
                                            "Serial_No='" & TempSerial & "'"
                                    Set rsCek = Db.Execute(sql1)
                                    If rsCek.EOF Then
'                                        lblErrMsg = "[000] - Serial Number " & TempSerial & " doesn't Exist at Order Data !!! "
'                                        Me.MousePointer = vbDefault
'                                        Exit Sub

                                        sql1 = "Insert Into Serial_Detail (Item_Code,Serial_No,Po_No,PO_SeqNo,Serial_Status) Values ('" & Trim(cboitemcode) & "'," & _
                                        vbLf & "'" & Trim(TempSerial) & "','',0,'1')"
                                        dbw.Execute (sql1)
                                        
                                        SerialStatus = True
                                    End If
                                End If
                            Next X
                        End If
                        ' -------------------------
                       
                            sqlGrid = "select * from daily_production where seq_no='" & txtseqno.Text & "' " & _
                                      "order by schedule_date, item_code, lot_no, seq_no"
                            If rsUpdate.State <> adStateClosed Then rsUpdate.Close
                            rsUpdate.Open sqlGrid, dbw, adOpenKeyset, adLockOptimistic
                            rsUpdate("prod_barcode") = Trim(cbocust.Text) & Trim(cbolinecd.Text) & Format(scheduledate.Value, "YYYYMMDD") & txtseqno.Text
                       End If
                       
                       rsUpdate("item_Code") = cboitemcode.Text
                       rsUpdate("lot_no") = txtlotno.Text
                       rsUpdate("schedule_date") = Format(scheduledate.Value, "YYYY-MM-DD")
                       rsUpdate("qty") = txtqty.Text
                       rsUpdate("unit_cls") = lblunit1.Caption
                       rsUpdate("remark") = txtremark1.Text
                       rsUpdate("SerialNoFrom") = TxtSerialFrom     ' Add 20090207
                       rsUpdate("SerialNoTo") = TxtSerialTo           ' Add 20090207
                       rsUpdate("auto_cls") = "0"
                       rsUpdate("Last_Update") = Now
                       rsUpdate("Last_User") = userLogin
                       rsUpdate.update
                       
                        ' Save and update Serial Number Status - 20090210
                        
                        sql1 = "Update Serial_Detail Set Product_No=NULL, " & _
                                 " Serial_Status = '1' Where item_Code='" & cboitemcode.Text & "' And " & _
                                 " Product_No='" & rsUpdate("seq_no") & "'"
                        dbw.Execute (sql1)
                        
                        If Trim(TxtSerialFrom) <> "" And Trim(TxtSerialTo) <> "" Then
                            ' Add Serial No Data base on Order Data
                           For X = awal To akhir
                                TempSerial = Depan & Format(X, String(Panjang - 1, "0"))
                                sql1 = "Update Serial_Detail Set Product_No='" & rsUpdate("seq_no") & "' , " & _
                                         " Serial_Status = '2' Where item_Code='" & cboitemcode.Text & "' And " & _
                                         " Serial_No='" & TempSerial & "'"
                                dbw.Execute (sql1)
                            Next X
                        End If
                        
                    ' --------------------
                       'ditambahkan insert ke bom production
                        On Error GoTo ErrHandler
                        
                        ' Susun SQL dengan menggandakan tanda petik tunggal untuk menghindari error/injeksi
                        sql = "EXEC dbo.sp_BomMasterProduction_InsDel " & _
                              "@Seqno = '" & Replace(rsUpdate("seq_no"), "'", "''") & "', " & _
                              "@UserID = '" & Replace(userLogin, "'", "''") & "', " & _
                              "@Type = '1'"
                        
                        ' Eksekusi SQL
                        dbw.Execute sql
                        
                                
                       dbw.CommitTrans
                       dbw.Close
                       
                       If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                            NO = seqNo
                            rsUpdate.update
                       End If
                                                
                        If CDate(scheduledate.Value) > CDate(scheduledate1.Value) Then
                            If CDate(scheduledate.Value) >= CDate(scheduledate2.Value) Then
                                scheduledate2.Value = Format(scheduledate.Value, "dd MMM yyyy")
                            End If
                        Else
                            scheduledate1.Value = Format(scheduledate.Value, "dd MMM yyyy")
                        End If
                    End If
                    lblErrMsg = IIf(ubah, DisplayMsg(1101), DisplayMsg(1000))
                    kosongBwh
                    Browse
                    ubah = True
                End If
    Case 1
        Command1_Click 3
        cbocust = ""
        lblcust = ""
        cboCust_Click
        cbolinecd = ""
        lbllinecd = ""
        scheduledate1.Value = Date
        scheduledate2.Value = Date
    Case 2:
            If cbocust.Text <> "" Then
            
            Dim application As New CRAXDDRT.application
            Dim report As New CRAXDDRT.report
            Dim rsRpt As New ADODB.Recordset
            Dim SqlRpt As String
            Dim Rpt As New FrmRpt3
            
            Me.MousePointer = vbHourglass
            
            SqlRpt = "select rtrim(a.factory_code) as factory_code, rtrim(a.line_code) as line_code, a.schedule_date, " & _
                     "rtrim(a.item_code) as item_code, rtrim(a.lot_no) as lot_no, a.seq_no, a.qty,rtrim(a.serialNoFrom) as serialNoFrom,rTrim(a.SerialNoTo) as SerialNoTo, rtrim(a.unit_cls) as unit_cls, (select description from unit_cls uc where uc.unit_cls=a.unit_cls) unit_desc,rtrim(a.remark) as remark1, " & _
                     "b.makeritem_code, b.item_name, (select trade_name from trade_master where trade_code = '" & cbocust.Text & "') factory_name " & _
                     "from daily_production a left join item_master b on b.item_code=a.item_code " & _
                     "where a.factory_code='" & cbocust.Text & "' and " & _
                     "a.line_code='" & cbolinecd.Text & "' and a.schedule_date>='" & Format(scheduledate1.Value, "yyyymmdd") & _
                     "' and a.schedule_date<='" & Format(scheduledate2.Value, "yyyymmdd") & "' order by a.schedule_date, a.item_code, a.lot_no, a.seq_no"

            If rsRpt.State <> adStateClosed Then rsRpt.Close
            rsRpt.Open SqlRpt, Db, adOpenDynamic, adLockOptimistic
            
            sqlprint = SqlRpt
            reportcode = "dailyproduction"
            printorient = 2
            If rsRpt.EOF Then lblErrMsg = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
            
            Set report = application.OpenReport(App.path & "\Reports\rptdailyproduction.rpt")
            report.Database.Tables(1).SetDataSource rsRpt
            report.FormulaFields(1).Text = "'" & Format(scheduledate1.Value, "dd MMM yyyy") & "'"
            report.FormulaFields(2).Text = "'" & Format(scheduledate2.Value, "dd MMM yyyy") & "'"
            
            '#####################################################################
            '# Qty Digit and decimal
            report.FormulaFields(4).Text = "" & gi_decimalDigitQty & ""
            report.FormulaFields(5).Text = "" & gi_decimalDigitQty & ""
            '#####################################################################
            
            Fbulan = scheduledate1.Value
            Ftahun = scheduledate2.Value
            
            Rpt.CRViewer1.ReportSource = report
            Rpt.CRViewer1.ViewReport
            Rpt.CRViewer1.Zoom 1
            
            Rpt.WindowState = 2
            Rpt.Show 1
            
            Me.MousePointer = vbDefault
            End If
            
    Case 3:
            kosongColGrid
            kosongBwh
            Browse

ErrHandler:
                            MsgBox "Terjadi kesalahan saat menjalankan prosedur:" & vbCrLf & _
                                   err.Description, vbCritical, "Database Error"
End Select
End Sub

Private Sub command3_Click()
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



Private Sub txtQty_LostFocus()
    If IsNumeric(txtqty.Text) Then
        txtqty.Text = Format(txtqty.Text, gs_formatQty)
    Else
        txtqty.Text = Format(0, gs_formatQty)
    End If
End Sub

Private Sub TxtSerialFrom_LostFocus()
If TxtSerialFrom <> "" Then TxtSerialTo.Text = GetSerialTo(Trim(TxtSerialFrom), txtqty)
If TxtSerialFrom = "" Then TxtSerialTo = ""
End Sub

Private Sub StartSerial()
Dim RsSerial As New ADODB.Recordset
Dim strSQL As String
Dim X As Integer
Dim strAwal As String
Dim LongAwal As Long

'On Error Resume Next

strSQL = "Select * From Item_Master Where Item_Code='" & Trim(cboitemcode) & "'"
Set RsSerial = Db.Execute(strSQL)

If RsSerial.EOF Then
    RsSerial.Close
    Exit Sub
End If

If RsSerial("Finishgoodpart_Cls") = "02" Then
    RsSerial.Close
    Exit Sub
End If
RsSerial.Close

strSQL = "Select Max(SerialNoTo) From Daily_Production Where Left(SerialNoTo,1)='G'"
strAwal = "G"

'For x = 1 To Len(Trim(LblDesc))
'    If Mid(Trim(LblDesc), x, 1) = "D" Then
'        strAwal = "E"
'        StrSql = "Select Max(SerialNoTo) From Daily_Production Where Left(SerialNoTo,1)='E'"
'        Exit For
'    End If
'Next

If Right(Trim(lbldesc), 1) = "D" Then
        strAwal = "E"
        strSQL = "Select Max(SerialNoTo) From Daily_Production Where Left(SerialNoTo,1)='E'"
End If

Set RsSerial = Db.Execute(strSQL)

If IsNull(RsSerial(0)) Then
    StrStartSerial = strAwal & Right(Str(1000001), 6)
Else
    StrStartSerial = strAwal & Right(Str(Val(Right(RsSerial(0), 6)) + 1000001), 6)
End If

End Sub



Private Function uf_ValidasiSerialNo() As String

Dim RsSerial As New ADODB.Recordset
Dim strSQL As String, seqNo As Integer
lblErrMsg.Caption = ""


If txtseqno.Text = "" Then
    seqNo = 0
Else
    seqNo = txtseqno.Text
End If

strSQL = "exec SP_DailyProductionEntry_Validasi_SerialNo " & seqNo & ",'" & Trim(TxtSerialFrom.Text) & "','" & Trim(TxtSerialTo.Text) & "'    "

Set RsSerial = Db.Execute(strSQL)
i = 1
If Not (RsSerial.BOF And RsSerial.EOF) Then
        With grid
            Do While Not RsSerial.EOF
                If i = 1 Then
                    uf_ValidasiSerialNo = "Serial No Already Exist..! " & uf_ValidasiSerialNo & "Item Code :" & RsSerial("Item_code") & " - Date : " & Format(RsSerial("Schedule_Date"), "dd MMM yyyy")
                Else
                    uf_ValidasiSerialNo = uf_ValidasiSerialNo & "," & RsSerial("Serial_No")
                End If
                RsSerial.MoveNext
                i = i + 1
            Loop
        End With
End If
    RsSerial.Close


End Function
