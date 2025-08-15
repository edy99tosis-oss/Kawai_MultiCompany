VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOrderEntry 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Entry (Update)"
   ClientHeight    =   10605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "frmOrderEntry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10605
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   10035
      Width           =   1425
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Stuffing Report"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   10035
      Width           =   1485
   End
   Begin VB.TextBox txtErr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   79
      Top             =   7080
      Width           =   1785
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Upload Detail"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   10035
      Width           =   1485
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Loading Form"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   10035
      Width           =   1485
   End
   Begin VB.TextBox txtDestination 
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
      Height          =   270
      Left            =   8490
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   8850
      Width           =   1545
   End
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
      Left            =   10230
      MaxLength       =   7
      TabIndex        =   15
      Top             =   7890
      Width           =   1455
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
      Left            =   8730
      MaxLength       =   7
      TabIndex        =   14
      Top             =   7890
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker delDate 
      Height          =   315
      Left            =   11730
      TabIndex        =   16
      Top             =   7890
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
      CustomFormat    =   "dd MMM yyyy"
      Format          =   131465219
      CurrentDate     =   39777
   End
   Begin VB.TextBox txtremarks 
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
      Left            =   4650
      MaxLength       =   35
      TabIndex        =   21
      Top             =   8820
      Width           =   2775
   End
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "..."
      Height          =   315
      Left            =   2340
      TabIndex        =   23
      Top             =   7890
      Width           =   300
   End
   Begin VB.TextBox txtRevisi 
      Alignment       =   2  'Center
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
      Left            =   5925
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2340
      Width           =   450
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
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   8820
      Width           =   2775
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
      Left            =   10140
      Locked          =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   8820
      Width           =   2025
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FDDFE3&
      Caption         =   "No Commercial"
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
      Left            =   9540
      TabIndex        =   9
      Top             =   2400
      Width           =   1620
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
      Height          =   330
      Left            =   10470
      TabIndex        =   54
      Top             =   10065
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox lbldesc 
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
      Height          =   315
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   7890
      Width           =   4155
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
      Left            =   11670
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   10035
      Width           =   1125
   End
   Begin VB.TextBox txtpono 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   3000
      MaxLength       =   25
      TabIndex        =   6
      Top             =   2370
      Width           =   1995
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
      Left            =   11310
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2310
      Width           =   1125
   End
   Begin VB.TextBox txtamount 
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
      Left            =   2850
      Locked          =   -1  'True
      MaxLength       =   24
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   8820
      Width           =   1665
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   480
      Left            =   120
      TabIndex        =   36
      Top             =   9420
      Width           =   15075
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
         Left            =   60
         TabIndex        =   37
         Top             =   150
         Width           =   14940
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
      Left            =   6990
      MaxLength       =   9
      TabIndex        =   13
      Top             =   7890
      Width           =   975
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
      Left            =   14070
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   10035
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
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   10035
      Width           =   1125
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
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   10035
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
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   10035
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
      Left            =   2670
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   10035
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
      Left            =   1470
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   10035
      Visible         =   0   'False
      Width           =   1125
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
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   10035
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   4230
      Left            =   60
      TabIndex        =   11
      Top             =   2775
      Width           =   15075
      _cx             =   26591
      _cy             =   7461
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
      Begin VB.TextBox txtlocation 
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
         Left            =   5280
         MaxLength       =   35
         TabIndex        =   78
         Top             =   7440
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1185
      Left            =   90
      TabIndex        =   39
      Top             =   1020
      Width           =   15075
      Begin VB.CommandButton cmd_Browser 
         Caption         =   "..."
         Height          =   300
         Left            =   5895
         TabIndex        =   82
         Top             =   705
         Width           =   300
      End
      Begin VB.TextBox lbldelplace 
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
         Height          =   270
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   712
         Width           =   2355
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
         Height          =   285
         Index           =   1
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   270
         Width           =   5715
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
         Height          =   285
         Index           =   0
         Left            =   3390
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   300
         Width           =   4635
      End
      Begin VB.TextBox txtcontact 
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
         MaxLength       =   25
         TabIndex        =   2
         Top             =   705
         Width           =   2610
      End
      Begin MSComCtl2.DTPicker deliverydate1 
         Height          =   315
         Left            =   11730
         TabIndex        =   3
         Top             =   690
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
         Format          =   131465219
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker deliverydate2 
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
         Left            =   13350
         TabIndex        =   4
         Top             =   690
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
         Format          =   131465219
         CurrentDate     =   37798
      End
      Begin VB.Label Label3 
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
         Index           =   0
         Left            =   10440
         TabIndex        =   47
         Top             =   750
         Width           =   1185
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
         Left            =   12810
         TabIndex        =   46
         Top             =   750
         Width           =   165
      End
      Begin VB.Line Line1 
         X1              =   3390
         X2              =   5820
         Y1              =   990
         Y2              =   990
      End
      Begin MSForms.ComboBox cbodelplace 
         Height          =   315
         Left            =   1410
         TabIndex        =   1
         Top             =   690
         Width           =   1785
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "3149;556"
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
         Caption         =   "Consignee"
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
         Left            =   120
         TabIndex        =   45
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
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
         Left            =   6300
         TabIndex        =   44
         Top             =   750
         Width           =   1305
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   9090
         X2              =   14790
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Label Label3 
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
         Index           =   4
         Left            =   8190
         TabIndex        =   43
         Top             =   345
         Width           =   690
      End
      Begin MSForms.ComboBox cbocust 
         Height          =   315
         Left            =   1410
         TabIndex        =   0
         Top             =   285
         Width           =   1785
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "3149;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3390
         X2              =   8100
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Label LblCode 
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
         TabIndex        =   40
         Top             =   345
         Width           =   1170
      End
   End
   Begin MSComCtl2.DTPicker podate 
      Height          =   315
      Left            =   7800
      TabIndex        =   8
      Top             =   2340
      Width           =   1530
      _ExtentX        =   2699
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
      Format          =   131465219
      CurrentDate     =   37798
   End
   Begin MSComCtl2.DTPicker deltime 
      Height          =   315
      Left            =   13200
      TabIndex        =   17
      Top             =   7890
      Width           =   930
      _ExtentX        =   1640
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
      CustomFormat    =   "HH:mm "
      Format          =   131465219
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13320
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   240
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invalid [Item Code] or [Price]"
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
      Left            =   10560
      TabIndex        =   80
      Top             =   7080
      Width           =   2580
   End
   Begin VB.Line Line4 
      X1              =   8430
      X2              =   10110
      Y1              =   9135
      Y2              =   9135
   End
   Begin MSForms.ComboBox cboDestination 
      Height          =   315
      Left            =   7530
      TabIndex        =   74
      Top             =   8820
      Width           =   810
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "1429;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Final Place of Destination"
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
      Left            =   7530
      TabIndex        =   73
      Top             =   8490
      Width           =   2145
   End
   Begin MSForms.CheckBox cbEditStatus 
      Height          =   255
      Left            =   2340
      TabIndex        =   72
      Top             =   8850
      Width           =   225
      BackColor       =   16637923
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "397;450"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label15 
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
      Left            =   10230
      TabIndex        =   71
      Top             =   7530
      Width           =   1050
   End
   Begin VB.Label Label14 
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
      Left            =   8730
      TabIndex        =   70
      Top             =   7530
      Width           =   1275
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      TabIndex        =   69
      Top             =   8520
      Width           =   660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curr"
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
      Left            =   14220
      TabIndex        =   68
      Top             =   7530
      Width           =   390
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   780
      TabIndex        =   67
      Top             =   8490
      Width           =   420
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
      Index           =   0
      Left            =   11820
      TabIndex        =   66
      Top             =   7530
      Width           =   1185
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
      Left            =   7620
      TabIndex        =   65
      Top             =   7530
      Width           =   300
   End
   Begin VB.Label Label6 
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
      Index           =   0
      Left            =   240
      TabIndex        =   64
      Top             =   7530
      Width           =   1080
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
      Left            =   2760
      TabIndex        =   63
      Top             =   7530
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Index           =   0
      Left            =   8040
      TabIndex        =   62
      Top             =   7530
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   2190
      TabIndex        =   61
      Top             =   8490
      Width           =   540
   End
   Begin MSForms.ComboBox cboServices 
      Height          =   285
      Left            =   3090
      TabIndex        =   20
      Top             =   8820
      Visible         =   0   'False
      Width           =   1980
      VariousPropertyBits=   746604571
      MaxLength       =   16
      DisplayStyle    =   3
      Size            =   "3492;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
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
      Index           =   2
      Left            =   4680
      TabIndex        =   60
      Top             =   8490
      Width           =   765
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rev."
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
      Index           =   25
      Left            =   5415
      TabIndex        =   59
      Top             =   2400
      Width           =   390
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
      Left            =   13830
      TabIndex        =   57
      Top             =   8490
      Width           =   1140
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SI/PO No"
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
      Left            =   10170
      TabIndex        =   55
      Top             =   8490
      Width           =   795
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   60
      Top             =   8400
      Width           =   15075
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   540
      Index           =   0
      Left            =   60
      Top             =   8670
      Width           =   15075
   End
   Begin VB.Label lblfix 
      Alignment       =   1  'Right Justify
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
      Left            =   14100
      TabIndex        =   49
      Top             =   2400
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
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
      Left            =   13200
      TabIndex        =   48
      Top             =   7530
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SI/PO No"
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
      Left            =   1965
      TabIndex        =   42
      Top             =   2400
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SI/PO Date"
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
      Left            =   6675
      TabIndex        =   41
      Top             =   2400
      Width           =   975
   End
   Begin MSForms.ComboBox cbopono 
      Height          =   315
      Left            =   2970
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2340
      Width           =   2340
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "4128;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox combo1 
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   2340
      Width           =   1290
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "2275;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboprice 
      Height          =   285
      Left            =   120
      TabIndex        =   19
      Top             =   8820
      Width           =   1980
      VariousPropertyBits=   746604571
      MaxLength       =   16
      DisplayStyle    =   3
      Size            =   "3492;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cbocurr 
      Height          =   315
      Left            =   14190
      TabIndex        =   18
      Top             =   7890
      Width           =   855
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "1508;556"
      TextColumn      =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cbounit 
      Height          =   315
      Left            =   8010
      TabIndex        =   28
      Top             =   7890
      Width           =   630
      VariousPropertyBits=   746604575
      BackColor       =   14737632
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "1111;556"
      BoundColumn     =   2
      TextColumn      =   2
      ListRows        =   0
      MatchEntry      =   1
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line3 
      X1              =   2760
      X2              =   6900
      Y1              =   8220
      Y2              =   8220
   End
   Begin MSForms.ComboBox cboitemcode 
      Height          =   315
      Left            =   180
      TabIndex        =   12
      Top             =   7890
      Width           =   2085
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "3678;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label LblPage 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
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
      TabIndex        =   38
      Top             =   9360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Order Entry (Update)"
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
      Left            =   90
      TabIndex        =   35
      Top             =   360
      Width           =   15075
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   540
      Index           =   2
      Left            =   60
      Top             =   7800
      Width           =   15075
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   60
      Top             =   7440
      Width           =   15075
   End
End
Attribute VB_Name = "frmOrderEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Update By Dudi November 25 2008, menambah Services

Option Explicit
  
Dim sql As String, sqlGrid As String
Dim ubah As Boolean, ubahgrid As Boolean, ada As Boolean
Dim Tgl As Date
Dim lblQty As Double
Dim statusfix As String
Dim Lotforecast As String
Public frmpanggil As String

Dim ls_PathExcel As String

Dim bteColSelect As Byte
Dim bteColItemCode As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColQty As Byte
Dim bteColUnitCls As Byte
Dim bteColUnit As Byte
Dim BteColSerialFrom As Byte
Dim BteColSerialTo As Byte
Dim bteColDate As Byte
Dim bteColTime As Byte
Dim bteColCurrCode As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColRemark As Byte
Dim bteColServices As Byte
Dim bteColCur1 As Byte
Dim bteColCalculate As Byte
Dim bteColGenerate As Byte
Dim bteColSeqNo As Byte
Dim bteColFinalDestination As Byte
Dim bteColCurrStock As Byte

Dim bteColStatusEdit As Byte 'added 23/082016 by do

Dim bteHakPrice As Byte
'Dim CalculateCls  As Byte

Sub Header()
        
    Dim i As Integer
    
    bteColSelect = 0
    bteColItemCode = 1
    bteColPartNo = 2
    bteColDesc = 3
    bteColQty = 4
    bteColCurrStock = 5
    bteColUnitCls = 6
    bteColUnit = 7
    BteColSerialFrom = 8
    BteColSerialTo = 9
    bteColDate = 7 + 3
    bteColTime = 8 + 3
    bteColCurrCode = 9 + 3
    bteColCurr = 10 + 3
    bteColPrice = 11 + 3
    bteColServices = 12 + 3
    'bteColCur1 = 13+2
    bteColAmount = 13 + 3
    bteColRemark = 14 + 3
    bteColCalculate = 15 + 3
    bteColGenerate = 16 + 3
    bteColSeqNo = 17 + 3
    bteColStatusEdit = 18 + 3
    bteColFinalDestination = 22
    
    With grid
        .clear
        .Rows = 1
        .ColS = 19 + 4
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColItemCode) = "Item Code"
        .TextMatrix(0, bteColPartNo) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColCurrStock) = "Current Stock"
        .TextMatrix(0, bteColUnitCls) = "Unit Cls"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, BteColSerialFrom) = "Serial From"
        .TextMatrix(0, BteColSerialTo) = "Serial To"
        .TextMatrix(0, bteColDate) = "Delivery Date"
        .TextMatrix(0, bteColTime) = "Time"
        .TextMatrix(0, bteColCurrCode) = "Curr Code"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        'dudi update
        .TextMatrix(0, bteColServices) = "Services"
        '.TextMatrix(0, bteColCur1) = "Curr"
        'end dudi Update
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColRemark) = "Remarks"
        .TextMatrix(0, bteColCalculate) = "Calculate"
        .TextMatrix(0, bteColGenerate) = "Generate"
        .TextMatrix(0, bteColSeqNo) = "SeqNo"
        .TextMatrix(0, bteColStatusEdit) = "Price Status"
        .TextMatrix(0, bteColFinalDestination) = "Final Destination"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColPartNo) = 2000
        .ColWidth(bteColDesc) = 4000
        .ColWidth(bteColQty) = 850
        .ColWidth(bteColCurrStock) = 1300
        .ColWidth(bteColUnit) = 600
        .ColWidth(BteColSerialFrom) = 1100
        .ColWidth(BteColSerialTo) = 1100
        .ColWidth(bteColDate) = 1400
        .ColWidth(bteColTime) = 900
        .ColWidth(bteColCurr) = 600
        .ColWidth(bteColPrice) = 1400
        .ColWidth(bteColServices) = 1400
       ' .ColWidth(bteColCur1) = 700
        .ColWidth(bteColAmount) = 1500
        .ColWidth(bteColRemark) = 2000
        .ColWidth(bteColCalculate) = 1000
        .ColWidth(bteColGenerate) = 1000
        .ColWidth(bteColStatusEdit) = 2000
        .ColWidth(bteColFinalDestination) = 1000
        
        .ColHidden(bteColItemCode) = True
        .ColHidden(bteColUnitCls) = True
        .ColHidden(bteColCurrCode) = True
        '.ColHidden(bteColCur1) = True
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColCalculate) = True
        .ColHidden(bteColGenerate) = True
        .ColHidden(bteColServices) = True
        .ColHidden(bteColStatusEdit) = True
        .ColHidden(bteColFinalDestination) = True
        
        .ColHidden(bteColCurr) = (bteHakPrice = 0)
        .ColHidden(bteColPrice) = (bteHakPrice = 0)
        .ColHidden(bteColAmount) = (bteHakPrice = 0)
        
        
        '.ColHidden(bteColServices) = (bteHakPrice = 0)
        .ColHidden(bteColCur1) = (bteHakPrice = 0)
        
        
        '.Cell(flexcpAlignment, 0, 0, 0, bteColRemark) = flexAlignLeftCenter
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColPartNo) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        
        For i = bteColUnit To bteColCurr
        .ColAlignment(i) = flexAlignLeftCenter
        Next i
        
        .ColAlignment(bteColRemark) = flexAlignLeftCenter
        .ColAlignment(bteColCalculate) = flexAlignCenterCenter
        .ColAlignment(bteColGenerate) = flexAlignCenterCenter
        
        .ColAlignment(bteColPrice) = flexAlignRightCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColServices) = flexAlignRightCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
        
        .EditMaxLength = 1
    End With
End Sub

Sub Kosong()
    txtPoNo.Text = ""
    deliverydate1.Value = Format(Now + 1, "dd MMM yyyy")
    deliverydate2.Value = Format(Now + 1, "dd MMM yyyy")
    txtRevisi.Text = ""
    PODate.Value = Format(Now, "dd MMM yyyy")
    lblcust(0).Text = ""
    lblcust(1).Text = ""
    cboCust.Text = ""
    
    txtcontact.Text = ""
    cboDelPlace.Text = ""
    lblDelPlace.Text = ""
    Check1.Value = False
    ubah = False
    ada = False
    
    Call up_FillCombo(cbounit, "unit_cls")
    cbounit.TextColumn = 2
    LblErrMsg = ""
    DelDate.Value = Format(Now + 1, "dd MMM yyyy")
    deltime.Value = Format("00:00", "hh:mm")
    statusfix = 0
        
    PODate.DataChanged = False
    txtcontact.DataChanged = False
         
    kunci (False)
    kosongBwh
    Header
End Sub

Sub kosongBwh()
    CboItemCode.Text = ""
    lbldesc.Text = ""
    txtQty.Text = ""
    
    TxtSerialFrom = ""
    TxtSerialTo = ""
    
    
    cbounit.ListIndex = -1
    cbocurr.ListIndex = -1
    cboprice.clear
    cboprice.ListIndex = -1
    cboDestination.ListIndex = 1
    
    cboServices.clear
    cboServices.ListIndex = -1
    txtamount.Text = 0
    txtremarks.Text = ""
    Txtdisplay(0) = txtPoNo.Text
    Txtdisplay(1) = Format(0, gs_formatAmountIDR)
    CboItemCode.Enabled = True
    
    
End Sub

Sub adtocboCust()
Dim sqlcust As String
Dim RsCust As New Recordset
Dim i As Integer
    
    ' Customer adalah yang mempunyai Sales Price di Price Master
    ' Update 20090203
    
    sqlcust = "select trade_code, trade_name, address1 from trade_master " & _
                vbLf & " Where Trade_Cls IN ('2','3') "
                
'                vbLf & " where trade_code in " & _
'                vbLf & " (Select distinct Trade_Code From Price_Master Where Price_Cls='02')"
                
    Set RsCust = Db.Execute(sqlcust)
    
    With cboCust
        .clear
        .columnCount = 3
        .ColumnWidths = "50pt;300pt;0pt"
        .ListWidth = 350
        .ListRows = 15
        
        i = 0
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("Trade_code"))
            .List(i, 1) = IIf(IsNull(RsCust("trade_name")), " ", Trim(RsCust("Trade_Name")))
            .List(i, 2) = IIf(IsNull(RsCust("address1")), " ", Trim(RsCust("Address1")))
            RsCust.MoveNext
            i = i + 1
        Loop
    End With
End Sub

Sub adtocbodelplace()
    
    Dim sqlplace As String, sqlcust As String
    Dim RsCust As New Recordset
    Dim i As Integer

    ' Consignee adalah Trade yang mempunyai Sales Price di Price Master & memiliki Trade_Cls=4
    ' Update 20090203
    
'    sqlcust = "select trade_code, trade_name, address1 from trade_master " & _
'                vbLf & " where trade_code in " & _
'                vbLf & " (Select distinct Trade_Code From Price_Master Where Price_Cls='02')" & _
'                vbLf & " Union All " & _
'                vbLf & " Select Trade_Code,Trade_Name,Address1 From Trade_Master " & _
'                vbLf & " Where Trade_Cls='4'"
    
    sqlcust = "select trade_code, trade_name, address1 from trade_master " & _
                vbLf & " Where Trade_Cls IN ('2','4','3') "
    
    Set RsCust = Db.Execute(sqlcust)
    
    With cboDelPlace
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;300pt"
        .ListWidth = 350
        .ListRows = 15
        
        i = 0
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("Trade_code"))
            .List(i, 1) = IIf(IsNull(RsCust("trade_name")), " ", Trim(RsCust("Trade_Name")))
            RsCust.MoveNext
            i = i + 1
        Loop
    End With
    
End Sub

Sub adtocboitem(Optional Row As Long)
Dim sqlitem As String
Dim RsItem As New Recordset
Dim i As Double

' Item yang ditampilkan di Combo Box adalah item dijual kepada Customer yang terpilih di Combo Box
' berdasarkan tanggal Aktif dan Price Master
' Update 20090204

    If grid.TextMatrix(Row, bteColSelect) = "S" Then
    
'        sqlitem = " select * From ( " & _
'            vbLf & " select item_code, item_name, unit_cls, makeritem_code from item_master " & _
'            vbLf & " where use_endday >= convert(char(8), getdate(), 112) and item_Code in " & _
'            vbLf & " (Select Item_Code From Price_Master Where price_cls='02' " & _
'            vbLf & " And Trade_Code in ('00000','" & Trim(cboCust.Text) & "'))) A" & _
'            vbLf & "union " & _
'            vbLf & "select item_code, item_name, unit_cls, makeritem_code from item_master " & _
'            vbLf & "where makeritem_code = '" & grid.TextMatrix(Row, bteColPartNo) & "' " & _
'            vbLf & ")Z order by makeritem_code "

        sqlitem = " EXEC dbo.sp_OrderEntry_LoadItem @CustomerCode = '" & grid.TextMatrix(Row, bteColPartNo) & "', " & _
            vbLf & " @ConsigneeCode = '" & Trim(cboDelPlace.Text) & "', " & _
            vbLf & " @ItemCode = '" & grid.TextMatrix(Row, bteColPartNo) & "', " & _
            vbLf & " @Type = 'S' "
    Else
'        sqlitem = " select * From ( " & _
'            vbLf & " select item_code, item_name, unit_cls, makeritem_code from item_master " & _
'            vbLf & " where use_endday >= convert(char(8), getdate(), 112) and item_Code in " & _
'            vbLf & " (Select Item_Code From dbo.Part_Receipt Where " & _
'            vbLf & "  Supplier_Code in ('" & Trim(cbodelplace.Text) & "'))) A"

        sqlitem = " EXEC dbo.sp_OrderEntry_LoadItem @CustomerCode = '" & grid.TextMatrix(Row, bteColPartNo) & "', " & _
            vbLf & " @ConsigneeCode = '" & Trim(cboDelPlace.Text) & "', " & _
            vbLf & " @ItemCode = '" & grid.TextMatrix(Row, bteColPartNo) & "', " & _
            vbLf & " @Type = '' "
    End If

    Set RsItem = Db.Execute(sqlitem)
    
    With CboItemCode
        .clear
        .columnCount = 4
        .ColumnWidths = "120pt;120pt;250pt;0pt"
        .ListWidth = 490
        .ListRows = 15
        
        i = 0
        Do While Not RsItem.EOF
            .AddItem
            .List(i, 0) = Trim(RsItem("makeritem_code"))
            .List(i, 1) = Trim(RsItem("item_code"))
            .List(i, 2) = Trim(RsItem("item_Name"))
            .List(i, 3) = Trim(RsItem("unit_cls"))
            
            RsItem.MoveNext
            i = i + 1
        Loop
    End With
End Sub

Sub adtocbopono(p As String)
    
    Dim i As Integer
    Dim sqlno As String
    Dim rsno As New Recordset

    sqlno = "select distinct a.po_no, a.rev_no " & _
        "from orderentry_master a " & _
        "inner join orderentry_detail b on a.cust_code = b.cust_code and a.po_no = b.po_no " & _
        "where b.delivery_date>='" & Format(deliverydate1.Value, "YYYY-MM-DD") & "' " & _
        "and b.delivery_date<='" & Format(deliverydate2.Value, "YYYY-MM-DD") & "' " & p
        
    sqlno = sqlno & _
        " union select a.po_no, a.rev_no from orderentry_master a where po_no not in (select po_no from orderentry_detail) " & p
        
    Set rsno = Db.Execute(sqlno)
    
    With CboPOnO
        .clear
        .columnCount = 2
        .ColumnWidths = "115pt;0pt"
        .ListWidth = 115
        .ListRows = 15
        
        i = 0
        Do While Not rsno.EOF
            .AddItem ""
            .List(i, 0) = Trim(rsno("po_no") & "")
            .List(i, 1) = Trim(rsno("rev_no") & "")
            rsno.MoveNext
            i = i + 1
        Loop
    End With
    
End Sub

Sub kunci(l As Boolean)
    PODate.Enabled = Not l
    cboDelPlace.locked = l
    txtcontact.locked = l
    grid.Editable = Not l
    Command1(0).Enabled = Not l
    lblFix.Visible = l
End Sub

Function seqNo(ByVal cust$, ByVal PONO$) As Integer
Dim sqlseqno As String
Dim rsseqno As New Recordset

sqlseqno = "select * from orderentry_detail where cust_Code='" & cust & "' and po_no='" & PONO & "' order by seq_no"
If rsseqno.State <> adStateClosed Then rsseqno.Close
rsseqno.Open sqlseqno, Db, adOpenKeyset, adLockOptimistic

If Not (rsseqno.BOF And rsseqno.EOF) Then
    rsseqno.MoveLast
    seqNo = rsseqno!Seq_no + 1
Else
    seqNo = 1
End If
End Function

Sub browsecust()
Dim sql1 As String
Dim rs1 As New Recordset
    
    sql1 = "select contact_person from trade_master where trade_code='" & cboCust.Text & "' "
    Set rs1 = Db.Execute(sql1)
    
    If Not (rs1.BOF And rs1.EOF) Then
        txtcontact.Text = Trim(IIf(IsNull(rs1(0)), "", rs1(0)))
    End If
    
End Sub
Sub BrowseService()
    Dim sql2 As String
    Dim rs2 As New Recordset
    On Error GoTo handlar
    
    If CboItemCode = "" Or IsNull(CboItemCode) Then Exit Sub
    
    sql2 = "select trade_code, priority_cls, currency_code, price from price_master where " & _
           "item_code='" & CboItemCode.Text & "' and price_cls='05' and (trade_code='" & cboCust.Text & _
           "' or trade_code='000000') and start_date<='" & Format(DelDate.Value, "yyyymmdd") & "' and end_date>='" & _
           Format(DelDate.Value, "yyyymmdd") & "'"
    
    
    If cbocurr <> "" Or Not IsNull(cbocurr) Then
    sql2 = sql2 & " AND Currency_Code='" & cbocurr.Column(0) & "'"
    End If
    
    sql2 = sql2 & " order by trade_code desc, priority_cls desc"
    Set rs2 = Db.Execute(sql2)
    
    
    With cboServices
    
        .clear
        .columnCount = 3
        .ColumnWidths = "70pt;70pt;0pt"
        .ListWidth = 140
        .ListRows = 4
        
        i = 0
        Do While Not rs2.EOF
            .AddItem
            If Trim(rs2("Currency_Code")) = "03" Then
                .List(i, 0) = Format(Trim(rs2("price")), gs_formatPriceIDR)
            Else
                .List(i, 0) = Format(Trim(rs2("price")), gs_formatPrice)
            End If
            If rs2("trade_code") = "000000" Then
              .List(i, 1) = "Common " & Trim(rs2("priority_cls"))
            Else
              .List(i, 1) = "Priority " & Trim(rs2("priority_cls"))
            End If
            .List(i, 2) = Trim(rs2("Currency_Code"))
            
            rs2.MoveNext
            i = i + 1
        
        Loop
    End With
    If cboServices.ListCount > 0 Then
        cboServices.ListIndex = 0
        cboServices_Click
        cboServices.locked = True 'jika ada harga maka di lock
    ElseIf cboServices.Text = "" Then
        cboServices.ListIndex = -1
        cboServices.locked = False 'jika tidak ada
    End If
Exit Sub

handlar:
LblErrMsg.Caption = err.Description
err.clear


End Sub
Sub browseprice()
    Dim sql2 As String
    Dim rs2 As New Recordset
    
    On Error GoTo Err_Curr
    ' Lookup Price Base On PO Date
    ' Update 20090205
    
    'Lookup ke part receipt jika tidak ada baru ambil ke price master
'    sql2 = "select trade_code, priority_cls, currency_code, price from price_master where " & _
'           "item_code='" & CboItemCode.Column(1) & "' and price_cls='02' and (trade_code='" & cboCust.Text & _
'           "' or trade_code='000000') and start_date<='" & Format(PODate.Value, "yyyymmdd") & "' and end_date>='" & _
'           Format(PODate.Value, "yyyymmdd") & "'"
           
'    If cbocurr <> "" Or Not IsNull(cbocurr) Then
'        sql2 = sql2 & " AND Currency_Code='" & cbocurr.Column(0) & "'"
'    End If
'
'    sql2 = sql2 & "  order by trade_code desc, priority_cls desc"

 If cbocurr <> "" Or Not IsNull(cbocurr) Then
    sql2 = "EXEC dbo.sp_OrderEntry_BrowsePrice @ItemCode = '" & CboItemCode.Column(1) & "'," & _
           " @SupplierCode = '" & cboCust.Text & "'," & _
           " @ConsigneeCode = '" & cboDelPlace.Text & "'," & _
           " @StartDate = '" & Format(PODate.Value, "yyyymmdd") & "'," & _
           " @EndDate = '" & Format(PODate.Value, "yyyymmdd") & "'," & _
           " @CurrencyCode = '" & cbocurr.Column(0) & "' "
 Else
    sql2 = "EXEC dbo.sp_OrderEntry_BrowsePrice @ItemCode = '" & CboItemCode.Column(1) & "'," & _
           " @SupplierCode = '" & cboCust.Text & "'," & _
           " @ConsigneeCode = '" & cboDelPlace.Text & "'," & _
           " @StartDate = '" & Format(PODate.Value, "yyyymmdd") & "'," & _
           " @EndDate = '" & Format(PODate.Value, "yyyymmdd") & "'," & _
           " @CurrencyCode = '' "
 End If
     
    Set rs2 = Db.Execute(sql2)
        
    
    With cboprice
        .clear
        .columnCount = 3
        .ColumnWidths = "70pt;70pt;0pt"
        .ListWidth = 140
        .ListRows = 4
        
        i = 0
        Do While Not rs2.EOF
            .AddItem
            If Trim(rs2("Currency_Code")) = "03" Then
                .List(i, 0) = Format(Trim(rs2("price")), gs_formatPriceIDR)
            Else
                .List(i, 0) = Format(Trim(rs2("price")), gs_formatPrice)
            End If
            If rs2("trade_code") = "000000" Then
              .List(i, 1) = "Common " & Trim(rs2("priority_cls"))
            Else
              .List(i, 1) = "Priority " & Trim(rs2("priority_cls"))
            End If
            .List(i, 2) = Trim(rs2("Currency_Code"))
            
            rs2.MoveNext
            i = i + 1
        
        Loop
    End With
    If cboprice.ListCount > 0 Then
        cboprice.ListIndex = 0
        cboprice_Click
        cboprice.locked = False
    ElseIf cboprice.Text = "" Then
        cbocurr.ListIndex = -1
        cboprice.locked = True
    End If

Exit Sub
Err_Curr:

End Sub

Sub Browse()
Dim RS As New ADODB.Recordset
Dim t, p As String
Dim sqltgl As String
Dim rstgl As New Recordset

    LblErrMsg = ""
    
    sql = "select * from orderentry_master where po_no='" & txtPoNo.Text & "' "
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
        
    If Not (RS.BOF And RS.EOF) Then
        ada = True
        ubah = True
        
        sqltgl = "select max(delivery_date) as max, min(delivery_date) as min from orderentry_detail where po_no='" & _
                  txtPoNo.Text & "' "
        Set rstgl = Db.Execute(sqltgl)
        If IsNull(rstgl("max")) = False Then
            If CDate(rstgl("max")) < CDate(deliverydate2.Value) Then
            Else
                deliverydate2.Value = Format(rstgl("max"), "dd MMM yyyy")
            End If
        End If
        
        If IsNull(rstgl("min")) = False Then
            If CDate(rstgl("min")) > CDate(deliverydate1.Value) Then
            Else
                deliverydate1.Value = Format(rstgl("min"), "dd MMM yyyy")
            End If
        End If
        
        txtRevisi.Text = Trim(RS("rev_no") & "")
        PODate.Value = IIf(IsNull(RS("po_date")), " ", Format(Trim(RS("po_date")), "dd MMM yyyy"))
        p = IIf(IsNull(RS("location_code")), " ", Trim(RS("location_code")))
        t = IIf(IsNull(RS("contact_person")), " ", Trim(RS("contact_person")))
        statusfix = IIf(IsNull(RS("fix_cls")), 0, RS("fix_cls"))
        cboCust.Text = Trim(RS("Cust_code"))
        
        
        If RS("NoCommercial_Cls") = "1" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
        
        txtcontact.Text = t
        cboDelPlace.Text = p
        cbodelplace_Click
        
        PODate.DataChanged = False
        txtcontact.DataChanged = False

        BrowseGrid
        
        If statusfix = 1 Then
            kunci (True)
        Else
            kunci (False)
        End If
        
    Else
        ada = False
    End If
    
End Sub

Sub BrowseGrid()
    
    Dim rsGrid As New ADODB.Recordset
    Dim dblTotAmount As Double
    
    Header
    kosongBwh
    
    sqlGrid = "select a.*, b.item_name, (SELECT TM_Current FROM Stock_Master WHERE Item_Code =A.Item_Code " & _
              " AND Warehouse_Code='WH-002' )Current_Stock  from orderentry_detail a " & _
              " INNER JOIN item_master b on a.item_code = b.item_code " & _
        " WHERE a.po_no='" & txtPoNo.Text & "' ORDER BY a.delivery_date, a.makeritem_code, a.seq_no"
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
                        
    i = 1
    dblTotAmount = 0
    With grid
    Do While Not rsGrid.EOF
        .Rows = .Rows + 1
        
        .TextMatrix(i, bteColItemCode) = Trim(rsGrid("Item_Code") & "")
        .TextMatrix(i, bteColPartNo) = Trim(rsGrid("MakerItem_code") & "")
        .TextMatrix(i, bteColDesc) = Trim(rsGrid("Item_Code") & "") & " " & Trim(rsGrid("Item_name") & "")
        .TextMatrix(i, 4) = IIf(IsNull(rsGrid("Qty")), 0, Format(Trim(rsGrid("Qty")), gs_formatQty))
        .TextMatrix(i, 5) = IIf(IsNull(rsGrid("Current_Stock")), 0, Format(Trim(rsGrid("Current_Stock")), gs_formatQty))
        If IsNull(rsGrid("unit_cls")) Then
          .TextMatrix(i, bteColUnitCls) = ""
          .TextMatrix(i, bteColUnit) = ""
        Else
          .TextMatrix(i, bteColUnitCls) = Trim(rsGrid("Unit_cls"))
          .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(Trim(rsGrid("Unit_Cls")))
        End If
        .TextMatrix(i, bteColDate) = Format(Trim(rsGrid("delivery_date")), "dd MMM yyyy")
        .TextMatrix(i, bteColTime) = IIf(IsNull(rsGrid("delivery_time")), "", Trim(rsGrid("delivery_time")))
        If IsNull(rsGrid("currency_code")) Then
           .TextMatrix(i, bteColCurrCode) = ""
           .TextMatrix(i, bteColCurr) = ""
           .TextMatrix(i, bteColCur1) = ""
        Else
          .TextMatrix(i, bteColCurrCode) = Trim(rsGrid("currency_code"))
          .TextMatrix(i, bteColCurr) = uf_GetCurrencyDescription(Trim(rsGrid("Currency_code")))
      '    .TextMatrix(i, bteColCur1) = uf_GetCurrencyDescription(Trim(rsGrid("Currency_code")))
        End If
        If Trim(rsGrid("currency_code")) = "03" Then
            .TextMatrix(i, bteColPrice) = IIf(IsNull(rsGrid("price")), 0, Format(Trim(rsGrid("price")), gs_formatPriceIDR))
            'tambahan dudi
            .TextMatrix(i, bteColServices) = IIf(IsNull(rsGrid("service")), 0, Format(Trim(rsGrid("service")), gs_formatPriceIDR))
            'end tambahan dudi
        Else
            .TextMatrix(i, bteColPrice) = IIf(IsNull(rsGrid("price")), 0, Format(Trim(rsGrid("price")), gs_formatPrice))
            'Tambahan dudi
            .TextMatrix(i, bteColServices) = IIf(IsNull(rsGrid("Service")), 0, Format(Trim(rsGrid("Service")), gs_formatPrice))
            'End tambahan dudi
        End If
        ' ---
        .TextMatrix(i, BteColSerialFrom) = IIf(IsNull(rsGrid("SerialNoFrom")), "", Trim(rsGrid("SerialNoFrom")))
        .TextMatrix(i, BteColSerialTo) = IIf(IsNull(rsGrid("SerialNoTo")), "", Trim(rsGrid("SerialNoTo")))
        ' ---
        .TextMatrix(i, bteColAmount) = IIf(IsNull(rsGrid("amount")), 0, Format(Trim(rsGrid("amount")), gs_formatAmount)) 'gs_formatAmountIDR
        .TextMatrix(i, bteColRemark) = IIf(IsNull(rsGrid("remarks")), "", Trim(rsGrid("remarks")))
        .TextMatrix(i, bteColFinalDestination) = IIf(IsNull(rsGrid("PlaceOfDestination_Cls")), "2", Trim(rsGrid("PlaceOfDestination_Cls")))
        .TextMatrix(i, bteColSeqNo) = Trim(rsGrid("seq_no"))
        
        .TextMatrix(i, bteColStatusEdit) = Trim(rsGrid("Edit_Price_cls") & "")
        .TextMatrix(i, bteColFinalDestination) = Trim(rsGrid("PlaceOfDestination_Cls") & "")
        
        
        'If Val(rsGrid("Calculate_Cls") & "") = 0 Then .TextMatrix(i, bteColCalculate) = "No" Else .TextMatrix(i, bteColCalculate) = "Yes"
        'If Val(rsGrid("Generate_Cls") & "") = 0 Then .TextMatrix(i, bteColGenerate) = "No" Else .TextMatrix(i, bteColGenerate) = "Yes"
        
        dblTotAmount = dblTotAmount + Val(rsGrid("Amount") & "")
        
        .Cell(flexcpBackColor, i, 0) = &HFFFFFF
        rsGrid.MoveNext
        i = i + 1
    Loop
    End With
    
    Txtdisplay(1) = Format(dblTotAmount, gs_formatAmount) 'gs_formatAmountIDR
    
End Sub



'added by do 23/08/16
'price bisa diedit jika status edit nya di ceklist
Private Sub cbEditStatus_Click()
    
    If cbEditStatus.Value = True Then
        cboprice.Enabled = True
    ElseIf cbEditStatus.Value = False Then
        cboprice.Enabled = False
    Else
        cboprice.Enabled = False
    End If
    
End Sub

Private Sub cbocurr_Click()
    cboPrice_LostFocus
End Sub

Private Sub cboDestination_Change()
    If cboDestination.ListIndex = 0 Then
        txtDestination.Text = "JAPAN"
    ElseIf cboDestination.ListIndex = 1 Then
        txtDestination.Text = "OVERSEAS"
    End If
End Sub

Private Sub cboServices_Change()
If InStr(1, cboServices.Text, ",") = 1 Then cboServices.Text = Right(cboServices, Len(cboServices) - 1)
    If cboServices.Text <> "" And cboprice.Text <> "" Then
        If txtQty <> "" Then txtamount.Text = (CDec(cboprice.Text) + CDec(Trim(cboServices.Text))) * CDec(txtQty.Text)
        'Format(uf_Trunc((CDbl(cboprice.Text) + CDbl(cboServices)) * CDbl(txtQty.Text), gi_decimalDigitAmountIDR), gs_formatAmountIDR)
    End If
End Sub

Private Sub cboServices_Click()
 If cboServices.ListIndex <> -1 Then
    
        If txtQty <> "" And cboprice.Text <> "" And cboServices.Text <> "" Then txtamount.Text = (CDec(cboprice.Text) + CDec(Trim(cboServices.Text))) * CDec(txtQty.Text)
        'Format(uf_Trunc((CDbl(cboprice.Text) + CDbl(cboServices.Text)) * CDbl(txtQty.Text), gi_decimalDigitAmountIDR), gs_formatAmountIDR)
    End If
End Sub

Private Sub cboServices_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then cboServices_Click
End Sub

Private Sub cboServices_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
          KeyAscii = 0
    End If
    On Error Resume Next
    If CDbl(cboServices.Text & Chr(KeyAscii)) > gd_MaxPrice Then KeyAscii = 0
End Sub

Private Sub cboServices_LostFocus()
Dim z As Double
If cboServices.Text <> "" Then
    z = CDbl(cboServices.Text)
    If z > gd_MaxPrice Then
        cboServices.Text = Left(z, 10)
    End If
End If

If cbocurr.Text = "IDR" Then
    cboServices.Text = Format(cboServices.Text, gs_formatPriceIDR)
Else
    cboServices.Text = Format(cboServices.Text, gs_formatPrice)
End If
End Sub

Private Sub cmd_Browser_Click()

  Me.MousePointer = vbHourglass
  frm_BrowseCon.getItemCode = cboDelPlace.Text
  frm_BrowseCon.Show 1
  cboDelPlace.Text = frm_BrowseCon.getPartNumber
  cboDelPlace.SetFocus
  
  Me.MousePointer = vbDefault
End Sub

Private Sub cmdBrowser_Click()
 If CboItemCode.Enabled = True Then
  Me.MousePointer = vbHourglass
  frm_BrowseItem.getPartNumber = CboItemCode.Text
  frm_BrowseItem.Show 1
  CboItemCode.Text = frm_BrowseItem.getPartNumber
  Me.MousePointer = vbDefault
 End If
End Sub

Private Sub ComboBox2_Change()

End Sub

Private Sub cmdReport_Click()
 Me.MousePointer = vbHourglass
 
 Call loadingformReport
 
 Me.MousePointer = vbDefault
 
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
            
    bteHakPrice = hakPrice(Me.Name)
    cbocurr.Visible = (bteHakPrice = 1)
    cboprice.Visible = (bteHakPrice = 1)

    cboServices = (bteHakPrice = 1)
    
    txtamount.Visible = (bteHakPrice = 1)
    Txtdisplay(1).Visible = (bteHakPrice = 1)
    
    Label7.Visible = (bteHakPrice = 1)
    Label11.Visible = (bteHakPrice = 1)
    Label12.Visible = (bteHakPrice = 1)
    Label5(1).Visible = (bteHakPrice = 1)
    
    adtocboCust
    adtocbodelplace
    
    ' Item Order harus ada pada Price_Master
    adtocboitem
    
    cboDestination.AddItem "1"
    cboDestination.AddItem "2"
    
    cboDestination.Text = "2"
    
    'do 23/08/16
    cboprice.Enabled = False

    combo1.AddItem "Create"
    combo1.AddItem "Update"
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    cbounit.ListRows = 9
    Call up_FillCombo(cbocurr, "curr_cls")
    cbocurr.TextColumn = 2
    Kosong
    combo1.ListIndex = 1
    
End Sub

Private Sub Combo1_Click()

Dim ketemu As Boolean
    
    ketemu = False
    LblErrMsg = ""
    kunci (False)
    kosongBwh
    Header
    
    If combo1.ListIndex = 0 Then
        
        ClearData
        Command1(2).Caption = "&Create"
        ubah = False
        deliverydate1.Enabled = False
        deliverydate2.Enabled = False
        CboPOnO.locked = True
        txtPoNo.Text = "KI3-"
        txtRevisi.Text = ""
        PODate.Value = Format(Now, "dd MMM yyyy")
        PODate.DataChanged = False
        Call up_FillCombo(cbounit, "unit_cls")
        cbounit.TextColumn = 2
        Txtdisplay(0) = ""
        Txtdisplay(1) = ""

    Else
        
        If cboCust.Text = "" Then
            CboPOnO.clear
            txtPoNo.Text = ""
        Else
            sql = " and a.Cust_Code='" & cboCust.Text & "' "
            adtocbopono (sql)
        End If
        
        ubah = True
        Command1(2).Caption = "&Update"
        deliverydate1.Enabled = True
        deliverydate2.Enabled = True
        CboPOnO.locked = False
        
        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next
        
        If ketemu = False Then txtPoNo.Text = "": txtRevisi.Text = "": PODate.Value = Format(Now, "dd MMM yyyy")
        
    End If

End Sub

Private Sub combo1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then Combo1_Click
End Sub

Private Sub cbopono_Click()
    LblErrMsg = ""
    txtPoNo.Text = CboPOnO.Text
    If CboPOnO.MatchFound Then txtRevisi.Text = CboPOnO.List(CboPOnO.ListIndex, 1)
    Header
    kosongBwh
End Sub

Private Sub cbopono_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then cbopono_Click
End Sub


Private Sub txtpono_Change()
Dim ketemu As Boolean

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

Private Sub txtPONo_GotFocus()
'SendKeys "{end}"
End Sub

Private Sub txtpono_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys vbTab
    If combo1.ListIndex = 1 Then
        Header
        kosongBwh
      End If
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboCust_Click()

    Dim ketemu As Boolean

    ketemu = False
    LblErrMsg = ""
    kunci (False)
    ClearData
    If cboCust.ListIndex <> -1 Then
        lblcust(0).Text = cboCust.Column(1)
        lblcust(1).Text = cboCust.Column(2)
        browsecust
        If combo1.ListIndex = 1 Then
            sql = " and a.Cust_Code='" & cboCust.Text & "' "
            adtocbopono (sql)
            
            For i = 0 To CboPOnO.ListCount - 1
                If txtPoNo.Text = CboPOnO.List(i) Then
                    ketemu = True
                    CboPOnO.ListIndex = i
                    Exit For
                End If
            Next
            If ketemu = False Then txtPoNo.Text = "": txtRevisi.Text = "": PODate.Value = Format(Now, "dd MMM yyyy")
            kosongBwh
            Header
        End If
                
        If CboItemCode.Text <> "" Then
            For i = 0 To CboItemCode.ListCount - 1
                If CboItemCode.Text = CboItemCode.List(i) Then
                    If CboItemCode.Column(1) = CboItemCode.List(i, 1) Then
                        CboItemCode.ListIndex = i
                        browseprice
                        Exit For
                    End If
                End If
            Next
        End If

    Else
        lblcust(0).Text = ""
        lblcust(1).Text = ""
        txtcontact = ""
        txtcontact.DataChanged = False
        CboPOnO.clear
        If combo1.ListIndex = 1 Then
            txtPoNo.Text = ""
            txtRevisi.Text = ""
            PODate.Value = Format(Now, "dd MMM yyyy")
            kosongBwh
            Header
        End If
        LblErrMsg.Caption = DisplayMsg(4011)
        cboCust.SetFocus
        Exit Sub
    End If

End Sub

Private Sub cboCust_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then
    For i = 0 To cboCust.ListCount - 1
        If cboCust.Text = cboCust.List(i) Then
            cboCust.ListIndex = i
            Exit For
        End If
    Next
  
    cboCust_Click
  End If
End Sub

Private Sub cbocust_KeyPress(KeyAscii As MSForms.ReturnInteger)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub cbodelplace_Click()
LblErrMsg = ""
    If cboDelPlace.ListIndex <> -1 Then
        lblDelPlace.Text = cboDelPlace.Column(1)
    Else
        If RTrim(cboDelPlace.Text) <> "" Then
            lblDelPlace.Text = ""
            lblDelPlace.DataChanged = False
            LblErrMsg.Caption = DisplayMsg(4014)
            cboDelPlace.SetFocus
            Exit Sub
        Else
            lblDelPlace = ""
        End If
    End If
End Sub

Private Sub cbodelplace_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
    For i = 0 To cboDelPlace.ListCount - 1
        If cboDelPlace.Text = cboDelPlace.List(i) Then
            cboDelPlace.ListIndex = i
            Exit For
        End If
    Next
    cbodelplace_Click
  End If
End Sub

Private Sub cbodelplace_KeyPress(KeyAscii As MSForms.ReturnInteger)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub deliverydate1_Change()
   
   If CDate(deliverydate1) > CDate(deliverydate2) Then
      LblErrMsg.Caption = DisplayMsg("4076") & " " & Format(deliverydate2, "dd MMM yyyy")
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
    
    If cboCust.Text = "" Then
        CboPOnO.clear
        txtPoNo.Text = ""
    Else
        sql = " and a.Cust_Code='" & cboCust.Text & "' "
        adtocbopono (sql)
    End If
    
End Sub

Private Sub deliverydate2_Change()
   
   If CDate(deliverydate2) < CDate(deliverydate1) Then
      LblErrMsg.Caption = DisplayMsg("4077") & " " & Format(deliverydate1, "dd MMM yyyy")
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
   
    If cboCust.Text = "" Then
        CboPOnO.clear
        txtPoNo.Text = ""
    Else
        sql = " and a.Cust_Code='" & cboCust.Text & "' "
        adtocbopono (sql)
    End If
   
End Sub

Private Sub cboitemcode_Click()
    LblErrMsg = ""
    If CboItemCode.ListIndex <> -1 Then
        lbldesc.Text = CboItemCode.Column(1) & " " & CboItemCode.Column(2)
        cbounit.TextColumn = 2
        cbounit.Text = uf_GetUnitDescription(Trim(CboItemCode.Column(3)))
        browseprice
        BrowseService
    Else
        lbldesc.Text = ""
        Call up_FillCombo(cbounit, "unit_cls")
        cbounit.TextColumn = 2
        cbounit.ListIndex = -1
        cbocurr.ListIndex = -1
        cboprice.clear
        cboprice.ListIndex = -1
        LblErrMsg.Caption = DisplayMsg(4003)
        CboItemCode.SetFocus
        Exit Sub
    End If
'    CalculateCls = 0
End Sub

Private Sub cboitemcode_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then
    For i = 0 To CboItemCode.ListCount - 1
        If CboItemCode.Text = CboItemCode.List(i) Then
            If CboItemCode.Column(1) = CboItemCode.List(i, 1) Then
                CboItemCode.ListIndex = i
                Exit For
            End If
        End If
    Next
    cboitemcode_Click
  End If
End Sub

Private Sub deldate_Change()
    If CboItemCode.Text <> "" Then
        For i = 0 To CboItemCode.ListCount - 1
            If CboItemCode.Text = CboItemCode.List(i) Then
                If CboItemCode.Column(1) = CboItemCode.List(i, 1) Then
                    CboItemCode.ListIndex = i
                    browseprice
                    Exit For
                End If
            End If
        Next
    End If
End Sub

Private Sub cboprice_Change()
If InStr(1, cboprice.Text, ",") = 1 Then cboprice.Text = Right(cboprice, Len(cboprice) - 1)
    If cboprice.Text <> "" Then
    If txtQty <> "" Then txtamount.Text = (CDbl(IfNol(cboprice.Text)) + CDbl(IfNol(cboServices))) * CDbl(IfNol(txtQty.Text))
    'Format(uf_Trunc((CDbl(IfNol(cboprice.Text)) + CDbl(IfNol(cboServices))) * CDbl(IfNol(txtQty.Text)), gi_decimalDigitAmountIDR), gs_formatAmountIDR)
    End If
End Sub
Function IfNol(Angka)
IfNol = IIf(Angka = "", 0, Angka)
End Function

Private Sub cboprice_Click()
    If cboprice.ListIndex <> -1 Then
        cbocurr.Text = uf_GetCurrencyDescription(Trim(cboprice.Column(2)))
        On Error Resume Next
        If txtQty <> "" Then txtamount.Text = (CDec(cboprice.Text) + CDec(Trim(cboServices.Text))) * CDec(txtQty.Text)
        'Format(uf_Trunc((CDbl(cboprice.Text) + IfNol(CDbl(cboServices))) * CDbl(txtQty.Text), gi_decimalDigitAmountIDR), gs_formatAmountIDR)
    End If
End Sub
Private Sub cboprice_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then
  cboprice_Click
 End If
End Sub

Private Sub CboPrice_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
          KeyAscii = 0
    End If
    
    On Error Resume Next
    'If CDbl(cboPrice.Text & Chr(KeyAscii)) > gd_MaxPrice Then KeyAscii = 0
End Sub

Private Sub cboPrice_LostFocus()
Dim z As Double
If cboprice.Text <> "" Then
    z = CDbl(cboprice.Text)
    If z > gd_MaxPrice Then
        cboprice.Text = Left(z, 10)
    End If
End If
BrowseService
'browseprice
If cbocurr.Text = "IDR" Then
    cboprice.Text = Format(cboprice.Text, gs_formatPriceIDR)
Else
    cboprice.Text = Format(cboprice.Text, gs_formatPrice)
End If
End Sub

Private Sub txtqty_Change()
On Error GoTo bugs

If InStr(1, txtQty.Text, ",") = 1 Then txtQty.Text = Right(txtQty, Len(txtQty) - 1)
  If txtQty <> "" Then
    If cboprice.Text <> "" Then
        If cboServices.Text = "" Then cboServices.Text = "0"
        txtamount.Text = (CDec(cboprice.Text) + CDec(Trim(cboServices.Text))) * CDec(txtQty.Text)
'        Format(uf_Trunc((Format(CDbl(cboprice.Text), gs_formatAmountIDR) + CDbl(Trim(cboServices.Text))) * CDbl(txtQty.Text), gi_decimalDigitAmountIDR), gs_formatAmountIDR) 20241216
    End If
    ' Get Serial To Automatic
    If TxtSerialFrom <> "" Then TxtSerialTo = GetSerialTo(Trim(TxtSerialFrom), txtQty)
        
  End If
  
bugs:
  LblErrMsg.Caption = err.Description
  
  
End Sub

Private Sub txtqty_GotFocus()
'SendKeys "{home}"
'SendKeys "+{End}"
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
          KeyAscii = 0
    End If
    'If KeyAscii = Asc(".") Then KeyAscii = 0

    If Val(txtQty) > gd_MaxQty And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtQty_LostFocus()
    txtQty.Text = Format(txtQty.Text, gs_formatQty)
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim TextGrid As String
Dim k As Boolean
LblErrMsg.Caption = ""
k = False

On Error GoTo handler


With grid
    TextGrid = grid.Text

    If TextGrid = "S" Then
      'Call adtocboitem(Row)
        CboItemCode.ListIndex = -1
        
        For i = 0 To CboItemCode.ListCount - 1
            If .TextMatrix(Row, bteColItemCode) = CboItemCode.List(i, 1) Then
                CboItemCode.ListIndex = i
                Exit For
            End If
        Next
        CboItemCode.Enabled = False
        lbldesc.Text = .TextMatrix(Row, bteColDesc)
        txtQty.Text = Format(.TextMatrix(Row, bteColQty), gs_formatQty)
        lblQty = Format(.TextMatrix(Row, bteColQty), gs_formatQty)
        cbounit.ListIndex = -1
        For i = 0 To cbounit.ListCount - 1
            If .TextMatrix(Row, bteColUnitCls) = cbounit.List(i, 0) Then
                cbounit.ListIndex = i
                Exit For
            End If
        Next
        '---
        TxtSerialFrom = .TextMatrix(Row, BteColSerialFrom)
        TxtSerialTo = .TextMatrix(Row, BteColSerialTo)
        DelDate.Value = .TextMatrix(Row, bteColDate)
        CboItemCode.Text = .TextMatrix(Row, bteColItemCode)
        
        Tgl = .TextMatrix(Row, bteColDate)
        
        If .TextMatrix(Row, bteColTime) = "" Then
            deltime.Value = Format("00:00", "hh:mm")
        Else
            deltime.Value = .TextMatrix(Row, bteColTime)
        End If
        
        browseprice
        BrowseService
        
        cbocurr.ListIndex = -1
        For i = 0 To 4
            If .TextMatrix(Row, bteColCurrCode) = cbocurr.List(i, 0) Then
                cbocurr.ListIndex = i
                Exit For
            End If
        Next
        
        
        If Trim(cbocurr.Text) = "IDR" Then
            cboprice.Text = Format(.TextMatrix(Row, bteColPrice), gs_formatPriceIDR)
            'Tambahan Dudi
            cboServices.Text = Format(.TextMatrix(Row, bteColServices), gs_formatPriceIDR)
            
        Else
            cboprice.Text = Format(.TextMatrix(Row, bteColPrice), gs_formatPrice)
            'Tambahan Dudi
            cboServices.Text = Format(.TextMatrix(Row, bteColServices), gs_formatPriceIDR)
        End If

        txtamount.Text = Format(.TextMatrix(Row, bteColAmount), gs_formatAmount) 'gs_formatAmountIDR
        txtremarks.Text = .TextMatrix(Row, bteColRemark)
        'cboDestination.Text = .TextMatrix(Row, bteColFinalDestination)
        TxtSeqNo.Text = .TextMatrix(Row, bteColSeqNo)
        
        'added by do 23/08/16
        'untuk mengetahui price sudah pernah diedit atau belum
        If .TextMatrix(Row, bteColStatusEdit) = "" Then
            cbEditStatus.Value = False
        ElseIf .TextMatrix(Row, bteColStatusEdit) = 0 Then
            cbEditStatus.Value = False
        ElseIf .TextMatrix(Row, bteColStatusEdit) = 1 Then
            cbEditStatus.Value = True
        End If
        
       Call kosongColGrid
    ElseIf TextGrid = "D" Then
       Call kosongColGrid("S")
    End If

    .TextMatrix(Row, Col) = TextGrid
    'If .TextMatrix(Row, bteColCalculate) = "Yes" Then CalculateCls = 1 Else CalculateCls = 0
    
End With

handler:
LblErrMsg.Caption = err.Description

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
On Error GoTo handler


  If grid.Col <> bteColSelect Then Cancel = True
handler:
LblErrMsg.Caption = err.Description
  
End Sub

Private Sub grid_Click()
  With grid
    If .Row = 1 And .Col <> bteColSelect Then
      If .Col = bteColDesc Or .Col = bteColCurr Or .Col = bteColPrice Then
        If .ColSort(.Col) = flexSortNumericAscending Then
          .ColSort(.Col) = flexSortNumericDescending
        Else
          .ColSort(.Col) = flexSortNumericAscending
        End If
      Else
        If .ColSort(.Col) = flexSortStringAscending Then
          .ColSort(.Col) = flexSortStringDescending
        Else
          .ColSort(.Col) = flexSortStringAscending
        End If
      End If
      .Sort = .ColSort(.Col)
    End If
  End With
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
'On Error GoTo ErrMsg
    Dim RS As New ADODB.Recordset
    Dim Sqlc As String, sqllot As String, tmpLotNo As String, sql1 As String
    Dim rsc As New Recordset, rs1 As New Recordset
    Dim rslot As New Recordset
    Dim rsUpdate As New Recordset
    Dim tempitem As String
    Dim dbw As New Connection
    Dim tanya
    Dim hapus As Boolean
    Dim changeQty As Double, j As Integer
    Dim CekSr As String
    Dim X As Long
    Dim awal As Long, akhir As Long, Panjang As Integer
    Dim TempSerial As String, Depan As String
    Dim rsCek As New ADODB.Recordset
    
    dbw.ConnectionString = Db.ConnectionString
    LblErrMsg = ""
    ubahgrid = False
    
    Select Case Index
    Case 0:
        Me.MousePointer = vbHourglass
        If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
        If txtPoNo.Text = "" Then
            txtPoNo.SetFocus
            LblErrMsg = DisplayMsg(1046) 'Please Input PO No
            Me.MousePointer = vbDefault
            Exit Sub
        ElseIf cboCust.Text = "" Then
            cboCust.SetFocus
            LblErrMsg = DisplayMsg(1045) 'Please Select Customer Code
            Me.MousePointer = vbDefault
            Exit Sub
        ElseIf cboDelPlace.Text = "" Then
            cboCust.SetFocus
            LblErrMsg = "Please Select Consignee Code !" 'Please Select Customer Code
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        
        If TxtSerialFrom.Text <> "" Then
        
            CekSr = ""
            CekSr = uf_ValidasiSerialNo
                                        
            If CekSr <> "" Then
                
               LblErrMsg.Caption = CekSr
               Me.MousePointer = vbDefault
               Exit Sub
            End If
            
        End If
        
        If grid.Rows = 2 And grid.TextMatrix(grid.Row, bteColSelect) = "S" Then
            'Lanjut
        Else
            For i = 1 To grid.Rows - 1
                If Trim(grid.TextMatrix(i, bteColCurrCode)) <> Trim(cbocurr) Then
                    LblErrMsg = DisplayMsg(4084)
                    cbocurr.SetFocus
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
            Next
            
        End If
        
        sql = "select * from orderentry_master where po_no='" & txtPoNo.Text & "' "
        If RS.State <> adStateClosed Then RS.Close
        RS.Open sql, Db, adOpenKeyset, adLockOptimistic
        
        If RS.BOF And RS.EOF Then
            LblErrMsg.Caption = DisplayMsg(4015)
            txtPoNo.SetFocus
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        If ubah = True Then
        
            If cboDelPlace.Text <> "" Then
                cboDelPlace.MatchEntry = 1
                cboDelPlace.Text = cboDelPlace.Text
                If cboDelPlace.MatchFound = False Then
                    LblErrMsg = DisplayMsg("0039")
                    cboDelPlace.SetFocus
                    cboDelPlace.MatchEntry = 2
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
                cboDelPlace.MatchEntry = 2
            End If
            cboDelPlace.MatchEntry = 1
            
'            If CalculateCls = 1 Then
'                LblErrMsg = DisplayMsg("8110")
'                Me.MousePointer = vbDefault
'                Exit Sub
'            End If
            
            RS("rev_no") = Trim(txtRevisi.Text)
            RS("po_date") = Format(PODate.Value, "YYYY-MM-DD")
            If Check1.Value = 1 Then
                RS("NoCommercial_Cls") = "1"
            Else
                RS("NoCommercial_Cls") = "0"
            End If
            RS("location_code") = cboDelPlace.Text
            RS("contact_person") = txtcontact.Text
            RS("last_update") = Now
            RS("last_user") = userLogin
            If Trim(txtRevisi.Text) <> "" Then
                RS("po_no") = Trim(Replace(RS("po_no"), "(RE)", "")) & " (RE)"
            Else
                RS("po_no") = Trim(Replace(RS("po_no"), "(RE)", ""))
            End If
            RS.update
            
            If CboPOnO.ListIndex >= 0 Then
                If CboPOnO.List(CboPOnO.ListIndex, 1) <> Trim(txtRevisi.Text) Then
                    CboPOnO.List(CboPOnO.ListIndex, 1) = Trim(txtRevisi.Text)
                    CboPOnO.List(CboPOnO.ListIndex, 0) = Trim(RS("po_no"))
                    BrowseGrid
                    LblErrMsg.Caption = DisplayMsg(1000)
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            
        End If
        
        With grid
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, bteColSelect) = "D" Then
                    If IsEmpty(tanya) Then tanya = MsgBox("Do You Really Want To Delete This Data ?", vbQuestion & vbYesNo, "Confirmation")
                    If tanya = vbYes Then
                        sql1 = "select * from Delivery_Order where PO_No = '" & txtPoNo.Text & "' and " & _
                        "item_code='" & .TextMatrix(i, bteColItemCode) & "' and delivery_date='" & Format(.TextMatrix(i, bteColDate), "YYYY-MM-DD") & "' and " & _
                        "seq_no='" & .TextMatrix(i, bteColSeqNo) & "' "
                        Set rs1 = Db.Execute(sql1)
                        If Not (rs1.BOF And rs1.EOF) Then
                            LblErrMsg.Caption = DisplayMsg(1204)
                            .Row = i
                            .SetFocus
                            Me.MousePointer = vbDefault
                            Exit Sub
'                        ElseIf .TextMatrix(i, bteColCalculate) = "Yes" Then
'                            lblErrMsg.Caption = DisplayMsg(8110)
'                            .Row = i
'                            .SetFocus
'                            Me.MousePointer = vbDefault
'                            Exit Sub
                        Else
                            sql1 = "SELECT Lot_No FROM Production_Planning " & _
                            "WHERE Item_Code = '" & .TextMatrix(i, bteColItemCode) & "' AND month(Production_Date) = " & Month(.TextMatrix(i, bteColDate)) & " AND year(Production_Date) = " & Year(.TextMatrix(i, bteColDate))
                            Set rs1 = Db.Execute(sql1)
                            If Not (rs1.BOF And rs1.EOF) Then tmpLotNo = rs1("Lot_No")
            
                            dbw.Open
                            dbw.BeginTrans
            
                        ' Add Validation for Serial Number on Process before Delete
                        ' Update 20090205
'                        If .TextMatrix(i, BteColSerialFrom) <> "" And .TextMatrix(i, BteColSerialTo) <> "" Then
'                            Panjang = Len(Trim(.TextMatrix(i, BteColSerialFrom)))
'
'                            awal = Val(Mid(.TextMatrix(i, BteColSerialFrom), 2, (Panjang - 1)))
'                            akhir = Val(Mid(.TextMatrix(i, BteColSerialTo), 2, (Panjang - 1)))
'
'                            For x = awal To akhir
'                                TempSerial = Left(.TextMatrix(i, BteColSerialFrom), 1) & Format(x, String(Panjang - 1, "0"))
'                                sql1 = "Select * from Serial_Detail where item_code='" & .TextMatrix(i, bteColItemCode) & "' and " & _
'                                    vbLf & " Serial_No='" & TempSerial & "' And Serial_Status<>'1'"
'                                Set rsCek = dbw.Execute(sql1)
'                                If Not rsCek.EOF Then
'                                    lblErrMsg = "[000]- Serial Number " & Trim(TempSerial) & " has been processed and can't delete !!! "
'                                    rsCek.Close
'                                    Me.MousePointer = vbDefault
'                                    Exit Sub
'                                End If
'                            Next x
'                        End If
                        ' ----------------------------
                            
                            sql1 = "delete from orderentry_detail where cust_Code='" & cboCust.Text & "' and " & _
                            "po_no='" & txtPoNo.Text & "' and item_code='" & .TextMatrix(i, bteColItemCode) & "' and " & _
                            "delivery_date='" & Format(.TextMatrix(i, bteColDate), "YYYY-MM-DD") & "' and seq_no='" & .TextMatrix(i, bteColSeqNo) & "' "
                            dbw.Execute sql1
                            
                        ' Delete Serial Number base on Order Delete
                        ' Update 20090204
'                        If .TextMatrix(i, BteColSerialFrom) <> "" And .TextMatrix(i, BteColSerialTo) <> "" Then
'                            Panjang = Len(Trim(.TextMatrix(i, BteColSerialFrom)))
'
'                            awal = Val(Mid(.TextMatrix(i, BteColSerialFrom), 2, (Panjang - 1)))
'                            akhir = Val(Mid(.TextMatrix(i, BteColSerialTo), 2, (Panjang - 1)))
'
'                            For x = awal To akhir
'                                TempSerial = Left(.TextMatrix(i, BteColSerialFrom), 1) & Format(x, String(Panjang - 1, "0"))
'                                sql1 = "delete from Serial_Detail where po_no='" & txtpono.Text & "' and item_code='" & .TextMatrix(i, bteColItemCode) & "' and " & _
'                                " Serial_No='" & TempSerial & "' And Serial_Status='1'"
'                                dbw.Execute sql1
'                            Next x
'                        End If
                        ' --------------------------------------
                            
                            dbw.CommitTrans
                            dbw.Close
            
                            hapus = True
                        End If
                    Else
                        Exit For
                    End If
                ElseIf .TextMatrix(i, bteColSelect) = "S" Then
                    ubahgrid = True
                End If
            Next i
            
            If (hapus) Then BrowseGrid: LblErrMsg = DisplayMsg(1201): Me.MousePointer = vbDefault: Exit Sub
        
        End With
        
        If CboItemCode.Text = "" Then
            'CboItemCode.SetFocus
            LblErrMsg = DisplayMsg(1047) 'Please Select Part Number
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        If txtQty.Text = "" Then
            txtQty.SetFocus
            LblErrMsg = DisplayMsg(1012) 'Please Input Quantity
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        If cbounit.Text = "" Then
            cbounit.SetFocus
            LblErrMsg = DisplayMsg(1030)
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        
        If cbocurr.Text = "" Then
            If bteHakPrice = 0 Then
                cbocurr.ListIndex = uf_GetComboListIndex(cbocurr, gs_DefaultCurrencyCode)
            Else
                cbocurr.SetFocus
                LblErrMsg = DisplayMsg(1028)
                Me.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        'Tambahan Dudi,Nov 2009
         
        If cboprice.Text = "" Then
            If bteHakPrice = 0 Then
                cboprice = 0
            Else
                cboprice.SetFocus
                LblErrMsg = DisplayMsg(1029)
                Me.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        
        'Tambahan dudi, Nov 2008, mengecek apakah services nya terisi atau tidak
        
        If CboItemCode.Text <> "" Then
            CboItemCode.MatchEntry = 1
            CboItemCode.Text = CboItemCode.Text
'            If CboItemCode.MatchFound = False Then
'                lblErrMsg = DisplayMsg(4062) 'Record with This Part Number not Exist !
'                'CboItemCode.SetFocus
'                CboItemCode.MatchEntry = 2
'                Me.MousePointer = vbDefault
'                Exit Sub
'            End If
            CboItemCode.MatchEntry = 2
        End If
        CboItemCode.MatchEntry = 1

        'GET LOT NO
        sqllot = "select lot_no, isnull(qty,0) qty from production_planning where item_code='" & CboItemCode.Text & "' and" & _
            " prod_year='" & Year(DelDate.Value) & "' and prod_month='" & Month(DelDate.Value) & "' "
        Set rslot = Db.Execute(sqllot)
        If Not (rslot.BOF And rslot.EOF) Then
            If IsNull(rslot(0)) Or rslot(0) = "" Then
                Lotforecast = ""
            ElseIf IsNull(rslot(1)) Or rslot(1) = 0 Then
                Lotforecast = ""
            Else
                Lotforecast = Trim(rslot!Lot_no)
            End If
        Else
            Lotforecast = ""
        End If
        '------------------------------------------------------------

        dbw.Open
        dbw.BeginTrans
        Set rsUpdate = Nothing
        
        ' Validation of Serial No
        ' Update 20090204
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
            LblErrMsg = "[000] - Data doesn't match between Qty and Serial No ! "
            Me.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    ' ----------------------------
    
        'INSERT ORDER ENTRY DETAIL
        If ubahgrid = False Then
        
            ' Check For Exist Serial No
            ' When New Record Check Serial No in All Record
'        If Trim(TxtSerialFrom) <> "" And Trim(TxtSerialTo) <> "" Then
'            For x = awal To akhir
'                TempSerial = Depan & Format(x, String(Panjang - 1, "0"))
'                sql1 = "Select * from Serial_Detail where item_code='" & cboItemCode.Text & "' and " & _
'                " Serial_NO='" & TempSerial & "'"
'                Set rsCek = Db.Execute(sql1)
'                If Not rsCek.EOF Then
'                    lblErrMsg = "[000] - Serial Number Already order at Po Number : " & Trim(rsCek("Po_no")) & " Detail No : " & rsCek("PO_SeqNo")
'                    Me.MousePointer = vbDefault
'                    Exit Sub
'                End If
'            Next x
'        End If
            ' -------------------------
            
            sqlGrid = "select * from orderentry_detail where po_no='" & txtPoNo & "' order by delivery_date, makeritem_code, seq_no"
            If rsUpdate.State <> adStateClosed Then rsUpdate.Close
            rsUpdate.Open sqlGrid, dbw, adOpenKeyset, adLockOptimistic
        
            rsUpdate.AddNew
            rsUpdate("cust_Code") = cboCust.Text
            rsUpdate("po_no") = txtPoNo.Text
            rsUpdate("seq_no") = seqNo(cboCust.Text, txtPoNo.Text)
            changeQty = CDbl(txtQty.Text)
            
        'UPDATE ORDER ENTRY DETAIL
        Else
            
            ' Check For Exist Serial No
            ' When Update Record Check Serial No in All Record ( Not Include This Record )
'        If Trim(TxtSerialFrom) <> "" And Trim(TxtSerialTo) <> "" Then
'            For x = awal To akhir
'                TempSerial = Depan & Format(x, String(Panjang - 1, "0"))
'
'                sql1 = " Select * From " & _
'                    vbLf & " (Select * from Serial_Detail where item_code='" & cboItemCode.Text & "' and " & _
'                    vbLf & "  Serial_No not in (Select Serial_No From Serial_Detail Where PO_No='" & txtpono & "' and " & _
'                    vbLf & " Po_SeqNo=" & TxtSeqNo & ")" & _
'                    vbLf & " ) chk Where Serial_No='" & TempSerial & "' "
'
'                Set rsCek = Db.Execute(sql1)
'                If Not rsCek.EOF Then
'                    lblErrMsg = "[000] - Serial Number " & TempSerial & " Already order at Po Number : " & Trim(rsCek("Po_no")) & " Detail No : " & rsCek("PO_SeqNo")
'                    Me.MousePointer = vbDefault
'                    Exit Sub
'                End If
'            Next x
'        End If
        
        ' Check For Exist Serial No
        ' When Serial No on Process
'        If Trim(TxtSerialFrom) <> "" And Trim(TxtSerialTo) <> "" Then
'            For x = awal To akhir
'                TempSerial = Depan & Format(x, String(Panjang - 1, "0"))
'                sql1 = "Select * from Serial_Detail where item_code='" & cboItemCode.Text & "' and " & _
'                " Serial_NO='" & TempSerial & "' and serial_Status<>'1'"
'                Set rsCek = Db.Execute(sql1)
'                If Not rsCek.EOF Then
'                    lblErrMsg = "[000] - Serial Number " & TempSerial & " is on process and can't update !!! "
'                    Me.MousePointer = vbDefault
'                    Exit Sub
'                End If
'            Next x
'        End If
            ' -------------------------
        
        
        
        
        ' -------------------------
            
            Sqlc = "select sum(Qty) from Delivery_Order where PO_No = '" & txtPoNo.Text & "' and " & _
            "item_code='" & CboItemCode.Text & "' and delivery_date='" & Format(DelDate.Value, "YYYY-MM-DD") & "' and " & _
            "seq_no='" & TxtSeqNo.Text & "' "
            Set rsc = Db.Execute(Sqlc)
            If Not (rsc.BOF And rsc.EOF) Then
                If CDbl(txtQty.Text) < rsc(0) Then
                    LblErrMsg.Caption = DisplayMsg("4043") & " " & rsc(0)
                    txtQty.SetFocus
                    Me.MousePointer = vbDefault
                    Exit Sub
                Else
                    sqlGrid = "select * from orderentry_detail where cust_Code='" & cboCust.Text & "' and po_no='" & txtPoNo.Text & "' and " & _
                        "item_code='" & CboItemCode.Text & "' and delivery_date='" & Format(Tgl, "YYYY-MM-DD") & "' and seq_no='" & _
                        TxtSeqNo.Text & "' order by delivery_date, makeritem_code, seq_no"
                    If rsUpdate.State <> adStateClosed Then rsUpdate.Close
                    rsUpdate.Open sqlGrid, dbw, adOpenKeyset, adLockOptimistic
                End If
            End If
        End If

        rsUpdate("Lot_no") = Lotforecast
        rsUpdate("Makeritem_code") = CboItemCode.Text
        rsUpdate("item_Code") = CboItemCode.Text
        rsUpdate("delivery_date") = Format(DelDate.Value, "YYYY-MM-DD")
        rsUpdate("delivery_time") = Format(deltime.Value, "HH:mm")
        rsUpdate("price") = cboprice.Text
        rsUpdate("Service") = cboServices.Text
        rsUpdate("currency_code") = cbocurr.Column(0)
        rsUpdate("unit_cls") = cbounit.Column(0)
        rsUpdate("qty") = txtQty.Text
        rsUpdate("SerialNoFrom") = TxtSerialFrom.Text
        rsUpdate("SerialNoTo") = TxtSerialTo.Text
        rsUpdate("amount") = txtamount.Text
        rsUpdate("remarks") = txtremarks.Text
        rsUpdate("PlaceOfDestination_Cls") = Trim(cboDestination.Text)
        rsUpdate("last_update") = Now
        rsUpdate("last_user") = userLogin
        
        If cbEditStatus.Value = True Then
            rsUpdate("Edit_Price_cls") = 1
        Else
            rsUpdate("Edit_Price_cls") = 0
        End If
        
        rsUpdate.update
    
        ' Delete Serial_No For Update
'        sql1 = "Delete From Serial_Detail Where Item_Code='" & Trim(cboItemCode) & "' And " & _
'                vbLf & " PO_No='" & Trim(cbopono) & "' And PO_SeqNo=" & rsUpdate("Seq_no") & _
'                vbLf & " And Serial_Status='1'"
'                dbw.Execute (sql1)
'
'    If Trim(TxtSerialFrom) <> "" And Trim(TxtSerialTo) <> "" Then
'        ' Add Serial No Data base on Order Data
'       For x = awal To akhir
'            TempSerial = Depan & Format(x, String(Panjang - 1, "0"))
'            sql1 = "Insert Into Serial_Detail (Item_Code,Serial_No,Po_No,PO_SeqNo,Serial_Status) Values ('" & Trim(cboItemCode) & "'," & _
'            vbLf & "'" & Trim(TempSerial) & "','" & Trim(cbopono) & "'," & rsUpdate("Seq_No") & ",'1')"
'            dbw.Execute (sql1)
'        Next x
'    End If
    
       ' --------------------
        
        If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
            LblErrMsg = DisplayMsg(1023): DelDate.SetFocus: Me.MousePointer = vbDefault: Exit Sub
        End If

        dbw.CommitTrans
        dbw.Close
        
        If CDate(DelDate.Value) > CDate(deliverydate1.Value) Then
            If CDate(DelDate.Value) < CDate(deliverydate2.Value) Then
            Else
                deliverydate2.Value = Format(DelDate.Value, "dd MMM yyyy")
            End If
        Else
            deliverydate1.Value = Format(DelDate.Value, "dd MMM yyyy")
        End If
        
        tempitem = CboItemCode.Text
        combo1.ListIndex = 1
        BrowseGrid
    
        For j = 1 To grid.Rows - 1
            If grid.TextMatrix(j, bteColItemCode) = tempitem And grid.TextMatrix(j, bteColDate) = Format(DelDate.Value, "dd MMM yyyy") Then
                grid.Row = j
            End If
        Next j
        ubahgrid = True
        
        LblErrMsg = DisplayMsg(1101)
        
        'Call adtocboitem
    Case 1: Kosong
        Me.MousePointer = vbHourglass
        'Call adtocboitem
        combo1.ListIndex = 1
        Call Combo1_Click
        cboDestination.Text = "2"
        cboCust.SetFocus
    Case 2:
        Me.MousePointer = vbHourglass
        
        ' Mengubah List Item sesuai dengan Customern_Code - 20090204
            Call adtocboitem
        ' ---------

        If combo1.ListIndex = 0 Then
            
            If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
        
            If txtPoNo.Text = "" Then
                txtPoNo.SetFocus
                LblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                Me.MousePointer = vbDefault
                Exit Sub
            ElseIf cboCust.Text = "" Then
                cboCust.SetFocus
                LblErrMsg = DisplayMsg(1045) '"Please Select Customer Code"
                Me.MousePointer = vbDefault
                Exit Sub
            Else
        
                If cboCust.Text <> "" Then
                    cboCust.MatchEntry = 1
                    cboCust.Text = cboCust.Text
                    If cboCust.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4011)
                        cboCust.SetFocus
                        cboCust.MatchEntry = 2
                        Me.MousePointer = vbDefault
                        Exit Sub
                    End If
                    cboCust.MatchEntry = 2
                End If
                cboCust.MatchEntry = 1
        
                If cboDelPlace.Text <> "" Then
                    cboDelPlace.MatchEntry = 1
                    cboDelPlace.Text = cboDelPlace.Text
                    If cboDelPlace.MatchFound = False Then
                        LblErrMsg = DisplayMsg("0039")
                        cboDelPlace.SetFocus
                        cboDelPlace.MatchEntry = 2
                        Me.MousePointer = vbDefault
                        Exit Sub
                    End If
                    cboDelPlace.MatchEntry = 2
                End If
                cboDelPlace.MatchEntry = 1
        
                sql = "select * from orderentry_master where po_no='" & txtPoNo.Text & "' "
                If RS.State <> adStateClosed Then RS.Close
                RS.Open sql, Db, adOpenKeyset, adLockOptimistic
        
                If ubah = False Then
                    If Not (RS.BOF And RS.EOF) Then
                        LblErrMsg.Caption = DisplayMsg(1023)
                        txtPoNo.SetFocus
                        Me.MousePointer = vbDefault
                        Exit Sub
                    Else
                        RS.AddNew
                        RS("po_no") = txtPoNo.Text
                        RS("cust_code") = cboCust.Text
                    End If
                End If
        
                RS("rev_no") = Trim(txtRevisi.Text)
                RS("po_date") = Format(PODate.Value, "YYYY-MM-DD")
                If Check1.Value = 1 Then
                    RS("NoCommercial_Cls") = "1"
                Else
                    RS("NoCommercial_Cls") = "0"
                End If
                
                RS("location_code") = cboDelPlace.Text
                RS("contact_person") = txtcontact.Text
                RS("last_update") = Now
                RS("last_user") = userLogin
                RS.update
        
                PODate.DataChanged = False
                txtcontact.DataChanged = False
                
                combo1.ListIndex = 1
                
                LblErrMsg.Caption = DisplayMsg(1000)
                ubah = True
        
            End If
        
        Else
            
            If txtPoNo.Text = "" Then
                txtPoNo.SetFocus
                LblErrMsg = DisplayMsg(1048) '"Please Select PO No"
                Me.MousePointer = vbDefault
                Exit Sub
            Else
        
                Browse
                If ada = False Then
                    txtRevisi.Text = ""
                    PODate.Value = Format(Now, "dd MMM yyyy")
                    PODate.DataChanged = False
                    LblErrMsg.Caption = DisplayMsg(4015)
                    txtPoNo.SetFocus
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
            
            End If
        
        End If
        
    Case 3:
        Me.MousePointer = vbHourglass
        kosongColGrid
        kosongBwh
        'Call adtocboitem
        
    Case 4:
       Dim xdo As Recordset
       Dim po_no As String
            MousePointer = vbHourglass
            If Trim(CboPOnO) <> "" Then
                CboPOnO = CboPOnO
                If CboPOnO.MatchFound Then
                    sql = "select top 1 po_no from orderentry_detail where po_no ='" & CboPOnO & "'"
                    Set xdo = New Recordset
                    xdo.Open sql, Db, adOpenDynamic, adLockOptimistic
                    If xdo.EOF Then
                        LblErrMsg = DisplayMsg(4071)
                        MousePointer = 1
                        Exit Sub
                    Else
                        po_no = "'" & CboPOnO & "'"
                        Call LoadForm(po_no)
                    End If
                Else
                    LblErrMsg = DisplayMsg(4015)
                End If
            End If
            MousePointer = vbDefault
    Case 5:
    Stuffing
   
    Case 6:
    
        If cboCust.Text = "" Then
        
            LblErrMsg.Caption = "Please Select Customer Code !"
            Exit Sub
            
        ElseIf CboPOnO.Text = "" Then
            
            LblErrMsg.Caption = "Please Select or Create SI/PO No !"
            Exit Sub
            
        ElseIf cboDelPlace.Text = "" Then
            
            LblErrMsg.Caption = "Please Select Consignee Code !"
            Exit Sub
                       
        ElseIf uf_validasi_upload(Trim(cboCust.Text), Trim(CboPOnO.Text)) = False Then
            LblErrMsg.Caption = "Data SI/PO No : " & Trim(CboPOnO.Text) & " Has Been Upload !!"
            Exit Sub
        
        End If
            
        FrmUploadDetailItem.txtTradeCode.Text = Trim(cboCust.Text)
        FrmUploadDetailItem.txtPoNo.Text = Trim(CboPOnO.Text)
        FrmUploadDetailItem.txtConsCode.Text = Trim(cboDelPlace.Text)
        FrmUploadDetailItem.Show 1
        Command1_Click (2)
        
        

    End Select
    
    Me.MousePointer = vbDefault
    
    Exit Sub
ErrMsg:
    Me.MousePointer = vbDefault
    Set RS = Nothing
    Set rsc = Nothing
    Set rslot = Nothing
    Set rsUpdate = Nothing
    LblErrMsg = err.number & " " & err.Description

    
    
End Sub

Sub Stuffing()
    Dim objExcel As New Excel.application
    Dim objWorkSheet As New Worksheet
    Dim objWorkBook As Workbook
    Dim rsStuffing As New ADODB.Recordset
    Dim RS As New ADODB.Recordset
    Dim PathExcel As String
    
    Dim SerialFrom As String
    Dim SerialTo As String
    Dim ls_PONo As String
    Dim Tgl As String
    Dim ls_consignee As String
    Dim ls_ItemCode As String
    Dim ls_desc As String
    
    Dim iRowExl As Integer
    Dim iRowExlStart As Integer
    Dim iSerial As Long
    Dim Idx As Integer
    Dim iLoop As Integer
    Dim iSheet As Integer
    
    Dim cell1 As String, cell2 As String, cell3 As String, cell4 As String
    Dim changeCell As Boolean
    Dim chsngeSheet As Boolean
    
    Dim Cekserial As Boolean
    
    
    If Trim(txtPoNo.Text) <> "" Then
    
        sql = " SELECT OD.Cust_Code,OD.PO_No,OD.Item_Code,IM.Item_Name,OD.Delivery_Date,OD.SerialNoFrom,OD.SerialNoto,TM.Trade_Name, " & vbCrLf & _
              " CASE WHEN LEN(RTRIM(ITEM_NAME))<=15 THEN " & vbCrLf & _
              " SUBSTRING(Right(rtrim(item_name),6), " & vbCrLf & _
              " CHARINDEX(' ',Right(rtrim(item_name),6),0)+1,5) " & vbCrLf & _
              " ELSE " & vbCrLf & _
              " SUBSTRING(Right(rtrim(item_name),20), " & vbCrLf & _
              " CHARINDEX(' ',Right(rtrim(item_name),20),0)+1,19) " & vbCrLf & _
              " END MODEL " & vbCrLf & _
              " FROM dbo.OrderEntry_Detail OD " & vbCrLf & _
              " LEFT JOIN dbo.Item_Master IM ON OD.Item_Code = IM.Item_Code " & vbCrLf & _
              " LEFT JOIN dbo.OrderEntry_Master OM ON OD.PO_No = OM.PO_No " & vbCrLf & _
              " LEFT JOIN dbo.Trade_Master TM ON OM.Location_Code = TM.Trade_Code " & vbCrLf & _
              " WHERE OD.Cust_Code = '" & Trim(cboCust.Text) & "' AND OD.PO_No = '" & Trim(txtPoNo.Text) & "' ORDER BY Item_Code"
        
        If rsStuffing.State <> adStateClosed Then rsStuffing.Close
        rsStuffing.Open sql, Db, adOpenKeyset, adLockOptimistic
        
        If RS.State <> adStateClosed Then RS.Close
        RS.Open sql, Db, adOpenKeyset, adLockOptimistic
        
        Cekserial = False
        
        If Not rsStuffing.EOF Then
            
            Do While Not RS.EOF
                If Trim(RS!SerialNoFrom) <> "" Then
                    Cekserial = True
                End If
                RS.MoveNext
            Loop
            
            If Cekserial = False Then
                LblErrMsg = "This SI/PO No have not serial number"
                Exit Sub
            End If
            
            cdg.filter = "Excel Files (*.xls)|*.xls"
            cdg.filename = "Stuffing Report"
            cdg.ShowSave
            If cdg.FileTitle = "" Then Exit Sub
            If Len(cdg.filename) = 0 Then Exit Sub
            If Dir(cdg.filename) <> "" Then
                If MsgBox("Overwrite existing file?", vbExclamation + vbYesNo, "Overwrite") = vbNo Then Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
            
            PathExcel = App.path & "\Reports\Stuffing Template.xlsx"
            If Dir(cdg.filename) <> "" Then Kill cdg.filename
            
            objExcel.Workbooks.Open PathExcel
            
            objExcel.Worksheets(Array("Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5", "Cplan")).Copy
            
            With objExcel

                .ActiveWorkbook.SaveAs filename:= _
                    cdg.filename, _
                    FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                    ReadOnlyRecommended:=False, CreateBackup:=False
            
                .Workbooks.Close
                .Workbooks.Open cdg.filename
            End With
            
            ls_PONo = Trim(rsStuffing!po_no)
            Tgl = Format(rsStuffing!delivery_Date, "dd-MM-yy")
            ls_consignee = Trim(rsStuffing!trade_name)
            iRowExl = 11
            iRowExlStart = 0
            iSheet = 1
'            objExcel.Visible = True
            
            With objExcel
                
                .Sheets("Cplan").Select
                .Range("C4") = ls_PONo
                .Range("J4") = Tgl
                .Range("M4") = ls_consignee
                
                
                .Sheets("Sheet1").Select
                
                .Range("B5") = "S/I NO  : " & ls_PONo
                .Range("B5", "D5").Merge
                .Range("I5") = "TGL : " & Tgl
                .Range("I5", "J5").Merge
                .Range("K5") = "Dest : " & ls_consignee
                .Range("K5", "R5").Merge

                
                cell1 = "B": cell2 = "C": cell3 = "D": cell4 = "F"
                'iRowExlStart = 1
                Do While Not rsStuffing.EOF
                    Dim rowStatus As Boolean
                    
                    If Trim(rsStuffing!SerialNoFrom) <> "" Then
                                        
                        SerialFrom = Mid(rsStuffing!SerialNoFrom, 2, 6)
                        SerialTo = Mid(rsStuffing!SerialNoTo, 2, 6)
                        iLoop = SerialTo - SerialFrom
                        ls_ItemCode = Trim(rsStuffing!item_name)
                        'ls_desc = Trim(rsStuffing!Model)
                        Idx = 1
                        iRowExl = iRowExl + Idx
                        iSerial = 0
                        iRowExlStart = iRowExlStart + 1

                        For i = 0 To iLoop
                                               
                            rowStatus = False
                            
                            Select Case iRowExlStart
                                Case 49, 95, 142, 189, 236, 283, 330, 377, 424, 471, 518, 565, 612, 659
                                    rowStatus = True
                            End Select
                            
                            If iRowExlStart <> 1 And rowStatus = False Then
                                    iRowExlStart = iRowExlStart + 1
                            End If
                            
                            If iRowExlStart = 1 Then
                                cell1 = "B": cell2 = "C": cell3 = "E" ': cell4 = "F"
                                iRowExl = 12
                            ElseIf iRowExlStart = 49 Then
                                cell1 = "H": cell2 = "I": cell3 = "K" ': cell4 = "M"
                                iRowExl = 12
                                
                                iRowExlStart = iRowExlStart + 1
                            ElseIf iRowExlStart = 95 And rowStatus <> False Then
                                cell1 = "N": cell2 = "O": cell3 = "Q" ': cell4 = "T"
                                iRowExl = 12
                               
                                If rowStatus <> False Then
                                    iRowExlStart = iRowExlStart + 1
                                End If
                            ElseIf iRowExlStart = 142 Then
                                
                                .Range("Q4") = "1 of 2"
                                
                                .Sheets("Sheet2").Select
                                'cell1 = "B": cell2 = "C": cell3 = "D": cell4 = "F"
                                cell1 = "B": cell2 = "C": cell3 = "E"
                                iRowExl = 12
                                
                                .Range("B5") = "S/I NO  : " & ls_PONo
                                .Range("B5", "D5").Merge
                                .Range("I5") = "TGL : " & Tgl
                                .Range("I5", "J5").Merge
                                .Range("K5") = "Dest : " & ls_consignee
                                .Range("K5", "R5").Merge

                                
                                .Range("Q4") = "2 of 2"
                                
                                iRowExlStart = iRowExlStart + 1
                                
                            ElseIf iRowExlStart = 189 Then
                                'cell1 = "I": cell2 = "J": cell3 = "K": cell4 = "M"
                                cell1 = "H": cell2 = "I": cell3 = "K"
                                iRowExl = 12
                                
                                iRowExlStart = iRowExlStart + 1
                                
                            ElseIf iRowExlStart = 236 Then
                                'cell1 = "P": cell2 = "Q": cell3 = "R": cell4 = "T"
                                cell1 = "N": cell2 = "O": cell3 = "Q"
                                iRowExl = 12
                                
                                iRowExlStart = iRowExlStart + 1
                                
                            ElseIf iRowExlStart = 283 Then
                                
                                .Range("Q4") = "2 of 3"
                                .Sheets("Sheet1").Select
                                .Range("Q4") = "1 of 3"
                                                                
                                .Sheets("Sheet3").Select
                                'cell1 = "B": cell2 = "C": cell3 = "D": cell4 = "F"
                                cell1 = "B": cell2 = "C": cell3 = "E"
                                iRowExl = 12
                                
                                .Range("B5") = "S/I NO  : " & ls_PONo
                                .Range("B5", "D5").Merge
                                .Range("I5") = "TGL : " & Tgl
                                .Range("I5", "J5").Merge
                                .Range("K5") = "Dest : " & ls_consignee
                                .Range("K5", "R5").Merge

                                
                                .Range("Q4") = "3 of 3"
                                
                                iRowExlStart = iRowExlStart + 1
                                
                            ElseIf iRowExlStart = 330 Then
                                'cell1 = "I": cell2 = "J": cell3 = "K": cell4 = "M"
                                cell1 = "H": cell2 = "I": cell3 = "K"
                                iRowExl = 12
                                
                                iRowExlStart = iRowExlStart + 1
                                
                            ElseIf iRowExlStart = 377 Then
                                'cell1 = "P": cell2 = "Q": cell3 = "R": cell4 = "T"
                                cell1 = "N": cell2 = "O": cell3 = "Q"
                                
                                iRowExl = 12
                                
                                iRowExlStart = iRowExlStart + 1
                                
                            ElseIf iRowExlStart = 424 Then
                            
                                .Range("Q4") = "3 of 4"
                                .Sheets("Sheet4").Select
                                .Range("Q4") = "1 of 4"
                                .Sheets("Sheet2").Select
                                .Range("Q4") = "2 of 4"
                                
                                .Sheets("Sheet4").Select
                                'cell1 = "B": cell2 = "C": cell3 = "D": cell4 = "F"
                                cell1 = "B": cell2 = "C": cell3 = "E"
                                iRowExl = 12
                                
                                .Range("B5") = "S/I NO  : " & ls_PONo
                                .Range("B5", "D5").Merge
                                .Range("I5") = "TGL : " & Tgl
                                .Range("I5", "J5").Merge
                                .Range("K5") = "Dest : " & ls_consignee
                                .Range("K5", "R5").Merge
                                
                                .Range("Q4") = "4 of 4"
                                
                                iRowExlStart = iRowExlStart + 1
                                
                            ElseIf iRowExlStart = 471 Then
                                'cell1 = "I": cell2 = "J": cell3 = "K": cell4 = "M"
                                cell1 = "H": cell2 = "I": cell3 = "K"
                                iRowExl = 12
                            ElseIf iRowExlStart = 518 Then
                                'cell1 = "P": cell2 = "Q": cell3 = "R": cell4 = "T"
                                 cell1 = "N": cell2 = "O": cell3 = "Q"
                                iRowExl = 12
                            
                            ElseIf iRowExlStart = 565 Then
                            
                                .Range("Q4") = "4 of 5"
                                .Sheets("Sheet5").Select
                                .Range("Q4") = "1 of 5"
                                .Sheets("Sheet5").Select
                                .Range("Q4") = "2 of 5"
                                .Sheets("Sheet5").Select
                                .Range("Q4") = "3 of 5"
                                
                                .Sheets("Sheet5").Select
                                'cell1 = "B": cell2 = "C": cell3 = "D": cell4 = "F"
                                cell1 = "B": cell2 = "C": cell3 = "E"
                                iRowExl = 12
                                
                                .Range("B5") = "S/I NO  : " & ls_PONo
                                .Range("B5", "D5").Merge
                                .Range("I5") = "TGL : " & Tgl
                                .Range("I5", "J5").Merge
                                .Range("K5") = "Dest : " & ls_consignee
                                .Range("K5", "R5").Merge

                                
                                .Range("Q4") = "5 of 5"
                                
                                iRowExlStart = iRowExlStart + 1
                                
                            ElseIf iRowExlStart = 612 Then
                                'cell1 = "I": cell2 = "J": cell3 = "K": cell4 = "M"
                                cell1 = "H": cell2 = "I": cell3 = "K"
                                iRowExl = 12
                            ElseIf iRowExlStart = 659 Then
                                'cell1 = "P": cell2 = "Q": cell3 = "R": cell4 = "T"
                                 cell1 = "N": cell2 = "O": cell3 = "Q"
                                iRowExl = 12
                                
                                iRowExlStart = iRowExlStart + 1
                                
                            End If
                            
                            .Range(cell1 & iRowExl) = Idx
                            .Range(cell2 & iRowExl) = "'" & ls_ItemCode
                            '.Range(cell3 & iRowExl) = ls_desc
                            '.Range(cell3 & iRowExl).HorizontalAlignment = xlHAlignLeft
                            .Range(cell3 & iRowExl) = Left(rsStuffing!SerialNoFrom, 1) & Mid((1000000 + SerialFrom + iSerial), 2, 6)
                            .Range(cell3 & iRowExl).horizontalAlignment = xlHAlignLeft
                            
                            If iRowExl = 57 Then
                                iRowExl = iRowExl
                            End If
                            
                            Idx = Idx + 1
                            iSerial = iSerial + 1
                            iRowExl = iRowExl + 1
                            

                            If iRowExlStart = 1 Then
                                iRowExlStart = iRowExlStart + 1
                            End If
                            
'                            If i <> iLoop Then
'                                iRowExlStart = iRowExlStart + 1
'                            End If
'                            iRowExlStart = iRowExlStart + 1
                            
                        Next i
                    End If
                    rsStuffing.MoveNext
                Loop
                
                
                
            End With
            
               
            
            objExcel.Sheets("Sheet1").Select
            objExcel.Visible = True
            objExcel.ActiveWorkbook.save
            
            Me.MousePointer = vbDefault
            
        Else
            LblErrMsg = DisplayMsg(4015)
        End If
        
    Else
        LblErrMsg = DisplayMsg(1048)
    End If
    
err:
    'objExcel.Workbooks.Close
    Set objWorkBook = Nothing
    Set objWorkSheet = Nothing
    Set objExcel = Nothing
    LblErrMsg = err.Description
    Me.MousePointer = vbDefault

End Sub
Private Sub command2_Click(Index As Integer)
'Dim Atas As Integer
'
'LblErrMsg.Caption = ""
'
'If txtPONo.Text <> "" Then rsGrid.Find "po_no='" & txtPONo.Text & "' "
'
'Select Case Index
'
'    Case 1:
'            jmlpage = rsGrid.PageCount
'            If intpage = 1 Then
'               LblErrMsg.Caption = DisplayMsg(4020) '"This is the first page !"
'            ElseIf jmlpage > 1 Then
'               intpage = 1
'               Call BrowseGrid
'               LblErrMsg.Caption = ""
'            End If
'
'            On Error Resume Next
'            If cboPONo.Text <> "" Then
'               Atas = 1
'               grid.TopRow = 1
'            End If
'
'    Case 2:
'            jmlpage = rsGrid.PageCount
'            If intpage = 1 Then
'               LblErrMsg = DisplayMsg(4020) '"This is the first page !"
'            Else
'               intpage = intpage - 1
'               Call BrowseGrid
'               LblErrMsg = ""
'            End If
'            On Error Resume Next
'            Atas = grid.TopRow
'
'            grid.TopRow = grid.TopRow - 17
'            If Atas = grid.TopRow Then grid.TopRow = 1
'    Case 3:
'            rsGrid.PageSize = 17
'            jmlpage = rsGrid.PageCount
'            If intpage < jmlpage Then
'              intpage = intpage + 1
'              Call BrowseGrid
'              LblErrMsg.Caption = ""
'            Else
'              LblErrMsg.Caption = DisplayMsg(4021) '"This is the last page !"
'            End If
'
'            On Error Resume Next
'            Atas = grid.TopRow
'            grid.TopRow = grid.TopRow + 17
'    Case 4:
'            rsGrid.PageSize = 17
'            jmlpage = rsGrid.PageCount
'            If intpage = jmlpage Then
'              LblErrMsg.Caption = DisplayMsg(4021) '"This is the last page !"
'            ElseIf intpage < jmlpage Then
'              intpage = jmlpage
'              Call BrowseGrid
'              LblErrMsg.Caption = ""
'            End If
'
'            On Error Resume Next
'            grid.TopRow = grid.Rows
'End Select
End Sub

Private Sub command3_Click()
            
    ClearData
    If command3.Caption = "Back" And frmpanggil = "orderinquiry" Then
        Unload Me
        Call frm_order_inquiry.cmdSearch_Click(0)
        frm_order_inquiry.Show
    ElseIf command3.Caption = "Back" And frmpanggil = "orderinquirydate" Then
        Unload Me
        Call frm_order_inquiry_date.cmdSearch_Click(0)
        frm_order_inquiry_date.Show
    Else
        Unload Me
        frmMainMenu.Show
    End If

End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Public Sub dr_orderInquiry(p_custCode As String, p_poNo As String, p_itemCode As String, date1 As String, date2 As String, date3 As String, p_seqNo As String)
    combo1.ListIndex = 1
    cboCust.Text = p_custCode
    Call cboCust_Click
    txtPoNo.Text = p_poNo
    Call txtpono_KeyPress(13)
    deliverydate1.Value = date1
    Call deliverydate1_Change
    deliverydate2.Value = date2
    Call deliverydate2_Change

    Call Command1_Click(2)
    If statusfix = 0 Then
        For i = 1 To grid.Rows - 1
            If grid.TextMatrix(i, bteColItemCode) = p_itemCode And grid.TextMatrix(i, bteColSeqNo) = p_seqNo And grid.TextMatrix(i, bteColDate) = Format(date3, "dd MMM yyyy") Then
                grid.TextMatrix(i, bteColSelect) = "S"
                grid.Row = i
                grid.Col = bteColSelect
                grid.Text = "S"
                Call Grid_AfterEdit(i, bteColSelect)
                Exit For
             End If
        Next i
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub txtRevisi_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub ClearData()
    sql = "delete from orderentry_master where not exists(select po_no from orderentry_detail where po_no = orderentry_master.po_no) and po_no = '" & txtPoNo & "'"
    Db.Execute sql
End Sub

Private Sub TxtSerialFrom_LostFocus()
If TxtSerialFrom <> "" Then TxtSerialTo.Text = GetSerialTo(Trim(TxtSerialFrom), txtQty)
If TxtSerialFrom = "" Then TxtSerialTo = ""
End Sub


Private Function uf_validasi_upload(TradeCode As String, PONO As String) As Boolean
                           
   Dim adoCmd As New Command
   Dim RS As New Recordset
                        
    adoCmd.ActiveConnection = Db.ConnectionString
    adoCmd.CommandTimeout = 120
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = "sp_validasi_upload_OrderEntry"
    adoCmd.Parameters(1) = TradeCode
    adoCmd.Parameters(2) = PONO
                                 
    Set RS = adoCmd.Execute
    If Not RS.EOF Then
        
        If RS("Result") = "NG" Then
            uf_validasi_upload = False
        Else
            uf_validasi_upload = True
        End If
            
    End If
                        

End Function
Private Sub loadingformReport()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3


 LblErrMsg = ""
            MousePointer = vbHourglass
            
           
            
    sql = " SELECT a.Cust_Code, a.item_code, d.item_name, Delivery_date, " & _
          " a.Po_No, Qty, SerialNoFrom, SerialNoTo, Remarks ,b.location_code,trade_name " & _
          " from orderentry_detail a " & _
          " INNER JOIN orderentry_master b ON a.po_no=b.po_no " & _
          " INNER JOIN trade_master c ON b.location_code=c.trade_code " & _
          " INNER JOIN item_master d ON a.item_code = d.item_code " & _
          " where a.po_no= '" & CboPOnO.Text & "' " & _
          " order by a.delivery_date, a.makeritem_code, a.seq_no "
   
            
          sqlprint = sql
          If rsRpt.State <> adStateClosed Then rsRpt.Close
         
          Set rsRpt = Db.Execute(sql)
         
         
         
         If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
    
         'Set report = application.OpenReport(App.path & "\Reports\rptSupplyList.rpt")
          Set report = application.OpenReport(App.path & "\Reports\LoadingForm.rpt")
          
         report.Database.Tables(1).SetDataSource rsRpt
         
    
    report.ReportTitle = "Loading Form"
    printorient = 1 'Landscape
    reportcode = "LoadingForm"
        
    Rpt.CRViewer1.ReportSource = report
    Rpt.CRViewer1.ViewReport
    Rpt.CRViewer1.Zoom (75)
        
    Rpt.WindowState = 2
    Rpt.Show 1
    
    Me.MousePointer = vbDefault
End Sub



Private Function uf_ValidasiSerialNo() As String
 Dim adoCmd As New Command
Dim RsSerial As New ADODB.Recordset
Dim strSQL As String
Dim seqNo As Integer
LblErrMsg.Caption = ""




If TxtSeqNo.Text = "" Then
    seqNo = 0
Else
    seqNo = TxtSeqNo.Text
End If


strSQL = "exec SP_OrderEntry_Validasi_SerialNo '" & CboPOnO.Text & "'," & seqNo & ",'" & Trim(TxtSerialFrom.Text) & "','" & Trim(TxtSerialTo.Text) & "'    "


Set RsSerial = Db.Execute(strSQL)
i = 1
If Not (RsSerial.BOF And RsSerial.EOF) Then
        With grid
            Do While Not RsSerial.EOF
                If i = 1 Then
                    uf_ValidasiSerialNo = "Serial No Already Exist..! " & uf_ValidasiSerialNo & RsSerial("PO_No")
                Else
                    uf_ValidasiSerialNo = uf_ValidasiSerialNo & "," & RsSerial("PO_No")
                End If
                RsSerial.MoveNext
                i = i + 1
            Loop
        End With
End If
    RsSerial.Close


End Function


