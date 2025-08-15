VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPart_RecUn 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Part (Material) Receipt Unscheduled"
   ClientHeight    =   11070
   ClientLeft      =   1125
   ClientTop       =   1845
   ClientWidth     =   15285
   Icon            =   "FrmPart_RecUn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11070
   ScaleWidth      =   15285
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtNoSeri 
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
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   72
      Tag             =   "TTFF*/"
      Top             =   8760
      Width           =   1005
   End
   Begin VB.TextBox txtRegisterNo 
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
      Left            =   9000
      MaxLength       =   25
      TabIndex        =   70
      Tag             =   "TTFF*/"
      Top             =   8760
      Width           =   1725
   End
   Begin VB.TextBox txtRec_Status 
      Enabled         =   0   'False
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
      Left            =   0
      MaxLength       =   25
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox LblItemName 
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
      Height          =   255
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   8250
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1215
      Left            =   345
      TabIndex        =   58
      Top             =   915
      Width           =   14595
      Begin VB.CommandButton CmdSearch 
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
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   690
         Width           =   1185
      End
      Begin VB.TextBox Lblsupp 
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
         Height          =   255
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   285
         Width           =   7155
      End
      Begin MSComCtl2.DTPicker MRecDate 
         Height          =   345
         Left            =   1800
         TabIndex        =   2
         Top             =   705
         Width           =   1515
         _ExtentX        =   2672
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
         CustomFormat    =   "MMM yyyy"
         Format          =   130220035
         UpDown          =   -1  'True
         CurrentDate     =   37868
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month Period"
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
         Left            =   270
         TabIndex        =   60
         Top             =   780
         Width           =   1110
      End
      Begin VB.Line Line4 
         X1              =   3420
         X2              =   10560
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code "
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
         Left            =   285
         TabIndex        =   59
         Top             =   315
         Width           =   1275
      End
      Begin MSForms.ComboBox CboSupp 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   255
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2672;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   345
      TabIndex        =   55
      Top             =   9165
      Width           =   14595
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
         Left            =   30
         TabIndex        =   56
         Top             =   210
         Width           =   14235
      End
   End
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "..."
      Height          =   300
      Left            =   2610
      TabIndex        =   11
      Top             =   8220
      Width           =   300
   End
   Begin VB.TextBox txtPackage 
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
      Left            =   8145
      MaxLength       =   12
      TabIndex        =   31
      Top             =   10515
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox txtBC40 
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
      Left            =   5820
      MaxLength       =   30
      TabIndex        =   6
      Top             =   7290
      Width           =   1410
   End
   Begin VB.CommandButton CmdBC40 
      BackColor       =   &H0080FFFF&
      Caption         =   "B.C 4.0"
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
      Left            =   9915
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9900
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
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
      Left            =   11195
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9900
      Width           =   1185
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
      Left            =   12475
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9900
      Width           =   1185
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      Left            =   12990
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8220
      Width           =   1815
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
      Left            =   8245
      MaxLength       =   12
      TabIndex        =   14
      Top             =   8220
      Width           =   1035
   End
   Begin VB.CommandButton CmdData 
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
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9900
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdData 
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
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9900
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdData 
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9900
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdData 
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
      Left            =   7170
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9900
      Visible         =   0   'False
      Width           =   1185
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
      Left            =   13755
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9900
      Width           =   1185
   End
   Begin VB.CommandButton CmdMenu 
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
      Left            =   345
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   9900
      Width           =   1185
   End
   Begin VB.TextBox TxtRemarks 
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
      Left            =   1305
      MaxLength       =   50
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   8760
      Width           =   4350
   End
   Begin VB.TextBox TxtSj 
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
      Left            =   945
      MaxLength       =   25
      TabIndex        =   5
      Top             =   7290
      Width           =   1500
   End
   Begin MSComCtl2.DTPicker TglReceipt 
      Height          =   315
      Left            =   13395
      TabIndex        =   9
      Top             =   7290
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
      Format          =   130220035
      CurrentDate     =   37868
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13050
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   300
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4920
      Left            =   345
      TabIndex        =   4
      Top             =   2250
      Width           =   14595
      _cx             =   25744
      _cy             =   8678
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
      BackColorFixed  =   12640511
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
      SelectionMode   =   1
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
   Begin MSComCtl2.DTPicker DtBCDate 
      Height          =   315
      Left            =   8010
      TabIndex        =   7
      Top             =   7290
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
      Format          =   130220035
      CurrentDate     =   37868
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Seri"
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
      Left            =   5760
      TabIndex        =   73
      Top             =   8835
      Width           =   690
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Register No."
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
      Left            =   7800
      TabIndex        =   71
      Top             =   8835
      Width           =   1050
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BC Type"
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
      Left            =   2520
      TabIndex        =   68
      Top             =   7350
      Width           =   735
   End
   Begin MSForms.ComboBox CbotypeBC 
      Height          =   315
      Left            =   3270
      TabIndex        =   67
      Top             =   7290
      Width           =   1575
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2778;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BC Date"
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
      Left            =   7260
      TabIndex        =   66
      Top             =   7350
      Width           =   720
   End
   Begin MSForms.ComboBox cbocurr 
      Height          =   330
      Left            =   10320
      TabIndex        =   16
      Top             =   8220
      Width           =   855
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "1508;582"
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line3 
      X1              =   3000
      X2              =   5400
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Label LblItemCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ItemCode"
      Height          =   195
      Left            =   3300
      TabIndex        =   65
      Top             =   9990
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label LblRecDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rec date"
      Height          =   195
      Left            =   3300
      TabIndex        =   64
      Top             =   9990
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label LblRecCls 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rec Cls"
      Height          =   195
      Left            =   3300
      TabIndex        =   63
      Top             =   9990
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label LblWHCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WHCode"
      Height          =   195
      Left            =   3300
      TabIndex        =   62
      Top             =   9990
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label LblQty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   3300
      TabIndex        =   61
      Top             =   9990
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Part (Material) Receipt Unscheduled"
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
      Left            =   345
      TabIndex        =   57
      Top             =   330
      Width           =   14595
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   345
      Top             =   8100
      Width           =   14595
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transport By"
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
      Left            =   11025
      TabIndex        =   54
      Top             =   8865
      Width           =   1110
   End
   Begin MSForms.ComboBox cboTransport 
      Height          =   315
      Left            =   12345
      TabIndex        =   21
      Top             =   8805
      Width           =   900
      VariousPropertyBits=   746604571
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "1587;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Package"
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
      Left            =   7170
      TabIndex        =   53
      Top             =   10575
      Visible         =   0   'False
      Width           =   720
   End
   Begin MSForms.ComboBox cboPackage 
      Height          =   315
      Left            =   9090
      TabIndex        =   32
      Top             =   10515
      Visible         =   0   'False
      Width           =   1005
      VariousPropertyBits=   746604571
      MaxLength       =   4
      DisplayStyle    =   3
      Size            =   "1773;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line5 
      X1              =   13350
      X2              =   14865
      Y1              =   9090
      Y2              =   9090
   End
   Begin VB.Label lblTransport 
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
      Height          =   255
      Left            =   13350
      TabIndex        =   52
      Top             =   8835
      Width           =   1530
   End
   Begin VB.Line Line6 
      Visible         =   0   'False
      X1              =   10125
      X2              =   11640
      Y1              =   10815
      Y2              =   10815
   End
   Begin VB.Label lblPackage 
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
      Height          =   255
      Left            =   10170
      TabIndex        =   51
      Top             =   10545
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDFE3&
      BackStyle       =   0  'Transparent
      Caption         =   "BC No."
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
      Left            =   4890
      TabIndex        =   50
      Top             =   7350
      Width           =   600
   End
   Begin VB.Label LblRec 
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
      Height          =   255
      Left            =   11280
      TabIndex        =   49
      Top             =   7290
      Width           =   930
   End
   Begin VB.Line Line1 
      X1              =   11310
      X2              =   12270
      Y1              =   7560
      Y2              =   7560
   End
   Begin MSForms.ComboBox cboprice 
      Height          =   315
      Left            =   11235
      TabIndex        =   18
      Top             =   8220
      Width           =   1665
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "2937;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label14 
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
      Left            =   9630
      TabIndex        =   48
      Top             =   7830
      Width           =   330
   End
   Begin MSForms.ComboBox cbocurr1 
      Height          =   315
      Left            =   10305
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8220
      Width           =   855
      VariousPropertyBits=   746604575
      BackColor       =   -2147483624
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "1508;556"
      BoundColumn     =   2
      MatchEntry      =   1
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox Cbounit 
      Height          =   315
      Left            =   9360
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8220
      Width           =   855
      VariousPropertyBits=   746604575
      BackColor       =   -2147483624
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "1508;556"
      BoundColumn     =   2
      MatchEntry      =   1
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CboItem 
      Height          =   315
      Left            =   480
      TabIndex        =   10
      Top             =   8220
      Width           =   2055
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "3625;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label13 
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
      Left            =   480
      TabIndex        =   47
      Top             =   7830
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
      Left            =   8610
      TabIndex        =   46
      Top             =   7830
      Width           =   300
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
      Left            =   11865
      TabIndex        =   45
      Top             =   7830
      Width           =   420
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
      Left            =   10530
      TabIndex        =   44
      Top             =   7830
      Width           =   390
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
      Left            =   13560
      TabIndex        =   43
      Top             =   7830
      Width           =   660
   End
   Begin VB.Label Label6 
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
      Left            =   3000
      TabIndex        =   42
      Top             =   7830
      Width           =   960
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
      Left            =   7155
      TabIndex        =   41
      Top             =   7830
      Width           =   690
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WH Code"
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
      Left            =   5505
      TabIndex        =   40
      Top             =   7830
      Width           =   795
   End
   Begin MSForms.ComboBox CboRecCls 
      Height          =   315
      Left            =   10530
      TabIndex        =   8
      Top             =   7290
      Width           =   735
      VariousPropertyBits=   746604571
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1296;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CboWHCode 
      Height          =   315
      Left            =   5505
      TabIndex        =   13
      Top             =   8220
      Width           =   1245
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "2196;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Lbladdress 
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
      Height          =   255
      Left            =   6825
      TabIndex        =   39
      Top             =   8250
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   6825
      X2              =   8145
      Y1              =   8490
      Y2              =   8490
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Cls"
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
      TabIndex        =   38
      Top             =   7350
      Width           =   960
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Date"
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
      Left            =   12270
      TabIndex        =   37
      Top             =   7350
      Width           =   1095
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   1200
      TabIndex        =   36
      Top             =   450
      Width           =   1185
   End
   Begin VB.Label LblPart 
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
      Index           =   5
      Left            =   450
      TabIndex        =   35
      Top             =   8835
      Width           =   765
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SJ No"
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
      Left            =   390
      TabIndex        =   34
      Top             =   7350
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   345
      Top             =   7740
      Width           =   14595
   End
End
Attribute VB_Name = "FrmPart_RecUn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rst As Recordset, RsW As Recordset, RsI As Recordset
Dim SQLT As String, SqlW As String, SqlI As String
Dim RsPM As Recordset, RsPD As Recordset, RsPr As Recordset
Dim SqlPM As String, SqlPD As String

Dim baru As Boolean
Dim Pos As Integer, jml As Integer
Dim StaErr As Boolean, nErr As Integer, StrWDel As String, nOK As Integer
Dim HakU As Integer
Dim staBaru As String
Dim SeqID As Long
Dim seqNo As New clsMRP

Dim KeyProd As String
Dim tampungQty As Double
Dim blnFix As Integer, thnFix As Integer
Dim newDb As New ADODB.Connection
Dim provisionCls As String, l_SJNo As String
Dim xcurr As String, blnshow As Boolean
Dim DateActual, receiptDate As Date
Dim validate As Boolean

Dim bteColSelect As Byte
Dim bteColSJNo As Byte
Dim bteColReceiptDate As Byte
Dim bteColProdCode As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColWHCode As Byte
Dim bteColAddress As Byte
Dim bteColReceipt As Byte
Dim bteColQty As Byte
Dim bteColUnit As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColUnitCls As Byte
Dim bteColCurrCode As Byte
Dim bteColRemark As Byte
Dim bteColSeqNo As Byte
Dim bteColBC As Byte
Dim bteColTransport As Byte
Dim bteColPackageQty As Byte
Dim bteColPackage As Byte
Dim bteColBctype As Byte
Dim bteColRecStatus As Byte
Dim bteColNoRegister As Byte
Dim bteColNoSeri As Byte

Dim bteColBCDate As Byte

Dim bteHakPrice As Byte

Dim dateUp As Date

Sub Header()
    Dim C As Byte
    
    bteColSelect = 0
    bteColSJNo = 1
    bteColReceiptDate = 2
    bteColProdCode = 3
    bteColPartNo = 4
    bteColDesc = 5
    bteColWHCode = 6
    bteColAddress = 7
    bteColReceipt = 8
    bteColQty = 9
    bteColUnit = 10
    bteColCurr = 11
    bteColPrice = 12
    bteColAmount = 13
    bteColUnitCls = 14
    bteColCurrCode = 15
    bteColRemark = 16
    bteColSeqNo = 17
    bteColBC = 18
    bteColTransport = 19
    bteColPackageQty = 20
    bteColPackage = 21
    bteColBCDate = 22
    bteColBctype = 23
    bteColRecStatus = 24
    bteColNoRegister = 25
    bteColNoSeri = 26
    
    With Grid
        .ColS = 27
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColSJNo) = "SJ No"
        .TextMatrix(0, bteColReceiptDate) = "Receipt Date"
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColPartNo) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColWHCode) = "WH Code"
        .TextMatrix(0, bteColAddress) = "Address"
        .TextMatrix(0, bteColReceipt) = "Receipt"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColUnitCls) = "Unit"
        .TextMatrix(0, bteColCurrCode) = "Curr"
        .TextMatrix(0, bteColRemark) = "Remarks"
        .TextMatrix(0, bteColSeqNo) = "Seqno"
        .TextMatrix(0, bteColBC) = "BC No"
        .TextMatrix(0, bteColTransport) = "Transport Cls"
        .TextMatrix(0, bteColPackageQty) = "Package Qty"
        .TextMatrix(0, bteColPackage) = "Package Cls"
        .TextMatrix(0, bteColBctype) = "BC Type"
        .TextMatrix(0, bteColRecStatus) = "Rec. Status"
        .TextMatrix(0, bteColNoRegister) = "No. Register"
        .TextMatrix(0, bteColNoSeri) = "No."
        
        .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
        .ColAlignment(bteColSJNo) = flexAlignCenterCenter
        .ColAlignment(bteColReceiptDate) = flexAlignCenterCenter
        .ColAlignment(bteColProdCode) = flexAlignCenterCenter
        .ColAlignment(bteColPartNo) = flexAlignCenterCenter
        .ColAlignment(bteColDesc) = flexAlignCenterCenter
        .ColAlignment(bteColWHCode) = flexAlignCenterCenter
        .ColAlignment(bteColAddress) = flexAlignCenterCenter
        .ColAlignment(bteColReceipt) = flexAlignCenterCenter
        .ColAlignment(bteColQty) = flexAlignCenterCenter
        .ColAlignment(bteColUnit) = flexAlignCenterCenter
        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        .ColAlignment(bteColPrice) = flexAlignCenterCenter
        .ColAlignment(bteColAmount) = flexAlignCenterCenter
        .ColAlignment(bteColUnitCls) = flexAlignCenterCenter
        .ColAlignment(bteColCurrCode) = flexAlignCenterCenter
        .ColAlignment(bteColRemark) = flexAlignLeftCenter
        .ColAlignment(bteColBctype) = flexAlignCenterCenter
        .ColAlignment(bteColRecStatus) = flexAlignCenterCenter
        .ColAlignment(bteColNoRegister) = flexAlignCenterCenter
        .ColAlignment(bteColNoSeri) = flexAlignRightCenter
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColSJNo) = 2000
        .ColWidth(bteColReceiptDate) = 1300
        .ColWidth(bteColProdCode) = 2000
        .ColWidth(bteColPartNo) = 2000
        .ColWidth(bteColDesc) = 2500
        .ColWidth(bteColWHCode) = 1000
        .ColWidth(bteColReceipt) = 800
        .ColWidth(bteColQty) = 1000
        .ColWidth(bteColUnit) = 500
        .ColWidth(bteColCurr) = 500
        .ColWidth(bteColPrice) = 1400
        .ColWidth(bteColAmount) = 1600
        .ColWidth(bteColRemark) = 1600
        .ColWidth(bteColBC) = 1500
        .ColWidth(bteColBctype) = 1500
        .ColWidth(bteColNoRegister) = 1700
        .ColWidth(bteColNoSeri) = 500
        
        .ColHidden(bteColUnitCls) = True
        .ColHidden(bteColCurrCode) = True
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColTransport) = True
        .ColHidden(bteColPackageQty) = True
        .ColHidden(bteColPackage) = True
        .ColHidden(bteColAddress) = True
        .ColHidden(bteColBC) = True
        .ColHidden(bteColBCDate) = True
        .ColHidden(bteColBctype) = True
        .ColHidden(bteColRecStatus) = False
        
        .ColHidden(bteColCurr) = (bteHakPrice = 0)
        .ColHidden(bteColPrice) = (bteHakPrice = 0)
        .ColHidden(bteColAmount) = (bteHakPrice = 0)
        
        
        .EditMaxLength = 1
    End With
End Sub

Private Sub cbocurr_Change()
    If cbocurr.ListCount > 0 Then
        cbocurr = Trim(cbocurr)
        If cbocurr.MatchFound Then
            xcurr = cbocurr.List(, 0)
            blnshow = False
            Call browseprice(CboItem, TglReceipt, CboSupp, xcurr)
        Else
            xcurr = ""
        End If
    End If
End Sub

Private Sub cbocurr_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub CboItem_Change()
    cboitem_Click
End Sub
Private Sub comboBCtype()
Dim ls_sql As String
Dim rs_combo As New ADODB.Recordset
Dim i As Integer


CbotypeBC.columnCount = 1
CbotypeBC.clear

ls_sql = "select bc_type from BC_master"
rs_combo.Open ls_sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
i = 0

Do While Not rs_combo.EOF
CbotypeBC.AddItem rs_combo("Bc_type")
rs_combo.MoveNext
Loop

CbotypeBC.ColumnWidths = "90"
CbotypeBC.ListWidth = 90
CbotypeBC.ListRows = 7


End Sub
Private Sub cboitem_Click()
    Dim RSX As Recordset
    Dim RsI2 As Recordset
    
    On Error GoTo X
    
    If CboItem.MatchFound Then
        LblItemName = Trim$(uf_GetItemDescription(Trim$(CboItem)))
        Lbladdress = ""
        cboprice = ""
    
        Set RsI2 = Db.Execute("Select WH_code,unit_cls from Item_master where item_code='" & Trim(CboItem) & "'")
        If Not RsI2.EOF Then
            CboWHCode = Trim$(RsI2!wh_code)
            CboWHCode_Click
            Set RSX = Db.Execute("Select unit_cls UC,Currency_code CC from Price_master " & _
                "where (price_cls = '01' or price_cls = '05') and item_code='" & Trim$(CboItem) & "' and (trade_code = '" & CboSupp & "' or  trade_code='000000')")
            If Not RSX.EOF Then
                cbocurr.Text = uf_GetCurrencyDescription(Trim(RSX!CC))
                cbocurr.TextColumn = 2
            Else
            cbocurr.Text = ""
            End If
            Call filterCboUnit(RsI2!Unit_cls, Cbounit)
            Cbounit.TextColumn = 2
            Cbounit.Text = uf_GetUnitDescription(Trim(RsI2!Unit_cls))
            txtPackage = Format(0, gs_formatBox)
        End If
        blnshow = True
        browseprice Trim$(CboItem), Format(TglReceipt, "YYYY-MM-DD"), Trim$(CboSupp), xcurr
        
        If baru = True Then up_GetNoSeri (CboItem.Text)
        
    End If
    Exit Sub
X:
    Lbladdress = ""
End Sub

Private Sub cboitem_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboPackage_Change()
    If cboPackage.MatchFound Then
        lblPackage = cboPackage.Column(1)
    Else
        lblPackage = ""
    End If
End Sub

Private Sub cboprice_Change()
    If Trim(cboprice) = "." Then cboprice = "": Exit Sub
    If Trim(cboprice) = "" Then txtamount = "0": Exit Sub
    If Left(Trim(cboprice), 1) = "," Then cboprice = Right(Trim(cboprice), Len(Trim(cboprice)) - 1)
    If cboprice.Text <> "" And IsNumeric(cboprice) = True Then
        If txtqty <> "" And IsNumeric(txtqty) = True Then
            txtamount.Text = Format(CDbl(cboprice.Text) * CDbl(txtqty.Text), gs_formatAmount)
        Else
            txtamount = Format(0, gs_formatAmount)
        End If
    Else
        txtamount = Format(0, gs_formatAmount)
    End If
    If Right(Trim(txtamount), 1) = "." Then txtamount = Left(Trim(txtamount), Len(Trim(txtamount)) - 1)
End Sub

Private Sub cboprice_Click()
    If Val(cboprice) = 0 Then txtamount = 0: Exit Sub
    If Val(txtqty) = 0 Then txtamount = 0: Exit Sub
    If CDbl(txtqty) > 0 And CDbl(cboprice) > 0 Then
        txtamount = CDbl(txtqty) * CDbl(cboprice)
        If Round(CDbl(txtamount)) / CDbl(txtamount) = 1 Then
            txtamount = Format(CDbl(txtamount), gs_formatAmount)
        Else
            txtamount = Format(CDbl(txtamount), gs_formatAmount)
        End If
    Else
        txtamount = 0
    End If
    txtamount = Format(CDbl(txtamount), gs_formatAmount)
End Sub

Private Sub CboPrice_KeyPress(KeyAscii As MSForms.ReturnInteger)
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub cboPrice_LostFocus()
    cboprice = Format(cboprice, gs_formatPrice)
    DoEvents
End Sub

Private Sub CboRecCls_Change()
    CboRecCls_Click
End Sub

Private Sub CboRecCls_Click()
    CboRecCls = CboRecCls
    If CboRecCls.MatchFound Then LblRec = CboRecCls.List(CboRecCls.ListIndex, 1)
End Sub

Private Sub CboRecCls_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub CboSupp_Change()
    CboSupp_Click
End Sub

Private Sub CboSupp_Click()
    CboSupp = CboSupp
    Lblsupp = ""
    CbotypeBC = ""
    GetBCType
    If CboSupp.MatchFound = True Then Lblsupp = CboSupp.List(CboSupp.ListIndex, 1) ': display
End Sub

Private Sub CboSupp_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboTransport_Change()
    If cboTransport.MatchFound Then
        lblTransport = cboTransport.Column(1)
    Else
        lblTransport = ""
    End If
End Sub


Private Sub CbotypeBC_Change()
    LblErr.Caption = ""
    
    up_GetNoSeri (CboItem.Text)
    
    If CbotypeBC.Text = "4.0" Or CbotypeBC.Text = "2.6.2" Or CbotypeBC.Text = "2.3" Then
        txtRec_Status.Text = "02"
    Else
        txtRec_Status.Text = "01"
    End If
    
   TxtSj_LostFocus
End Sub

Private Sub Cbounit_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub CboWHCode_Change()
    CboWHCode_Click
End Sub

Private Sub CboWHCode_Click()
    Dim RsIA As New Recordset
    CboWHCode = CboWHCode
    CboItem = CboItem
    Lbladdress = ""
    If CboWHCode.MatchFound = True And CboItem.MatchFound = True Then
    
        If RsIA.State <> adStateClosed Then RsIA.Close
        RsIA.Open "Select Address from item_master where item_code='" & CboItem & "' and WH_code='" & CboWHCode & "'", Db, adOpenKeyset, adLockOptimistic
        If Not RsIA.EOF Then Lbladdress = Trim$(RsIA!Address & "")
    End If
    If RsIA.State <> adStateClosed Then RsIA.Close
End Sub

Private Sub CboWHCode_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub CmdBC40_Click()

    Me.MousePointer = vbHourglass
    
    If Grid.Rows > 1 And Trim(Grid.TextMatrix(Grid.Row, bteColBC)) <> "" Then
        BC40 Grid.TextMatrix(Grid.Row, bteColBC)
    Else
        LblErr = DisplayMsg("0013")
    End If
    
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdBrowser_Click()
 If CboItem.Enabled = True Then
  Me.MousePointer = vbHourglass
  frm_BrowseItem.getItemCode = CboItem.Text
  frm_BrowseItem.Show 1
  CboItem.Text = frm_BrowseItem.getItemCode
  Me.MousePointer = vbDefault
 End If
End Sub

Private Sub cmdCancel_Click()
    Dim Ikd As Long
    For Ikd = 1 To Grid.Rows - 1
        Grid.TextMatrix(Ikd, 0) = ""
    Next
    Kosong
End Sub

Private Sub cmdClear_Click()
    Header
    CboItem = ""
    CboWHCode = ""
    Lbladdress = ""
    txtqty = ""
    Cbounit = ""
    cbocurr = ""
    cboprice = ""
    txtamount = ""
    TxtSj = ""
    txtBC40 = ""
    CbotypeBC = ""
    cboTransport = ""
    txtPackage = ""
    cboPackage = ""
    CboRecCls = ""
    TglReceipt = Now()
    DtBCDate = Now()
    TxtRemarks = ""
    CboSupp = ""
    LblErr = ""
    LblRecCls = ""
    LblWHCode = ""
    LblItemCode = ""
    LblQty = ""
    LblRecDate = ""
    LblRec = ""
    txtRegisterNo = ""
    txtNoSeri.Text = ""
    
    baru = True
End Sub

Private Sub CmdData_Click(Index As Integer)
    If RsPM.RecordCount = 0 Then Exit Sub
    Select Case Index
    Case 0
        RsPM.MoveFirst
        LblErr = DisplayMsg(4020) '"This is the first page of file !"
    Case 1
        If Not RsPM.BOF And RsPM.AbsolutePosition > 1 Then
            RsPM.MovePrevious
            LblErr = ""
        Else
            RsPM.MoveFirst
            LblErr = DisplayMsg(4020) '""This is the first page of file !"
        End If
    Case 2
        If Not RsPM.EOF And RsPM.AbsolutePosition < RsPM.RecordCount Then
            RsPM.MoveNext
            LblErr = ""
        Else
            RsPM.MoveLast
            LblErr = DisplayMsg(4021) '"This is the last page of file !"
        End If
    Case 3
        RsPM.MoveLast
        LblErr = DisplayMsg(4021) '""This is the last page of file !"
    End Select
    CboSupp.Text = Trim$(RsPM!Supplier_Code)
    CboRecCls = Trim$(RsPM!receipt_cls)
    TglReceipt = Format(RsPM!Receipt_Date, "dd mmm yyyy")
    RsPM.filter = ""
    display
End Sub

Private Sub CmdMenu_Click()
    rst.Close
    Set rst = Nothing
    RsW.Close
    Set RsW = Nothing
    frmMainMenu.Show
    Unload Me
    DoEvents
End Sub

Private Sub cmdSearch_Click()
    CboSupp = CboSupp
    If CboSupp.MatchFound Then display
End Sub

Private Sub CmdSubmit_Click()
    Dim strS As Integer, Jawab As Integer, CekS As Boolean, CekD As Boolean
    Dim strD As Integer, ie As Integer
    Dim FixMonth As String * 2
    Dim FixYear As String * 4
    Dim NoSeq As Long, PosRec As Long
    Dim rsProv As New ADODB.Recordset
        
    If HakU = 0 Then LblErr = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
     If up_Validasi_Item(Trim(CboItem.Text)) = False Then up_SendEmail
    
    CekS = False
    CekD = False
    StaErr = False
    
    Dim li_diff As Integer
    li_diff = DateDiff("M", uf_GetLastClosing("fulldate"), TglReceipt)
    If li_diff > 2 Or li_diff < 0 Then
        LblErr = DisplayMsg(8022) '"Please Select Valid Time Range!"
        Exit Sub
    End If
    
    strS = 0
    strD = 0
    
    Dim tampungBln As String
    tampungBln = seqNo.blnAkhir()
    blnFix = Split(tampungBln, ",")(0)
    thnFix = Split(tampungBln, ",")(1)
    
    If baru = False Then
        
        '# Yudha 2007-12-11
        '# Cek AP Invoice, jika sudah dibuat tidak bisa di-update
        Dim adoInv As New ADODB.Recordset
        sql = "select invoice_no from invoicesupplier_detail where receiptseq_no = '" & SeqID & "'"
        adoInv.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not adoInv.EOF Then
            LblErr = DisplayMsg(4110) & " (Invoice No : " & Trim(adoInv.Fields("invoice_no")) & ")": Me.MousePointer = vbDefault: adoInv.Close: Exit Sub
        End If
        adoInv.Close
        
        strS = Grid.FindRow("S", , bteColSelect, False)
        strD = Grid.FindRow("D", , bteColSelect, False, False)
        
        If strD > 0 Then CekD = True: Jawab = MsgBox(DisplayMsg(8071), vbInformation + vbYesNo + vbDefaultButton2, "Confirmation")
        If Jawab = vbYes Then DataGrid
        
        If strS > 0 And cek Then
            DataGrid
            If StaErr = False Then
                LblErr = DisplayMsg(1101)
                baru = True
                PosRec = strS
                NoSeq = Grid.TextMatrix(PosRec, bteColSeqNo)
                GridAct PosRec, NoSeq, 0
                Kosong
                CboItem = ""
                txtqty = ""
                Cbounit = ""
                cbocurr = ""
                cboprice = ""
                txtamount = ""
                Dim IK As Long
                For IK = 1 To Grid.Rows - 1
                    Grid.TextMatrix(IK, bteColSelect) = ""
                Next
            Else
                LblErr = DisplayMsg(1102)
            End If
        End If
        
        Dim PRec As Integer
        If CekD Then
            If Jawab = vbYes Then
                If StaErr = False Then
                    LblErr = DisplayMsg(1201)
                    baru = True
                Else
                    LblErr = DisplayMsg(1202)
                End If
            ElseIf Jawab = vbNo Then
                LblErr = ""
                baru = True
                Kosong
                Dim Ikd As Long
                For Ikd = 1 To Grid.Rows - 1
                    Grid.TextMatrix(Ikd, bteColSelect) = ""
                Next
            End If
        End If
        strS = 0
    
    Else
        
        Dim SqlU As String
        If cek Then
           err.clear
            'Simpan data
            KeyProd = seqNo.keyReceipt
            If Val(cboprice) = 0 Then
                'SqlU = "insert into Part_Receipt (Seq_No, Supplier_Code, PO_No, Warehouse_Code, Address, Receipt_Cls, Receipt_Date, " & _
                    "Item_Code, Qty, Unit_Cls, Currency_Code, Price, Amount, SuratJalan_No, ProductionResult_Cls, DailySeq_No, Remarks, " & _
                    "BC40_No, Transport_Cls, Package_Qty, Package_Cls, " & _
                    "Last_Update, Last_User) " & _
                    "Values (" & KeyProd & ",'" & Trim$(CboSupp) & "','0','" & Trim$(cboWhCode) & "','" & Trim$(Lbladdress) & "','" & Trim$(CboRecCls) & " ','" & Format(TglReceipt.Value, "yyyy-mm-dd") & "','" & _
                    Trim$(CboItem) & "'," & CDbl(txtQty) & ",'" & Trim$(cboUnit.List(cboUnit.ListIndex, 0)) & "','" & Trim$(cbocurr.List(cbocurr.ListIndex, 0)) & "'," & Val(cboprice) & "," & Val(txtAmount) & ",'" & Trim$(TxtSj) & "','0',Null,'" & Trim$(txtRemarks) & "', '" & _
                    Trim$(txtBC40) & "', '" & Trim(cboTransport) & "', " & CDbl(txtPackage) & ", '" & Trim(cboPackage) & "', " & _
                    "getdate(), '" & userLogin & "')"
              'Update  BY Dudi 25 Desember 2008
              SqlU = "insert into Part_Receipt (Seq_No, Supplier_Code, PO_No, Warehouse_Code, Address, Receipt_Cls, Receipt_Date, " & _
                    "BC40_No, BC40_Date,BC_type , Item_Code, Qty, Unit_Cls, Currency_Code, Price, Amount, SuratJalan_No, ProductionResult_Cls, DailySeq_No, Remarks, " & _
                    "Last_Update, Last_User, Receipt_Status, No_Register, No_Seri) " & _
                    "Values (" & KeyProd & ",'" & Trim$(CboSupp) & "','0','" & Trim$(CboWHCode) & "','" & Trim$(Lbladdress) & "','" & Trim$(CboRecCls) & " ','" & Format(TglReceipt.Value, "yyyy-mm-dd") & "','" & _
                    Trim$(txtBC40) & "','" & Format(DtBCDate.Value, "yyyy-mm-dd") & "','" & Trim(CbotypeBC.Text) & "','" & Trim$(CboItem) & "'," & CDbl(txtqty) & ",'" & Trim$(Cbounit.List(Cbounit.ListIndex, 0)) & "','" & Trim$(cbocurr.List(cbocurr.ListIndex, 0)) & "'," & Val(cboprice) & "," & Val(txtamount) & ",'" & Trim$(TxtSj) & "','0',Null,'" & Trim$(TxtRemarks) & "', " & _
                    "getdate(), '" & userLogin & "', '" & txtRec_Status.Text & "', '" & txtRegisterNo.Text & "', '" & txtNoSeri.Text & "')"
                    
            ElseIf Val(cboprice) > 0 Then
            '    SqlU = "insert into Part_Receipt (Seq_No, Supplier_Code, PO_No, Warehouse_Code, Address, Receipt_Cls, Receipt_Date, " & _
                    "Item_Code, Qty, Unit_Cls, Currency_Code, Price, Amount, SuratJalan_No, ProductionResult_Cls, DailySeq_No, Remarks, " & _
                    "BC40_No, Transport_Cls, Package_Qty, Package_Cls, " & _
                    "Last_Update, Last_User) " & _
                    "values (" & KeyProd & ",'" & Trim$(CboSupp) & "','0','" & Trim$(cboWhCode) & "','" & Trim$(Lbladdress) & "','" & Trim$(CboRecCls) & " ','" & Format(TglReceipt.Value, "yyyy-mm-dd") & "','" & _
                    Trim$(CboItem) & "'," & CDbl(txtQty) & ",'" & Trim$(cboUnit.List(cboUnit.ListIndex, 0)) & "','" & Trim$(cbocurr.List(cbocurr.ListIndex, 0)) & "'," & CDbl(cboprice) & "," & CDbl(txtAmount) & ",'" & Trim$(TxtSj) & "','0',Null,'" & Trim$(txtRemarks) & "', '" & _
                    Trim$(txtBC40) & "', '" & Trim(cboTransport) & "', " & CDbl(txtPackage) & ", '" & Trim(cboPackage) & "', " & _
                    "getdate(), '" & userLogin & "')"
                'Update BY dudi 25 Desember 2008
                SqlU = "insert into Part_Receipt (Seq_No, Supplier_Code, PO_No, Warehouse_Code, Address, Receipt_Cls, Receipt_Date, " & _
                    "BC40_No, BC40_Date,BC_Type, Item_Code, Qty, Unit_Cls, Currency_Code, Price, Amount, SuratJalan_No, ProductionResult_Cls, DailySeq_No, Remarks, " & _
                    "Last_Update, Last_User, Receipt_Status, No_Register, No_Seri) " & _
                    "values (" & KeyProd & ",'" & Trim$(CboSupp) & "','0','" & Trim$(CboWHCode) & "','" & Trim$(Lbladdress) & "','" & Trim$(CboRecCls) & " ','" & Format(TglReceipt.Value, "yyyy-mm-dd") & "','" & _
                    Trim$(txtBC40) & "','" & Format(DtBCDate.Value, "yyyy-mm-dd") & "','" & Trim(CbotypeBC.Text) & "','" & Trim$(CboItem) & "'," & CDbl(txtqty) & ",'" & Trim$(Cbounit.List(Cbounit.ListIndex, 0)) & "','" & Trim$(cbocurr.List(cbocurr.ListIndex, 0)) & "'," & CDbl(cboprice) & "," & CDbl(txtamount) & ",'" & Trim$(TxtSj) & "','0',Null,'" & Trim$(TxtRemarks) & "'" & _
                    ",getdate(), '" & userLogin & "', '" & txtRec_Status.Text & "', '" & txtRegisterNo.Text & "', '" & txtNoSeri.Text & "')"
            End If
            
            PosRec = Grid.Rows - 1
            NoSeq = 0
            
            'On Error GoTo Errhandle

            
            Db.BeginTrans
            Db.Execute SqlU

Errhandle:
            If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                err.clear
                KeyProd = seqNo.keyReceipt
                If Val(cboprice) = 0 Then
                    SqlU = "insert into Part_Receipt (Seq_No, Supplier_Code, PO_No, Warehouse_Code, Address, Receipt_Cls, Receipt_Date, " & _
                        "BC40_No, BC40_Date,BC_type, Item_Code, Qty, Unit_Cls, Currency_Code, Price, Amount, SuratJalan_No, ProductionResult_Cls, DailySeq_No, Remarks, " & _
                        "Last_Update, Last_User, Receipt_Status, No_Register, No_Seri) " & _
                        "Values (" & KeyProd & ",'" & Trim$(CboSupp) & "','0','" & Trim$(CboWHCode) & "','" & Trim$(Lbladdress) & "','" & Trim$(CboRecCls) & " ','" & Format(TglReceipt.Value, "yyyy-mm-dd") & "','" & _
                        Trim$(txtBC40) & "','" & Format(DtBCDate.Value, "yyyy-mm-dd") & "','" & Trim(CbotypeBC.Text) & "','" & Trim$(CboItem) & "'," & CDbl(txtqty) & ",'" & Trim$(Cbounit.List(Cbounit.ListIndex, 0)) & "','" & Trim$(cbocurr.List(cbocurr.ListIndex, 0)) & "'," & Val(cboprice) & "," & Val(txtamount) & ",'" & Trim$(TxtSj) & "','0',Null,'" & Trim$(TxtRemarks) & "', " & _
                        "getdate(), '" & userLogin & "', '" & txtRec_Status.Text & "', '" & txtRegisterNo.Text & "', '" & txtNoSeri.Text & "')"
                ElseIf Val(cboprice) > 0 Then
                    SqlU = "insert into Part_Receipt (Seq_No, Supplier_Code, PO_No, Warehouse_Code, Address, Receipt_Cls, Receipt_Date, " & _
                        "BC40_No, BC40_Date,BC_type, Item_Code, Qty, Unit_Cls, Currency_Code, Price, Amount, SuratJalan_No, ProductionResult_Cls, DailySeq_No, Remarks, " & _
                        "Last_Update, Last_User, Receipt_Status, No_Register) " & _
                        "values (" & KeyProd & ",'" & Trim$(CboSupp) & "','0','" & Trim$(CboWHCode) & "','" & Trim$(Lbladdress) & "','" & Trim$(CboRecCls) & " ','" & Format(TglReceipt.Value, "yyyy-mm-dd") & "','" & _
                        Trim$(txtBC40) & "','" & Format(DtBCDate.Value, "yyyy-mm-dd") & "','" & Trim(CbotypeBC.Text) & "','" & Trim$(CboItem) & "'," & CDbl(txtqty) & ",'" & Trim$(Cbounit.List(Cbounit.ListIndex, 0)) & "','" & Trim$(cbocurr.List(cbocurr.ListIndex, 0)) & "'," & CDbl(cboprice) & "," & CDbl(txtamount) & ",'" & Trim$(TxtSj) & "','0',Null,'" & Trim$(TxtRemarks) & "', " & _
                        "getdate(), '" & userLogin & "', '" & txtRec_Status.Text & "', '" & txtRegisterNo.Text & "', '" & txtNoSeri.Text & "')"
                End If
                PosRec = Grid.Rows - 1
                NoSeq = 0
                Db.Execute SqlU
                If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then GoTo Errhandle
            End If
            
            If err.number = 0 Then
                Db.CommitTrans
                
                If txtRec_Status.Text = "01" Then
                    ProsesStock 1, Trim$(CboItem), Trim$(CboWHCode), Trim$(CboWHCode), Format(TglReceipt, "YYYYMM"), Trim$(txtqty), ""
                End If
                
            
                '*********** Proses Insert ke Supply **************
                sql = "Select Provision_Cls from Item_MaSter where ITem_Code = '" & CboItem & "'"
                Set rsProv = Db.Execute(sql)
                If Not rsProv.EOF Then provisionCls = Trim(rsProv(0)) Else provisionCls = ""
                If provisionCls = "01" Then
                    KeyProd = "R" & KeyProd
                    Call inputSupply(CboItem, CDbl(txtqty))
                End If
                '*************************************************
            
                LblErr = DisplayMsg(1000)
                Dim rsSeq As Recordset
                Set rsSeq = Db.Execute("select top 1 seq_no from part_receipt order by seq_no desc")
                If Not rsSeq.EOF Then
                    NoSeq = Val(rsSeq!Seq_no)
                Else
                    NoSeq = 1
                End If
            
                GridAct PosRec + 1, NoSeq, 1
                Kosong
                CboSupp.SetFocus
            Else
                Db.RollbackTrans
                LblErr = err.Description
                err.clear
            End If
        End If
        baru = True
        SqlU = ""
    End If
End Sub

Function isiPrice(ItemCode As String, tglDO As String, currCode As String) As String
    Dim rsPrice As New ADODB.Recordset
    sql = "select top 1 currency_code,isnull(price,0) from price_master where " & _
        "item_code='" & ItemCode & _
        "' and price_cls='01' " & _
        "and start_date<='" & Format(tglDO, "yyyymmdd") & _
        "' and end_date>='" & Format(tglDO, "yyyymmdd") & _
        "' order by trade_code desc, priority_cls desc"
    Set rsPrice = newDb.Execute(sql)
    If rsPrice.EOF Then
        isiPrice = currCode & ",0"
    Else
        isiPrice = Trim(rsPrice(0)) & "," & CDbl(rsPrice(1))
    End If
    Set rsPrice = Nothing
End Function

Sub inputSupply(ibu As String, Qty As Double, Optional UnitAnak As String, Optional TglRequirement As String, Optional tampungParent As String)
    Dim rsAnak As New ADODB.Recordset
    Dim fromWHCode As String, fromAddress As String, toWHCode As String, toAddress As String
    Dim UnitCls As String, currCD As String, Price As Double, Amount As Double
    Dim itemAnak As String, nilPrice As String, currAnak As String, qtyAnak As Double
    Dim stockWH As String, stockItem As String
    
    '*********Update Supply Anak2nya diambil dr BOM MaSter ***********
    sql = "Select c.Manufacture_Code as Factory_Code,a.Item_Code,a.Qty as qtyAnak,a.Unit_Cls," & _
        "b.WH_Code,b.Address,b.Stockcontrol_Cls as stockItem,d.Stockcontrol_Cls as stockWH " & _
        "from BOM_Master a,Item_Master b, Item_Master c, Warehouse_Master d " & _
        "where a.Item_Code = b.Item_Code " & _
        "And a.Parent_ItemCode = c.Item_Code " & _
        "And b.WH_Code = d.WH_Code " & _
        "And a.Parent_ItemCode = '" & ibu & _
        "' And Start_Date <='" & Format(TglReceipt, "yyyyMMdd") & _
        "' And End_Date >= '" & Format(TglReceipt, "yyyyMMdd") & "'"
    Set rsAnak = newDb.Execute(sql)
    
    If Not rsAnak.EOF Then
        Do While Not rsAnak.EOF
        fromWHCode = rsAnak("Wh_Code")
        fromAddress = IIf(IsNull(rsAnak("Address")), "", rsAnak("Address"))
        toWHCode = IIf(IsNull(rsAnak("Factory_Code")), "", rsAnak("Factory_Code"))
        itemAnak = rsAnak("Item_Code")
        qtyAnak = rsAnak("QtyAnak") * CDbl(Qty)
        UnitCls = rsAnak("Unit_Cls")
        nilPrice = isiPrice(itemAnak, Format(TglReceipt, "yyyy-MM-dd"), currCD)
        currAnak = Split(nilPrice, ",")(0)
        Price = Split(nilPrice, ",")(1)
        Amount = CDbl(qtyAnak) * CDbl(Price)
        stockWH = rsAnak("StockWH")
        stockItem = rsAnak("StockItem")
    
        sql = "insert into Part_Supply(FromWarehouse_Code,From_Address,ToWarehouse_Code,ChildSupply_date,ChildItem_Code,Supply_Cls," & _
            "ChildRequirement_Qty,ChildUnit_Cls,Currency_Code,Price,Amount,ParentItem_Code,Lot_No,Production_Date,Remarks,Do_NO,Last_Update,Last_User) " & _
            "values ('" & fromWHCode & "','" & fromAddress & "','" & toWHCode & "','" & Format(TglReceipt, "yyyy-MM-dd") & "','" & itemAnak & "','S'," & _
            CDbl(qtyAnak) & ",'" & UnitCls & "','" & currAnak & "'," & Price & "," & Amount & ",'" & CboItem & "','',Null,'" & TxtRemarks & "','" & KeyProd & "',getdate(),'" & userLogin & "')"
    
        Db.Execute sql
    
        If stockWH = "01" And stockItem = "01" Then _
            Call seqNo.updateStock(fromWHCode, itemAnak, qtyAnak, "", Format(TglReceipt, "yyyy-MM-dd"), blnFix, thnFix, Db, "Supply", 0, 1)
            If Not (rsAnak.EOF) Then rsAnak.MoveNext
            Loop
        End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    Dim ir As Long, strUnit As String
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    HakU = hakUpdate(Me.Name)
    
    If newDb.State <> adStateClosed Then newDb.Close
    newDb.Open Db.ConnectionString
    
    bteHakPrice = (hakPrice(Me.Name))
    cbocurr.Visible = (bteHakPrice = 1)
    cboprice.Visible = (bteHakPrice = 1)
    txtamount.Visible = (bteHakPrice = 1)
    Label7.Visible = (bteHakPrice = 1)
    Label11.Visible = (bteHakPrice = 1)
    Label12.Visible = (bteHakPrice = 1)
    
    baru = True
    Header
    staBaru = ""
    LblRec = ""
    
    Call comboBCtype
    '##Tampilkan Combo Customer code dari trade_master
    SQLT = "Select rtrim(trade_code) as TC,Trade_name as TN, Address1 as A from trade_master where trade_cls in ('2', '3') order by trade_code"
    
    Set rst = New Recordset
    rst.Open SQLT, Db, adOpenKeyset, adLockOptimistic
    CboSupp.clear
    CboSupp.columnCount = 2
    CboSupp.TextColumn = 1
    ir = 0
    While Not rst.EOF
        CboSupp.AddItem ""
        CboSupp.List(ir, 0) = rst!TC
        CboSupp.List(ir, 1) = Trim$(rst!TN)
        ir = ir + 1
        rst.MoveNext
    Wend
    CboSupp.ColumnWidths = "60 pt; 300 pt"
    CboSupp.ListWidth = 360
    CboSupp.ListRows = 15
    
    '##Tampilkan Warehouse code dari warehouse_master
    SqlW = "Select rtrim(wh_code) as WC,wh_name as WN from warehouse_master order by wh_code"
    
    Set RsW = New Recordset
    RsW.Open SqlW, Db, adOpenKeyset, adLockOptimistic
    CboWHCode.clear
    CboWHCode.columnCount = 2
    CboWHCode.TextColumn = 1
    ir = 0
    While Not RsW.EOF
        CboWHCode.AddItem ""
        CboWHCode.List(ir, 0) = RsW!wC
        CboWHCode.List(ir, 1) = Trim$(RsW!wn)
        ir = ir + 1
        RsW.MoveNext
    Wend
    CboWHCode.ColumnWidths = "60 pt; 180 pt"
    CboWHCode.ListWidth = 240
    CboWHCode.ListRows = 15
    
    '## receipt Combo
    CboRecCls.clear
    CboRecCls.columnCount = 2
    CboRecCls.TextColumn = 1
    CboRecCls.AddItem ""
    CboRecCls.List(0, 0) = "R"
    CboRecCls.List(0, 1) = "Receipt"
    CboRecCls.AddItem ""
    CboRecCls.List(1, 0) = "R1"
    CboRecCls.List(1, 1) = "Return"
    CboRecCls.ColumnWidths = "30 pt; 60 pt"
    CboRecCls.ListWidth = 90
    CboRecCls.ListRows = 4
    
    '##Tampilkan Combo Item code dari item_master
    SqlI = "Select rtrim(item_code) as icd,item_name as iNm,makeritem_code,rtrim(wh_code) as WHCode,rtrim(Address) as address from item_master where use_endday > convert(char(8), getdate(), 112) order by item_code"
    
    Set RsI = New Recordset
    RsI.Open SqlI, Db, adOpenKeyset, adLockOptimistic
    CboItem.clear
    CboItem.columnCount = 2
    CboItem.TextColumn = 1
    ir = 0
    While Not RsI.EOF
        CboItem.AddItem ""
        CboItem.List(ir, 0) = RsI!ICd
        CboItem.List(ir, 1) = Trim$(RsI!inm)
        ir = ir + 1
        RsI.MoveNext
    Wend
    CboItem.ColumnWidths = "130 pt; 250 pt"
    CboItem.ListWidth = 380
    CboItem.ListRows = 15
    
    '## Unit Combo
    Call up_FillCombo(Cbounit, "unit_cls")
    Cbounit.TextColumn = 2
    '## Curr Combo
    Call up_FillCombo(cbocurr, "curr_cls")
    cbocurr.TextColumn = 2
    TglReceipt = Now()
    MRecDate = Now()
    DtBCDate = Now()
    
    CboRecCls = "R"
    dateUp = MRecDate.Value
    Call up_FillCombo(cboTransport, "Transport_Cls")
    With cboTransport
        .ColumnWidths = "30pt,90pt"
        .ListWidth = 120
    End With
    Call up_FillCombo(cboPackage, "Package_Cls")
    With cboPackage
        .ColumnWidths = "60pt,90pt"
        .ListWidth = 150
    End With
End Sub

Sub display()
    Dim ig As Long
    
    LblErr = ""
    
    '## Part Detil
'    SqlPD = "select PR.Seq_No as SeqNo,PR.Supplier_code,PR.PO_no,PR.SuratJalan_no, PR.BC40_No, PR.BC40_Date, PR.BC_type, " & _
'        "PR.Transport_Cls, PR.Package_Qty, PR.Package_Cls, " & _
'        "PR.Receipt_date,PR.item_code,IM.makeritem_code,IM.PRovision_Cls,PR.warehouse_code,PR.Address," & _
'        "PR.Receipt_cls, PR.price,PR.qty,PR.Amount,PR.Unit_cls,PR.Currency_code,remarks " & _
'        "from part_receipt PR,item_master IM  where PR.item_code=IM.item_code"
'
'    SqlPD = SqlPD & " and PR.Supplier_code='" & CboSupp & "' and (PO_No='0' or po_no is null) and " & _
'        "(ProductionResult_cls='0' or ProductionResult_cls is null) and (year(Receipt_date)='" & MRecDate.Year & "' and month(receipt_date)='" & MRecDate.Month & "')"
    
    SqlPD = " SELECT  PR.Seq_No AS SeqNo , " & vbCrLf & _
                    "         PR.Supplier_Code , " & vbCrLf & _
                    "         PR.PO_No , " & vbCrLf & _
                    "         PR.SuratJalan_No , " & vbCrLf & _
                    "         PR.BC40_No , " & vbCrLf & _
                    "         PR.BC40_Date , " & vbCrLf & _
                    "         PR.BC_Type , " & vbCrLf & _
                    "         PR.Transport_Cls , " & vbCrLf & _
                    "         PR.Package_Qty , " & vbCrLf & _
                    "         PR.Package_Cls , " & vbCrLf & _
                    "         PR.Receipt_Date , "
    
    SqlPD = SqlPD + "         PR.Item_Code , " & vbCrLf & _
                    "         IM.MakerItem_Code , " & vbCrLf & _
                    "         IM.Provision_Cls , " & vbCrLf & _
                    "         PR.Warehouse_Code , " & vbCrLf & _
                    "         PR.Address , " & vbCrLf & _
                    "         PR.Receipt_Cls , " & vbCrLf & _
                    "         PR.Price , " & vbCrLf & _
                    "         PR.Qty , " & vbCrLf & _
                    "         PR.Amount , " & vbCrLf & _
                    "         PR.Unit_Cls , " & vbCrLf & _
                    "         PR.Currency_Code , "
    
    SqlPD = SqlPD + "         Remarks , " & vbCrLf & _
                    "         ISNULL(PR.Receipt_Status, '02') Receipt_Status , " & vbCrLf & _
                    "         ISNULL(PR.No_Register, '') No_Register, " & vbCrLf & _
                    "         ISNULL(PR.No_Seri, '') No_Seri " & vbCrLf & _
                    " FROM    Part_Receipt PR , " & vbCrLf & _
                    "         Item_Master IM " & vbCrLf & _
                    " WHERE   PR.Item_Code = IM.Item_Code " & vbCrLf & _
                    "         AND PR.Supplier_Code = '" & CboSupp & "' " & vbCrLf & _
                    "         AND ( PO_No = '0' " & vbCrLf & _
                    "               OR PO_No IS NULL " & vbCrLf & _
                    "             ) " & vbCrLf & _
                    "         AND ( ProductionResult_Cls = '0' " & vbCrLf & _
                    "               OR ProductionResult_Cls IS NULL "
    
    SqlPD = SqlPD + "             ) " & vbCrLf & _
                    "         AND ( YEAR(Receipt_Date) = '" & MRecDate.Year & "' " & vbCrLf & _
                    "               AND MONTH(Receipt_Date) = '" & MRecDate.Month & "' " & vbCrLf & _
                    "             ) "
                
    Set RsPD = New Recordset
    If RsPD.State = adStateOpen Then RsPD.Close
    RsPD.Open SqlPD, Db, adOpenDynamic, adLockOptimistic
    
    With Grid
        ig = 0
        Header
        While Not RsPD.EOF
            ig = ig + 1
            .AddItem ""
            .TextMatrix(ig, bteColSJNo) = Trim$(RsPD!SuratJalan_No)
            .Cell(flexcpAlignment, ig, bteColSJNo) = flexAlignLeftCenter
            .TextMatrix(ig, bteColReceiptDate) = Format(RsPD!Receipt_Date, "dd mmm yyyy")
            .Cell(flexcpAlignment, ig, bteColReceiptDate) = flexAlignLeftCenter
            .TextMatrix(ig, bteColProdCode) = Trim$(RsPD!Item_Code)
            .Cell(flexcpAlignment, ig, bteColProdCode) = flexAlignLeftCenter
            .TextMatrix(ig, bteColPartNo) = Trim$(RsPD!MakerItem_Code)
            .Cell(flexcpAlignment, ig, bteColPartNo) = flexAlignLeftCenter
            .TextMatrix(ig, bteColDesc) = Trim$(uf_GetItemDescription(Trim$(RsPD!Item_Code)))
            .Cell(flexcpAlignment, ig, bteColDesc) = flexAlignLeftCenter
            .TextMatrix(ig, bteColWHCode) = Trim$(RsPD!Warehouse_Code)
            .Cell(flexcpAlignment, ig, bteColWHCode) = flexAlignLeftCenter
            .TextMatrix(ig, bteColAddress) = Trim$(RsPD!Address)
            .Cell(flexcpAlignment, ig, bteColAddress) = flexAlignLeftCenter
            .TextMatrix(ig, bteColReceipt) = Trim$(RsPD!receipt_cls)
            .Cell(flexcpAlignment, ig, bteColReceipt) = flexAlignLeftCenter
            .TextMatrix(ig, bteColBctype) = Trim(RsPD!BC_Type & "")
            .Cell(flexcpAlignment, ig, bteColBctype) = flexAlignLeftCenter
            
            If InStr(1, CDbl(RsPD!Qty), ".") = 0 Then
                .TextMatrix(ig, bteColQty) = Format(RsPD!Qty, gs_formatQty)
            ElseIf InStr(1, CDbl(RsPD!Qty), ".") > 0 Then
                If Split(CDbl(RsPD!Qty), ".")(1) <> "" Then .TextMatrix(ig, bteColQty) = Format(RsPD!Qty, gs_formatQty)
            End If
            .Cell(flexcpAlignment, ig, bteColQty) = flexAlignRightCenter
    
            .TextMatrix(ig, bteColUnit) = uf_GetUnitDescription(Trim(Trim$(RsPD!Unit_cls)))
            .Cell(flexcpAlignment, ig, bteColUnit) = flexAlignLeftCenter
            .TextMatrix(ig, bteColCurr) = uf_GetCurrencyDescription(Trim(Trim$(RsPD!currency_code)))
            .Cell(flexcpAlignment, ig, bteColCurr) = flexAlignLeftCenter
            .TextMatrix(ig, bteColPrice) = Format(Trim(RsPD!Price), gs_formatPrice)
            .Cell(flexcpAlignment, ig, bteColPrice) = flexAlignRightCenter
                
            If InStr(1, CDbl(RsPD!Amount), ".") = 0 Then
                .TextMatrix(ig, bteColAmount) = Format(RsPD!Amount, gs_formatAmount)
            ElseIf InStr(1, CDbl(RsPD!Amount), ".") > 0 Then
                If Split(CDbl(RsPD!Amount), ".")(1) <> "" Then .TextMatrix(ig, bteColAmount) = Format(RsPD!Amount, gs_formatAmount)
            End If
            
            .Cell(flexcpAlignment, ig, bteColAmount) = flexAlignRightCenter
            
            .TextMatrix(ig, bteColUnitCls) = Trim$(RsPD!Unit_cls)
            .TextMatrix(ig, bteColCurrCode) = Trim$(RsPD!currency_code)
            .TextMatrix(ig, bteColRemark) = Trim$(RsPD!Remarks)
            .TextMatrix(ig, bteColSeqNo) = RsPD!seqNo
            .TextMatrix(ig, bteColBC) = Trim$(RsPD!BC40_No & "")
            .TextMatrix(ig, bteColBCDate) = Format(RsPD!BC40_Date, "dd mmm yyyy")
           ' .TextMatrix(ig, bteColBctype) = Trim(RsPD!BC_Type)
            .TextMatrix(ig, bteColTransport) = Trim$(RsPD!Transport_Cls & "")
            .TextMatrix(ig, bteColPackageQty) = Val(RsPD!Package_Qty & "")
            .TextMatrix(ig, bteColPackage) = Trim$(RsPD!Package_Cls & "")
            .TextMatrix(ig, bteColRecStatus) = Trim(RsPD!Receipt_Status)
            .Cell(flexcpAlignment, ig, bteColRecStatus) = flexAlignCenterCenter
            .TextMatrix(ig, bteColNoRegister) = Trim(RsPD!No_Register)
            .Cell(flexcpAlignment, ig, bteColNoRegister) = flexAlignLeftCenter
            .TextMatrix(ig, bteColNoSeri) = Trim(RsPD!No_Seri)
            .Cell(flexcpAlignment, ig, bteColNoSeri) = flexAlignRightCenter
            
            .Cell(flexcpBackColor, ig, bteColSelect) = vbWhite
            RsPD.MoveNext
        Wend
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If newDb.State <> adStateClosed Then newDb.Close
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim StrGrid As String
    Dim AdaS As Boolean, brs As Integer, id As Long
    
    StrGrid = Grid.Text
    AdaS = False
    brs = 0
    If StrGrid = "S" Then
    
        For id = 1 To Grid.Rows - 1
            If id <> Row Then Grid.TextMatrix(id, bteColSelect) = ""
        Next id
    
        CboItem = Grid.TextMatrix(Grid.Row, bteColProdCode)
        TxtSj = Grid.TextMatrix(Grid.Row, bteColSJNo)
        If UCase(Trim(Grid.TextMatrix(Grid.Row, bteColReceipt))) = "R" Then
            CboRecCls = "R"
        ElseIf UCase(Trim(Grid.TextMatrix(Grid.Row, bteColReceipt))) = "R1" Then
            CboRecCls = "R1"
        End If
        
        DtBCDate = IIf(Grid.TextMatrix(Grid.Row, bteColBCDate) = "", Grid.TextMatrix(Grid.Row, bteColReceiptDate), Grid.TextMatrix(Grid.Row, bteColBCDate))
        
        TglReceipt = Grid.TextMatrix(Grid.Row, bteColReceiptDate)
        CboWHCode = Grid.TextMatrix(Grid.Row, bteColWHCode)
        
        If IsNull(Grid.TextMatrix(Grid.Row, bteColQty)) Then
            txtqty = 0
        Else
            If InStr(1, CDbl(Grid.TextMatrix(Grid.Row, bteColQty)), ".") = 0 Then
                txtqty = Format(Grid.TextMatrix(Grid.Row, bteColQty), gs_formatQty)
            ElseIf InStr(1, CDbl(Grid.TextMatrix(Grid.Row, bteColQty)), ".") > 0 Then
                If Split(CDbl(Grid.TextMatrix(Grid.Row, bteColQty)), ".")(1) <> "" Then txtqty = Format(Grid.TextMatrix(Grid.Row, bteColQty), gs_formatQty)
            End If
        End If
    
        Cbounit = uf_GetUnitDescription(Trim$(Grid.TextMatrix(Grid.Row, bteColUnitCls)))
        cbocurr.Text = uf_GetCurrencyDescription(Trim$(Grid.TextMatrix(Grid.Row, bteColCurrCode)))
    
        If IsNull(Grid.TextMatrix(Grid.Row, bteColPrice)) Then
            cboprice = 0
        Else
            cboprice = Trim(Grid.TextMatrix(Grid.Row, bteColPrice))
        End If
    
        If IsNull(Grid.TextMatrix(Grid.Row, bteColAmount)) Then
            txtamount = 0
        Else
            txtamount = Grid.TextMatrix(Grid.Row, bteColAmount)
        End If
    
        TxtRemarks = Trim$(Grid.TextMatrix(Grid.Row, bteColRemark))
        LblRecCls = CboRecCls
        LblRecDate = Grid.TextMatrix(Grid.Row, bteColReceiptDate)
        LblWHCode = Grid.TextMatrix(Grid.Row, bteColWHCode)
        LblItemCode = Grid.TextMatrix(Grid.Row, bteColProdCode)
        LblQty = CDbl(Grid.TextMatrix(Grid.Row, bteColQty))
        SeqID = Grid.TextMatrix(Grid.Row, bteColSeqNo)
        txtBC40 = Grid.TextMatrix(Grid.Row, bteColBC)
        CbotypeBC = Grid.TextMatrix(Grid.Row, bteColBctype)
        cboTransport = Grid.TextMatrix(Grid.Row, bteColTransport)
        txtPackage = Val(Grid.TextMatrix(Grid.Row, bteColPackageQty))
        cboPackage = Grid.TextMatrix(Grid.Row, bteColPackage)
        txtRec_Status.Text = Grid.TextMatrix(Grid.Row, bteColRecStatus)
        txtRegisterNo.Text = Grid.TextMatrix(Grid.Row, bteColNoRegister)
        txtNoSeri.Text = Grid.TextMatrix(Grid.Row, bteColNoSeri)
        
'        up_GetNoSeri (CboItem.Text)
        
        baru = False
        LblErr = ""
    ElseIf StrGrid = "D" Then
        For id = 1 To Grid.Rows - 1
            If Grid.TextMatrix(id, bteColSelect) = "S" Then Grid.TextMatrix(id, bteColSelect) = "": Exit For
        Next id
        Kosong
        baru = False
        LblErr = ""
    ElseIf StrGrid = "" Then
        Kosong
        LblErr = ""
    End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim pesanTgl As String
    With Grid
    If .Col > bteColSelect Then Cancel = True: Exit Sub
        pesanTgl = up_ValidateDateRange(Format(.TextMatrix(Row, bteColReceiptDate), "yyyy-MM-dd"), True)
        If pesanTgl <> "" Then LblErr = pesanTgl: Cancel = 1 Else LblErr = ""
    End With
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Grid.Col = bteColSelect Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then
            KeyAscii = 0
        End If
        If KeyAscii = vbKeyEscape Then KeyAscii = 0
        If KeyAscii = Asc(".") Then KeyAscii = 0
    End If
End Sub

Function NamaItem(StrItem As String) As String
    Dim rsin As Recordset
    Set rsin = Db.Execute("Select item_name from item_master where item_code='" & StrItem & "'")
    If Not rsin.EOF Then
        NamaItem = rsin!item_name
    Else
        NamaItem = ""
    End If
    rsin.Close
End Function

Function MakerItem(StrItem As String) As String
    Dim RsMI As Recordset
    Set RsMI = Db.Execute("Select makeritem_code from item_master where item_code='" & StrItem & "'")
    If Not RsMI.EOF Then
        MakerItem = RsMI!MakerItem_Code
    Else
        MakerItem = ""
    End If
    RsMI.Close
End Function


Private Sub MRecDate_Change()
If Format(MRecDate.Value, "MM") < Format(dateUp, "MM") And Val(Format(MRecDate.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then _
            MRecDate.Year = MRecDate.Year + 1: GoTo pass
    If Format(MRecDate.Value, "MM") > Format(dateUp, "MM") And Val(Format(MRecDate.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then _
            MRecDate.Year = MRecDate.Year - 1
pass:
    dateUp = Format(MRecDate.Value, "dd MMM yyyy")
End Sub

Private Sub TglReceipt_Change()
If CboItem <> "" Then
    If cbocurr.ListCount > 0 Then
        cbocurr = Trim(cbocurr)
        If cbocurr.MatchFound Then
            xcurr = cbocurr.List(0)
            blnshow = False
            If CboItem <> "" Then Call browseprice(CboItem, TglReceipt, CboSupp, xcurr)
        Else
            xcurr = ""
            LblErr = DisplayMsg(4005): cbocurr.SetFocus
        End If
    End If
Else
     LblErr = ""
End If
End Sub

Private Sub txtNoSeri_KeyPress(KeyAscii As Integer)
    ' Hanya izinkan angka (48-57) dan tombol Backspace (8)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' Batalkan input
    End If
End Sub

Private Sub txtPackage_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtPackage_LostFocus()
If IsNumeric(txtPackage) = True Then
    txtPackage.Text = Format(txtPackage.Text, gs_formatBox)
Else
    txtPackage.Text = Format(0, gs_formatBox)
End If
End Sub

Private Sub txtqty_Change()
    If InStr(1, txtqty.Text, ",") = 1 Then txtqty.Text = Mid(txtqty.Text, 2, Len(txtqty.Text))
    If InStr(1, txtqty.Text, ".") = 1 Then txtqty = ""
    If txtqty.Text <> "" Then
    End If
    If Val(cboprice) = 0 Then txtamount = 0: Exit Sub
    If Val(txtqty) = 0 Then txtamount = 0: Exit Sub
    If IsNumeric(txtqty) = True And IsNumeric(cboprice) = True Then
        If CDbl(txtqty) > 0 And CDbl(cboprice) > 0 Then
            txtamount = CDbl(txtqty) * CDbl(cboprice)
            If Round(CDbl(txtamount)) / CDbl(txtamount) = 1 Then
                txtamount = Format(CDbl(txtamount), gs_formatAmount)
            Else
                txtamount = Format(CDbl(txtamount), gs_formatAmount)
            End If
        Else
            txtamount = "0"
        End If
    Else
        txtamount = "0"
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Sub DataGrid()
    Dim kode As String, Sta As String
    Dim strSQL As String
    Dim PosS As Integer
    
    PosS = Grid.FindRow("S", , bteColSelect, False)
    If PosS > 0 Then
        Db.BeginTrans
        kode = Trim$(Grid.TextMatrix(PosS, 1))
    
        strSQL = "update part_receipt " & _
            "set Supplier_code='" & Trim$(CboSupp.Text) & "', " & _
            "PO_no='0', " & _
            "warehouse_code='" & Trim$(CboWHCode) & "', " & _
            "receipt_cls='" & Trim$(CboRecCls) & "', " & _
            "receipt_date='" & Format(TglReceipt, "yyyy-mm-dd") & "', " & _
            "item_code='" & Trim$(CboItem) & "', " & _
            "qty=" & CDbl(txtqty) & ", " & _
            "Unit_cls='" & Trim$(Cbounit.List(Cbounit.ListIndex, 0)) & "', " & _
            "currency_code='" & Trim$(cbocurr.List(cbocurr.ListIndex, 0)) & "', " & _
            "BC40_Date='" & Format(DtBCDate, "yyyy-mm-dd") & "', " & _
            "BC40_No='" & Trim$(txtBC40) & "', " & _
            "BC_type = '" & Trim(CbotypeBC.Text) & "', "

        If Val(cboprice) = 0 Then
            strSQL = strSQL & "price=0, amount=0, "
        ElseIf Val(cboprice) > 0 Then
            strSQL = strSQL & "price=" & CDbl(cboprice) & ", amount=" & Trim$(CDbl(txtamount)) & ", "
        End If
        
        strSQL = strSQL & " suratjalan_no='" & Trim$(TxtSj) & "', Remarks='" & Trim$(TxtRemarks) & "', No_Register='" & Trim$(txtRegisterNo) & "', No_Seri='" & Trim$(txtNoSeri) & "',  " & _
            "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
            "where Seq_No='" & SeqID & "'"
        
        Db.Execute (strSQL)
        
        If err.number <> 0 Then
            StaErr = True
            Db.RollbackTrans
            err.clear
        Else
            StaErr = False
            '*********** Proses Update ke Supply **************
            Dim rsProv As New ADODB.Recordset
            KeyProd = "R" & SeqID
            Call seqNo.HapusDataSupp(Db, "'" & KeyProd & "'", blnFix, thnFix)
            sql = "Select Provision_Cls from Item_MaSter where ITem_Code = '" & CboItem & "'"
            Set rsProv = Db.Execute(sql)
            If Not rsProv.EOF Then provisionCls = Trim(rsProv(0)) Else provisionCls = ""
            If provisionCls = "01" Then Call inputSupply(CboItem, CDbl(txtqty))
            '*************************************************
        
            Db.CommitTrans
            
            If txtRec_Status.Text = "01" Then
                If Format(LblRecDate, "YYYYMMDD") <> Format(TglReceipt, "YYYYMMDD") Then
                    If Trim$(CboItem) <> Trim$(LblItemCode) Then
                        If Trim$(LblWHCode) <> Trim$(CboWHCode.List(CboWHCode.ListIndex, 0)) Then
                            ProsesStock 2, Trim$(LblItemCode), Trim$(LblWHCode), Trim$(LblWHCode), Format(LblRecDate, "YYYYMM"), 0, CDbl(LblQty) 'Update yg lama
                            ProsesStock 2, Trim$(CboItem), Trim$(CboWHCode.List(CboWHCode.ListIndex, 0)), Trim$(CboWHCode.List(CboWHCode.ListIndex, 0)), Format(TglReceipt, "YYYYMM"), CDbl(txtqty), 0 'Update yg baru
                        Else
                            ProsesStock 2, Trim$(LblItemCode), Trim$(LblWHCode), Trim$(LblWHCode), Format(LblRecDate, "YYYYMM"), 0, CDbl(LblQty) 'Update yg lama
                            ProsesStock 2, Trim$(CboItem), Trim$(LblWHCode), Trim$(LblWHCode), Format(TglReceipt, "YYYYMM"), CDbl(txtqty), 0 'Update yg baru
                        End If
                        '##
                    Else
                        If Trim$(LblWHCode) <> Trim$(CboWHCode.List(CboWHCode.ListIndex, 0)) Then
                            ProsesStock 2, Trim$(LblItemCode), Trim$(LblWHCode), Trim$(LblWHCode), Format(LblRecDate, "YYYYMM"), 0, CDbl(LblQty) 'Update yg lama
                            ProsesStock 2, Trim$(LblItemCode), Trim$(CboWHCode.List(CboWHCode.ListIndex, 0)), Trim$(CboWHCode.List(CboWHCode.ListIndex, 0)), Format(TglReceipt, "YYYYMM"), CDbl(txtqty), 0 'Update yg baru
                        Else
                            ProsesStock 2, Trim$(LblItemCode), Trim$(LblWHCode), Trim$(LblWHCode), Format(LblRecDate, "YYYYMM"), 0, CDbl(LblQty) 'Update yg lama
                            ProsesStock 2, Trim$(LblItemCode), Trim$(LblWHCode), Trim$(LblWHCode), Format(TglReceipt, "YYYYMM"), CDbl(txtqty), 0 'Update yg baru
                        End If
                    End If
                Else
                    If Trim$(CboItem) <> Trim$(LblItemCode) Then
                        If Trim$(LblWHCode) <> Trim$(CboWHCode.List(CboWHCode.ListIndex, 0)) Then
                            ProsesStock 2, Trim$(LblItemCode), Trim$(LblWHCode), Trim$(LblWHCode), Format(LblRecDate, "YYYYMM"), 0, CDbl(LblQty) 'Update yg lama
                            ProsesStock 2, Trim$(CboItem), Trim$(CboWHCode.List(CboWHCode.ListIndex, 0)), Trim$(CboWHCode.List(CboWHCode.ListIndex, 0)), Format(LblRecDate, "YYYYMM"), CDbl(txtqty), 0 'Update yg baru
                        Else
                            ProsesStock 2, Trim$(LblItemCode), Trim$(LblWHCode), Trim$(LblWHCode), Format(LblRecDate, "YYYYMM"), 0, CDbl(LblQty) 'Update yg lama
                            ProsesStock 2, Trim$(CboItem), Trim$(LblWHCode), Trim$(LblWHCode), Format(LblRecDate, "YYYYMM"), CDbl(txtqty), 0 'Update yg baru
                        End If
                        '##
                    Else
                        If Trim$(LblWHCode) <> Trim$(CboWHCode.List(CboWHCode.ListIndex, 0)) Then
                            ProsesStock 2, Trim$(LblItemCode), Trim$(LblWHCode), Trim$(LblWHCode), Format(LblRecDate, "YYYYMM"), 0, CDbl(LblQty) 'Update yg lama
                            ProsesStock 2, Trim$(LblItemCode), Trim$(CboWHCode.List(CboWHCode.ListIndex, 0)), Trim$(CboWHCode.List(CboWHCode.ListIndex, 0)), Format(TglReceipt, "YYYYMM"), CDbl(txtqty), 0 'Update yg baru
                        Else
                            ProsesStock 2, Trim$(CboItem), Trim$(LblWHCode), Trim$(CboWHCode), Format(TglReceipt, "YYYYMM"), CDbl(txtqty), CDbl(LblQty)
                        End If
                    End If
                End If
            End If
            
        End If
        Exit Sub
    End If
    
    nErr = 0
    nOK = 0
    StrWDel = ""
    
    Db.BeginTrans
    Dim ix As Long
    Dim strKey As String
    
    strKey = ""
    For ix = 1 To Grid.Rows - 1
        kode = Trim$(Grid.TextMatrix(ix, bteColQty))
        If UCase(Grid.TextMatrix(ix, bteColReceipt)) = "R" Or UCase(Grid.TextMatrix(ix, bteColReceipt)) = "R" Then
            LblRecCls = "R"
        ElseIf UCase(Grid.TextMatrix(ix, bteColReceipt)) = "R1" Or UCase(Grid.TextMatrix(ix, bteColReceipt)) = "R1" Then
            LblRecCls = "R1"
        End If
        LblQty = Grid.TextMatrix(ix, bteColQty)
        LblRecDate = Grid.TextMatrix(ix, bteColReceiptDate)
        LblWHCode = Grid.TextMatrix(ix, bteColWHCode)
        LblItemCode = Grid.TextMatrix(ix, bteColProdCode)
        Sta = Trim$(Grid.TextMatrix(ix, bteColSelect))
    
        If Sta = "D" Then
    
            strSQL = "delete from  part_receipt where Seq_No='" & Grid.TextMatrix(ix, bteColSeqNo) & "'"
            If strSQL <> "" Then Db.Execute strSQL
        
            If err.number <> 0 Then
                StrWDel = StrWDel & kode & ","
                nErr = nErr + 1
                err.clear
            Else
                If txtRec_Status.Text = "01" Then
                    nOK = nOK + 1
                    ProsesStock 3, Trim$(LblItemCode), Trim$(LblWHCode), Trim$(LblWHCode), Format(LblRecDate, "YYYYMM"), CDbl(LblQty), ""
                End If
            End If
            strKey = strKey & "'R" & Grid.TextMatrix(ix, bteColSeqNo) & "',"
        End If
        strSQL = ""
    Next ix
    
    If Len(StrWDel) > 1 Then StrWDel = Mid(StrWDel, 1, Len(StrWDel) - 1)
    Call seqNo.HapusDataSupp(Db, Left(strKey, Len(strKey) - 1), blnFix, thnFix)
    
    If nErr > 0 Then
        StaErr = True
    End If
    
    If err.number = 0 Then
        Db.CommitTrans
    Else
        Db.RollbackTrans
        err.clear
    End If
    kode = ""
    Sta = ""
    strSQL = ""
    display
End Sub

Sub Kosong()
    CboItem = ""
    LblItemName = ""
    CboWHCode = ""
    Lbladdress = ""
    Cbounit = ""
    cbocurr = ""
    cboprice = ""
    txtqty = ""
    txtamount = ""
    TxtRemarks = ""
    LblWHCode = ""
    LblItemCode = ""
    LblRecCls = ""
    LblRecDate = ""
    txtNoSeri.Text = ""
'    txtBC40 = ""
'    CbotypeBC = ""
'    TxtSj = ""
End Sub

Function cek() As Boolean
    cek = True
    
    If Trim$(CboSupp) = "" Then
        LblErr = DisplayMsg(1054) '""Please input Supplier Code"
        CboSupp.SetFocus: cek = False: Exit Function
    End If
    CboSupp = CboSupp
    If CboSupp.MatchFound = False Then
        LblErr = DisplayMsg(4050) '"Supplier Code is invalid ! "
        CboSupp.SetFocus: cek = False: Exit Function
    End If
    
    If Trim(TxtSj) = "" Then
        LblErr = DisplayMsg(1036) '"Please input Surat Jalan No"
        TxtSj.SetFocus:  cek = False: Exit Function
    End If
    
    If Trim(CboRecCls) = "" Then
        LblErr = DisplayMsg(8088) '"Please input Receipt Code"
        CboRecCls.SetFocus: cek = False: Exit Function
    End If
    CboRecCls = CboRecCls
    If CboRecCls.MatchFound = False Then
        LblErr = DisplayMsg(8088) '"Receipt is invalid !"
        CboRecCls.SetFocus: cek = False: Exit Function
    End If
    
    If Trim$(CboItem) = "" Then
        LblErr = DisplayMsg(1009) '"Please input Product Code"
        CboItem.SetFocus: cek = False: Exit Function
    End If
    CboItem = CboItem
    If CboItem.MatchFound = False Then
        '**** cek Data cocok / tdk dgn Database
        Dim rsDB As New ADODB.Recordset
        rsDB.Open "select Item_Code from Item_Master where Item_Code='" & Trim(CboItem) & "'", Db, adOpenKeyset, adLockOptimistic
        If rsDB.EOF = True Then
            LblErr = DisplayMsg(4061) '"Product Code is invalid! "
            CboItem.SetFocus: cek = False: Exit Function
        End If
        rsDB.Close
    End If
    
    If Trim$(CboWHCode) & "" = "" Then LblErr = DisplayMsg(1042): CboWHCode.SetFocus: cek = False: Exit Function
    CboWHCode = CboWHCode
    If CboWHCode.MatchFound = False Then LblErr = DisplayMsg(4023): CboWHCode.SetFocus:  cek = False: Exit Function
    
    If Trim(txtqty) = "" Or IsNumeric(txtqty) = False Then
        LblErr = DisplayMsg(4044) '"Please input Quantity"
        txtqty.SetFocus: cek = False: Exit Function
    End If
    
    If CDbl(txtqty) > gd_MaxQty Then
        LblErr = DisplayMsg(4045) & " " & gd_MaxQty
        txtqty.SetFocus: cek = False: Exit Function
    End If
    
    If Trim$(Cbounit & "") = "" Then LblErr = DisplayMsg(1030):  cek = False: Exit Function 'Cbounit.SetFocus
    Cbounit = Cbounit
    If Cbounit.MatchFound = False Then LblErr = DisplayMsg(1030):  cek = False: Exit Function 'Cbounit.SetFocus
    
    If Trim$(cbocurr & "") = "" Then
        If bteHakPrice = 0 Then
            cbocurr = uf_GetCurrencyDescription(gs_DefaultCurrencyCode)
        Else
            LblErr = DisplayMsg(1011)
            cbocurr.SetFocus
            cek = False
            Exit Function
        End If
    End If
    cbocurr = cbocurr
    If cbocurr.MatchFound = False Then LblErr = DisplayMsg(1028): cbocurr.SetFocus: cek = False: Exit Function
    
    If Trim(cboprice) = "" Then
        If bteHakPrice = 0 Then
            cboprice = 0
        Else
            If Trim(cboprice) = "" Or IsNumeric(cboprice) = False Then
                LblErr = DisplayMsg(8023)
                cboprice.SetFocus
                cek = False
                Exit Function
            End If
            If CDbl(cboprice) > gd_MaxPrice Then
                LblErr = DisplayMsg(4048) & " " & gd_MaxPrice
                cboprice.SetFocus
                cek = False
                Exit Function
            End If
            
            If CDbl(cboprice) * CDbl(txtqty) > gd_MaxAmount Then
                LblErr = DisplayMsg(4051) & " " & gd_MaxAmount
                cboprice.SetFocus
                cek = False
                Exit Function
            End If
        End If
    End If
    
    If txtNoSeri.Text = "" Then LblErr = DisplayMsg("0001") & " No Seri ! ":  txtNoSeri.SetFocus: cek = False: Exit Function
    
    LblErr = up_ValidateSuratJalan()
    If validate <> True Then
        LblErr = DisplayMsg(9013)
        cek = False
        Exit Function
    End If
    
    LblErr = up_ValidateDateRange(Format(TglReceipt, "yyyy-MM-dd"), True)
    If LblErr <> "" Then
        cek = False
        Exit Function
    End If
    
    LblErr = uf_ValidClosingReceipt(Format(TglReceipt, "yyyy-MM-dd"))
    If LblErr <> "" Then
        cek = False
        Exit Function
    End If
    
    
End Function

Private Sub txtQty_LostFocus()
If IsNumeric(txtqty) = True Then
    txtqty.Text = Format(txtqty.Text, gs_formatQty)
Else
    txtqty.Text = Format(0, gs_formatQty)
End If
End Sub

Sub ProsesStock(nTipe As Byte, ItemCode As String, OldWHCode As String, NewWHCode As String, RecDate As String, QtyX As String, OldQty As String)
    '## nTipe=1 ----> Insert
    '## nTipe=2 ----> Update
    '## nTipe=3 ----> Delete
    Dim Sqlc As String, RsInvControl As New ADODB.Recordset
    Dim RsWHS As Recordset, RSIS As Recordset
    Dim WHStock_ctrl As Boolean
    Dim ItemStock_ctrl As Boolean
    
    ItemStock_ctrl = False
    WHStock_ctrl = False
    QtyX = Format(QtyX, gs_formatQty)
    OldQty = Format(OldQty, gs_formatQty)
    
    Dim FlagU As Byte
    
    'Proses Cek nya adalah dari warehouse dulu baru cek item nya
    Select Case nTipe
    'If nTipe = 2 Then
    Case 1 'insert
    
        Set RsWHS = Db.Execute("Select stockcontrol_cls,WH_name from warehouse_master where WH_code='" & Trim$(OldWHCode) & "'")
        If Not RsWHS.EOF Then
            If RsWHS!stockcontrol_cls <> "01" Then
                'Stock tidak boleh di update
                Exit Sub
            Else
                'Stock boleh di update
                'Cek apakah item boleh  update stock atau tidak
                Set RSIS = Db.Execute("Select StockControl_cls from item_master where item_code='" & Trim$(ItemCode) & "'") ' and (stockcontrol_cls='01')")
                If Not RSIS.EOF Then
                    If Trim$(RSIS!stockcontrol_cls) <> "01" Then
                        'Stock tidak boleh di update
                        Exit Sub
                    End If
                    '## Update Stock Master
                    If Trim$(CboRecCls) = "R" Then
                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(QtyX), 0
                    Else
                        Dim xz As Long
                        Dim ZZZ As Integer
                        Db.Execute "select * from stock_master where warehouse_code = '" & NewWHCode & "' and item_code = '" & ItemCode & "'", xz
                        If xz <> 0 Then
                            updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), 0, CDbl(QtyX)
                        Else
                            updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(-(QtyX)), 0
                        End If
                    End If
                    ItemStock_ctrl = True
                Else
                    ItemStock_ctrl = False
                    Exit Sub
                End If
                RSIS.Close
                Set RSIS = Nothing
            End If
            WHStock_ctrl = True
        Else
            WHStock_ctrl = False
        End If
        RsWHS.Close
        Set RsWHS = Nothing
    
    Case 2 'Update
        
        'Jika Update Cek Apakah WH_code nya berubah
        'Jika iya maka cek masing2 WH Code
        'Jika kedua WH Code tsb boleh update stock maka proses kedua WH COde tsb
        If OldWHCode <> NewWHCode Then 'WH_code berubah
            'Cek Apakah OldWH_code yg dipilih StockControl_cls nya ='01'
            Set RsWHS = Db.Execute("Select stockcontrol_cls,WH_name from warehouse_master where WH_code='" & Trim$(OldWHCode) & "'")
            If Not RsWHS.EOF Then
                If RsWHS!stockcontrol_cls <> "01" Then
                    'Stock tidak boleh di update
                    Exit Sub
                Else
                    'Stock boleh di update
                    'Cek apakah item boleh  update stock atau tidak
                    Set RSIS = Db.Execute("Select StockControl_cls from item_master where item_code='" & Trim$(ItemCode) & "'") ' and (stockcontrol_cls='01')")
                    If Not RSIS.EOF Then
                        If Trim$(RSIS!stockcontrol_cls) <> "01" Then
                            'Stock tidak boleh di update
                            Exit Sub
                        End If
                        '## Update Stock master
                        'updateStock Trim$(OldWHCode), itemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(OldQty), 0
                        If Trim$(LblRecCls) = "R" Then
                            If Trim$(CboRecCls) = "R" Then
                                updateStock Trim$(OldWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), 0, CDbl(OldQty)
                            Else
                                updateStock Trim$(OldWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), 0, CDbl(OldQty)
                            End If
                        Else
                            If Trim$(CboRecCls) = "R" Then
                                updateStock Trim$(OldWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), 0, (0 - CDbl(OldQty))
                            Else
                                updateStock Trim$(OldWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), 0, (0 - CDbl(OldQty))
                            End If
                        End If
                        ItemStock_ctrl = True
                    Else
                        ItemStock_ctrl = False
                        Exit Sub
                    End If
                    RSIS.Close
                    Set RSIS = Nothing
                End If
                WHStock_ctrl = True
            Else
                WHStock_ctrl = False
            End If
            RsWHS.Close
            Set RsWHS = Nothing
        End If
        
        Set RsWHS = Db.Execute("Select stockcontrol_cls,WH_name from warehouse_master where WH_code='" & Trim$(OldWHCode) & "'")
        If Not RsWHS.EOF Then
            If RsWHS!stockcontrol_cls <> "01" Then
                'Stock tidak boleh di update
                Exit Sub
            Else
                'Stock boleh di update
                'Cek apakah item boleh  update stock atau tidak
                Set RSIS = Db.Execute("Select StockControl_cls from item_master where item_code='" & Trim$(ItemCode) & "'") ' and (stockcontrol_cls='01')")
                If Not RSIS.EOF Then
                    If Trim$(RSIS!stockcontrol_cls) <> "01" Then
                    'Stock tidak boleh di update
                    Exit Sub
                End If
                '## Update Stock Master
                If Trim$(LblRecCls) = "R" Then
                    If Trim$(CboRecCls) = "R" Then
                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(QtyX), CDbl(OldQty)
                    Else
                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), (0 - CDbl(OldQty)), 0
                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), (0 - CDbl(QtyX)), 0
                    End If
                Else
                    If Trim$(CboRecCls) = "R" Then
                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(OldQty), 0
                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(QtyX), 0
                    Else
                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(OldQty), CDbl(QtyX)
                    End If
                End If
                    ItemStock_ctrl = True
                Else
                    ItemStock_ctrl = False
                    Exit Sub
                End If
                RSIS.Close
                Set RSIS = Nothing
            End If
            WHStock_ctrl = True
        Else
            WHStock_ctrl = False
        End If
        RsWHS.Close
        Set RsWHS = Nothing
    
    Case 3 'Delete
    
        Set RsWHS = Db.Execute("Select stockcontrol_cls,WH_name from warehouse_master where WH_code='" & Trim$(OldWHCode) & "'")
        If Not RsWHS.EOF Then
            If RsWHS!stockcontrol_cls <> "01" Then
                'Stock tidak boleh di update
                Exit Sub
            Else
                'Stock boleh di update
                'Cek apakah item boleh  update stock atau tidak
                Set RSIS = Db.Execute("Select StockControl_cls from item_master where item_code='" & Trim$(ItemCode) & "'") ' and (stockcontrol_cls='01')")
                If Not RSIS.EOF Then
                    If Trim$(RSIS!stockcontrol_cls) <> "01" Then
                        'Stock tidak boleh di update
                        Exit Sub
                    End If
                    '## Update Stock Master /Delete
                    If LblRecCls = "R1" Then
                        DeleteStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(-(QtyX))
                    Else
                        DeleteStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(QtyX)
                    End If
                    ItemStock_ctrl = True
                Else
                    ItemStock_ctrl = False
                    Exit Sub
                End If
                RSIS.Close
                Set RSIS = Nothing
            End If
            WHStock_ctrl = True
        Else
            WHStock_ctrl = False
        End If
        RsWHS.Close
        Set RsWHS = Nothing
    
    End Select
End Sub

Sub updateStock(WHCode As String, ItemCode As String, RecMonth As String, RecYear As String, QtyX As Double, OldQty As Double)
    Dim RsI As Recordset
    Dim DBi As New Connection
    Dim FixMonth As String
    Dim FixYear As String
    
    Dim LM_P As Double, LM_S As Double, LM_RJ As Double, LM_I As Double
    Dim TM_P As Double, TM_S As Double, TM_RJ As Double, TM_I As Double
    Dim NM_P As Double, NM_S As Double, NM_RJ As Double, NM_I As Double
    
    DBi.ConnectionTimeout = 0
    DBi.Open Db.ConnectionString
    DBi.BeginTrans
    
    If GetLastMonthStock = "" Then
        LblErr = DisplayMsg(4019) ' "There's no Fix Status Stock Master!"
        Exit Sub
    End If
    
    FixMonth = Right(GetLastMonthStock, 2)
    FixYear = Mid(GetLastMonthStock, 1, 4)
    
    Set RsI = New Recordset
    RsI.Open "select * from stock_master (updlock) where " & _
        " warehouse_code='" & Trim(WHCode) & "' and " & _
        " item_code='" & Trim(ItemCode) & "'", DBi, adOpenDynamic, adLockOptimistic
    
    If RsI.EOF Then
        RsI.AddNew
        'Receipt
        RsI!Item_Code = Trim$(ItemCode)
        RsI!Warehouse_Code = Trim$(WHCode)
        RsI!lm_premonth = "0"
        RsI!tm_premonth = "0"
        RsI!nm_premonth = "0"
        
        RsI!lm_supply = "0"
        RsI!tm_supply = "0"
        RsI!nm_supply = "0"
        RsI!lm_lossreject = "0"
        RsI!tm_lossreject = "0"
        RsI!nm_lossreject = "0"
        
        RsI!lm_current = "0"
        RsI!tm_current = "0"
        RsI!nm_current = "0"
        
        RsI!lm_inventory = Null
        RsI!tm_inventory = Null
        RsI!nm_inventory = Null
        
        If Val(FixYear) = Val(RecYear) Then
            RsI!lm_receipt = IIf(Val(RecMonth) = Val(FixMonth), QtyX, 0)
            RsI!tm_receipt = IIf(Val(RecMonth) = Val(FixMonth) + 1, QtyX, 0)
            RsI!nm_receipt = IIf(Val(RecMonth) = Val(FixMonth) + 2, QtyX, 0)
            'Current
            RsI!lm_current = Val(RsI!lm_premonth) + Val(RsI!lm_receipt) - Val(RsI!lm_supply) - Val(RsI!lm_lossreject)
            RsI!tm_current = Val(RsI!tm_premonth) + Val(RsI!tm_receipt) - Val(RsI!tm_supply) - Val(RsI!tm_lossreject)
            'Next Proses
            RsI!nm_premonth = RsI!tm_current
            NM_P = RsI!nm_premonth
            RsI!nm_current = Val(RsI!nm_premonth) + Val(RsI!nm_receipt) - Val(RsI!nm_supply) - Val(RsI!nm_lossreject)
        ElseIf Val(FixYear) < Val(RecYear) Then
            RsI!tm_receipt = RsI!tm_receipt + IIf((Val(RecMonth) + 12) - Val(FixMonth) = 1, QtyX, 0)
            RsI!nm_receipt = RsI!nm_receipt + IIf((Val(RecMonth) + 12) - Val(FixMonth) = 2, QtyX, 0)
            'Current
            RsI!lm_current = Val(RsI!lm_premonth) + Val(RsI!lm_receipt) - Val(RsI!lm_supply) - Val(RsI!lm_lossreject)
            RsI!tm_current = Val(RsI!tm_premonth) + Val(RsI!tm_receipt) - Val(RsI!tm_supply) - Val(RsI!tm_lossreject)
            'Next Proses
            RsI!nm_current = Val(RsI!nm_premonth) + Val(RsI!nm_receipt) - Val(RsI!nm_supply) - Val(RsI!nm_lossreject)
        End If
        RsI!Last_Update = Now
        RsI!last_user = userLogin
        RsI.update
    Else
        'LM Null offset
        If IsNull(RsI!lm_premonth) Then
            LM_P = 0
        Else
            LM_P = CDbl(RsI!lm_premonth)
        End If
        If IsNull(RsI!lm_supply) Then
            LM_S = 0
        Else
            LM_S = CDbl(RsI!lm_supply)
        End If
        If IsNull(RsI!lm_lossreject) Then
            LM_RJ = 0
        Else
            LM_RJ = CDbl(RsI!lm_lossreject)
        End If
        'TM Null Offset
        If IsNull(RsI!tm_premonth) Then
            TM_P = 0
        Else
            TM_P = CDbl(RsI!tm_premonth)
        End If
        If IsNull(RsI!tm_supply) Then
            TM_S = 0
        Else
            TM_S = CDbl(RsI!tm_supply)
        End If
        If IsNull(RsI!tm_lossreject) Then
            TM_RJ = 0
        Else
            TM_RJ = CDbl(RsI!tm_lossreject)
        End If
        'NM Null Offset
        If IsNull(RsI!nm_premonth) Then
            NM_P = 0
        Else
            NM_P = CDbl(RsI!nm_premonth)
        End If
        If IsNull(RsI!nm_supply) Then
            NM_S = 0
        Else
            NM_S = CDbl(RsI!nm_supply)
        End If
        If IsNull(RsI!nm_lossreject) Then
            NM_RJ = 0
        Else
            NM_RJ = CDbl(RsI!nm_lossreject)
        End If
    
        If Val(FixYear) = Val(RecYear) Then
            RsI!lm_receipt = RsI!lm_receipt - IIf(Val(RecMonth) = Val(FixMonth), (OldQty - QtyX), 0)
            RsI!tm_receipt = RsI!tm_receipt - IIf(Val(RecMonth) = Val(FixMonth) + 1, (OldQty - QtyX), 0)
            RsI!nm_receipt = RsI!nm_receipt - IIf(Val(RecMonth) = Val(FixMonth) + 2, (OldQty - QtyX), 0)
            'Current
            RsI!lm_current = (LM_P + Val(RsI!lm_receipt)) - LM_S - LM_RJ
            RsI!tm_current = (TM_P + Val(RsI!tm_receipt)) - TM_S - TM_RJ
            'Next Proses
            RsI!nm_premonth = RsI!tm_current
            NM_P = RsI!nm_premonth
            RsI!nm_current = (NM_P + Val(RsI!nm_receipt)) - NM_S - NM_RJ
        ElseIf Val(FixYear) < Val(RecYear) Then
            RsI!tm_receipt = RsI!tm_receipt - IIf((Val(RecMonth) + 12) - Val(FixMonth) = 1, (OldQty - QtyX), 0)
            RsI!nm_receipt = RsI!nm_receipt - IIf((Val(RecMonth) + 12) - Val(FixMonth) = 2, (OldQty - QtyX), 0)
            'Current
            RsI!lm_current = (LM_P + Val(RsI!lm_receipt)) - LM_S - LM_RJ
            RsI!tm_current = (TM_P + Val(RsI!tm_receipt)) - TM_S - TM_RJ
            'Next Proses
            RsI!nm_premonth = RsI!tm_current
            NM_P = RsI!nm_premonth
            RsI!nm_current = (NM_P + Val(RsI!nm_receipt)) - NM_S - NM_RJ
        End If
        RsI!Last_Update = Now
        RsI!last_user = userLogin
        RsI.update
    End If
    
    On Error GoTo erri
    DBi.CommitTrans
    RsI.Close
    Set RsI = Nothing
    DBi.Close
    Set DBi = Nothing
    Exit Sub
erri:
    DBi.RollbackTrans
    RsI.Close
    Set RsI = Nothing
    DBi.Close
    Set DBi = Nothing
End Sub

Sub DeleteStock(WHCode As String, ItemCode As String, RecMonth As String, RecYear As String, QtyX As Double)
    Dim RsI As Recordset
    Dim DBi As New Connection
    Dim FixMonth As String
    Dim FixYear As String
    Dim LM_P As Double, LM_S As Double, LM_RJ As Double, LM_I As Double
    Dim TM_P As Double, TM_S As Double, TM_RJ As Double, TM_I As Double
    Dim NM_P As Double, NM_S As Double, NM_RJ As Double, NM_I As Double
    
    DBi.ConnectionTimeout = 0
    DBi.Open Db.ConnectionString
    DBi.BeginTrans
    
    If GetLastMonthStock = "" Then
        LblErr = DisplayMsg(4019) '"There's no Fix Status Stock Master!"
        Exit Sub
    End If
    
    FixMonth = Right(GetLastMonthStock, 2)
    FixYear = Mid(GetLastMonthStock, 1, 4)
    
    Set RsI = New Recordset
    RsI.Open "select * from stock_master (updlock) where " & _
    " warehouse_code='" & Trim(WHCode) & "' and " & _
    " item_code='" & Trim(ItemCode) & "'", DBi, adOpenDynamic, adLockOptimistic
    
    If Not RsI.EOF Then
        'LM Null offset
        If IsNull(RsI!lm_premonth) Then
            LM_P = 0
        Else
            LM_P = CDbl(RsI!lm_premonth)
        End If
        If IsNull(RsI!lm_supply) Then
            LM_S = 0
        Else
            LM_S = CDbl(RsI!lm_supply)
        End If
        If IsNull(RsI!lm_lossreject) Then
            LM_RJ = 0
        Else
            LM_RJ = CDbl(RsI!lm_lossreject)
        End If
        'TM Null Offset
        If IsNull(RsI!tm_premonth) Then
            TM_P = 0
        Else
            TM_P = CDbl(RsI!tm_premonth)
        End If
        If IsNull(RsI!tm_supply) Then
            TM_S = 0
        Else
            TM_S = CDbl(RsI!tm_supply)
        End If
        If IsNull(RsI!tm_lossreject) Then
            TM_RJ = 0
        Else
            TM_RJ = CDbl(RsI!tm_lossreject)
        End If
        'NM Null Offset
        If IsNull(RsI!nm_premonth) Then
            NM_P = 0
        Else
            NM_P = CDbl(RsI!nm_premonth)
        End If
        If IsNull(RsI!nm_supply) Then
            NM_S = 0
        Else
            NM_S = CDbl(RsI!nm_supply)
        End If
        If IsNull(RsI!nm_lossreject) Then
            NM_RJ = 0
        Else
            NM_RJ = CDbl(RsI!nm_lossreject)
        End If
    
        If Val(FixYear) = Val(RecYear) Then
            RsI!lm_receipt = RsI!lm_receipt - IIf(Val(RecMonth) = Val(FixMonth), QtyX, 0)
            RsI!tm_receipt = RsI!tm_receipt - IIf(Val(RecMonth) = Val(FixMonth) + 1, QtyX, 0)
            RsI!nm_receipt = RsI!nm_receipt - IIf(Val(RecMonth) = Val(FixMonth) + 2, QtyX, 0)
            'Current
            RsI!lm_current = (LM_P + Val(RsI!lm_receipt)) - LM_S - LM_RJ
            RsI!tm_current = (TM_P + Val(RsI!tm_receipt)) - TM_S - TM_RJ
            'Next Proses
            RsI!nm_premonth = RsI!tm_current
            NM_P = RsI!nm_premonth
            RsI!nm_current = (NM_P + Val(RsI!nm_receipt)) - NM_S - NM_RJ
        ElseIf Val(FixYear) < Val(RecYear) Then
            RsI!tm_receipt = RsI!tm_receipt - IIf((Val(RecMonth) + 12) - Val(FixMonth) = 1, QtyX, 0)
            RsI!nm_receipt = RsI!nm_receipt - IIf((Val(RecMonth) + 12) - Val(FixMonth) = 2, QtyX, 0)
            'Current
            RsI!lm_current = (LM_P + Val(RsI!lm_receipt)) - LM_S - LM_RJ
            RsI!tm_current = (TM_P + Val(RsI!tm_receipt)) - TM_S - TM_RJ
            'Next Proses
            RsI!nm_premonth = RsI!tm_current
            NM_P = RsI!nm_premonth
            RsI!nm_current = (NM_P + Val(RsI!nm_receipt)) - NM_S - NM_RJ
        End If
        RsI!Last_Update = Now
        RsI!last_user = userLogin
        RsI.update
    End If
    On Error GoTo erri
    DBi.CommitTrans
    RsI.Close
    Set RsI = Nothing
    DBi.Close
    Set DBi = Nothing
    Exit Sub
erri:
    DBi.RollbackTrans
    RsI.Close
    Set RsI = Nothing
    DBi.Close
    Set DBi = Nothing
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub TxtSj_Change()
    txtRegisterNo.Text = ""
End Sub

Private Sub TxtSj_GotFocus()
    l_SJNo = TxtSj.Text
End Sub

Private Sub TxtSj_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Sub browseprice(itmCode As String, Tgl As String, TradeCode As String, Curr As String)
    Dim sql2 As String
    Dim rs2 As New Recordset
    Dim i As Integer, idxShow As Byte, blnpricez As Boolean
    Dim tmpprice As String
    
    sql2 = "select trade_code, priority_cls, currency_code, price, unit_cls " & _
        "from price_master " & _
        "where item_code='" & Trim$(itmCode) & "' and price_cls='01' " & _
        "and (trade_code='" & TradeCode & "' or trade_code='000000') " & _
        "and start_date<='" & Format(Tgl, "yyyymmdd") & "' and end_date>='" & Format(Tgl, "yyyymmdd") & "' " & _
        "and currency_code = '" & Curr & "' " & _
        "order by trade_code desc, priority_cls desc"
    
    sql2 = "select trade_code, priority_cls, currency_code, sum(price) price, unit_cls " & _
        "from( " & _
        "select trade_code, priority_cls, currency_code, price, unit_cls " & _
        "From price_master " & _
        "where item_code='" & Trim$(itmCode) & "' and price_cls='01' " & _
        "and (trade_code='" & TradeCode & "' or trade_code='000000') " & _
        "and start_date<='" & Format(Tgl, "yyyymmdd") & "' and end_date>='" & Format(Tgl, "yyyymmdd") & "' " & _
        "and currency_code = '" & Curr & "' " & _
        "Union " & _
        "select trade_code, priority_cls, currency_code, price, unit_cls " & _
        "From price_master " & _
        "where item_code='" & Trim$(itmCode) & "' and price_cls='05' " & _
        "and (trade_code='" & TradeCode & "' or trade_code='000000') " & _
        "and start_date<='" & Format(Tgl, "yyyymmdd") & "' and end_date>='" & Format(Tgl, "yyyymmdd") & "' " & _
        "and currency_code = '" & Curr & "') a " & _
        "group by trade_code, priority_cls, currency_code, unit_cls " & _
        "order by trade_code desc, priority_cls desc "
    
    Set rs2 = Db.Execute(sql2)
    
    With cboprice
        tmpprice = cboprice
        .clear
        .columnCount = 4
        .ColumnWidths = "70pt;70pt;0pt;0pt"
        .ListWidth = 140
        .ListRows = 4
        idxShow = 0
        i = 0
        blnpricez = False
        If Not rs2.EOF Then
            Do While Not rs2.EOF
                .AddItem
                .List(i, 0) = Format(rs2("price"), gs_formatPrice)
                If rs2("trade_code") = "000000" Then
                    .List(i, 1) = "Common " & Trim(rs2("priority_cls"))
                Else
                    .List(i, 1) = "Priority " & Trim(rs2("priority_cls"))
                    If blnshow Then idxShow = i: blnshow = False: blnpricez = True
                End If
                .List(i, 2) = Trim(rs2("Currency_Code"))
                .List(i, 3) = Trim(rs2("unit_cls"))
                rs2.MoveNext
                i = i + 1
            Loop
            cboprice = tmpprice
            If idxShow >= 0 And blnpricez Then cboprice.Text = cboprice.List(idxShow, 0)
        End If
    End With
    rs2.Close
    Set rs2 = Nothing
End Sub

Function CariMakerItem(ItemCode As String) As String
    Dim RsCM As Recordset
    Set RsCM = Db.Execute("Select MakerItem_code IMC from item_master where item_code='" & ItemCode & "'")
    If Not RsCM.EOF Then
        If IsNull(RsCM!imc) Then
            CariMakerItem = ""
        Else
            CariMakerItem = Trim$(RsCM!imc)
        End If
    Else
        CariMakerItem = ""
    End If
    RsCM.Close
End Function

Sub GridAct(PosRec As Long, seqNo As Long, opt As Byte)
    With Grid
        If opt = 1 Then .AddItem "", PosRec
        .TextMatrix(PosRec, bteColSJNo) = Trim$(TxtSj)
        .Cell(flexcpAlignment, PosRec, bteColSJNo) = flexAlignLeftCenter
        .TextMatrix(PosRec, bteColReceiptDate) = Format(TglReceipt, "dd mmm yyyy")
        .Cell(flexcpAlignment, PosRec, bteColReceiptDate) = flexAlignLeftCenter
        .TextMatrix(PosRec, bteColProdCode) = Trim$(CboItem)
        .Cell(flexcpAlignment, PosRec, bteColProdCode) = flexAlignLeftCenter
        .TextMatrix(PosRec, bteColPartNo) = CariMakerItem(CboItem)
        .Cell(flexcpAlignment, PosRec, bteColPartNo) = flexAlignLeftCenter
        .TextMatrix(PosRec, bteColDesc) = Trim$(uf_GetItemDescription(Trim$(CboItem)))
        .Cell(flexcpAlignment, PosRec, bteColDesc) = flexAlignLeftCenter
        .TextMatrix(PosRec, bteColWHCode) = Trim$(CboWHCode)
        .Cell(flexcpAlignment, PosRec, bteColWHCode) = flexAlignLeftCenter
        .TextMatrix(PosRec, bteColAddress) = Trim$(Lbladdress)
        .Cell(flexcpAlignment, PosRec, bteColAddress) = flexAlignLeftCenter
        .TextMatrix(PosRec, bteColReceipt) = Trim$(CboRecCls)
        .Cell(flexcpAlignment, PosRec, bteColReceipt) = flexAlignLeftCenter
        
        If InStr(1, CDbl(txtqty), ".") = 0 Then
            .TextMatrix(PosRec, bteColQty) = Format(txtqty, gs_formatQty)
        ElseIf InStr(1, CDbl(txtqty), ".") > 0 Then
            If Split(CDbl(txtqty), ".")(1) <> "" Then .TextMatrix(PosRec, bteColQty) = Format(txtqty, gs_formatQty)
        End If
        
        .Cell(flexcpAlignment, PosRec, bteColQty) = flexAlignRightCenter
        .TextMatrix(PosRec, bteColUnit) = uf_GetUnitDescription(Trim$(Cbounit.List(Cbounit.ListIndex, 0)))
        .Cell(flexcpAlignment, PosRec, bteColUnit) = flexAlignLeftCenter
        
        .TextMatrix(PosRec, bteColCurr) = uf_GetCurrencyDescription(Trim$(cbocurr.List(cbocurr.ListIndex, 0)))
        .Cell(flexcpAlignment, PosRec, bteColCurr) = flexAlignLeftCenter
        
        .TextMatrix(PosRec, bteColPrice) = Format(Trim(cboprice), gs_formatPrice)
        .Cell(flexcpAlignment, PosRec, bteColPrice) = flexAlignRightCenter

        If InStr(1, CDbl(txtamount), ".") = 0 Then
            .TextMatrix(PosRec, bteColAmount) = Format(txtamount, gs_formatAmount)
        ElseIf InStr(1, CDbl(txtamount), ".") > 0 Then
            If Split(CDbl(txtamount), ".")(1) <> "" Then .TextMatrix(PosRec, bteColAmount) = Format(txtamount, gs_formatAmount)
        End If

        .Cell(flexcpAlignment, PosRec, bteColAmount) = flexAlignRightCenter
        .TextMatrix(PosRec, bteColUnitCls) = Trim$(Cbounit.List(Cbounit.ListIndex, 0))
        .TextMatrix(PosRec, bteColCurrCode) = Trim$(cbocurr.List(cbocurr.ListIndex, 0))
        .TextMatrix(PosRec, bteColRemark) = Trim$(TxtRemarks)
        .TextMatrix(PosRec, bteColSeqNo) = seqNo
        .TextMatrix(PosRec, bteColBC) = Trim$(txtBC40)
        .TextMatrix(PosRec, bteColBctype) = Trim(CbotypeBC)
        .TextMatrix(PosRec, bteColTransport) = Trim$(cboTransport)
        .TextMatrix(PosRec, bteColPackageQty) = CDbl(txtPackage)
        .TextMatrix(PosRec, bteColPackage) = Trim$(cboPackage)
        .TextMatrix(PosRec, bteColRecStatus) = Trim$(txtRec_Status.Text)
        .TextMatrix(PosRec, bteColNoRegister) = Trim$(txtRegisterNo.Text)
        .TextMatrix(PosRec, bteColNoSeri) = Trim$(txtNoSeri.Text)
        
        .Cell(flexcpBackColor, PosRec, 0) = vbWhite
        .Row = PosRec
    End With
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblErr.Caption = ErrMsg
    End If
End Sub


Private Sub up_SendEmail()
    'Recordset
    Dim rsmailconfig    As New ADODB.Recordset
    Dim rsHeader        As New ADODB.Recordset
    Dim rsdetail        As New ADODB.Recordset
    
    'Excel Application
    Dim xlapp As New Excel.application
    
    Dim ls_SupplierCodeGrid As String
    Dim ls_TicketDateGrid   As String
    
    'Variable
    Dim ls_sql As String
    Dim li_Row As Integer
    Dim ls_smtpserver As String
    Dim li_smtpport     As Double
    Dim li_smtptimeout  As Double
    Dim ls_sender       As String
    Dim ls_username     As String
    Dim ls_password     As String
    Dim li_timer        As Double
    Dim ls_subject      As String
    Dim ls_cc           As String
    Dim ls_bcc          As String
    Dim ls_message      As String
    Dim ls_recipient    As String
    Dim ls_SupplierCode As String
    Dim ls_TicketNo     As String
    Dim ls_SampleNo     As String
    Dim ls_ticketdate   As String
    Dim ls_ItemCode     As String
    Dim li_wet          As Double
    Dim li_drc          As Double
    Dim li_dry          As Double
    Dim li_price        As Double
    Dim li_amount       As Double
    Dim ls_log          As String
    
    Dim Idx                 As Integer
    Dim li_totalwet         As Double, li_totaldry          As Double
    Dim li_totalamout       As Double, li_totaladvanced     As Double
    Dim li_totaltaxamount   As Double, li_totalfinalpaid    As Double
    
    Dim strFile         As String
    
    
    Dim strAttachFile   As String
    Const clrGrey = 15
    
    Dim ls_desc As String
    
    On Local Error GoTo errHandler
    
        'Set Timer
    ls_sql = " select smtp_server, port_number, smtp_timeout, email_address, user_email, pass_email, timer, subject, " & vbCrLf & _
            "       ccemail_address, bccemail_address,mail_header,Mail_Content,Mail_Footer " & vbCrLf & _
            " from email_config "
            
    If rsmailconfig.State <> adStateClosed Then rsmailconfig.Close
    Set rsmailconfig = Db.Execute(ls_sql)
    
    If Not rsmailconfig.EOF Then
    
        ls_smtpserver = Trim(rsmailconfig!smtp_server & "")
        li_smtpport = IIf(IsNull(rsmailconfig!port_number), 0, rsmailconfig!port_number)
        li_smtptimeout = IIf(IsNull(rsmailconfig!smtp_timeout), 0, rsmailconfig!smtp_timeout)
        ls_sender = Trim(rsmailconfig!Email_Address & "")
        ls_username = Trim(rsmailconfig!user_email & "")
        ls_password = Trim(rsmailconfig!pass_email & "")
        li_timer = IIf(IsNull(rsmailconfig!Timer), 0, rsmailconfig!Timer)
        ls_subject = Trim(rsmailconfig!Subject & "")
        ls_cc = Trim(rsmailconfig!ccemail_address & "")
        ls_bcc = Trim(rsmailconfig!bccemail_address & "")
        ls_message = Trim(rsmailconfig!mail_header & "") & vbCrLf & vbCrLf
        ls_message = ls_message & Trim(rsmailconfig!Mail_Content & "") & vbCrLf & vbCrLf
        ls_message = ls_message & "Part No" & vbTab & vbTab & ":" & Trim(CboItem.Text) & vbCrLf
        ls_message = ls_message & "Part Name" & vbTab & ":" & Trim(LblItemName.Text) & vbCrLf
        ls_message = ls_message & "Qty" & vbTab & vbTab & ":" & Trim(txtqty.Text) & vbCrLf
        ls_message = ls_message & "Vendor" & vbTab & vbTab & ":" & Trim(CboSupp.Text) & " - " & Trim(Lblsupp.Text) & vbCrLf
        ls_message = ls_message & "Surat Jalan" & vbTab & ":" & Trim(TxtSj.Text) & vbCrLf & vbCrLf
        ls_message = ls_message & "Kepada yang berkepentingan Mohon Segera Pastikan Barang Tersebut." & vbCrLf
        ls_message = ls_message & "Terima Kasih" & vbCrLf & vbCrLf & vbCrLf
        ls_message = ls_message & "Note :" & vbCrLf
        ls_message = ls_message & Trim(rsmailconfig!Mail_Footer & "")
        
        
        
    End If
    
    rsmailconfig.Close
    Set rsmailconfig = Nothing
    
    
      '========================================================================
      'start send email function
      '========================================================================

      Call up_cdoSendEmail(ls_smtpserver, li_smtpport, li_smtptimeout, ls_username, ls_password, ls_sender, ls_recipient, ls_subject & ls_TicketNo, ls_message, ls_cc, ls_bcc, strAttachFile)

      '========================================================================
      'end send email function
      '========================================================================
      
    
ErrExit:
    Exit Sub
errHandler:
    'Call SaveToTxt("SendEmail_ErrorLog", " Ticket No : " & ls_TicketNo & " ( " & ls_ticketdate & " ) " & err.Description, CurrDate)
    err.clear
    Resume ErrExit
    
End Sub



Public Sub up_cdoSendEmail(ByVal ls_smtpserver As String, ByVal li_smtpport As Double, ByVal li_smtptimeout As Double, _
                    ByVal ls_username As String, ByVal ls_password As String, ByVal ls_sender As String, _
                    ByVal ls_recipient As String, ByVal ls_subject As String, ByVal ls_body As String, _
                    Optional ByVal ls_cc As String, Optional ByVal ls_bcc As String, Optional ByVal ls_attachment As String)
    
    Dim cdoMsg As New CDO.message
    Dim cdoConf As New CDO.Configuration
    Dim Flds
    Dim attachment
   
    On Error GoTo errHandler
    
    Set Flds = cdoConf.Fields
       
    With Flds
        .Item(cdoSendUsingMethod) = cdoSendUsingPort
        .Item(cdoSMTPServer) = ls_smtpserver
        '.Item(cdoSMTPServerPort) = IIf(li_smtpport = 0, 25, li_smtpport)
        .Item(cdoSMTPServerPort) = li_smtpport
        .Item(cdoSMTPConnectionTimeout) = IIf(li_smtptimeout = 0, 10, li_smtptimeout)
        .Item(cdoSMTPAuthenticate) = 1 'default
        .Item(cdoSendUserName) = ls_username
        .Item(cdoSendPassword) = ls_password
        .Item(cdoSMTPUseSSL) = 1
        
        
        
        
        
        
    .update
    End With
   
   
    ' Apply the settings to the message.
    With cdoMsg
        Set .Configuration = cdoConf
        .To = ls_cc
        .From = ls_sender
        .Subject = ls_subject
        .TextBody = ls_body
        
'        If Not colAttachments Is Nothing Then
'            For Each attachment In colAttachments
'                .AddAttachment attachment
'            Next
'        End If
        'If ls_cc <> "" Then .CC = ls_cc
        If ls_bcc <> "" Then .BCC = ls_bcc
        .Send
    End With
    LblErr.Caption = "Send Successful"
    
   
ErrExit:
    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set Flds = Nothing
    
    Exit Sub

errHandler:
LblErr.Caption = err.Description
'Call SaveToTxt("SendEmail_ErrorLog", "SEND MAIL - " & err.number & " : " & err.Description, CurrDate)
GoTo ErrExit
    
End Sub


Public Function up_Validasi_Item(Item_Code As String) As Boolean
Dim sqlcek As String

Dim RsCekItem As New ADODB.Recordset

    sqlcek = "select tOP 1 Item_Code From Part_ReceipT WHERE Item_Code='" & Item_Code & "' "
    Set RsCekItem = Db.Execute(sqlcek)
    
    If Not RsCekItem.EOF Then
        up_Validasi_Item = True
    Else
        up_Validasi_Item = False
    End If

End Function

Private Sub GetBCType()
Dim RS As New ADODB.Recordset
    RS.Open "select Type_BC from Trade_Master where Trade_Code = '" & CboSupp.Text & "' ", Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If RS.EOF = False Then
        CbotypeBC.Text = IIf(IsNull(Trim(RS!Type_BC)), "", Trim(RS!Type_BC))
    End If
    RS.Close
End Sub

Public Function up_ValidateSuratJalan() As Boolean
Dim sqlcek As String
Dim rsCek As New ADODB.Recordset
    
    DateActual = Format(TglReceipt, "MM/DD/YYYY")
    
    sqlcek = " SELECT  * FROM Part_Receipt WHERE SuratJalan_No='" & TxtSj.Text & "' "
    Set rsCek = Db.Execute(sqlcek)
    
    If Not rsCek.EOF Then
    
        sqlcek = " SELECT DISTINCT Receipt_Date FROM Part_Receipt WHERE SuratJalan_No='" & TxtSj.Text & "' "
        Set rsCek = Db.Execute(sqlcek)
        
        If Not rsCek.EOF Then
            receiptDate = Trim(rsCek!Receipt_Date & "")
        End If
        
        If receiptDate = DateActual Then
            validate = True
        Else
            validate = False
        End If
    Else
        validate = True
    End If
    
End Function

Private Sub TxtSj_LostFocus()
     Dim rsGetNoSeri As New ADODB.Recordset
        
    If l_SJNo <> TxtSj Then
        If Trim(TxtSj.Text) <> "" Then
            sql = "EXEC dbo.sp_GetNoRegister '" & TglReceipt.Value & "', 'S1', '" & userLogin & "' "
                    
            If rsGetNoSeri.State <> adStateClosed Then rsGetNoSeri.Close
            rsGetNoSeri.Open sql, Db, adOpenForwardOnly, adLockReadOnly
            
            If Not rsGetNoSeri.EOF Then
               txtRegisterNo.Text = rsGetNoSeri.Fields("No_Register")
            End If
        End If
    End If
End Sub

Private Sub up_GetNoSeri(pItemCode As String)
Dim rsGetNo As New ADODB.Recordset

    If Trim(CbotypeBC.Text) = "4.0" Then
            sql = "EXEC dbo.sp_PartReceipt_GetNoSeri @SJNo = '" & Trim(TxtSj.Text) & "', " & vbCrLf & _
                "  @ReceiptDate = '" & Format(TglReceipt.Value, "yyyy-mm-dd") & "', " & vbCrLf & _
                "  @ItemCode ='" & pItemCode & "' "
                
                If rsGetNo.State <> adStateClosed Then rsGetNo.Close
                rsGetNo.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                
                If Not rsGetNo.EOF Then
                   txtNoSeri.Text = Trim(rsGetNo.Fields("No_Seri"))
                End If
    ElseIf Trim(CbotypeBC.Text) = "2.3" Then
    
        sql = "EXEC dbo.sp_PartReceipt_GetNoSeri @SJNo = '" & Trim(TxtSj.Text) & "', " & vbCrLf & _
                "  @ReceiptDate = '" & Format(TglReceipt.Value, "yyyy-mm-dd") & "', " & vbCrLf & _
                "  @ItemCode ='" & pItemCode & "' "
                        
                If rsGetNo.State <> adStateClosed Then rsGetNo.Close
                rsGetNo.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                
                If Not rsGetNo.EOF Then
                   txtNoSeri.Text = Trim(rsGetNo.Fields("No_Seri"))
                End If
            CboRecCls.Text = "R"
    Else
        sql = "EXEC dbo.sp_PartReceipt_GetNoSeri @SJNo = '" & Trim(TxtSj.Text) & "', " & vbCrLf & _
                "  @ReceiptDate = '" & Format(TglReceipt.Value, "yyyy-mm-dd") & "', " & vbCrLf & _
                "  @ItemCode ='" & pItemCode & "' "
                
                If rsGetNo.State <> adStateClosed Then rsGetNo.Close
                rsGetNo.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                
                If Not rsGetNo.EOF Then
                   txtNoSeri.Text = Trim(rsGetNo.Fields("No_Seri"))
                End If
    End If
    
End Sub
