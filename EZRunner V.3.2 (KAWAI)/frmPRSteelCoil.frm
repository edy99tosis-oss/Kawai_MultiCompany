VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPRSteelCoil 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Request (Steel/Coil)"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "frmPRSteelCoil.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13343
      TabIndex        =   32
      Top             =   240
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "Previe&w"
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   10080
      Width           =   1125
   End
   Begin VB.TextBox txtRequestNo 
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
      MaxLength       =   25
      TabIndex        =   8
      Top             =   2400
      Width           =   1600
   End
   Begin VB.CommandButton Command1 
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
      Left            =   11007
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10080
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
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
      Index           =   0
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2355
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   563
      TabIndex        =   21
      Top             =   9360
      Width           =   14145
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
         TabIndex        =   22
         Top             =   180
         Width           =   13905
      End
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   13583
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10080
      Width           =   1125
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
      Left            =   563
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   10080
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
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
      Index           =   2
      Left            =   12294
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10080
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1515
      Left            =   83
      TabIndex        =   23
      Top             =   720
      Width           =   15105
      Begin VB.TextBox lblsec 
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
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   735
         Width           =   3645
      End
      Begin VB.TextBox lblDept 
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
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   300
         Width           =   3645
      End
      Begin VB.TextBox lblPerson 
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
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Width           =   3645
      End
      Begin MSComCtl2.DTPicker Period 
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Top             =   660
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   141230083
         UpDown          =   -1  'True
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker RequestDate1 
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Top             =   1080
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
      Begin MSComCtl2.DTPicker RequestDate2 
         Height          =   315
         Left            =   4020
         TabIndex        =   6
         Top             =   1080
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
      Begin MSForms.ComboBox CboSec 
         Height          =   315
         Left            =   8385
         TabIndex        =   37
         Top             =   675
         Width           =   975
         VariousPropertyBits=   746604571
         MaxLength       =   2
         DisplayStyle    =   3
         Size            =   "1720;556"
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
         Caption         =   "Section"
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
         Left            =   7200
         TabIndex        =   36
         Top             =   720
         Width           =   630
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   9480
         X2              =   13185
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   9480
         X2              =   13185
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Left            =   7200
         TabIndex        =   33
         Top             =   280
         Width           =   1020
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   8385
         TabIndex        =   2
         Top             =   240
         Width           =   975
         VariousPropertyBits=   746604571
         MaxLength       =   2
         DisplayStyle    =   3
         Size            =   "1720;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboPerson 
         Height          =   315
         Left            =   2040
         TabIndex        =   0
         Top             =   240
         Width           =   975
         VariousPropertyBits=   746604571
         MaxLength       =   2
         DisplayStyle    =   3
         Size            =   "1720;556"
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
         Caption         =   "Person in Charge"
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
         Left            =   195
         TabIndex        =   31
         Top             =   280
         Width           =   1725
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3120
         X2              =   6825
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to "
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
         Left            =   3652
         TabIndex        =   30
         Top             =   1110
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Request Date From "
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
         Left            =   195
         TabIndex        =   29
         Top             =   1130
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
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
         Left            =   195
         TabIndex        =   26
         Top             =   705
         Width           =   1710
      End
   End
   Begin MSComCtl2.DTPicker RequestDate 
      Height          =   315
      Left            =   6240
      TabIndex        =   10
      Top             =   2400
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
   Begin MSComCtl2.FlatScrollBar hscrollbar 
      Height          =   255
      Left            =   90
      TabIndex        =   34
      Top             =   8940
      Visible         =   0   'False
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Max             =   1
      Orientation     =   1638401
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   6060
      Left            =   90
      TabIndex        =   13
      Top             =   2880
      Width           =   15105
      _cx             =   26644
      _cy             =   10689
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
         Left            =   0
         TabIndex        =   14
         Top             =   600
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
   End
   Begin MSForms.ComboBox cboAlarm 
      Height          =   315
      Left            =   8700
      TabIndex        =   11
      Top             =   2400
      Width           =   855
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
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
      Caption         =   "Alarm"
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
      Index           =   5
      Left            =   8070
      TabIndex        =   28
      Top             =   2445
      Width           =   600
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
      Left            =   9780
      TabIndex        =   27
      Top             =   2430
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Index           =   1
      Left            =   1680
      TabIndex        =   25
      Top             =   2445
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Request Date"
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
      Left            =   4950
      TabIndex        =   24
      Top             =   2445
      Width           =   1155
   End
   Begin MSForms.ComboBox cboRequestNo 
      Height          =   315
      Left            =   2730
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1860
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "3281;556"
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
      Left            =   83
      TabIndex        =   7
      Top             =   2400
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Request (Steel/Coil)"
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
      Left            =   158
      TabIndex        =   20
      Top             =   240
      Width           =   14955
   End
End
Attribute VB_Name = "frmPRSteelCoil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'0: direct , 1: others
Option Explicit
Dim sql As String, sqlGrid As String
Dim RS As New ADODB.Recordset, rsGrid As New ADODB.Recordset
Dim i As Long, orderawal As Double, poqty As Double, Counti As Byte
Dim ubah As Boolean, ada As Boolean, statusfix As Byte, temptgl As Byte
Dim tempperiod2 As String, isirequestdate As Date, OldDelDate As Date

Dim bteColSelect As Byte
Dim bteColProdCode As Byte
Dim bteColSteelKind As Byte
Dim bteColMatT As Byte
Dim bteColMatW As Byte
Dim bteColMatL As Byte
Dim bteColUnitCls As Byte
Dim bteColUnitDes As Byte
Dim bteColUnit As Byte
Dim bteColStock As Byte
Dim bteColOPQty As Byte
Dim bteColFixOrder As Byte
Dim bteColRemReq As Byte
Dim bteColReq As Byte
Dim bteColReqQty As Byte
Dim bteColEnd As Byte
Dim bteColDelDate As Byte
Dim bteColPur As Byte
Dim bteColAcc As Byte

Dim bteHakPrice As Byte

Sub Kosong()
    LblErrMsg = ""
    cboPerson.Text = "": lblPerson.Text = ""
    cboDept.Text = "": lblDept.Text = ""
    CboSec.Text = "": lblsec.Text = ""
    Period.Value = Format(Now, "MMM yyyy")
    temptgl = Period.Month
    requestdate1.Value = Format(Now, "yyyy-mm-01")
    requestdate1.Enabled = True
    requestdate2.Value = Format(Now, "dd MMM yyyy")
    requestdate2.Enabled = True
    txtRequestNo.Text = ""
    cborequestno.clear
    RequestDate.Value = Format(Now, "dd MMM yyyy")
    RequestDate.Enabled = True
    isirequestdate = Format(RequestDate, "yyyy-mm-dd")
    DelDate.Value = Format(Now, "dd MMM yyyy")
    DelDate.Visible = False
    cboAlarm.ListIndex = 1
    
    grid.FocusRect = flexFocusNone
    ubah = False: ada = False
    statusfix = 0: Call kunci(False)
    Call Header
End Sub

Sub adtocboperson()
Dim sqlperson As String
Dim rsperson As New Recordset

    sqlperson = "select * from PersonInCharge_Cls order by PersonInCharge_cls"
    Set rsperson = Db.Execute(sqlperson)
    
    With cboPerson
        .clear
        .columnCount = 2
        .ColumnWidths = "30pt;120pt"
        .ListWidth = 150
        .ListRows = 15
        
        i = 0
        Do While Not rsperson.EOF
            .AddItem
            .List(i, 0) = Trim(rsperson("PersonInCharge_cls"))
            .List(i, 1) = IIf(IsNull(rsperson("description")), "", Trim(rsperson("description")))
            rsperson.MoveNext
            i = i + 1
        Loop
    End With
    Set rsperson = Nothing
End Sub

Sub adtocboDept()
Dim sqldept As String
Dim rsdept As New Recordset

    sqldept = "select * from Department_Cls order by Department_cls"
    Set rsdept = Db.Execute(sqldept)
    
    With cboDept
        .clear
        .columnCount = 2
        .ColumnWidths = "30pt;120pt"
        .ListWidth = 150
        .ListRows = 15
        
        i = 0
        Do While Not rsdept.EOF
            .AddItem
            .List(i, 0) = Trim(rsdept("Department_cls"))
            .List(i, 1) = IIf(IsNull(rsdept("Description")), "", Trim(rsdept("Description")))
            rsdept.MoveNext
            i = i + 1
        Loop
    End With
    Set rsdept = Nothing
End Sub

Sub adtocboSec()
Dim sqlsec As String
Dim rssec As New Recordset

    sqlsec = "select * from Section_Cls order by Section_cls"
    Set rssec = Db.Execute(sqlsec)
    
    With CboSec
        .clear
        .columnCount = 2
        .ColumnWidths = "30pt;120pt"
        .ListWidth = 150
        .ListRows = 15
        
        i = 0
        Do While Not rssec.EOF
            .AddItem
            .List(i, 0) = Trim(rssec("Section_cls"))
            .List(i, 1) = IIf(IsNull(rssec("Description")), "", Trim(rssec("Description")))
            rssec.MoveNext
            i = i + 1
        Loop
    End With
    Set rssec = Nothing
End Sub

Sub adtocborequestno()
Dim sqlno As String
Dim rsno As New Recordset
    
    sqlno = "select PORequest_no from PORequest_Master " & _
            "where PORequest_date >= '" & Format(requestdate1.Value, "yyyy-mm-dd") & "' " & _
            "and PORequest_date <= '" & Format(requestdate2.Value, "yyyy-mm-dd") & "' " & _
            "and PersonInCharge_Cls = '" & Trim(cboPerson.Text) & "' and others_cls = '0' " & _
            "and isnull(SheetCoil_Cls, '0') = '1' " & _
            "order by PORequest_date desc, PORequest_No desc "
    Set rsno = Db.Execute(sqlno)

    With cborequestno
        .clear
        Do While Not rsno.EOF
            .AddItem Trim(rsno("PORequest_No"))
            rsno.MoveNext
        Loop
        
        .ColumnWidths = "90pt"
        .ListWidth = 90
        .ListRows = 15
    End With
    Set rsno = Nothing
End Sub

Sub requestno(ByVal thn As String, ByVal bln As String)
Dim sqlno As String, SqlS As String
Dim rsno As New Recordset, rsS As New Recordset
    'PRYYMM99999
'    If Format(RequestDate, "YYYY-MM-01") > "2006-07-30" Then
'        sqlno = "select top 1 rtrim(PORequest_No) from PORequest_Master " & _
'                "where substring(rtrim(PORequest_No),3,2) = '" & thn & "' and substring(rtrim(PORequest_No),5,2) > '07' " & _
'                "order by right(rtrim(PORequest_No),5) desc"
'    Else
        sqlno = "select top 1 rtrim(PORequest_No) from PORequest_Master " & _
                "where substring(rtrim(PORequest_No),3,2) = '" & thn & "' " & _
                "order by right(rtrim(PORequest_No),5) desc"
'    End If
    Set rsno = Db.Execute(sqlno)
    If Not (rsno.BOF And rsno.EOF) Then
        txtRequestNo.Text = Left(Trim(rsno(0)), 4) & bln & Format(Right(Trim(rsno(0)), 5) + 1, "0000#")
    Else
''        SqlS = "select top 1 PORequest_No from Initial_No "
''        Set rsS = Db.Execute(SqlS)
''        If Not (rsS.BOF And rsS.EOF) Then
'            txtRequestNo.Text = Left(Trim(rsS(0)), 2) & thn & bln & Right(Trim(rsS(0)), 5)
'        Else
            txtRequestNo.Text = "PR" & thn & bln & "00001"
'        End If
'        Set rsS = Nothing
    End If
    txtRequestNo.locked = True
    Set rsno = Nothing
End Sub

Sub kunci(l As Boolean)
    Period.Enabled = Not l
    RequestDate.Enabled = Not l
    cboDept.Enabled = Not l
    CboSec.Enabled = Not l
    grid.Editable = Not l
    Command1(1).Enabled = Not l
    lblFix.Caption = "Status Fix"
    lblFix.Visible = l
End Sub

Sub Header()
    
    bteColSelect = 0
    bteColProdCode = 1
    bteColSteelKind = 2
    bteColMatT = 3
    bteColMatW = 4
    bteColMatL = 5
    bteColUnitCls = 6
    bteColUnitDes = 7
    bteColUnit = 8
    bteColStock = 9
    bteColOPQty = 10
    bteColFixOrder = 11
    bteColRemReq = 12
    bteColReq = 13
    bteColReqQty = 14
    bteColEnd = 15
    bteColDelDate = 16
    bteColPur = 17
    bteColAcc = 25
        
    With grid
        
        .clear
        .Rows = 2
        .ColS = 26
        
        .TextMatrix(0, bteColSelect) = " "
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColSteelKind) = "Steel Kind"
        .TextMatrix(0, bteColMatT) = "Material Measure"
        .TextMatrix(0, bteColMatW) = "Material Measure"
        .TextMatrix(0, bteColMatL) = "Material Measure"
        .TextMatrix(0, bteColUnitDes) = "Unit Cls"
        .TextMatrix(0, bteColUnit) = "Qty Unit"
        .TextMatrix(0, bteColStock) = "Current Stock"
        .TextMatrix(0, bteColOPQty) = "Order Point Qty"
        .TextMatrix(0, bteColFixOrder) = "Fix Order (Receipt Schd)"
        .TextMatrix(0, bteColRemReq) = "Remaining Request"
        .TextMatrix(0, bteColReq) = "Req"
        .TextMatrix(0, bteColReqQty) = "Request Qty"
        .TextMatrix(0, bteColEnd) = "End Stock"
        .TextMatrix(0, bteColDelDate) = "Delivery Date"
        .TextMatrix(0, bteColPur) = "Purpose"
        .TextMatrix(0, bteColAcc) = "Account No"
        
        .TextMatrix(1, bteColSelect) = " "
        .TextMatrix(1, bteColProdCode) = "Product Code"
        .TextMatrix(1, bteColSteelKind) = "Steel Kind"
        .TextMatrix(1, bteColMatT) = "T"
        .TextMatrix(1, bteColMatW) = "W"
        .TextMatrix(1, bteColMatL) = "L"
        .TextMatrix(1, bteColUnitDes) = "Unit Cls"
        .TextMatrix(1, bteColUnit) = "Qty Unit"
        .TextMatrix(1, bteColStock) = "Current Stock"
        .TextMatrix(1, bteColOPQty) = "Order Point Qty"
        .TextMatrix(1, bteColFixOrder) = "Fix Order (Receipt Schd)"
        .TextMatrix(1, bteColRemReq) = "Remaining Request"
        .TextMatrix(1, bteColReq) = "Req"
        .TextMatrix(1, bteColReqQty) = "Request Qty"
        .TextMatrix(1, bteColEnd) = "End Stock"
        .TextMatrix(1, bteColDelDate) = "Delivery Date"
        .TextMatrix(1, bteColPur) = "Purpose"
        .TextMatrix(1, bteColAcc) = "Account No"
                
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColProdCode) = 2500
        .ColWidth(bteColSteelKind) = 3215
        .ColWidth(bteColMatT) = 1000
        .ColWidth(bteColMatW) = 1000
        .ColWidth(bteColMatL) = 1000
        .ColWidth(bteColUnitDes) = 500
        .ColWidth(bteColUnit) = 585
        .ColWidth(bteColStock) = 1395
        .ColWidth(bteColOPQty) = 1000
        .ColWidth(bteColFixOrder) = 1400
        .ColWidth(bteColRemReq) = 1400
        .ColWidth(bteColReq) = 1000
        .ColWidth(bteColReqQty) = 1000
        .ColWidth(bteColEnd) = 1500
        .ColWidth(bteColDelDate) = 1350
        .ColWidth(bteColPur) = 1500
        .ColWidth(bteColAcc) = 1500
        
        .MergeRow(bteColSelect) = True
        .MergeRow(bteColProdCode) = True
        For i = 0 To .ColS - 1
            .MergeCol(i) = True
        Next i
        .MergeCells = flexMergeFixedOnly
        
        .Cell(flexcpAlignment, 0, 0, 1, .ColS - 1) = flexAlignCenterCenter
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColSteelKind) = flexAlignLeftCenter
        For i = bteColMatT To bteColMatL
            .ColAlignment(i) = flexAlignRightCenter
        Next i
        .ColAlignment(bteColUnit) = flexAlignCenterCenter
        For i = bteColStock To bteColAcc
            .ColAlignment(i) = flexAlignRightCenter
        Next i

        .RowHeight(0) = 225
        .RowHeight(1) = 225
        
        .FrozenCols = bteColMatT
    
        .ColHidden(bteColUnitCls) = True
        
        .ColHidden(18) = True
        .ColHidden(19) = True
        .ColHidden(20) = True
        .ColHidden(21) = True
        .ColHidden(22) = True
        .ColHidden(23) = True
        .ColHidden(24) = True
        .ColHidden(bteColAcc) = True

End With

End Sub

Sub browseitem()
Dim sqlitem As String, RsItem As New ADODB.Recordset
Dim sqlinvcon As String, rsinvcon As New Recordset
Dim tempperiod As Date, closingmonth As Date
Dim nextperiod As Date, endtgl As Byte, endperiod As Date
    
    Me.MousePointer = vbHourglass
    Call Header
    
    tempperiod = Format(Period, "yyyy-mm-01")
    nextperiod = DateAdd("m", 1, Period)
    endtgl = DateDiff("d", Format(Period, "yyyy-mm-01"), Format(nextperiod, "yyyy-mm-01"))
    endperiod = Year(Period) & "-" & Month(Period) & "-" & Format(endtgl, "0#")
    
    'Get Month of Closing
    sqlinvcon = "select * from inventory_control where fix_cls = 1"
    If rsinvcon.State <> adStateClosed Then rsinvcon.Close
    rsinvcon.Open sqlinvcon, Db, adOpenKeyset, adLockOptimistic
    If Not (rsinvcon.BOF And rsinvcon.EOF) Then
        rsinvcon.MoveLast
        closingmonth = Trim(rsinvcon("inventory_year")) & "-" & Format(Trim(rsinvcon("inventory_month")), "0#") & "-01"
    End If
    Set rsinvcon = Nothing
    
    sqlitem = "select *, (curstock + fixorder + remainingrequest - requirement) EndStock ," & _
              "(select description from unit_cls uc where uc.unit_cls= po.unit_cls ) unit_desc  " & _
              "From ( " & _
              "       select stockControl_cls, makebuy_cls,item_code, item_name, accounting_code, unit_cls, finishgoodpart_cls, number_entering, number_box, personincharge_cls, " & _
              "       lot_qty, orderpoint_qty, control_cls, use_endday, " & _
              "       FixOrder = " & _
              "           isnull((select sum(sisaQty) SisaPOQty " & _
              "               From ( " & _
              "                       select Item_Code, " & _
              "                       SisaQty = (case when (qtyPO - isnull(QtyR,0)) < 0 then 0 else (qtyPO - isnull(QtyR,0)) end) " & _
              "                       from ( " & _
              "                          select pod.Item_Code, pod.Qty qtyPO, " & _
              "                          QtyR = ISNULL((select sum(case receipt_cls when 'R1' then -qty else qty end) qty " & _
              "                                 from part_receipt pr Where pod.po_no = pr.po_no and pod.item_code = pr.item_code),0) " & _
              "                          from purchaseOrder_detail pod "
            sqlitem = sqlitem & _
              "                          where isnull(pod.complete_cls,'0') = '0' " & _
              "                                and year(pod.delivery_date)= '" & Year(Period.Value) & "' and month(pod.delivery_date)= '" & Month(Period.Value) & "' " & _
              "                           )dt " & _
              "                     )tbF Where tbF.item_code = item_master.item_code group by item_code " & _
              "           ),0) , "
    sqlitem = sqlitem & _
              "        RemainingRequest = " & _
              "           isnull((select sum(sisaQty) SisaRequestQty " & _
              "               from ( " & _
              "                       Select Item_Code, " & _
              "                       SisaQty = (case when (QtyReq - isnull(QtyPO,0)) < 0 then 0 else (QtyReq - isnull(QtyPO,0)) end) " & _
              "                       From ( " & _
              "                              Select Item_Code, QtyReq = PORD.Qty, " & _
              "                              QtyPO = ISNULL( (select sum(pod.qty) qty from purchaseorder_detail pod " & _
              "                                      where pod.porequest_no = pord.porequest_no and pod.item_code = pord.item_code) ,0) " & _
              "                              from porequest_detail pord INNER JOIN porequest_master porm on pord.porequest_no = porm.porequest_no " & _
              "                              where isnull(porm.fix_cls,'0') = '0' " & _
              "                                    and year(pord.reqdelivery_date)= '" & Year(Period.Value) & "' and month(pord.reqdelivery_date)= '" & Month(Period.Value) & "' "

    If (Format(tempperiod2, "MMM yyyy") <> Format(Period.Value, "MMM yyyy")) Then _
    sqlitem = sqlitem & " and porm.porequest_no <> '" & Trim(txtRequestNo.Text) & "' "
              
    sqlitem = sqlitem & _
              "                             ) dt " & _
              "                     ) tbRR where tbRR.item_code = item_master.item_code group by item_code " & _
              "            ),0), "
    sqlitem = sqlitem & _
              "        Requirement = " & _
              "           isnull((Select sum(sisaReqQty) sisaReqQty " & _
              "                   from ( " & _
              "                          select childItem_code, sum(childRequirement_qty) plans, sum(childRequirementResult_qty) Results, " & _
              "                          (case when sum(childRequirement_qty) - sum(childRequirementResult_qty) - sum(offchildrequirement_qty) < 0 then 0 " & _
              "                           else Sum(childRequirement_qty) - Sum(childRequirementResult_qty) - sum(offchildrequirement_qty) end) As SisaReqQty " & _
              "                          From Requirement " & _
              "                          where isnull(complete_cls,'0') = '0' " & _
              "                                and year(childrequirement_date) = '" & Year(Period.Value) & "' and month(childrequirement_date) = '" & Month(Period.Value) & "' " & _
              "                          group by parentitem_code, factory_code, line_code, production_date, lot_no, " & _
              "                             cast(year(childrequirement_date) as char(4)) + '-' + cast(month(childrequirement_date)as char(2)), " & _
              "                             childItem_code " & _
              "                   )tbR where tbR.childitem_code = item_master.item_code  group by childItem_code " & _
              "           ),0) , "
    sqlitem = sqlitem & _
              "        CurStock  = " & _
              "           isnull((select isnull(tbCS.stockMaster_stock,0) + isnull(tbCS.sisaPOqty,0) +  isnull(tbCS.sisaRequestQty,0) - " & _
              "           isnull(tbCS.sisaReqQty,0) StockMaster_Stock1 " & _
              "               from ( " & _
              "                      Select " & _
              "                      ISNULL((select isnull(case when datediff(month, ClosingDate, StartDate) = 0  then sum(lm_inventory) " & _
              "                                             when datediff(month, ClosingDate, StartDate) = 1  then sum(tm_current) " & _
              "                                             when datediff(month, ClosingDate, StartDate) >= 2 then sum(nm_current) " & _
              "                              end,0) StockMaster_Stock " & _
              "                              From ( " & _
              "                                     select " & _
              "                                     (select cast( cast(year as varchar(4)) + case when month < 10 then '0' else '' end + cast(month as varchar(2)) + '01' as dateTime) ClosingDate " & _
              "                                      from (select top 1 max(inventory_month) month, inventory_year year from inventory_control " & _
              "                                            where fix_cls='1' group by inventory_year  order by inventory_year desc ) tbA " & _
              "                                     )ClosingDate, StartDate = '" & Format(tempperiod, "yyyy-mm-dd") & "', * " & _
              "                                     From stock_master " & _
              "                              ) tbA " & _
              "                              Where tbA.Item_Code = Item_Master.Item_Code group by ClosingDate, Item_code, StartDate " & _
              "                      ),0) StockMaster_Stock, "
    sqlitem = sqlitem & _
              "                      ISNULL((select ISNULL(sum(sisaQty),0) SisaPOQty " & _
              "                              From ( " & _
              "                                     select Item_Code, " & _
              "                                     SisaQty = (case when qtyPO - isnull(QtyR,0) < 0 then 0 else qtyPO - isnull(QtyR,0) end) " & _
              "                                     from ( " & _
              "                                            select pod.Item_Code, pod.Qty qtyPO, " & _
              "                                            QtyR = ISNULL((select sum(case receipt_cls when 'R1' then -qty else qty end)qty " & _
              "                                                           from part_receipt pr Where pod.po_no = pr.po_no and pod.item_code = pr.item_code),0) " & _
              "                                            from purchaseOrder_detail pod "
            sqlitem = sqlitem & _
              "                                            where isnull(pod.complete_cls,'0') = '0' " & _
              "                                                  And pod.delivery_date >= '" & Format(CDate(closingmonth), "yyyy-mm-dd") & "' And pod.delivery_date < '" & Format(tempperiod, "yyyy-mm-dd") & "' " & _
              "                                     )dt " & _
              "                              )tbB Where tbB.Item_Code = Item_Master.Item_Code group by item_code " & _
              "                      ),0) SisaPOQty, "
    sqlitem = sqlitem & _
              "                      ISNULL((select ISNULL(sum(sisaReqQty),0) sisaReqQty " & _
              "                              from ( " & _
              "                                     select childItem_code, sum(childRequirement_qty) plans, sum(childRequirementResult_qty) Results, " & _
              "                                     (case when sum(childRequirement_qty)-sum(childRequirementResult_qty)- sum(offchildrequirement_qty) < 0 then 0 " & _
              "                                     Else: Sum (childRequirement_qty) - Sum(childRequirementResult_qty) - Sum(offchildrequirement_qty) " & _
              "                                     end) as SisaReqQty " & _
              "                                     From Requirement " & _
              "                                     where isnull(complete_cls,'0') = '0' " & _
              "                                           and childrequirement_date >= '" & Format(CDate(closingmonth), "yyyy-mm-dd") & "' and childrequirement_date < '" & Format(tempperiod, "yyyy-mm-dd") & "' " & _
              "                                     group by parentitem_code, factory_code, line_code, production_date, lot_no, childItem_code, " & _
              "                                     cast(year(childrequirement_date) as varchar(4)) + '-' + cast(month(childrequirement_date)as varchar(2)) " & _
              "                              )tbC Where tbC.childItem_code = Item_Master.Item_Code group by childItem_code " & _
              "                      ),0) sisaReqQty, "
    sqlitem = sqlitem & _
              "                      ISNULL((select ISNULL(sum(sisaQty),0) SisaRequestQty " & _
              "                              from ( " & _
              "                                     Select Item_Code, " & _
              "                                     SisaQty = (case when QtyReq - isnull(QtyPO,0) < 0 then 0 else QtyReq - isnull(QtyPO,0) end) " & _
              "                                     From ( " & _
              "                                            Select Item_Code, QtyReq = PORD.Qty, " & _
              "                                            QtyPO = ISNULL((select sum(pod.qty) qty from PurchaseOrder_Detail pod " & _
              "                                                    where pod.porequest_no = pord.porequest_no and pod.item_code = pord.item_code),0) " & _
              "                                            from PORequest_Detail pord INNER JOIN PORequest_Master porm on pord.porequest_no=porm.porequest_no " & _
              "                                            where isnull(porm.fix_cls,'0') = '0' " & _
              "                                                  and pord.ReqDelivery_Date >= '" & Format(CDate(closingmonth), "yyyy-mm-dd") & "' and pord.ReqDelivery_Date < '" & Format(tempperiod, "yyyy-mm-dd") & "' "
              
    If (Format(tempperiod2, "MMM yyyy") <> Format(Period.Value, "MMM yyyy")) Then _
    sqlitem = sqlitem & " and porm.porequest_no <> '" & Trim(txtRequestNo.Text) & "' "
    
    sqlitem = sqlitem & _
              "                                     ) dt " & _
              "                              )tbD Where tbD.Item_Code = Item_Master.Item_Code group by item_code " & _
              "                      ),0) SisaRequestQty " & _
              "               )tbCS " & _
              "           ),0) " & _
              "       From Item_Master " & _
              ") PO "
              
' ----- Tambahan
        sqlitem = sqlitem & " inner join item_master iim on po.item_code=iim.item_code "
        sqlitem = sqlitem & " inner join Sheetcoil_cls Sccm on iim.sheetcoil_cls=Sccm.Sheetcoil_cls "

' --------------
        sqlitem = sqlitem & "where po.use_endday >= '" & Format(endperiod, "yyyymmdd") & "' "

    If cboPerson.Text <> "" Then _
        sqlitem = sqlitem '& " and PersonInCharge_cls='" & Trim(cboPerson.Text) & "' "
    If cboAlarm.Text = "Yes" Then _
        sqlitem = sqlitem & " and (curstock + fixorder + remainingrequest - requirement) < (case po.control_cls when '03' then po.orderpoint_qty else 0 end) "
    
    sqlitem = sqlitem & " and po.makebuy_cls ='02' and po.stockcontrol_cls ='01'"
   
    Set RsItem = Db.Execute(sqlitem)
    If Not (RsItem.BOF And RsItem.EOF) Then
        i = 2
        With grid
        Do While Not RsItem.EOF
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, i, 0) = &HFFFFFF
            .Cell(flexcpChecked, i, 0) = flexUnchecked
            .TextMatrix(i, 1) = Trim(RsItem("Item_Code"))
            .TextMatrix(i, 2) = IIf(IsNull(RsItem("item_name")), "", Trim(RsItem("item_name")))
            .TextMatrix(i, 3) = IIf(IsNull(RsItem("Thickness")), 0, Trim(RsItem("Thickness")))
            .TextMatrix(i, 4) = IIf(IsNull(RsItem("Length")), 0, Trim(RsItem("Length")))
            .TextMatrix(i, 5) = IIf(IsNull(RsItem("Width")), 0, Trim(RsItem("width")))
            
            If IsNull(RsItem("unit_cls")) Then
              .TextMatrix(i, 6) = ""
              .TextMatrix(i, 7) = ""
            Else
              .TextMatrix(i, 6) = Trim(RsItem("Unit_cls"))
              '.TextMatrix(i, 7) = Split(isiunit, ",")(Val(Trim(RsItem("Unit_Cls"))) - 1)
              .TextMatrix(i, 7) = Trim(RsItem("Unit_desc"))
            End If
            If RsItem("finishgoodpart_cls") = "01" Then
                .TextMatrix(i, 8) = IIf(IsNull(RsItem("number_entering")), 0, Format(RsItem("number_entering"), "##,##0"))
            Else
                .TextMatrix(i, 8) = IIf(IsNull(RsItem("number_box")), 0, Format(RsItem("number_box"), "##,##0"))
            End If
            .TextMatrix(i, 9) = IIf(IsNull(RsItem("curstock")), 0, Format(RsItem("curstock"), "##,##0.#0"))
            .TextMatrix(i, 21) = IIf(IsNull(RsItem("curstock")), 0, Format(RsItem("curstock"), "##,##0.#0"))
            .TextMatrix(i, 10) = IIf(IsNull(RsItem("orderpoint_qty")), 0, Format(RsItem("orderpoint_qty"), "##,##0.#0"))
            .TextMatrix(i, 11) = IIf(IsNull(RsItem("fixorder")), 0, Format(RsItem("fixorder"), "##,##0.#0"))
            .TextMatrix(i, 12) = IIf(IsNull(RsItem("remainingrequest")), 0, Format(RsItem("remainingrequest"), "#,##0.#0"))
            .TextMatrix(i, 22) = IIf(IsNull(RsItem("remainingrequest")), 0, Format(RsItem("remainingrequest"), "#,##0.#0"))
            .TextMatrix(i, 13) = IIf(IsNull(RsItem("requirement")), 0, Format(RsItem("requirement"), "##,##0.#0"))
            .TextMatrix(i, 14) = 0
            .TextMatrix(i, 23) = 0
            .Cell(flexcpBackColor, i, 14) = &HFFFFFF
            .TextMatrix(i, 15) = IIf(IsNull(RsItem("endstock")), 0, Format(RsItem("endstock"), "##,##0.#0"))
            .TextMatrix(i, 16) = ""
            .TextMatrix(i, 24) = ""
6            .Cell(flexcpBackColor, i, 16) = vbWhite
            .TextMatrix(i, 17) = ""
            .Cell(flexcpBackColor, i, 17) = vbWhite
            .TextMatrix(i, 18) = ""
            .TextMatrix(i, 19) = ""
            .TextMatrix(i, 20) = 0
            .TextMatrix(i, 25) = IIf(IsNull(RsItem("accounting_code")), "", Trim(RsItem("accounting_code")))
            
            .ColSort(1) = flexSortStringAscending
            RsItem.MoveNext
            i = i + 1
        Loop
        End With
    Else
        LblErrMsg.Caption = DisplayMsg(4006)
    End If
    Me.MousePointer = vbDefault
    Set RsItem = Nothing
End Sub

Sub Browse()
    LblErrMsg = ""

    sql = "select * from PORequest_Master " & _
          "where porequest_no = '" & txtRequestNo.Text & "' and others_Cls = '0'"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic

    If Not (RS.BOF And RS.EOF) Then
        ada = True: ubah = True
        tempperiod2 = IIf(IsNull(RS("porequest_period")), "", Left(Trim(RS("porequest_period")), 4) & "-" & Right(Trim(RS("porequest_period")), 2) & "-01")
        statusfix = IIf(IsNull(RS("fix_cls")), 0, RS("fix_cls"))
        Call browseitem
        Call BrowseGrid
        If statusfix = 1 Then Call kunci(True) Else Call kunci(False)
    Else
        ada = False
    End If
End Sub

Sub BrowseGrid()
Dim g As Integer

    sqlGrid = "select *, (select description from unit_cls uc where uc.unit_cls= PORequest_Detail.unit_cls ) unit_desc from PORequest_Detail where porequest_no = '" & txtRequestNo.Text & "' order by item_code"
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic

    With grid
    Do While Not rsGrid.EOF
        For g = 2 To .Rows - 1
            If Trim(.TextMatrix(g, 1)) = Trim(rsGrid("Item_Code")) Then
                .Cell(flexcpChecked, g, 0) = flexChecked
                .TextMatrix(g, 6) = Trim(rsGrid("Unit_cls"))
                .TextMatrix(g, 7) = Trim(rsGrid("Unit_desc"))
                '.TextMatrix(g, 7) = Split(isiunit, ",")(Val(Trim(rsGrid("Unit_Cls"))) - 1)
                .TextMatrix(g, 14) = IIf(IsNull(rsGrid("qty")), 0, Format(rsGrid("qty"), "##,##0.#0"))
                .TextMatrix(g, 23) = Format(.TextMatrix(g, 14), "##,##0.#0")
                poqty = cekpoqty(.TextMatrix(g, 1), txtRequestNo.Text)
                If (Format(tempperiod2, "MMM yyyy") <> Format(Period.Value, "MMM yyyy")) Then
                    If Year(Period) = Year(rsGrid("ReqDelivery_Date")) And Month(Period) = Month(rsGrid("ReqDelivery_Date")) Then
                        .TextMatrix(g, 12) = Format((CDbl(.TextMatrix(g, 12)) + CDbl(.TextMatrix(g, 14))) - poqty, "##,##0.#0")
                        .TextMatrix(g, 22) = Format(.TextMatrix(g, 12), "##,##0.#0")
                    ElseIf Format(Period, "yyyy-mm-01") > Format(rsGrid("ReqDelivery_Date"), "yyyy-mm-01") Then
                        .TextMatrix(g, 9) = Format((CDbl(.TextMatrix(g, 9)) + CDbl(.TextMatrix(g, 14))) - poqty, "##,##0.#0")
                        .TextMatrix(g, 21) = Format(.TextMatrix(g, 9), "##,##0.#0")
                    End If
                    .TextMatrix(g, 15) = Format((CDbl(.TextMatrix(g, 9)) + CDbl(.TextMatrix(g, 11)) + CDbl(.TextMatrix(g, 12)) - CDbl(.TextMatrix(g, 13))), "##,##0.#0")
                End If
                .TextMatrix(g, 16) = IIf(IsNull(rsGrid("ReqDelivery_Date")), "", Format(rsGrid("ReqDelivery_Date"), "dd MMM yyyy"))
                .TextMatrix(g, 24) = .TextMatrix(g, 16)
                .TextMatrix(g, 17) = IIf(IsNull(rsGrid("Purpose")), "", Trim(rsGrid("Purpose")))
                .TextMatrix(g, 18) = "D"
                .TextMatrix(g, 20) = rsGrid("PoReq_SeqNo")
                .TextMatrix(g, 25) = IIf(IsNull(rsGrid("AccountNo")), "", Trim(rsGrid("AccountNo")))
                .Cell(flexcpBackColor, g, 1, g, .ColS - 1) = &HC0FFC0
                .Cell(flexcpBackColor, g, 14) = vbWhite
                .Cell(flexcpBackColor, g, 16) = vbWhite
                .Cell(flexcpBackColor, g, 17) = vbWhite
            End If
        Next g
        rsGrid.MoveNext
    Loop
    End With
End Sub

Sub BrowseAtas()
Dim p As String

    sql = "select * from PORequest_Master where PORequest_No = '" & txtRequestNo.Text & "' and Others_Cls = '0'"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not (RS.BOF And RS.EOF) Then
        isirequestdate = Format(RequestDate, "yyyy-mm-dd")
        RequestDate.Value = IIf(IsNull(RS("porequest_date")), " ", Format(Trim(RS("porequest_date")), "dd MMM yyyy"))
        p = IIf(IsNull(RS("porequest_period")), " ", Left(Trim(RS("porequest_period")), 4) & "-" & Right(Trim(RS("porequest_period")), 2) & "-01")
        Period.Value = Format(p, "MMM yyyy")
        temptgl = Period.Month
        cboDept.Text = IIf(IsNull(RS("Department_Cls")), "", Trim(RS("Department_Cls")))
        CboSec.Text = IIf(IsNull(RS("Section_Cls")), "", Trim(RS("Section_Cls")))
        statusfix = IIf(IsNull(RS("fix_cls")), 0, RS("fix_cls"))
        If statusfix = 1 Then Call kunci(True) Else Call kunci(False)
    End If
End Sub

Function cekpoqty(ByVal ItemCode As String, ByVal requestno As String) As Double
Dim sqlcekpoqty As String
Dim rscekpoqty As New Recordset
    
    cekpoqty = 0
    sqlcekpoqty = "select pod.item_code, isnull(sum(pod.qty),0) poqty " & _
                  "from PurchaseOrder_Detail pod " & _
                  "where pod.porequest_no = '" & Trim(requestno) & "' " & _
                  "and pod.item_code='" & Trim(ItemCode) & "' " & _
                  "group by pod.item_code "
    If rscekpoqty.State <> adStateClosed Then rscekpoqty.Close
    rscekpoqty.Open sqlcekpoqty, Db, adOpenKeyset, adLockOptimistic
    If Not (rscekpoqty.BOF And rscekpoqty.EOF) Then _
        cekpoqty = CDbl(rscekpoqty("poqty"))
    
    Set rscekpoqty = Nothing
End Function

Function seqNo() As Long
Dim sqlseqno As String
Dim rsseqno As New Recordset

    sqlseqno = "select PoReq_SeqNo from PORequest_Detail order by PoReq_SeqNo desc"
    If rsseqno.State <> adStateClosed Then rsseqno.Close
    rsseqno.Open sqlseqno, Db, adOpenKeyset, adLockOptimistic
    
    If Not (rsseqno.BOF And rsseqno.EOF) Then
        seqNo = rsseqno!POReq_seqno + 1
    Else
        seqNo = 1
    End If
End Function

Private Sub CboSec_Change()
    lblsec.Text = ""
End Sub

Private Sub CboSec_Click()
    If CboSec.ListIndex <> -1 Then _
        lblsec.Text = CboSec.Column(1)

End Sub

Private Sub CboSec_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call CboSec_Click
End Sub

Private Sub Form_Load()
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    combo1.AddItem "Create"
    combo1.AddItem "Update"
    cboAlarm.AddItem "Yes"
    cboAlarm.AddItem "No"
    Call adtocboperson
    Call adtocboDept
    Call adtocboSec
    
    Call Kosong
    combo1.ListIndex = 1
    Header
End Sub

Private Sub cboperson_Change()
    lblPerson.Text = ""
End Sub

Private Sub cboperson_Click()
Dim ketemu As Boolean

    If cboPerson.ListIndex <> -1 Then _
        lblPerson.Text = cboPerson.Column(1)
    
    If combo1.ListIndex = 1 Then
        Call adtocborequestno
        For i = 0 To cborequestno.ListCount - 1
            If txtRequestNo.Text = cborequestno.List(i) Then
                ketemu = True
                cborequestno.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtRequestNo.Text = "": Call Header
    End If
End Sub

Private Sub cboperson_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboperson_Click
End Sub

Private Sub cboDept_Change()
    lblDept.Text = ""
End Sub

Private Sub cboDept_Click()
    If cboDept.ListIndex <> -1 Then _
        lblDept.Text = cboDept.Column(1)
End Sub

Private Sub cboDept_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboDept_Click
End Sub



Private Sub period_Change()
    Call period_Click
    temptgl = Period.Month
    If combo1.ListIndex = 1 Then Call Header
End Sub

Private Sub period_Click()
    If Period.Month = 1 And Val(temptgl) = 12 Then Period.Year = Period.Year + 1
    If Period.Month = 12 And Val(temptgl) = 1 Then Period.Year = Period.Year - 1
End Sub

Private Sub requestdate_Change()
    Dim t As String
    If combo1.ListIndex = 0 Then
        t = Right(Year(RequestDate), 2) & "-" & Format(Month(RequestDate), "0#")
        Call requestno(Right(Year(RequestDate), 2), Format(Month(RequestDate), "0#"))
    End If
End Sub

Private Sub requestdate1_Change()
Dim ketemu As Boolean
    
    LblErrMsg.Caption = ""
    If Format(requestdate1, "yyyy-mm-dd") > Format(requestdate2, "yyyy-mm-dd") Then
       LblErrMsg.Caption = DisplayMsg(4025) & " " & Format(requestdate2, "MMM yyyy")    '"Start Date must be lower than "
       Exit Sub
    End If

    If combo1.ListIndex = 1 Then
        Call adtocborequestno
        For i = 0 To cborequestno.ListCount - 1
            If txtRequestNo.Text = cborequestno.List(i) Then
                ketemu = True
                cborequestno.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtRequestNo.Text = "": Call Header
    End If
End Sub

Private Sub requestdate2_Change()
Dim ketemu As Boolean

    LblErrMsg.Caption = ""
    If Format(requestdate2, "yyyy-mm-01") < Format(requestdate1, "yyyy-mm-01") Then
       LblErrMsg.Caption = DisplayMsg(4024) & " " & Format(requestdate1, "MMM yyyy")    '"End Date must be higher than "
       Exit Sub
    End If

    If combo1.ListIndex = 1 Then
        Call adtocborequestno
        For i = 0 To cborequestno.ListCount - 1
            If txtRequestNo.Text = cborequestno.List(i) Then
                ketemu = True
                cborequestno.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtRequestNo.Text = "": Call Header
    End If
End Sub

Private Sub Combo1_Click()
Dim ketemu As Boolean, t As String

    ketemu = False
    LblErrMsg = ""
    Call kunci(False)
'    Call header

    If combo1.ListIndex = 0 Then    'CREATE
        Command1(0).Caption = "&Create"
        ubah = False
        requestdate1.Enabled = False
        requestdate2.Enabled = False
        RequestDate.Value = Format(Now, "dd MMM yyyy")
        RequestDate.Enabled = False
        cborequestno.locked = True
        txtRequestNo.Text = ""
        t = Right(Year(RequestDate), 2) & "-" & Format(Month(RequestDate), "0#")
        Call requestno(Right(Year(RequestDate), 2), Format(Month(RequestDate), "0#"))
    Else    'UPDATE
        Call adtocborequestno
        Command1(0).Caption = "&Update"
        ubah = True
        cborequestno.locked = False
        txtRequestNo.locked = False
        requestdate1.Enabled = True
        requestdate2.Enabled = True
        RequestDate.Enabled = True

        For i = 0 To cborequestno.ListCount - 1
            If txtRequestNo.Text = cborequestno.List(i) Then
                ketemu = True
                cborequestno.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtRequestNo.Text = ""
    End If
End Sub

Private Sub combo1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call Combo1_Click
End Sub

Private Sub cborequestno_Click()
    LblErrMsg = ""
    txtRequestNo.Text = cborequestno.Text
    Call Header
    Call BrowseAtas
End Sub

Private Sub cborequestno_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cborequestno_Click
End Sub

Private Sub txtrequestno_Change()
Dim ketemu As Boolean
    
    LblErrMsg = ""
    If combo1.ListIndex = 1 Then
        For i = 0 To cborequestno.ListCount - 1
            If txtRequestNo.Text = cborequestno.List(i) Then
                ketemu = True
                cborequestno.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then cborequestno.ListIndex = -1
    End If
End Sub

Private Sub txtrequestno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
    
    If KeyAscii = 13 Then
        If combo1.ListIndex = 0 Then
            SendKeys vbTab
        Else
            Call Header
            Call BrowseAtas
        End If
    End If
End Sub

Private Sub cboalarm_Change()
    If combo1.ListIndex = 1 Then Call Header
End Sub

Private Sub deldate_GotFocus()
    OldDelDate = DelDate.Value
End Sub

Private Sub deldate_Change()
    LblErrMsg = ""
    With grid
        .TextMatrix(.Row, 16) = Format(DelDate, "dd mmm yyyy")
        If Trim(.TextMatrix(.Row, 16)) <> "" And IsDate(.TextMatrix(.Row, 16)) Then
            If Trim(.TextMatrix(.Row, 24)) <> "" Then   'Update
                If Year(Period) = Year(.TextMatrix(.Row, 24)) And Month(Period) = Month(.TextMatrix(.Row, 24)) Then
                    .TextMatrix(.Row, 12) = Format((CDbl(.TextMatrix(.Row, 22)) - CDbl(.TextMatrix(.Row, 23))), "##,##0.#0")
                    .TextMatrix(.Row, 9) = Format(CDbl(.TextMatrix(.Row, 21)), "##,##0.#0")
                ElseIf Format(Period, "yyyy-mm-01") > Format(.TextMatrix(.Row, 24), "yyyy-mm-01") Then
                    .TextMatrix(.Row, 9) = Format((CDbl(.TextMatrix(.Row, 21)) - CDbl(.TextMatrix(.Row, 23))), "##,##0.#0")
                    .TextMatrix(.Row, 12) = Format(CDbl(.TextMatrix(.Row, 22)), "##,##0.#0")
                Else
                    .TextMatrix(.Row, 9) = Format(.TextMatrix(.Row, 21), "#,##0.#0")
                    .TextMatrix(.Row, 12) = Format(.TextMatrix(.Row, 22), "#,##0.#0")
                End If
            Else    'Insert
                If Year(Period) = Year(OldDelDate) And Month(Period) = Month(OldDelDate) Then
                    .TextMatrix(.Row, 12) = Format((CDbl(.TextMatrix(.Row, 12)) - CDbl(.TextMatrix(.Row, 14))), "##,##0.#0")
                ElseIf Format(Period, "yyyy-mm-01") > Format(OldDelDate, "yyyy-mm-01") Then
                    .TextMatrix(.Row, 9) = Format((CDbl(.TextMatrix(.Row, 9)) - CDbl(.TextMatrix(.Row, 14))), "##,##0.#0")
                End If
            End If
            
            If Year(Period) = Year(.TextMatrix(.Row, 16)) And Month(Period) = Month(.TextMatrix(.Row, 16)) Then
                .TextMatrix(.Row, 12) = Format((CDbl(.TextMatrix(.Row, 12)) + CDbl(.TextMatrix(.Row, 14))), "##,##0.#0")
            ElseIf Format(Period, "yyyy-mm-01") > Format(.TextMatrix(.Row, 13), "yyyy-mm-01") Then
                .TextMatrix(.Row, 9) = Format((CDbl(.TextMatrix(.Row, 9)) + CDbl(.TextMatrix(.Row, 14))), "##,##0.#0")
            End If
            .TextMatrix(.Row, 15) = Format((CDbl(.TextMatrix(.Row, 9)) + CDbl(.TextMatrix(.Row, 11)) + CDbl(.TextMatrix(.Row, 12)) - CDbl(.TextMatrix(.Row, 13))), "##,##0.#0")
        End If
    End With
    OldDelDate = DelDate.Value
End Sub

Private Sub deldate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub deldate_LostFocus()
    DelDate.Visible = False
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim tempd As Date

    With grid
        If Col = 14 Then
            If .TextMatrix(Row, 14) = "" Then .TextMatrix(Row, 14) = 0
            If IsNumeric(.TextMatrix(Row, 14)) = False Then .TextMatrix(Row, 14) = 0
            .TextMatrix(Row, 14) = Format(.TextMatrix(Row, 14), "#,##0.#0")
            If CDbl(.TextMatrix(Row, 14)) > 9999999.99 Then LblErrMsg = DisplayMsg(4045) & " 9,999,999.99": hscrollbar.Value = 0: .SetFocus: Exit Sub '"Quantity must be lower or equal than 9,999,999.99"
            
            If Trim(.TextMatrix(Row, 16)) <> "" And IsDate(.TextMatrix(Row, 16)) Then
                tempd = Format(.TextMatrix(Row, 16), "yyyy-mm-dd")
            Else
                tempd = Format(DelDate, "yyyy-mm-dd")
            End If
            If Year(Period) = Year(tempd) And Month(Period) = Month(tempd) Then
                .TextMatrix(Row, 12) = Format((CDbl(.TextMatrix(Row, 12)) + CDbl(.TextMatrix(Row, 14)) - orderawal), "##,##0.#0")
            ElseIf Format(Period, "yyyy-mm-01") > Format(tempd, "yyyy-mm-01") Then
                .TextMatrix(Row, 9) = Format((CDbl(.TextMatrix(Row, 9)) + CDbl(.TextMatrix(Row, 14)) - orderawal), "##,##0.#0")
            End If
            .TextMatrix(Row, 15) = Format((CDbl(.TextMatrix(Row, 9)) + CDbl(.TextMatrix(Row, 11)) + CDbl(.TextMatrix(Row, 12)) - CDbl(.TextMatrix(Row, 13))), "##,##0.#0")
        End If
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Counti = 0
    For i = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
            Counti = Counti + 1
        End If
    Next i
    If Counti >= 5 Then
        If grid.Cell(flexcpChecked, Row, 0) = flexUnchecked Then Cancel = True
    Else
        If grid.Cell(flexcpChecked, Row, 0) <> flexChecked Then
            If Col <> 0 Then Cancel = True
        Else
            If Col <> 0 And Col <> 14 And Col <> 16 And Col <> 17 Then Cancel = True
            If Col = 14 Then orderawal = CDbl(grid.TextMatrix(Row, 14))
            poqty = cekpoqty(grid.TextMatrix(Row, 1), txtRequestNo.Text)
        End If
    End If
End Sub

Private Sub Grid_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 14 Then _
        If InStr(1, grid.TextMatrix(Row, Col), ",") = 1 Then grid.TextMatrix(Row, Col) = Right(grid.TextMatrix(Row, Col), Len(grid.TextMatrix(Row, Col)) - 1)
End Sub

Private Sub grid_Click()
    LblErrMsg.Caption = ""
    With grid
        If statusfix = 0 Then
            If .Row > 1 Then
                If Counti >= 5 And .Col = 0 Then
                    If .Cell(flexcpChecked, .Row, 0) = flexUnchecked Then
                        LblErrMsg = "Can't select more than 5 products"
                        Exit Sub
                    End If
                End If
            
                If .Cell(flexcpChecked, .Row, 0) = flexChecked Then
                    If .Col = 14 Or .Col = 16 Or .Col = 17 Then
                        .SelectionMode = flexSelectionFree
                        .FocusRect = flexFocusInset
                    Else
                        .SelectionMode = flexSelectionByRow
                        .FocusRect = flexFocusNone
                    End If
                
                    If .Col = 16 Then
                        OldDelDate = DelDate.Value
                        DelDate.top = .Cell(flexcpTop, .Row, 16)
                        DelDate.Left = .Cell(flexcpLeft, .Row, 16)
                        DelDate.Width = .CellWidth + 30
                        DelDate.Height = .CellHeight + 30
                        If .TextMatrix(.Row, 16) <> "" Then
                            DelDate.Value = Format(.TextMatrix(.Row, 16), "yyyy-mm-dd")
                        Else
                            .TextMatrix(.Row, 16) = Format(DelDate, "dd mmm yyyy")
                            Call deldate_Change
                        End If
                        DelDate.Visible = True
                        DelDate.SetFocus
                    Else
                        DelDate.Visible = False
                    End If
                Else
                    .SelectionMode = flexSelectionByRow
                    .FocusRect = flexFocusNone
                End If
                
            End If
        End If
    End With
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    LblErrMsg = ""
    If Col = 14 Then
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
            KeyAscii = 0
    End If
End Sub

Private Sub Grid_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    DelDate.Visible = False
End Sub

Private Sub grid_AfterSort(ByVal Col As Long, Order As Integer)
    DelDate.Visible = False
End Sub

Private Sub Grid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call grid_Click
End Sub

Private Sub Command1_Click(Index As Integer)
Dim sql3 As String, sql4 As String, t As String
Dim rs4 As New Recordset

    LblErrMsg = ""
    Select Case Index
    Case 0: 'CREATE / UPDATE
            If hakUpdate(Me.Name) = 0 Then _
                LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
            'HEADER VALIDATION
            If cboPerson.Text = "" Then
                LblErrMsg = DisplayMsg(1070)
                cboPerson.SetFocus
                Exit Sub
            End If
            
            If cboPerson.Text <> "" Then
                If cboPerson.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4086)    'Record with this Person in Charge not found
                    cboPerson.SetFocus
                    Exit Sub
                End If
            Else
                LblErrMsg = "Please select Person in Charge first!"
                cboPerson.SetFocus
            End If
            
            If cboDept.Text <> "" Then
                If cboDept.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4142)    'Record with this Department not found
                    cboDept.SetFocus
                    Exit Sub
                End If
            Else
                LblErrMsg = "Please select department first!"
                cboDept.SetFocus
                Exit Sub
            End If
            
            If CboSec.Text <> "" Then
                If CboSec.MatchFound = False Then
                    LblErrMsg = "[4142] Record with this Section not found"
                    CboSec.SetFocus
                    Exit Sub
                End If
            Else
                LblErrMsg = "Please select section first!"
                CboSec.SetFocus
                Exit Sub
            End If
            
            If txtRequestNo.Text = "" Then
                LblErrMsg = DisplayMsg(1067) '"Please Input Request No"
                txtRequestNo.SetFocus
                Exit Sub
            End If
            
            If combo1.ListIndex = 0 Then    'CREATE
                If ubah = False Then
                    sql = "select * from PORequest_Master where porequest_no = '" & txtRequestNo.Text & "' "
                    If RS.State <> adStateClosed Then RS.Close
                    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
                    If Not (RS.BOF And RS.EOF) Then
                        LblErrMsg.Caption = DisplayMsg(1023)
                        txtRequestNo.SetFocus
                        Exit Sub
                    Else
                        RS.AddNew
                        RS("PORequest_No") = txtRequestNo.Text
                    End If
                End If
                RS("PORequest_Period") = Year(Period.Value) & Format(Month(Period.Value), "0#")
                RS("PORequest_Date") = Format(RequestDate.Value, "yyyy-mm-dd")
                RS("PersonInCharge_Cls") = Trim(cboPerson.Text)
                RS("Department_Cls") = Trim(cboDept.Text)
                RS("Section_Cls") = Trim(CboSec.Text)
                RS("Others_Cls") = "0"
                RS("SheetCoil_Cls") = "1"
                RS("Username") = userLogin
                RS("Last_Update") = Format(Now, "yyyy-mm-dd hh:mm:ss")
'On Error Resume Next
                RS.update
errHandler:
                If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                    t = Right(Year(RequestDate), 2) & "-" & Format(Month(RequestDate), "0#")
                    Call requestno(Right(Year(RequestDate), 2), Format(Month(RequestDate), "0#"))
                    RS("porequest_No") = txtRequestNo.Text
                    RS.update
                    If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                        GoTo errHandler
                    Else
                        If Trim$(err.Description) <> "" Then
                            LblErrMsg = Trim$(err.number) + " : " + Trim$(err.Description)
                            Exit Sub
                        End If
                    End If
                Else
                    If Trim$(err.Description) <> "" Then
                        LblErrMsg = Trim$(err.number) + " : " + Trim$(err.Description)
                        Exit Sub
                    End If
                End If
                
                If CDate(RequestDate.Value) > CDate(requestdate1.Value) Then
                    If CDate(RequestDate.Value) > CDate(requestdate2.Value) Then _
                        requestdate2.Value = Format(RequestDate.Value, "dd MMM yyyy")
                Else
                    requestdate1.Value = Format(RequestDate.Value, "dd MMM yyyy")
                End If
    
                combo1.Text = "Update"
                Call browseitem
                LblErrMsg.Caption = DisplayMsg(1000)
                ubah = True
            Else    'UPDATE
                Call Browse
                If ada = False Then
                    LblErrMsg.Caption = DisplayMsg(4144)    'Record with this Request No not found
                    txtRequestNo.SetFocus
                    Exit Sub
                End If
            End If

    Case 1: 'SUBMIT
            If hakUpdate(Me.Name) = 0 Then _
                LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
            
            'HEADER VALIDATION
            If cboPerson.Text = "" Then
                LblErrMsg = DisplayMsg(1070)
                cboPerson.SetFocus
                Exit Sub
            End If
            If cboPerson.Text <> "" Then
                If cboPerson.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4086)    'Record with this Person in Charge not found
                    cboPerson.SetFocus
                    Exit Sub
                End If
            End If
            If cboDept.Text <> "" Then
                If cboDept.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4142)    'Record with this Department not found
                    cboDept.SetFocus
                    Exit Sub
                End If
            End If
            
            If CboSec.Text <> "" Then
                If CboSec.MatchFound = False Then
                    LblErrMsg = "[4142] Record with this Section not found "
                    CboSec.SetFocus
                    Exit Sub
                End If
            End If

            If txtRequestNo.Text = "" Then
                LblErrMsg = DisplayMsg(1067) '"Please Input Request No"
                txtRequestNo.SetFocus
                Exit Sub
            End If
              
            sql = "select * from PORequest_Master where PORequest_No = '" & txtRequestNo.Text & "' and Others_Cls = '0'"
            If RS.State <> adStateClosed Then RS.Close
            RS.Open sql, Db, adOpenKeyset, adLockOptimistic
            If RS.BOF And RS.EOF Then
                LblErrMsg.Caption = DisplayMsg(4144)
                txtRequestNo.SetFocus
                Exit Sub
            End If

            If ubah = True Then
                RS("PORequest_Period") = Year(Period.Value) & Format(Month(Period.Value), "0#")
                RS("PORequest_Date") = Format(RequestDate.Value, "yyyy-mm-dd")
                RS("PersonInCharge_Cls") = Trim(cboPerson.Text)
                RS("Department_Cls") = Trim(cboDept.Text)
                RS("Section_Cls") = Trim(CboSec.Text)
                RS("Others_Cls") = "0"
                RS("Username") = userLogin
                RS("Last_Update") = Format(Now, "yyyy-mm-dd hh:mm:ss")
                RS.update

                With grid
                    'DETAIL VALIDATION
                    For i = 2 To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = flexChecked Then
                            If .TextMatrix(i, 14) = 0 Then
                                hscrollbar.Value = 0
                                .Col = 14
                                .SelectionMode = flexSelectionFree
                                .Row = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(1012) '"Please Input Quantity"
                                Exit Sub
                            ElseIf CDbl(.TextMatrix(i, 14)) > 9999999.99 Then
                                hscrollbar.Value = 0
                                .Col = 14
                                .SelectionMode = flexSelectionFree
                                .Row = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(4045) & " 9,999,999.99" '"Quantity must be lower or equal than 9,999,999.99"
                                Exit Sub
                            End If
                            
                            poqty = cekpoqty(.TextMatrix(i, 1), txtRequestNo.Text)
                            If CDbl(.TextMatrix(i, 14)) < poqty Then
                                hscrollbar.Value = 0
                                .Col = 14
                                .SelectionMode = flexSelectionFree
                                .Row = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(4036) & " " & poqty '"Quantity must be higher or equal than "
                                Exit Sub
                            End If
                            
                            If Trim(.TextMatrix(i, 16)) = "" Then
                                hscrollbar.Value = 1
                                .Col = 16
                                .SelectionMode = flexSelectionFree
                                .Row = i
                                .SetFocus
                                Call grid_Click
                                LblErrMsg = DisplayMsg(1096)    '"Please Input Request Delivery Date"
                                Exit Sub
                            End If
                            .TextMatrix(i, 19) = "S"
                        Else
                            sql4 = "select * from PurchaseOrder_Detail pd " & _
                                   "where pd.PORequest_No = '" & txtRequestNo.Text & "' and pd.POReq_SeqNo = '" & .TextMatrix(i, 20) & "' "
                            Set rs4 = Db.Execute(sql4)
                            If Not (rs4.BOF And rs4.EOF) Then
                                hscrollbar.Value = 0
                                .Row = i
                                .SetFocus
                                .Col = 0
                                LblErrMsg = DisplayMsg(1204)
                                Exit Sub
                            End If
                        End If
                    Next i
                                                            
                    For i = 2 To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = flexChecked Then
                            If Trim(.TextMatrix(i, 19) = "S") Then
                                sqlGrid = "select * from PORequest_Detail " & _
                                          "where PORequest_No = '" & txtRequestNo.Text & "' " & _
                                          "and Item_Code = '" & Trim(.TextMatrix(i, 1)) & "' " & _
                                          "and PoReq_SeqNo = " & IIf(Trim(.TextMatrix(i, 20)) = "", 0, .TextMatrix(i, 20)) & _
                                          " order by item_code"
                                Set rsGrid = Nothing
                                If rsGrid.State <> adStateClosed Then rsGrid.Close
                                rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
                                If rsGrid.BOF And rsGrid.EOF Then
                                    rsGrid.AddNew
                                    rsGrid("PoReq_SeqNo") = seqNo
                                    rsGrid("porequest_no") = Trim(txtRequestNo.Text)
                                    rsGrid("item_Code") = .TextMatrix(i, 1)
                                End If
                                rsGrid("unit_cls") = .TextMatrix(i, 6)
                                rsGrid("qty") = .TextMatrix(i, 14)
                                rsGrid("ReqDelivery_Date") = Format(.TextMatrix(i, 16), "yyyy-mm-dd")
                                rsGrid("Purpose") = Trim(.TextMatrix(i, 17))
                                rsGrid("username") = userLogin
                                rsGrid("last_update") = Format(Now, "yyyy-mm-dd hh:mm:ss")
                                rsGrid("accountno") = Trim(.TextMatrix(i, 22))
                                rsGrid.update
                            End If
                        Else
                            If Trim(.TextMatrix(i, 18) = "D") Then
                                sql3 = "delete from PORequest_Detail " & _
                                       "where PORequest_No = '" & txtRequestNo.Text & "' " & _
                                       "and Item_Code = '" & Trim(.TextMatrix(i, 1)) & "' " & _
                                       "and PoReq_SeqNo = " & .TextMatrix(i, 20)
                                Db.Execute sql3
                            End If
                        End If
                    Next i
                    
                    Call Browse
                End With
                LblErrMsg = DisplayMsg(1101)
            End If

    Case 2: 'CLEAR
            Call Kosong
            combo1.ListIndex = 1
            'Call Combo1_Click

    Case 3: 'CANCEL
            If Trim(txtRequestNo.Text) <> "" Then
                If cboPerson.Text = "" Then
                    LblErrMsg = DisplayMsg(1070)
                    cboPerson.SetFocus
                    Exit Sub
                End If
                If cboPerson.Text <> "" Then
                    If cboPerson.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4086)    'Record with this Person in Charge not found
                        cboPerson.SetFocus
                        Exit Sub
                    End If
                End If
                If cboDept.Text <> "" Then
                    If cboDept.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4142)    'Record with this Department not found
                        cboDept.SetFocus
                        Exit Sub
                    End If
                End If
                If CboSec.Text <> "" Then
                    If CboSec.MatchFound = False Then
                        LblErrMsg = "[4142]  Record with this Section not found"
                        CboSec.SetFocus
                        Exit Sub
                    End If
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
  
    If combo1.ListIndex = 1 And txtRequestNo.Text <> "" Then
        sqlcekdet = "select prd.porequest_no from PORequest_Detail prd " & _
                    "inner join PORequest_Master prm on prm.porequest_no = prd.porequest_no " & _
                    "where prd.porequest_no = '" & txtRequestNo.Text & "' "
        Set rscekdet = Db.Execute(sqlcekdet)
        If Not (rscekdet.BOF And rscekdet.EOF) Then
            Me.MousePointer = vbHourglass
            
'            SqlRpt = "select rtrim(pm.porequest_no) porequest_no, pm.porequest_Date, pm.department_cls, pm.section_cls,rtrim(dc.description) Department,rtrim(Scc.description) Section, " & _
'                     "pd.PoReq_seqno, rtrim(pd.item_code) item_code, iim.thickness, iim.width, iim.length, rtrim(im.item_name) item_name, rtrim(pd.class) class, isnull(pd.qty,0) qty, pd.unit_cls, pd.ReqDelivery_Date, " & _
'                     "rtrim(pd.Purpose) Purpose, rtrim(im.Accounting_Code) Accounting_Code, " & _
'                     "(select upper(rtrim(company_name)) from Company_Profile) comp_name, " & _
'                     "ISNULL((select isnull(case when datediff(month, ClosingDate, StartDate) = 0  then sum(lm_inventory) " & _
'                     "             when datediff(month, ClosingDate, StartDate) = 1  then sum(tm_current) " & _
'                     "             when datediff(month, ClosingDate, StartDate) >= 2 then sum(nm_current) end,0) Stock " & _
'                     "    From ( select " & _
'                     "            (select cast( cast(year as varchar(4)) + case when month < 10 then '0' else '' end + cast(month as varchar(2)) + '01' as dateTime) ClosingDate " & _
'                     "        from (select top 1 max(inventory_month) month, inventory_year year from inventory_control " & _
'                     "            where fix_cls='1' group by inventory_year order by inventory_year desc ) tbI " & _
'                     "            ) ClosingDate, StartDate = '" & Format(Now, "yyyy-mm-dd") & "', * " & _
'                     "        From Stock_Master " & _
'                     "    ) tbS " & _
'                     "    Where tbS.Item_Code = pd.Item_Code group by ClosingDate, Item_code, StartDate " & _
'                     "),0) Current_Stock, " & _
'                     "isnull((select top 1 p.Currency_Code from Price_Master p where p.Price_Cls='01' and p.Item_Code=pd.Item_Code " & _
'                     "and p.Trade_Code = im.Supplier_Code And p.Start_Date <= pm.PORequest_Date1 And p.End_Date >= pm.PORequest_Date1 " & _
'                     "order by p.Priority_cls desc),'') Currency_Code, " & _
'                     "isnull((select top 1 p.Price from Price_Master p where p.Price_Cls='01' and p.Item_Code=pd.Item_Code " & _
'                     "and p.Trade_Code = im.Supplier_Code And p.Start_Date <= pm.PORequest_Date1 And p.End_Date >= pm.PORequest_Date1 " & _
'                     "order by p.Priority_cls desc),0) Price, pm.Others_cls, rtrim((select isnull(PO_Person,'') from Company_Profile)) PO_Person "
'            SqlRpt = SqlRpt & _
'                     "from (select *, " & _
'                     "     cast(year(PORequest_date) as char(4)) + " & _
'                     "     cast((case when month(PORequest_date) < 10 then '0' else '' end) + cast(month(PORequest_date) as char) as char(2)) + " & _
'                     "     cast((case when day(PORequest_date) < 10 then '0' else '' end) + cast(day(PORequest_date) as char) as char(2)) " & _
'                     "     PORequest_Date1 " & _
'                     "from PORequest_Master) pm " & _
'                     "inner join PORequest_Detail pd on pd.porequest_no = pm.porequest_no " & _
'                     "left join (select item_code, item_name, accounting_Code, supplier_Code from Item_Master) im on im.item_code = pd.item_code " & _
'                     "left join Department_Cls dc on dc.Department_Cls = pm.Department_Cls " & _
'                     "left join Section_Cls Scc on Scc.Section_Cls = pm.Section_Cls "
'' ------ Tambahan
'                    SqlRpt = SqlRpt & " inner join item_master iim on pd.item_code=iim.item_code "
'                    SqlRpt = SqlRpt & " inner join Sheetcoil_cls Sccm on iim.sheetcoil_cls=Sccm.Sheetcoil_cls "
'' ------------------
'                    SqlRpt = SqlRpt & "where pm.PORequest_No = '" & Trim(txtRequestNo.Text) & "' and pm.Others_cls = '0' " & _
'                     "order by pd.item_code, pd.reqdelivery_date "
                     
                     
' PR Report New For Musashi 20090109
' Stock diambil dari Stock Master berdasarkan Inventory_Control

        Dim rsclosing As ADODB.Recordset
        Dim CloseThn As Long
        Dim CloseBln As Long
        Dim selisih As Long
        Dim FPilih As String
        
        Set rsclosing = New ADODB.Recordset
        rsclosing.Open "select * from inventory_control " & _
                              " Where Inventory_Month=(Select Max(Inventory_Month) from Inventory_Control " & _
                              " Where inventory_Year=(Select Max(Inventory_Year) from Inventory_Control)) ", Db
    
        CloseThn = rsclosing(0)
        CloseBln = rsclosing(1)
            
        selisih = (Year(RequestDate) * 12 + Month(RequestDate)) - (CloseThn * 12 + CloseBln)
        
        If selisih < 0 And selisih > 2 Then
            LblErrMsg = "Invalid Date Periode !"
        Else
            If selisih = 0 Then
                FPilih = "LM_Current"
            ElseIf selisih = 1 Then
                FPilih = "TM_Current"
            Else
                FPilih = "NM_Current"
            End If
        End If

SqlRpt = " Select PRM.PoRequest_No, PRM.PoRequest_Date, " & _
            vbLf & " PRM.Department_Cls,D.Description Department, " & _
            vbLf & " PRM.Section_Cls,S.Description Section, " & _
            vbLf & " PRM.PersonInCharge_Cls,P.Description PIC, " & _
            vbLf & " PRD.Item_Code, IM.Item_Name,IM.WH_Code,IM.Supplier_Code,IM.Control_Cls, " & _
            vbLf & " TM.Trade_Name Supplier_Name, " & _
            vbLf & " isnull(SM." & FPilih & ",0) Stock, " & _
            vbLf & " PRD.Qty,PRD.ReqDelivery_Date,dateadd(month,1,PRD.ReqDelivery_Date) BlnF1,dateadd(month,2,PRD.ReqDelivery_Date) BlnF2, " & _
            vbLf & " isnull((Select ChildRequirement_Qty from requirement_Master Where ChildRequirement_Month=month(PRD.ReqDelivery_Date)+1 and ChildRequirement_Year=year(PRD.ReqDelivery_Date) and ChildItem_Code=PRD.Item_code),0) F1, " & _
            vbLf & " isnull((Select ChildRequirement_Qty from requirement_Master Where ChildRequirement_Month=month(PRD.ReqDelivery_Date)+2 and ChildRequirement_Year=year(PRD.ReqDelivery_Date) and ChildItem_Code=PRD.Item_code),0) F2 " & _
            vbLf & " From PoRequest_Master PRM inner Join PORequest_Detail PRD " & _
            vbLf & " On PRM.PoREquest_No=PRD.PoRequest_No " & _
            vbLf & " Inner Join Department_Cls D on PRM.Department_Cls=D.Department_Cls " & _
            vbLf & " Inner Join Section_Cls S on PRM.Section_Cls=S.Section_Cls " & _
            vbLf & " Inner Join Item_Master IM on PRD.Item_Code=IM.Item_Code " & _
            vbLf & " Inner Join Trade_Master TM on IM.Supplier_Code=TM.Trade_Code " & _
            vbLf & " Inner Join PersonInCharge_Cls P on PRM.PersonInCharge_Cls=P.PersonInCharge_Cls " & _
            vbLf & " Left Join Stock_Master SM on PRD.Item_Code=SM.Item_Code and IM.WH_Code=SM.WareHouse_Code "
' ------ Tambahan
                    SqlRpt = SqlRpt & " inner join Sheetcoil_cls Sccm on im.sheetcoil_cls=Sccm.Sheetcoil_cls "
' ------------------
    SqlRpt = SqlRpt & _
            vbLf & " where PRM.PORequest_No = '" & Trim(txtRequestNo.Text) & "' and PRM.Others_cls = '0' " & _
            vbLf & " order by PRD.Item_Code, PRD.ReqDelivery_Date"
                     
' -----
                     
                     
            If rsRpt.State <> adStateClosed Then rsRpt.Close
            rsRpt.Open SqlRpt, Db, adOpenDynamic, adLockOptimistic
            
            sqlprint = SqlRpt
            reportcode = "PORequestSteel"
            printorient = 2
            
            If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
            
            Set report = application.OpenReport(App.path & "\Reports\rptPORequestDirectNewGroup.rpt")
            report.Database.Tables(1).SetDataSource rsRpt
            report.PaperOrientation = crLandscape
            report.PaperOrientation = crLandscape
            Rpt.CRViewer1.ReportSource = report
            Rpt.CRViewer1.ViewReport
            Rpt.CRViewer1.Zoom 1
            Rpt.WindowState = 2
            Rpt.Show 1
            Me.MousePointer = vbDefault
        Else
            LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault
        End If
    End If

    Set rscekdet = Nothing
    Set rsRpt = Nothing
End Sub

Private Sub CmdSubMenu_Click()
    sql = "delete from PORequest_Master " & _
          "where porequest_no not in (select porequest_no from PORequest_Detail) and others_cls = '0'"
    Db.Execute sql

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

Private Sub hscrollbar_Change()
'Dim k As Integer
'
'    For k = 6 To grid.ColS - 1
'       grid.ColHidden(k) = False
'    Next k
'
'    If hscrollbar.Value = 1 Then
'        For k = 6 To 15
'            grid.ColHidden(k) = True
'        Next k
'    End If
'    DelDate.Visible = False
'
'    grid.ColHidden(bteColUnitCls) = True
'    grid.ColHidden(18) = True
'    grid.ColHidden(19) = True
'    grid.ColHidden(20) = True 'Seq No
'    grid.ColHidden(21) = True 'Old Current Stock
'    grid.ColHidden(22) = True 'Old Remaining Req
'    grid.ColHidden(23) = True 'Old Request Qty
'    grid.ColHidden(24) = True 'Old Delivery Date
End Sub


Private Sub hscrollbar_Scroll()
'    Call hscrollbar_Change
End Sub



