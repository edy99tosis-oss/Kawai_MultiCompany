VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmNotaRetur 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota Retur"
   ClientHeight    =   11040
   ClientLeft      =   105
   ClientTop       =   345
   ClientWidth     =   15180
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmNotaRetur.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11040
   ScaleWidth      =   15180
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrev 
      BackColor       =   &H0080FFFF&
      Caption         =   "Pre&view"
      Height          =   375
      Left            =   9420
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10200
      Width           =   1200
   End
   Begin VB.TextBox txtBwhTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7725
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   44
      Top             =   8955
      Width           =   1860
   End
   Begin VB.TextBox txtBwhPPn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5325
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   43
      Top             =   8955
      Width           =   1860
   End
   Begin VB.TextBox txtBwhAmo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2955
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   42
      Top             =   8955
      Width           =   1860
   End
   Begin VB.TextBox txtBwhNota 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   615
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   41
      Top             =   8955
      Width           =   1860
   End
   Begin VB.TextBox txtNota 
      Height          =   315
      Left            =   2985
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   7
      Top             =   3360
      Width           =   2470
   End
   Begin VB.CommandButton cmdAtas 
      BackColor       =   &H0080FFFF&
      Caption         =   "Create"
      Height          =   375
      Left            =   13335
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3330
      Width           =   1200
   End
   Begin VB.TextBox txtRemarks 
      Height          =   315
      Left            =   9420
      MaxLength       =   100
      TabIndex        =   10
      Top             =   3360
      Width           =   3390
   End
   Begin VB.TextBox txtBwhAmoIdr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10485
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   32
      Top             =   8955
      Width           =   1860
   End
   Begin VB.TextBox txtBwhPPnIdr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   12750
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   30
      Top             =   8955
      Width           =   1860
   End
   Begin VB.Frame FrameHeader 
      BackColor       =   &H00FDDFE3&
      Height          =   2400
      Left            =   180
      TabIndex        =   23
      Top             =   825
      Width           =   14835
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   1845
         TabIndex        =   2
         Top             =   1110
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   334036995
         CurrentDate     =   38991
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   4110
         TabIndex        =   3
         Top             =   1110
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   334036995
         CurrentDate     =   38991
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retur Type"
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   53
         Top             =   735
         Width           =   945
      End
      Begin MSForms.ComboBox cbType 
         Height          =   315
         Left            =   1845
         TabIndex        =   1
         Top             =   690
         Width           =   1485
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "2619;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblS"
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
         Left            =   11940
         TabIndex        =   52
         Top             =   1470
         Width           =   840
      End
      Begin VB.Label lblEur 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblEur"
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
         Left            =   11940
         TabIndex        =   51
         Top             =   1125
         Width           =   840
      End
      Begin VB.Label lblCurrS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S $"
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
         Left            =   11355
         TabIndex        =   50
         Top             =   1470
         Width           =   300
      End
      Begin VB.Label lblCurrEuro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Euro"
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
         Left            =   11355
         TabIndex        =   49
         Top             =   1125
         Width           =   450
      End
      Begin VB.Label lblCurrUS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "US $"
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
         Left            =   11355
         TabIndex        =   48
         Top             =   465
         Width           =   435
      End
      Begin VB.Label lblCurrYen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "YEN"
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
         Left            =   11355
         TabIndex        =   47
         Top             =   810
         Width           =   375
      End
      Begin VB.Label lblUS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblUS"
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
         Left            =   11940
         TabIndex        =   46
         Top             =   465
         Width           =   840
      End
      Begin VB.Label lblYEN 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblYEN"
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
         Left            =   11940
         TabIndex        =   45
         Top             =   810
         Width           =   840
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         Height          =   195
         Index           =   15
         Left            =   225
         TabIndex        =   39
         Top             =   1995
         Width           =   1155
      End
      Begin MSForms.ComboBox cbPajak 
         Height          =   315
         Left            =   1845
         TabIndex        =   4
         Top             =   1530
         Width           =   2580
         VariousPropertyBits=   746604571
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "4551;556"
         ListRows        =   15
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
         Caption         =   "Faktur Pajak No"
         Height          =   195
         Index           =   14
         Left            =   225
         TabIndex        =   38
         Top             =   1590
         Width           =   1365
      End
      Begin MSForms.ComboBox cbProd 
         Height          =   315
         Left            =   1845
         TabIndex        =   5
         Top             =   1935
         Width           =   2580
         VariousPropertyBits=   746604571
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "4551;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line 
         Index           =   2
         X1              =   4605
         X2              =   8470
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Label lblProd 
         BackStyle       =   0  'Transparent
         Caption         =   "lblProd"
         Height          =   195
         Left            =   4605
         TabIndex        =   37
         Top             =   1995
         Width           =   3885
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return Date"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   36
         Top             =   1170
         Width           =   1035
      End
      Begin VB.Label lblCust 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblCust"
         Height          =   195
         Left            =   3510
         TabIndex        =   27
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   25
         Top             =   330
         Width           =   1350
      End
      Begin VB.Line Line 
         Index           =   0
         X1              =   3510
         X2              =   8710
         Y1              =   570
         Y2              =   570
      End
      Begin MSForms.ComboBox cbCust 
         Height          =   315
         Left            =   1845
         TabIndex        =   0
         Top             =   270
         Width           =   1515
         VariousPropertyBits=   612386843
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
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   195
         Index           =   1
         Left            =   3630
         TabIndex        =   24
         Top             =   1170
         Width           =   165
      End
   End
   Begin VB.Frame FrameFooter 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   533
      Left            =   173
      TabIndex        =   21
      Top             =   9435
      Width           =   14835
      Begin VB.Label lblErr 
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
         Height          =   225
         Left            =   105
         TabIndex        =   22
         Top             =   195
         Width           =   14610
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clea&r"
      Height          =   375
      Left            =   10890
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10200
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   12345
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10200
      Width           =   1200
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10200
      Width           =   1200
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10200
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13170
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   270
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   4530
      Left            =   180
      TabIndex        =   12
      Top             =   3825
      Width           =   14835
      _cx             =   26167
      _cy             =   7990
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
      GridColor       =   8421504
      GridColorFixed  =   4210752
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
      Rows            =   20
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   0
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
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   315
      Left            =   6690
      TabIndex        =   9
      Top             =   3360
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   334036995
      CurrentDate     =   38991
   End
   Begin MSForms.ComboBox cb1 
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   3360
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
   Begin MSForms.ComboBox cbNota 
      Height          =   315
      Left            =   2985
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3360
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
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Index           =   4
      Left            =   6075
      TabIndex        =   40
      Top             =   3420
      Width           =   405
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Index           =   13
      Left            =   8520
      TabIndex        =   35
      Top             =   3420
      Width           =   765
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Retur No"
      Height          =   195
      Index           =   3
      Left            =   1605
      TabIndex        =   34
      Top             =   3420
      Width           =   1200
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount IDR"
      Height          =   195
      Index           =   9
      Left            =   10515
      TabIndex        =   33
      Top             =   8580
      Width           =   1050
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PPN IDR"
      Height          =   195
      Index           =   12
      Left            =   12750
      TabIndex        =   31
      Top             =   8580
      Width           =   720
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Index           =   8
      Left            =   7740
      TabIndex        =   29
      Top             =   8580
      Width           =   1140
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PPN"
      Height          =   195
      Index           =   7
      Left            =   5325
      TabIndex        =   28
      Top             =   8580
      Width           =   330
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Retur"
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
      Left            =   480
      TabIndex        =   26
      Top             =   285
      Width           =   14190
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Index           =   6
      Left            =   2970
      TabIndex        =   20
      Top             =   8580
      Width           =   660
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Retur No"
      Height          =   195
      Index           =   5
      Left            =   615
      TabIndex        =   19
      Top             =   8580
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   180
      Top             =   8475
      Width           =   14835
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   1
      Left            =   180
      Top             =   8835
      Width           =   14835
   End
End
Attribute VB_Name = "FrmNotaRetur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqlprod As String, colOld As Integer, nomeR As String, actCurr As String, colNotaRetur As Integer, colTotRetur As Integer
Dim colOldRetur As Integer, colQtyFaktur As Integer
Dim ColCheck As Integer, colPart As Integer, ColDesc As Integer, colFaktur As Integer, colSJ As Integer, ColQty As Integer, colQtyRetur As Integer
Dim colUnitCls As Integer, ColUnit As Integer, colCurrCode As Integer, ColCurr As Integer, ColPrice As Integer, colAdj As Integer, colAmo As Integer, colReturnDate As Integer
'Update By Dudi
Dim colDnNo As Byte, ColRef As Byte, ColReturnSeq_no As Byte




Private Sub cb1_Change()
    If cb1.ListIndex = 0 Then
        cmdAtas.Caption = "Create"
        cbNota.Enabled = False
        Call addNo
        Call Header
        dtDate.Value = Format(Now, "dd MMM yyyy"): txtRemarks = ""
        nomeR = Trim$(txtNota.Text)
    Else
        cmdAtas.Caption = "Update"
        txtNota = ""
        cbNota.Enabled = True
        If cbCust.ListIndex <> -1 And cbPajak.ListIndex <> -1 And cbProd.ListIndex <> -1 Then Call addCbNota
        cbNota.Text = nomeR
    End If
    Call KosongBawah
    LblErr = ""
End Sub

Private Sub cb1_Click()
    'cb1_Change
End Sub

Private Sub cbCust_Change()
    If cbCust.MatchFound Then
        lblcust = cbCust.Column(1)
        If cbType.ListIndex <> -1 Then addCbPajak: Call Header: KosongBawah: cbNota.clear: txtNota = "": cbProd.clear
    Else
        lblcust = ""
        cbPajak.clear
    End If
    LblErr = ""
End Sub

Private Sub cbCust_Click()
    cbCust_Change
End Sub

Private Sub cbNota_Change()
Dim rstgl As New ADODB.Recordset
    If cbNota.MatchFound Then
        txtNota = cbNota.Column(0)
        Call Header
        sql = "select * from notaretur_master where notaretur_no='" & txtNota.Text & "' "
        rstgl.Open sql, Db, adOpenKeyset, adLockOptimistic
    
        If Not (rstgl.BOF And rstgl.EOF) Then
            dtDate.Value = Format(rstgl!notaretur_Date, "dd MMM yyyy")
            txtRemarks = Trim$(rstgl!Remarks)
'            p = IIf(IsNull(rs("notaretur_date")), " ", Left(Trim(rs("notaretur_date")), 4) & "-" & Right(Trim(rs("notaretur_date")), 2) & "-01")
'            Period.Value = Format(p, "MMM yyyy")
'            temptgl = Period.month
        End If
        rstgl.Close
        Set rstgl = Nothing
    Else
        txtNota = ""
    End If
    Call KosongBawah
    LblErr = ""
End Sub

Private Sub cbNota_Click()
    cbNota_Change
End Sub

Private Sub cbPajak_Change()
    If cbPajak.ListCount < 1 Then Exit Sub
    If cbPajak.MatchFound Then
        addCbProd
        If grid.Rows > 1 Then
            If cbType.ListIndex = 0 Then
                If Trim$(grid.TextMatrix(1, colFaktur)) <> Trim$(cbPajak.List(cbPajak.ListIndex, 0)) Then
                    LblErr = DisplayMsg("5555")
                    cbPajak.Text = Trim$(grid.TextMatrix(1, colFaktur))
                    Exit Sub
                End If
            ElseIf cbType.ListIndex = 1 Then
                If Trim$(grid.TextMatrix(1, colCurrCode)) <> Trim$(cbPajak.List(cbPajak.ListIndex, 1)) Then
                    LblErr = DisplayMsg("5405")
                    cbPajak.Text = Trim$(grid.TextMatrix(1, colFaktur))
                    Exit Sub
                End If
            End If
        End If
    Else
        Call Header
        cbProd.clear
        cbNota.clear
        txtNota = ""
    End If
'    lblErr = ""
End Sub

Private Sub cbPajak_Click()
'    cbPajak_Change
End Sub

Private Sub cbProd_Change()
    If cbProd.MatchFound Then
        lblProd = cbProd.Column(1)
        If cbProd.ListIndex = 0 Then sqlprod = "" Else sqlprod = " and sjr.item_code = '" & Trim$(cbProd.List(cbProd.ListIndex, 0)) & "' ": nomeR = Trim$(txtNota.Text): Call addCbNota: cbNota.Text = nomeR
    Else
        lblProd = "": sqlprod = ""
    End If
    LblErr = ""
End Sub
'
'Private Sub cbProd_Click()
'    cbProd_Change
'End Sub

Private Sub cbType_Change()
    If cbCust.ListIndex <> -1 Then addCbPajak: Call Header: KosongBawah: cbNota.clear: txtNota = "": cbProd.clear
End Sub

Private Sub cbType_Click()
    'cbType_Change
End Sub

Private Sub cmdCancel_Click()
    cbNota_Click
    cmdAtas_Click
End Sub

Private Sub cmdClear_Click()
    Kosong
    Call Header
    Call KosongBawah
End Sub

Private Sub cmdsub_Click()
    Db.Execute ("delete notaretur_master where notaretur_No not in (select notaretur_no from notaretur_Detail)")
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub DTDate_Change()
    cariRate
End Sub

Private Sub dtDate_Click()
    DTDate_Change
End Sub

Private Sub dtFrom_Change()
    If cbCust.ListIndex <> -1 And cbType.ListIndex <> -1 Then addCbPajak
End Sub

Private Sub dtTo_Change()
    If cbCust.ListIndex <> -1 And cbType.ListIndex <> -1 Then addCbPajak
End Sub

Private Sub Form_Load()
    Call iniSialisasi
    Call Header
    cb1.clear
    cb1.AddItem "Create"
    cb1.AddItem "Update"
    cbType.clear
    cbType.AddItem "Sales"
    cbType.AddItem "Faktur"
    Call Kosong
    Call KosongBawah
    Call cariRate
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    Call addCbCust
    DtFrom = Format(Now, "1 MMM yyyy")
    dtyo = Format(Now, "dd MMM yyyy")
End Sub

Private Sub addCbCust()
    Dim adoRs As New ADODB.Recordset
    Dim intCount As Integer
    
    Me.MousePointer = vbHourglass
    On Error GoTo errHandler
    
    With cbCust
        .clear
        .columnCount = 3
        .ColumnWidths = "50pt;300pt;0pt"
        .ListWidth = 350
        .ListRows = 15
        
        sql = "Select Trade_Code, Trade_Name, Address1 From Trade_Master Where Trade_Cls In (2, 3)"
        adoRs.Open sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
        While Not adoRs.EOF
            .AddItem ""
            .Column(0, intCount) = Trim(adoRs.Fields("Trade_Code"))
            .Column(1, intCount) = Trim(adoRs.Fields("Trade_Name"))
            .Column(2, intCount) = Trim(adoRs.Fields("Address1"))
            intCount = intCount + 1
            adoRs.MoveNext
        Wend
        adoRs.Close
    End With
    
ErrExit:
    Set adoRs = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErr.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = False Then Cancel = True
End Sub

Private Sub Kosong()
    sqlprod = ""
    DtFrom = Format(Now, "1 MMM yyyy")
    dtTo = Format(Now, "dd MMM yyyy")
    cbPajak.clear
    cbProd.clear
    lblProd = ""
    cbCust.ListIndex = -1
    cbType.ListIndex = -1
    lblcust = ""
    LblErr = ""
    cb1.ListIndex = 1
    cbNota.clear
    txtNota = ""
    dtDate = Format(Now, "dd MMM yyyy")
    txtRemarks = ""
    cmdAtas.Caption = "Update"
    Call KosongBawah
End Sub

Private Sub KosongBawah()
    txtBwhAmo = ""
    txtBwhAmoIdr = ""
    txtBwhNota = ""
    txtBwhPPn = ""
    txtBwhPPnIdr = ""
    txtBwhTotal = ""
End Sub

Private Sub addCbPajak()
Dim rspajak As New ADODB.Recordset, Q As Byte
cbPajak.clear
cbPajak.columnCount = 2
cbPajak.ColumnWidths = "120pt;0pt"
cbPajak.ListWidth = 120
cbPajak.ListRows = 15

rspajak.Open " select distinct fpm.fakturpajak_no, fpd.currency_code from fakturpajak_Master fpm " + _
            " inner join fakturpajak_Detail fpd on fpm.fakturpajak_no =fpd.fakturpajak_no " + _
            " where fpm.cust_Code = '" & Trim$(cbCust.List(cbCust.ListIndex, 0)) & "' " + _
            " and fpm.fakturpajak_Date<= '" & Format(dtTo.Value, "yyyy-mm-dd") & "' " + _
            " and fpm.fakturpajak_Date >='" & Format(DtFrom.Value, "yyyy-mm-dd") & "' ", Db, adOpenForwardOnly, adLockReadOnly

If Not (rspajak.EOF Or rspajak.BOF) Then
    Q = 0
    While Not rspajak.EOF
        cbPajak.AddItem ""
        cbPajak.List(Q, 0) = Trim(rspajak!fakturpajak_no)
        cbPajak.List(Q, 1) = Trim(rspajak!currency_code)
        rspajak.MoveNext
        Q = Q + 1
    Wend
End If
rspajak.Close
Set rspajak = Nothing
End Sub

Private Sub cariRate()
Dim rsrate As New ADODB.Recordset
lblUS = "--"
lblYEN = "--"
lblEur = "--"
lblS = "--"
rsrate.Open " select currency_code, tax_exchangerate from tax_Exchangerate where '" & Format(dtDate.Value, "yyyymmdd") & "' <= end_date and '" & Format(dtDate.Value, "yyyymmdd") & "'>=start_Date ", Db, adOpenForwardOnly, adLockReadOnly
    If Not (rsrate.EOF Or rsrate.BOF) Then
        While Not rsrate.EOF
            Select Case (rsrate!currency_code)
                Case "01": lblYEN = Format(Trim(rsrate!Tax_exchangerate & ""), "#,##0.#0")
                Case "02": lblUS = Format(Trim(rsrate!Tax_exchangerate & ""), "#,##0.#0")
                Case "04": lblEur = Format(Trim(rsrate!Tax_exchangerate & ""), "#,##0.#0")
                Case "05": lblS = Format(Trim(rsrate!Tax_exchangerate & ""), "#,##0.#0")
            End Select
            rsrate.MoveNext
        Wend
    End If
rsrate.Close
Set rsrate = Nothing
End Sub

Private Sub addCbProd()
Dim rsProd As New ADODB.Recordset, Baris As Integer, sQlproduk As String
cbProd.clear
cbProd.columnCount = 2
cbProd.ColumnWidths = "90pt;110pt"
cbProd.ListWidth = 200
cbProd.ListRows = 15

sQlproduk = " select distinct fpd.item_Code, im.item_Name " + _
                " from fakturpajak_detail fpd "

If cbType.ListIndex = 0 Then sQlproduk = sQlproduk + " inner join Delivery_Return sjr on sjr.item_Code = fpd.item_Code "
                
sQlproduk = sQlproduk + " left join item_Master im on fpd.item_Code = im.item_code " + _
                " where fpd.fakturpajak_No = '" & Trim(cbPajak.List(cbPajak.ListIndex, 0)) & "' "
                
If cbType.ListIndex = 0 Then sQlproduk = sQlproduk + " and sjr.return_Cls = 'D2' "

rsProd.Open sQlproduk, Db, 1, 3
    If Not (rsProd.EOF Or rsProd.BOF) Then
        cbProd.AddItem ""
        cbProd.List(0, 0) = "-- A L L --"
        cbProd.List(0, 1) = "ALL"
        Baris = 1
        While Not rsProd.EOF
            cbProd.AddItem ""
            cbProd.List(Baris, 0) = Trim(rsProd!Item_Code)
            cbProd.List(Baris, 1) = Trim(rsProd!item_name)
            Baris = Baris + 1
            rsProd.MoveNext
        Wend
        cbProd.ListIndex = 0
        If cbNota.ListCount < 1 Then Call addCbNota
    End If
rsProd.Close
Set rsProd = Nothing
End Sub

Private Sub addCbNota()
Dim RSA As New ADODB.Recordset, sqlNota As String
cbNota.clear
cbNota.ListRows = 15

If cbProd.ListIndex = 0 Then sqlprod = "" Else sqlprod = " and sjr.item_code = '" & Trim$(cbProd.List(cbProd.ListIndex, 0)) & "' "
If cbType.ListIndex = 0 Then
    sqlNota = " and nrm.notaretur_type=0 "
ElseIf cbType.ListIndex = 1 Then
    sqlNota = " and nrm.notaretur_type=1 "
End If

RSA.Open " select distinct nrm.NotaRetur_No from notaretur_master nrm " + _
        " left join notaretur_detail nrd on nrd.notaretur_no = nrm.notaretur_no " + _
        " where nrm.cust_Code = '" & Trim$(cbCust.List(cbCust.ListIndex, 0)) & "' " + _
        "     and nrm.notaretur_Date <= '" & Format(dtTo.Value, "yyyy-mm-dd") & "' " + _
        "     and nrm.notaretur_Date >= '" & Format(DtFrom.Value, "yyyy-mm-dd") & "' " + sqlNota, Db, 1, 3  '+ _ 'sqlProd, Db, 1, 3
        
    If Not (RSA.EOF Or RSA.BOF) Then
        While Not RSA.EOF
            cbNota.AddItem Trim(RSA!notaretur_no)
            RSA.MoveNext
        Wend
    End If
Set RSA = Nothing
End Sub

Private Sub addNo()
Dim rsno As New ADODB.Recordset, rsS As New Recordset
    'NR-YYMM-9999
    
    rsno.Open " select top 1 urut = right(rtrim(NotaRetur_No),4) from notaRetur_master " + _
        " where left(rtrim(NotaRetur_No),5) = 'NR-" & Trim$(Format(dtDate.Value, "yy")) & "' " + _
        " order by urut desc ", Db, 1, 3
    If Not (rsno.BOF And rsno.EOF) Then
        txtNota.Text = "NR-" + Format(dtDate.Value, "yymm") + "-" + Format(Str(Int(rsno!Urut) + 1), "0000")
    Else
        txtNota.Text = "NR-" & Format(dtDate.Value, "yymm") & "-" & "0001"
    End If
        
    rsno.Close
    Set rsno = Nothing
End Sub

Private Function cekAtas() As Boolean
    If cbCust.ListIndex < 0 Or cbCust.MatchFound = False Then
        LblErr = DisplayMsg("1045")
        cbCust.SetFocus
        GoTo salaH
    End If
    If cbType.ListIndex < 0 Or cbType.MatchFound = False Then
        LblErr = DisplayMsg("5464")
        cbType.SetFocus
        GoTo salaH
    End If
    If cbPajak.ListIndex < 0 Or cbPajak.MatchFound = False Then
        LblErr = DisplayMsg("5413")
        cbPajak.SetFocus
        GoTo salaH
    End If
    If cbProd.ListIndex < 0 Or cbProd.MatchFound = False Then
        LblErr = DisplayMsg("1024")
        cbProd.SetFocus
        GoTo salaH
    End If
    If Trim$(txtNota.Text) = "" Then
        LblErr = DisplayMsg("5412")
        txtNota.SetFocus
        GoTo salaH
    End If
    
    cekAtas = True
    Exit Function
salaH:
        cekAtas = False
        Exit Function
End Function

Private Sub cmdAtas_Click()
If cekAtas Then
    If cb1.ListIndex = 0 Then 'CREATE
        If hakUpdate(Me.Name) = 0 Then _
            LblErr = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
        Call savemaster(True)
'        Call addCbNota
        cb1.ListIndex = 1
        LblErr = DisplayMsg("1101")
        cbNota.Text = nomeR
        txtNota = nomeR
        cmdAtas_Click
        Exit Sub
    ElseIf cb1.ListIndex = 1 Then 'UPDATE
        MousePointer = vbHourglass
        Call Header
        Call browseitem
        If cbType.ListIndex = 0 Then
            If Not cekFaktur Then Call Header: LblErr = DisplayMsg("5555"): Exit Sub
        ElseIf cbType.ListIndex = 1 Then
            If Not cekCurr Then Call Header: LblErr = DisplayMsg("5405"): Exit Sub
        End If
        Call BrowseGrid
        MousePointer = vbDefault
    End If
    Call cariBawah
    LblErr = ""
End If
End Sub

Private Function cekFaktur() As Boolean
Dim Q As Byte
With grid
    For Q = 1 To .Rows - 1
        If Trim$(.TextMatrix(Q, colFaktur)) <> Trim$(cbPajak.Text) Then
        cekFaktur = False
        Exit Function
        End If
    Next Q
End With
cekFaktur = True
End Function

Private Function cekCurr() As Boolean
Dim Q As Byte
With grid
    For Q = 1 To .Rows - 1
        If Trim$(.TextMatrix(Q, colCurrCode)) <> Trim$(cbPajak.List(cbPajak.ListIndex, 1)) Then
        cekCurr = False
        Exit Function
        End If
    Next Q
End With
cekCurr = True
End Function

Private Sub iniSialisasi()
    
    ColCheck = 0
    colDnNo = 1
    ColRef = 2
    colReturnDate = 3
    colPart = 4
    ColDesc = 5
    colFaktur = 6
    colSJ = 7
    ColQty = 8
    colQtyFaktur = 9
    colQtyRetur = 10
    colUnitCls = 11
    ColUnit = 12
    colCurrCode = 13
    ColCurr = 14
    ColPrice = 15
    colAdj = 16
    colAmo = 17
    colOld = 18
    colNotaRetur = 19
    colTotRetur = 20
    colOldRetur = 21
    ColReturnSeq_no = 22
End Sub

Private Sub Header()
With grid
    .Rows = 1
    .ColS = 23
    
    .ColWidth(ColCheck) = 300
    If cbType.ListIndex = 0 Then
        .ColWidth(colReturnDate) = 1300
        .ColWidth(colDnNo) = 1500
    .ColWidth(ColRef) = 1500
    Else
        .ColWidth(colReturnDate) = 0
        .ColWidth(colDnNo) = 0
    .ColWidth(ColRef) = 0
        
    End If
    .ColWidth(colPart) = 1900
    .ColWidth(ColDesc) = 2700
    .ColWidth(colFaktur) = 2000
    .ColWidth(colSJ) = 0 '1900
    .ColWidth(ColQty) = 1400
    .ColWidth(colQtyFaktur) = 1400
    .ColWidth(colQtyRetur) = 1400
    .ColWidth(colUnitCls) = 0 '1000
    .ColWidth(ColUnit) = 600
    .ColWidth(colCurrCode) = 0 '1000
    .ColWidth(ColCurr) = 600
    .ColWidth(ColPrice) = 1800
    .ColWidth(colAdj) = 1800
    .ColWidth(colAmo) = 2000
    .ColWidth(colOld) = 1000
    .ColWidth(colNotaRetur) = 1500
    .ColWidth(colTotRetur) = 1500
    .ColWidth(colOldRetur) = 1500
    
    .TextMatrix(0, ColCheck) = ""
    .TextMatrix(0, colDnNo) = "DN Number"
    .TextMatrix(0, ColRef) = "Reference"
    .TextMatrix(0, colReturnDate) = "Return Date"
    .TextMatrix(0, colPart) = "Part Number"
    .TextMatrix(0, ColDesc) = "Description"
    .TextMatrix(0, colFaktur) = "Faktur Pajak No"
    .TextMatrix(0, colSJ) = "SJ No"
    .TextMatrix(0, ColQty) = "Qty SJ Return"
    .TextMatrix(0, colQtyFaktur) = "Qty Faktur"
    .TextMatrix(0, colQtyRetur) = "Qty Retur"
    .TextMatrix(0, colUnitCls) = "Unit_Cls"
    .TextMatrix(0, ColUnit) = "Unit"
    .TextMatrix(0, colCurrCode) = "Curr_Code"
    .TextMatrix(0, ColCurr) = "Curr"
    .TextMatrix(0, ColPrice) = "Price"
    .TextMatrix(0, colAdj) = "Adj Price"
    .TextMatrix(0, colAmo) = "Amount"
    .TextMatrix(0, colOld) = "Old"
    .TextMatrix(0, colNotaRetur) = "Nota Retur"
    .TextMatrix(0, colTotRetur) = "Total Retur"
    .TextMatrix(0, colOldRetur) = "Old Retur"
    
    .FixedRows = 1
    .FrozenCols = 3
    .Cell(flexcpAlignment, 0, 0, 0, colAmo) = flexAlignCenterCenter
    For i = colPart To colSJ
        .ColAlignment(i) = flexAlignLeftCenter
    Next i
    For i = ColQty To colAmo
        .ColAlignment(i) = flexAlignRightCenter
    Next i
    .ColAlignment(ColUnit) = flexAlignCenterCenter
    .ColAlignment(ColCurr) = flexAlignCenterCenter
    .ColHidden(ColReturnSeq_no) = True
End With
End Sub

Private Sub browseitem()
Dim sqlCari As String, rsCari As New ADODB.Recordset
If cbType.ListIndex = 0 Then
    sqlCari = " select return_date,item_Code, item_name, fakturpajak_no, sum(Return_qty)Qty, fQty, unit_cls, currency_Code, Price, total,DO_NO,reference,ReturnSeq_no   " + _
            " from( " + _
            "           select nrd.return_date,nrd.item_Code, im.item_Name, nrd.fakturPajak_No, sjr.Return_Qty, fQty=fpd.Qty, nrd.unit_Cls, nrd.currency_Code, nrd.Price ," + _
            "           total =    isnull((select sum(qty) from notaretur_detail a, notaretur_master b where a.notaretur_no = b.notaretur_no and a.item_code = sjr.item_code" & _
            "           and b.cust_code =  sjr.cust_code and return_date = sjr.return_date ),0) " & vbCrLf
    sqlCari = sqlCari & "           ,SJR.DO_NO,SJR.reference,SJR.ReturnSeq_no  from notaretur_Detail nrd " & _
            "           left join item_Master im  on nrd.item_Code = im.item_Code   " & _
            "           left join Delivery_return sjr on nrd.item_Code = sjr.item_Code AND nrd.ReturnSeq_no = SJR.ReturnSeq_no " + _
            "           left join fakturpajak_detail fpd on nrd.fakturpajak_No = fpd.fakturpajak_no and fpd.item_Code = nrd.item_Code " & vbCrLf
    sqlCari = sqlCari & "           where nrd.notaretur_no = '" & Trim$(txtNota.Text) & "' "
    sqlCari = sqlCari + "UNION "            ' Qty total -> gabungan beberapa FP
    sqlCari = sqlCari + "select sjr.return_date, fpd.item_Code, im.item_Name, fpd.fakturPajak_No, sjr.Return_Qty, fQty=fpd.qty, fpd.unit_Cls, currency_Code, fpd.Price ," + _
            "           total =    isnull((select sum(qty) from notaretur_detail a, notaretur_master b where a.notaretur_no = b.notaretur_no and a.item_code = sjr.item_code" & _
            "           and b.cust_code =  sjr.cust_code and return_date = sjr.return_date ),0), " & vbCrLf
    sqlCari = sqlCari & "           SJR.DO_NO,SJR.reference,SJR.ReturnSeq_no from fakturPajak_Detail fpd " + _
            "           left join item_Master im  on fpd.item_Code = im.item_Code " + _
            "           left join Delivery_return sjr on fpd.item_Code = sjr.item_Code" + _
            "           where sjr.return_date<= '" & Format(dtTo.Value, "yyyy-mm-dd") & "' " & vbCrLf
    sqlCari = sqlCari & "               and sjr.return_Date>= '" & Format(DtFrom.Value, "yyyy-mm-dd") & "' " + _
            "               and sjr.return_Cls = 'D2' and  ReturnSeq_NO NOT IN (SELECT ReturnSeq_No From notaretur_Detail) " + _
            "               and sjr.Cust_Code = '" & Trim$(cbCust.List(cbCust.ListIndex, 0)) & "' " + _
            "               and fpd.fakturpajak_No = '" & Trim(cbPajak.List(cbPajak.ListIndex, 0)) & "' "
            
    If cbProd.ListIndex > 0 Then sqlCari = sqlCari + " and sjr.item_Code = '" & Trim$(cbProd.List(cbProd.ListIndex, 0)) & "' "
               
    sqlCari = sqlCari + "       )s " + _
            " group by return_date, item_Code, item_name, fakturpajak_no, fQty, unit_cls, currency_Code, Price , total,DO_NO,reference,ReturnSeq_no"

Else
    
    sqlCari = " select item_Code, item_name, fakturpajak_no, sum(qty)Qty, fQty, unit_cls, currency_Code, Price, total " + _
            " from( " + _
            "           select nrd.item_Code, im.item_Name, nrd.fakturPajak_No, qty=0, fQty=fpd.Qty, nrd.unit_Cls, nrd.currency_Code, nrd.Price, " + _
            "           total =isnull( (select sum(qty) from notaretur_detail where fakturpajak_no = nrd.fakturpajak_no and item_Code = nrd.item_code) ,0)           " + _
            "            from notaretur_Detail nrd " & _
            "           left join item_Master im  on nrd.item_Code = im.item_Code   " & _
            "           left join fakturpajak_Detail fpd on fpd.fakturpajak_no = nrd.fakturpajak_No and fpd.item_Code = nrd.item_Code" + _
            "           where nrd.notaretur_no = '" & Trim$(txtNota.Text) & "' "

'            " and nrd.fakturpajak_No = '" & Trim(cbPajak.List(cbPajak.ListIndex, 0)) & "' "

    sqlCari = sqlCari + " UNION " + _
            "           select fpd.item_Code, im.item_Name, fpd.fakturPajak_No, qty=0, fQty=fpd.Qty, fpd.unit_Cls, fpd.currency_Code, fpd.Price,  " & _
            "           total = isnull( (select sum(qty) from notaretur_detail where fakturpajak_no = fpd.fakturpajak_no and item_Code = fpd.item_code) ,0)          " + _
            "           from fakturPajak_Detail fpd " & _
            "           left join fakturpajak_Master fpm on fpd.fakturpajak_No =  fpm.fakturpajak_no " & _
            "           left join item_Master im  on fpd.item_Code = im.item_Code   " & _
            "           where fpm.fakturpajak_date <= '" & Format(dtTo.Value, "yyyy-mm-dd") & "'       " & _
            "               and fpm.fakturpajak_date >= '" & Format(DtFrom.Value, "yyyy-mm-dd") & "'       " & _
            "               and fpm.Cust_Code = '" & Trim$(cbCust.List(cbCust.ListIndex, 0)) & "'       " & _
            "               and fpm.fakturpajak_No = '" & Trim(cbPajak.List(cbPajak.ListIndex, 0)) & "'  "
    
    If cbProd.ListIndex > 0 Then sqlCari = sqlCari + " and fpd.item_Code = '" & Trim$(cbProd.List(cbProd.ListIndex, 0)) & "' "
            
    sqlCari = sqlCari + "       ) a " + _
            "   group by item_Code, item_name, fakturpajak_no, fQty, unit_cls, currency_Code, Price, total  "

End If
rsCari.Open sqlCari, Db, 1, 3
If Not rsCari.BOF Or rsCari.EOF Then
    With grid
    While Not rsCari.EOF
        .Rows = .Rows + 1
        .Cell(flexcpChecked, .Rows - 1, ColCheck) = flexUnchecked
        .Cell(flexcpBackColor, .Rows - 1, ColCheck) = vbWhite
        If cbType.ListIndex = 0 Then
            .TextMatrix(.Rows - 1, colReturnDate) = Trim(rsCari!Return_Date & "")
            .TextMatrix(.Rows - 1, ColRef) = Trim(rsCari!Reference & "")
            .TextMatrix(.Rows - 1, colDnNo) = Trim(rsCari!do_no & "")
            .TextMatrix(.Rows - 1, ColReturnSeq_no) = Trim(rsCari!ReturnSeq_no)
        Else
            .TextMatrix(.Rows - 1, ColReturnSeq_no) = "0"
        End If
        
        .TextMatrix(.Rows - 1, colPart) = Trim(rsCari!Item_Code & "")
        .TextMatrix(.Rows - 1, ColDesc) = Trim(rsCari!item_name & "")
        .TextMatrix(.Rows - 1, colFaktur) = Trim(rsCari!fakturpajak_no & "")
'        .TextMatrix(.Rows - 1, colSJ) = Trim(rsCari!sj_No & "")
        .TextMatrix(.Rows - 1, ColQty) = Format(Trim(rsCari!Qty & ""), "#,##0.#0")
        .TextMatrix(.Rows - 1, colQtyFaktur) = Format(Trim(rsCari!fQty & ""), "#,##0.#0")
        .Cell(flexcpBackColor, .Rows - 1, colQtyRetur) = vbWhite
        If IsNull(rsCari("unit_cls")) Then
          .TextMatrix(.Rows - 1, colUnitCls) = " "
          .TextMatrix(.Rows - 1, ColUnit) = " "
        Else
          .TextMatrix(.Rows - 1, colUnitCls) = Trim(rsCari("Unit_cls"))
          '.TextMatrix(.Rows - 1, ColUnit) = Split(isiunit, ",")(Val(Trim(rsCari("Unit_Cls"))) - 1)
          .TextMatrix(.Rows - 1, ColUnit) = Trim(Get_Field("select * FROM Unit_Cls WHERE  Unit_Cls=" & .TextMatrix(.Rows - 1, colUnitCls), 1)) '(Val(Trim(rsCari("Unit_Cls")),) - 1)
        End If
        If IsNull(rsCari("currency_code")) Then
           .TextMatrix(.Rows - 1, colCurrCode) = ""
           .TextMatrix(.Rows - 1, ColCurr) = ""
        Else
          .TextMatrix(.Rows - 1, colCurrCode) = Trim(rsCari("currency_code"))
          '.TextMatrix(.Rows - 1, ColCurr) = Split(isiCurr, ",")(Val(Trim(rsCari("Currency_code"))) - 1)
          .TextMatrix(.Rows - 1, ColCurr) = Trim(Get_Field("SELECT * FROM Curr_Cls WHERE Curr_cls='" & .TextMatrix(.Rows - 1, colCurrCode) & "'", 1)) '(Val(Trim(rsCari("Currency_code"))) - 1))
        End If
        .TextMatrix(.Rows - 1, ColPrice) = Format(Trim(rsCari!Price & ""), "#,##0.00##")
        .Cell(flexcpBackColor, .Rows - 1, colAdj) = vbWhite
        .TextMatrix(.Rows - 1, colOld) = "0"
        .TextMatrix(.Rows - 1, colTotRetur) = Format(rsCari!Total, "#,##0.#0")
        .TextMatrix(.Rows - 1, colOldRetur) = "0"
        actCurr = Trim$(.TextMatrix(.Rows - 1, ColCurr))
        rsCari.MoveNext
    Wend
    End With
End If
rsCari.Close
Set rsCari = Nothing
End Sub
 Function Get_Field(sql, Field)
Dim Rdata As New ADODB.Recordset
Set Rdata = Db.Execute(sql)
Get_Field = ""
If Not Rdata.EOF Then
 Get_Field = IIf(IsNull(Rdata.Fields(Field)), "", Rdata.Fields(Field))
End If
End Function

Private Sub BrowseGrid()
Dim rsGrid As New ADODB.Recordset, sqlGrid As String
'sqlgrid = " select * from notaretur_detail where notaretur_no = '" & Trim$(txtNota.Text) & "' " ' + _
'            " and fakturpajak_No = '" & Trim(cbPajak.List(cbPajak.ListIndex, 0)) & "' " ' + sqlProd
sqlGrid = " select a.*,total  from notaretur_detail a  " & _
            " left join  " & _
            " ( select item_code, fakturpajak_no, sum (qty) total,ReturnSeq_NO  " & _
            "   from notaretur_detail  " & _
            "   group by item_code, fakturpajak_no,ReturnSeq_NO " & _
            " ) b " & _
            " on a.item_code=b.item_code  " & _
            "   and a.fakturpajak_no=b.fakturpajak_no " + _
            " where a.notaretur_No = '" & Trim$(txtNota.Text) & "'  " ' + _
'            "   and a.fakturpajak_No = '" & Trim(cbPajak.List(cbPajak.ListIndex, 0)) & "' "

If cbProd.ListIndex > 0 Then sqlCari = sqlCari + " and nrd.item_Code = '" & Trim$(cbProd.List(cbProd.ListIndex, 0)) & "' "
rsGrid.Open sqlGrid, Db, 1, 3
    If Not (rsGrid.EOF Or rsGrid.BOF) Then
        With grid
        While Not rsGrid.EOF
            For i = 1 To .Rows - 1
                If Trim$(rsGrid!Item_Code) = Trim$(.TextMatrix(i, colPart)) And Trim$(rsGrid!ReturnSeq_no) = Trim$(.TextMatrix(i, ColReturnSeq_no)) Then  'And Trim$(rsGrid!fakturpajak_No) = Trim$(.TextMatrix(i, colFaktur)) Then
                    .Cell(flexcpChecked, i, ColCheck) = flexChecked
                    .TextMatrix(i, colQtyRetur) = Format(Trim(rsGrid!Qty & ""), "#,##0.#0")
                    .TextMatrix(i, colAdj) = Format(Trim(rsGrid!adj_price & ""), "#,##0.00##")
                    .TextMatrix(i, colAmo) = Format(CDbl(.TextMatrix(i, colQtyRetur)) * CDbl(.TextMatrix(i, colAdj)), "#,##0.#0")
                    .TextMatrix(i, colOld) = "1"
                    .TextMatrix(i, colNotaRetur) = Trim(rsGrid!notaretur_no)
                    .TextMatrix(i, colTotRetur) = Trim(rsGrid!Total)
                    .TextMatrix(i, colOldRetur) = Format(Trim(rsGrid!Qty & ""), "#,##0.#0")
'                    actCurr = Trim$(.TextMatrix(i, colCurr))
                    Exit For
                End If
            Next i
            rsGrid.MoveNext
        Wend
        End With
    End If
rsGrid.Close
Set rsGrid = Nothing
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid
    If Col = colQtyRetur Then
        If Trim$(.TextMatrix(Row, colQtyRetur)) = "" Then .TextMatrix(Row, colQtyRetur) = "0.00"
        If IsNumeric(.TextMatrix(Row, colQtyRetur)) = False Then .TextMatrix(Row, colQtyRetur) = "0.00"
        If CDbl(.TextMatrix(Row, colQtyRetur)) + CDbl(.TextMatrix(Row, colTotRetur)) - CDbl(.TextMatrix(Row, colOldRetur)) > CDbl(.TextMatrix(Row, colQtyFaktur)) Then LblErr = "Qty Retur must be the same or lower than " & Format(CDbl(.TextMatrix(Row, colQtyFaktur)) - (CDbl(.TextMatrix(Row, colTotRetur)) - CDbl(.TextMatrix(Row, colOldRetur))), "#,##0")
        If cbType.ListIndex = 0 Then
            If CDbl(.TextMatrix(Row, colQtyRetur)) + CDbl(.TextMatrix(Row, colTotRetur)) - CDbl(.TextMatrix(Row, colOldRetur)) > CDbl(.TextMatrix(Row, ColQty)) Then LblErr = "Qty Retur must be the same or lower than " & Format(CDbl(.TextMatrix(Row, ColQty)) - (CDbl(.TextMatrix(Row, colTotRetur)) - CDbl(.TextMatrix(Row, colOldRetur))), "#,##0")
        End If
        .TextMatrix(Row, colQtyRetur) = Format(.TextMatrix(Row, colQtyRetur), "#,##0.#0")
    ElseIf Col = colAdj Then
        If Trim$(.TextMatrix(Row, colAdj)) = "" Then .TextMatrix(Row, colAdj) = "0.00"
        If IsNumeric(.TextMatrix(Row, colAdj)) = False Then .TextMatrix(Row, colAdj) = "0.00"
        .TextMatrix(Row, colAdj) = Format(.TextMatrix(Row, colAdj), "#,##0.#0")
    End If
    If Trim$(.TextMatrix(Row, colQtyRetur)) = "" Or Trim$(.TextMatrix(Row, colAdj)) = "" Then
    Else
        .TextMatrix(Row, colAmo) = Format(CDbl(.TextMatrix(Row, colQtyRetur)) * CDbl(.TextMatrix(Row, colAdj)), "#,##0.#0")
    End If
    Call cariBawah
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Cell(flexcpChecked, Row, ColCheck) <> flexChecked Then
        If Col <> ColCheck Then Cancel = True
    Else
        If Col <> ColCheck And Col <> colQtyRetur And Col <> colAdj Then Cancel = True
    End If
End Sub

Private Sub grid_Click()
With grid
If .Cell(flexcpChecked, .Row, ColCheck) = flexChecked Then
    If .Col = colQtyRetur Or .Col = colAdj Then
        .FocusRect = flexFocusInset
        .SelectionMode = flexSelectionFree
    Else
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
    End If
Else
    .FocusRect = flexFocusNone
    .SelectionMode = flexSelectionByRow
End If
End With
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
LblErr = ""
If Col = colQtyRetur Or Col = colAdj Then
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
        KeyAscii = 0
End If
End Sub

Private Sub savemaster(bool As Boolean)
Dim RS As New ADODB.Recordset
    If bool Then
        RS.Open " select * from notaretur_Master ", Db, 1, 3
        RS.AddNew
        RS!notaretur_no = Trim$(nomeR)
        RS!notaretur_type = Trim$(cbType.ListIndex)
        RS!Cust_CodE = Trim$(cbCust.List(cbCust.ListIndex, 0))
        RS!user_Entry = userLogin
        RS!date_entry = Now
    Else
        RS.Open " select * from notaretur_Master where notaretur_No = '" & Trim$(txtNota.Text) & "' ", Db, 1, 3
        If RS.EOF Or RS.BOF Then RS.Close: LblErr = "No Data with this Nota Retur No !": Exit Sub
        RS!user_update = userLogin
        RS!date_update = Now
    End If
        RS!notaretur_Date = Format(dtDate.Value, "yyyy-mm-dd")
        RS!Remarks = Trim$(txtRemarks.Text)
        RS!total_Qty = CDbl(0) 'hitungTotal
        If Trim$(txtBwhAmo.Text) = "" Then RS!Amount = CDbl(0) Else RS!Amount = CDbl(txtBwhAmo.Text)
        If Trim$(txtBwhPPn.Text) = "" Then RS!ppn = CDbl(0) Else RS!ppn = CDbl(txtBwhPPn.Text)
        If Trim$(txtBwhTotal.Text) = "" Then RS!total_amount = CDbl(0) Else RS!total_amount = CDbl(txtBwhTotal.Text)
        If Trim$(txtBwhAmoIdr.Text) = "" Then RS!amount_IDR = CDbl(0) Else RS!amount_IDR = CDbl(txtBwhAmoIdr.Text)
        If Trim$(txtBwhPPnIdr.Text) = "" Then RS!ppn_IDR = CDbl(0) Else RS!ppn_IDR = CDbl(txtBwhPPnIdr.Text)
'        rs!ppn = CDbl(0) 'IIf(Trim$(txtBwhPPn.Text) = "", 0, CDbl(txtBwhPPn.Text))
'        rs!total_amount = CDbl(0) 'IIf(Trim$(txtBwhTotal.Text) = "", 0, CDbl(txtBwhTotal.Text))
'        rs!amount_IDR = CDbl(0) 'IIf(Trim$(txtBwhAmoIdr.Text) = "", 0, CDbl(txtBwhAmoIdr.Text))
'        rs!ppn_IDR = CDbl(0) 'IIf(Trim$(txtBwhPPnIdr.Text) = "", 0, CDbl(txtBwhPPnIdr.Text))
    RS.update
    RS.Close
End Sub

Private Sub CmdSubmit_Click()
    If cekAtas Then
        If cekGrid Then
            Call savemaster(False)
            Call savedetail
            LblErr = DisplayMsg("1101")
            Call Header
            Call browseitem
            Call BrowseGrid
        End If
    End If
End Sub

Private Function cekGrid() As Boolean
Dim u As Integer
With grid
If .Rows > 1 Then
    For u = 1 To .Rows - 1
        If .Cell(flexcpChecked, u, ColCheck) = flexChecked Then
            If Trim$(.TextMatrix(u, colQtyRetur)) = "" Then .TextMatrix(u, colQtyRetur) = "0.00"
            If CDbl(.TextMatrix(u, colQtyRetur)) = CDbl(0) Then
                .Col = colQtyRetur: .Row = u: .TopRow = u: .LeftCol = .Col: .SetFocus: grid_Click
                LblErr = "Please Input Qty Retur!": GoTo salaH
            Else
                If CDbl(.TextMatrix(u, colQtyRetur)) + CDbl(.TextMatrix(u, colTotRetur)) - CDbl(.TextMatrix(u, colOldRetur)) > CDbl(.TextMatrix(u, colQtyFaktur)) Then
                    .Col = colQtyRetur: .Row = u: .TopRow = u: .LeftCol = .Col: .SetFocus: grid_Click
                    LblErr = "Qty Retur must be the same or lower than " & Format(CDbl(.TextMatrix(u, colQtyFaktur)) - (CDbl(.TextMatrix(u, colTotRetur)) - CDbl(.TextMatrix(u, colOldRetur))), "#,##0")
                    GoTo salaH
                End If
                If cbType.ListIndex = 0 Then
                    If CDbl(.TextMatrix(u, colQtyRetur)) + CDbl(.TextMatrix(u, colTotRetur)) - CDbl(.TextMatrix(u, colOldRetur)) > CDbl(.TextMatrix(u, ColQty)) Then
                        .Col = colQtyRetur: .Row = u: .TopRow = u: .LeftCol = .Col: .SetFocus: grid_Click
                        LblErr = "Qty Retur must be the same or lower than " & Format(CDbl(.TextMatrix(u, ColQty)) - (CDbl(.TextMatrix(u, colTotRetur)) - CDbl(.TextMatrix(u, colOldRetur))), "#,##0")
                        GoTo salaH
                    End If
                End If
            End If
            If Trim$(.TextMatrix(u, colAdj)) = "" Then .TextMatrix(u, colAdj) = "0.00"
            If CDbl(.TextMatrix(u, colAdj)) = CDbl(0) Then
                .Col = colAdj: .Row = u: .TopRow = u: .LeftCol = .Col: .SetFocus: grid_Click
                LblErr = "Please Input Price Adj !": GoTo salaH
            End If
        End If
    Next u
Else
    LblErr = DisplayMsg("4047")
    cekGrid = False
    Exit Function
End If
cekGrid = True
Exit Function

salaH:
If Trim$(.TextMatrix(u, colQtyRetur)) = "" Or Trim$(.TextMatrix(u, colAdj)) = "" Then
Else
    .TextMatrix(u, colAmo) = Format(CDbl(.TextMatrix(u, colQtyRetur)) * CDbl(.TextMatrix(u, colAdj)), "#,##0.#0")
End If
cekGrid = False
End With
End Function

Private Sub savedetail()
Dim rsS As New ADODB.Recordset, u As Integer
With grid
For u = 1 To .Rows - 1
If .Cell(flexcpChecked, u, ColCheck) = flexChecked Then
    If .TextMatrix(u, colOld) = "0" Then 'klo tadinya gada maka add
        rsS.Open " select * from notaretur_Detail ", Db, 1, 3
        rsS.AddNew
            rsS!notaretur_no = Trim$(cbNota.List(cbNota.ListIndex, 0))
            rsS!fakturpajak_no = Trim$(cbPajak.List(cbPajak.ListIndex, 0))
            rsS!Item_Code = Trim$(.TextMatrix(u, colPart))
            rsS!user_Entry = userLogin
            rsS!date_entry = Now
            
    Else    'sisanya update
        rsS.Open " select * from notaretur_Detail where notaretur_no = '" & Trim$(.TextMatrix(u, colNotaRetur)) & "' " + _
                    " and fakturpajak_no = '" & Trim$(.TextMatrix(u, colFaktur)) & "' " + _
                    " And item_code = '" & Trim$(.TextMatrix(u, colPart)) & "' AND ReturnSeq_no=" & .TextMatrix(u, ColReturnSeq_no), Db, 1, 3
        If rsS.BOF Or rsS.EOF Then rsS.Close: Exit Sub
        rsS!user_update = userLogin
        rsS!date_update = Now
    End If
        If Trim(.TextMatrix(u, colReturnDate)) <> "" Then
            rsS!Return_Date = Format(.TextMatrix(u, colReturnDate), "YYYY-MM-DD")
        Else
            rsS!Return_Date = Null
        End If
        
        rsS!ReturnSeq_no = IIf(.TextMatrix(u, ColReturnSeq_no) = "", 0, .TextMatrix(u, ColReturnSeq_no))
        
        rsS!Qty = CDbl(.TextMatrix(u, colQtyRetur))
        rsS!Unit_cls = Trim(.TextMatrix(u, colUnitCls))
        rsS!currency_code = Trim(.TextMatrix(u, colCurrCode))
        rsS!Price = CDbl(.TextMatrix(u, ColPrice))
        rsS!adj_price = CDbl(.TextMatrix(u, colAdj))
        rsS!Amount = CDbl(.TextMatrix(u, colAmo))
    rsS.update
    
    rsS.Close
Else
    If .TextMatrix(u, colOld) = "1" Then    'klo ada bekas data yg di uncheck maka di delete dari detail
        Db.Execute (" delete notaretur_Detail where notaretur_no = '" & Trim$(.TextMatrix(u, colNotaRetur)) & "' " + _
                    " and fakturpajak_no = '" & Trim$(.TextMatrix(u, colFaktur)) & "' " + _
                    " and item_code = '" & Trim$(.TextMatrix(u, colPart)) & "' ")
'Sql = " delete notaretur_Detail where notaretur_no = '" & Trim$(cbNota.List(cbNota.ListIndex, 0)) & "' " + _
                    " and fakturpajak_no = '" & Trim$(cbPajak.List(cbPajak.ListIndex, 0)) & "' " + _
                    " and item_code = '" & Trim$(.TextMatrix(u, colPart)) & "' "
    End If
End If
Next u
End With
Set rsS = Nothing
End Sub

Private Sub cariBawah()
Dim rsBwh As New ADODB.Recordset, rate As Double

    txtBwhNota.Text = Trim$(txtNota.Text)
    txtBwhAmo.Text = Format(hitungTotal(colAmo), "#,##0.#0")
    
    ' cek country_cLs=0
    rsBwh.Open " select country_Cls from trade_Master where trade_Code = '" & Trim$(cbCust.List(cbCust.ListIndex, 0)) & "' ", Db, adOpenForwardOnly, adLockReadOnly
    If rsBwh.EOF Or rsBwh.BOF Then bwhCek.Close: Exit Sub
    If rsBwh!country_cls = "0" Then
    txtBwhPPn.Text = Format(CDbl(txtBwhAmo.Text) / 10, "#,##0.#0")
    Else
    txtBwhPPn.Text = "0.00"
    End If
    rsBwh.Close
    
    txtBwhTotal = Format(CDbl(txtBwhAmo) + CDbl(txtBwhPPn), "#,##0.#0")
    
    'cek currCode
    Select Case actCurr
        Case "YEN": If Trim$(lblYEN) = "--" Then rate = 0 Else rate = CDbl(lblYEN)
        Case "US$": If Trim$(lblUS) = "--" Then rate = 0 Else rate = CDbl(lblUS)
        Case "IDR": rate = 1
        Case "EUR": If Trim$(lblEur) = "--" Then rate = 0 Else rate = CDbl(lblEur)
        Case "S$": If Trim$(lblS) = "--" Then rate = 0 Else rate = CDbl(lblS)
    End Select
    txtBwhAmoIdr = Format(CDbl(txtBwhAmo) * rate, "#,##0.#0")
    txtBwhPPnIdr = Format(CDbl(txtBwhPPn) * rate, "#,##0.#0")
End Sub

Private Function hitungTotal(Kolom As Integer) As Double
Dim Q As Integer
hitungTotal = 0
With grid
For Q = 1 To .Rows - 1
    If .Cell(flexcpChecked, Q, ColCheck) = flexChecked Then
        hitungTotal = hitungTotal + CDbl(IIf(Trim$(.TextMatrix(Q, Kolom)) = "", 0, .TextMatrix(Q, Kolom)))
    End If
Next Q
End With
End Function

Private Sub cmdPrev_Click()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3
Dim sqlcekdet As String, SqlRpt As String
Dim rscekdet As New Recordset
Dim rsCP As New ADODB.Recordset, rsSp As New ADODB.Recordset, rsPom As New ADODB.Recordset
  
    If grid.Rows > 1 Then
        sqlcekdet = " select notaretur_No=rtrim(nrm.notaretur_No) from notaretur_Detail nrd  " & _
            " inner join notaretur_Master nrm on nrd.notaretur_no=nrm.notaretur_No " & _
            " inner join trade_Master tm on nrm.cust_Code = tm.trade_Code "

        Set rscekdet = Db.Execute(sqlcekdet)
        If rscekdet.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
        Set rscekdet = Nothing
        
        Me.MousePointer = vbHourglass
        
        SqlRpt = " select notaretur_No=rtrim(nrm.notaretur_No), fakturpajak_No=rtrim(nrd.fakturpajak_No),  " & _
                " trade_Code=rtrim(tm.Trade_Code), trade_Name=rtrim(tm.Trade_Name), npwp_Name = rtrim(tm.NPWP_Name),  " & _
                " npwp_Address=rtrim(tm.NPWP_address), npwp_City=rtrim(tm.NPWP_city), isnull(tm.postal_code,'') zip,  " & _
                " npwp_No=rtrim(tm.NPWP_No), country_Cls, item_Code=rtrim(nrd.item_Code), item_Name=rtrim(im.item_Name), makeritem_Code=rtrim(im.makeritem_Code), " & _
                " nrd.Qty, nrd.currency_Code, nrd.price, nrd.adj_Price, nrd.amount, nrm.notaretur_Date, " & _
                " (select tax_Position from company_Profile) as Tax_Position, " + _
                " (select tax_Person from company_Profile) as Tax_Person " + _
                " from notaretur_Detail nrd  " & _
                " left join item_Master im on nrd.item_Code = im.item_Code " & _
                " left join notaretur_Master nrm on nrd.notaretur_no=nrm.notaretur_No " & _
                " left join trade_Master tm on nrm.cust_Code = tm.trade_Code " + _
                " where nrm.notaretur_No = '" & Trim$(txtNota.Text) & "' "
        
        If rsRpt.State <> adStateClosed Then rsRpt.Close
        rsRpt.CursorLocation = adUseClient
        rsRpt.Open SqlRpt, Db, adOpenDynamic, adLockOptimistic
        
        sqlprint = SqlRpt
        reportcode = "NotaRetur"
        printorient = 1
        
        If rsRpt.EOF Then LblErr.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
        Set report = application.OpenReport(App.path & "\Reports\rptnotaretur.rpt")
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

