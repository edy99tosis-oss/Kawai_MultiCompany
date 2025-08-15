VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPriceMaster 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Price Master"
   ClientHeight    =   10560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmPriceMaster.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10560
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdUpload 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Upload"
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
      TabIndex        =   42
      Top             =   10050
      Width           =   1125
   End
   Begin VB.TextBox txtremarks 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.#0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Left            =   1200
      MaxLength       =   16
      TabIndex        =   9
      Top             =   8970
      Width           =   7065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1155
      Left            =   195
      TabIndex        =   33
      Top             =   1245
      Width           =   14865
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Left            =   3630
         TabIndex        =   41
         Top             =   262
         Width           =   300
      End
      Begin VB.Label LblCode 
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
         Index           =   2
         Left            =   7545
         TabIndex        =   39
         Top             =   315
         Width           =   960
      End
      Begin VB.Label LblCode 
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
         Index           =   1
         Left            =   4170
         TabIndex        =   38
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label lblitem 
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
         Index           =   0
         Left            =   5385
         TabIndex        =   37
         Top             =   315
         Width           =   2010
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   5385
         X2              =   7380
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price Cls"
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
         Left            =   300
         TabIndex        =   36
         Top             =   720
         Width           =   750
      End
      Begin MSForms.ComboBox cboitem 
         Height          =   315
         Left            =   1620
         TabIndex        =   0
         Top             =   255
         Width           =   1965
         VariousPropertyBits=   612386843
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "3466;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   8670
         X2              =   13860
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label LblCode 
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
         Index           =   0
         Left            =   300
         TabIndex        =   35
         Top             =   315
         Width           =   1155
      End
      Begin VB.Label lblitem 
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
         Index           =   1
         Left            =   8670
         TabIndex        =   34
         Top             =   315
         Width           =   5205
      End
      Begin MSForms.ComboBox cbopricecls 
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   675
         Width           =   1965
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "3466;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.#0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Left            =   7530
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   8520
      Width           =   705
   End
   Begin VB.TextBox txtUnit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   10185
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   8460
      Width           =   765
   End
   Begin VB.CommandButton cmdreport 
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10050
      Width           =   1125
   End
   Begin VB.TextBox lbldesc 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.#0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "Description"
      Top             =   8490
      Width           =   3780
   End
   Begin VB.TextBox txtprice 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.#0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Left            =   8355
      MaxLength       =   16
      TabIndex        =   5
      Top             =   8460
      Width           =   1755
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   630
      Left            =   195
      TabIndex        =   23
      Top             =   9285
      Width           =   14865
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
         Height          =   315
         Left            =   105
         TabIndex        =   24
         Top             =   195
         Width           =   14640
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
      Left            =   13935
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10050
      Width           =   1125
   End
   Begin VB.CommandButton command2 
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
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10050
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10050
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5355
      Left            =   195
      TabIndex        =   12
      Top             =   2520
      Width           =   14865
      _cx             =   26220
      _cy             =   9446
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13215
      TabIndex        =   17
      Top             =   450
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin MSMask.MaskEdBox mask 
      Height          =   315
      Left            =   12585
      TabIndex        =   7
      Top             =   8460
      Width           =   1230
      _ExtentX        =   2170
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
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker dtsdate 
      Height          =   315
      Left            =   11025
      TabIndex        =   6
      Top             =   8460
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
      CurrentDate     =   37781
   End
   Begin MSComCtl2.DTPicker dtedate 
      Height          =   315
      Left            =   12585
      TabIndex        =   14
      Top             =   8460
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
      Format          =   141230081
      CurrentDate     =   37781
   End
   Begin VB.Label Label2 
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
      Left            =   270
      TabIndex        =   40
      Top             =   9030
      Width           =   765
   End
   Begin VB.Line Line3 
      X1              =   7530
      X2              =   8250
      Y1              =   8760
      Y2              =   8760
   End
   Begin MSForms.ComboBox cboreason 
      Height          =   315
      Left            =   14145
      TabIndex        =   8
      Top             =   8460
      Width           =   795
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "1402;556"
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
      Caption         =   "Reason"
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
      Left            =   14220
      TabIndex        =   29
      Top             =   8040
      Width           =   630
   End
   Begin MSForms.ComboBox cbocurr 
      Height          =   315
      Left            =   6570
      TabIndex        =   4
      Top             =   8460
      Width           =   870
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "1535;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cbosupplier 
      Height          =   315
      Left            =   330
      TabIndex        =   2
      Top             =   8460
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
   Begin MSForms.ComboBox cbounit 
      Height          =   315
      Left            =   8640
      TabIndex        =   13
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
      VariousPropertyBits=   746604569
      DisplayStyle    =   7
      Size            =   "2143;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cbopriority 
      Height          =   315
      Left            =   5595
      TabIndex        =   3
      Top             =   8460
      Width           =   915
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "1614;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   195
      Top             =   8340
      Width           =   14865
   End
   Begin VB.Label Label1 
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
      Index           =   0
      Left            =   1755
      TabIndex        =   28
      Top             =   8040
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Priority"
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
      Left            =   5715
      TabIndex        =   27
      Top             =   8040
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   1755
      X2              =   5520
      Y1              =   8730
      Y2              =   8730
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
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
      Left            =   12870
      TabIndex        =   26
      Top             =   8040
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
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
      Left            =   11265
      TabIndex        =   25
      Top             =   8040
      Width           =   885
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Code"
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
      Left            =   420
      TabIndex        =   22
      Top             =   8040
      Width           =   1005
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Currency"
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
      Left            =   7020
      TabIndex        =   21
      Top             =   8040
      Width           =   795
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
      Left            =   8355
      TabIndex        =   20
      Top             =   8040
      Width           =   420
   End
   Begin VB.Label Label12 
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
      Left            =   10395
      TabIndex        =   19
      Top             =   8040
      Width           =   330
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Price Master"
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
      Left            =   195
      TabIndex        =   18
      Top             =   465
      Width           =   14865
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   195
      Top             =   7980
      Width           =   14865
   End
End
Attribute VB_Name = "frmPriceMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim ubah As Boolean, hapus As Boolean, gavalid As Boolean, ubahedate As Boolean
Dim SDate, EDate, sdateawal, edateakhir
Dim tcode As String, priority As String

Const isiPart = "Finish Good,Parts/wip/material"
Const isiPrice = "Purchase,Sales,Supply,Inventory,Service"

Dim bteColSelect As Byte
Dim bteColTrade As Byte
Dim bteColName As Byte
Dim bteColPriority As Byte
Dim bteColCurrCode As Byte
Dim bteColPrice As Byte
Dim bteColUnitCls As Byte
Dim bteColDateStart As Byte
Dim bteColDateEnd As Byte
Dim bteColReason As Byte
Dim bteColRemarks As Byte
Dim bteColLastUpdate As Byte
Dim bteColLastuser As Byte
Dim bteColUnit As Byte
Dim bteColCurr As Byte

Sub Header()
    
    bteColSelect = 0
    bteColTrade = 1
    bteColName = 2
    bteColPriority = 3
    bteColCurr = 4
    bteColCurrCode = 5
    bteColPrice = 6
    bteColUnitCls = 7
    bteColUnit = 8
    bteColDateStart = 9
    bteColDateEnd = 10
    bteColReason = 11
    bteColRemarks = 12
    bteColLastUpdate = 13
    bteColLastuser = 14
    
    With grid
        .clear
        .Rows = 1
        .ColS = 15
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColTrade) = "Trade"
        .TextMatrix(0, bteColName) = "Name"
        .TextMatrix(0, bteColPriority) = "Priority"
        .TextMatrix(0, bteColCurrCode) = "CurrCode"
        .TextMatrix(0, bteColCurr) = "Currency"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColUnitCls) = "UnitCls"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColDateStart) = "Start Date"
        .TextMatrix(0, bteColDateEnd) = "End Date"
        .TextMatrix(0, bteColReason) = "Reason"
        .TextMatrix(0, bteColRemarks) = "Remarks"
        .TextMatrix(0, bteColLastUpdate) = "Last Update"
        .TextMatrix(0, bteColLastuser) = "Username"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColTrade) = 1000
        .ColWidth(bteColName) = 3000
        .ColWidth(bteColPriority) = 800
        .ColWidth(bteColCurr) = 900
        .ColWidth(bteColPrice) = 1700
        .ColWidth(bteColUnit) = 1000
        .ColWidth(bteColDateStart) = 1400
        .ColWidth(bteColDateEnd) = 1400
        .ColWidth(bteColReason) = 2200
        .ColWidth(bteColRemarks) = 2200
        .ColWidth(bteColLastUpdate) = 2200
        .ColWidth(bteColLastuser) = 1500
                
        .ColDataType(bteColPrice) = flexDTCurrency
        .ColDataType(bteColDateStart) = flexDTDate
        .ColDataType(bteColDateEnd) = flexDTDate
        
        .ColHidden(bteColCurrCode) = True
        .ColHidden(bteColUnitCls) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, bteColLastUpdate) = flexAlignCenterCenter
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColTrade) = flexAlignLeftCenter
        .ColAlignment(bteColName) = flexAlignLeftCenter
        .ColAlignment(bteColPriority) = flexAlignCenterCenter
        .ColAlignment(bteColCurrCode) = flexAlignCenterCenter
        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        .ColAlignment(bteColPrice) = flexAlignRightCenter
        .ColAlignment(bteColUnitCls) = flexAlignCenterCenter
        .ColAlignment(bteColUnit) = flexAlignCenterCenter
        .ColAlignment(bteColDateStart) = flexAlignCenterCenter
        .ColAlignment(bteColDateEnd) = flexAlignCenterCenter
        .ColAlignment(bteColReason) = flexAlignLeftCenter
        .ColAlignment(bteColRemarks) = flexAlignLeftCenter
        .ColAlignment(bteColLastUpdate) = flexAlignLeftCenter
        .ColAlignment(bteColLastuser) = flexAlignCenterCenter
        
        .EditMaxLength = 1
    End With
End Sub

Sub Kosong()
    'CboItem.ListIndex = -1
    lblitem(0).Caption = ""
    lblitem(1).Caption = ""
    'cbopricecls.ListIndex = -1
    cbounit = ""
    LblErrMsg.Caption = ""
    txtUnit = ""
    txtDesc = ""
    kosonggrid
    Header
End Sub

Sub kosonggrid()
    cboSupplier.ListIndex = -1
    cboSupplier.Enabled = True
    lbldesc.Text = ""
    cbopriority.ListIndex = -1
    cbopriority.Enabled = True
    dtsdate.Value = Format(Now, "dd MMM yyyy")
    dtedate.Value = Format(Now, "dd MMM yyyy")
    mask.Text = "99/99/9999"
    cbocurr.ListIndex = -1
    txtprice.Text = ""
    cboreason.ListIndex = -1
    txtremarks.Text = ""
    ubah = False
End Sub

Sub adtocbosupplier()
Dim sqlcust As String
Dim RsCust As New Recordset
Dim i As Integer

    sqlcust = "select trade_code, trade_name from trade_master"
    Set RsCust = Db.Execute(sqlcust)
    
    With cboSupplier
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;300pt"
        .ListWidth = 350
        .ListRows = 15
        
        .AddItem
        .List(0, 0) = "000000"
        .List(0, 1) = "Common"
        
        i = 1
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("Trade_code"))
            .List(i, 1) = Trim(RsCust("Trade_Name"))
            RsCust.MoveNext
            i = i + 1
        
        Loop
    End With
    RsCust.Close
    Set RsCust = Nothing
End Sub

Sub adtocboitem()
Dim sqlitem As String
Dim RsItem As New Recordset
Dim i As Long

   sqlitem = "select item_code, makeritem_code, item_name , finishgoodpart_cls from item_master " & _
          "where use_endday >= convert(char(8), getdate(), 112) "
    Set RsItem = Db.Execute(sqlitem)
    
    With cboitem
        .clear
        .columnCount = 3
        .ColumnWidths = "120pt;120pt;240pt;0pt"
        .ListWidth = 500
        .ListRows = 15
        
        i = 0
        Do While Not RsItem.EOF
            .AddItem
            .List(i, 0) = Trim(RsItem("item_code"))
            .List(i, 1) = Trim(RsItem("makeritem_code"))
            .List(i, 2) = Trim(RsItem("item_Name"))
            .List(i, 3) = Split(isiPart, ",")(Val(Trim(RsItem("finishgoodpart_cls"))) - 1)
            RsItem.MoveNext
            i = i + 1
        
        Loop
    End With
RsItem.Close
Set RsItem = Nothing
End Sub

Sub adtocboreason()
Dim sqlreason As String
Dim rsreason As New Recordset
Dim i As Integer

    sqlreason = "select * from reason_cls "
    Set rsreason = Db.Execute(sqlreason)
    
    With cboreason
        .clear
        .columnCount = 2
        .ColumnWidths = "20pt;100pt"
        .ListWidth = 120
        .ListRows = 15
                
        i = 0
         Do While Not rsreason.EOF
            .AddItem
            .List(i, 0) = Trim(rsreason("reason_cls"))
            .List(i, 1) = Trim(rsreason("description"))
            
            rsreason.MoveNext
            i = i + 1
        Loop
    End With
rsreason.Close
Set rsreason = Nothing
End Sub

Sub adtocombo(nmCombo, nmField, mulai As Integer, akhir As Integer, lebar As Integer)   'Isi Combo Unit
Dim j As Integer, i As Integer

With nmCombo
    .clear
    .columnCount = 1
    .TextColumn = 1
    
    j = 0
    For i = mulai To akhir
        .AddItem Format(i + 1, "0#") & " - " & Split(nmField, ",")(i)
        j = j + 1
    Next i
    .ListRows = 9
    .ListWidth = lebar
    .ColumnWidths = lebar & " pt"
End With
End Sub

Sub formatprice()
Dim p1 As Byte, p2 As String, p0 As String
Dim jmldigit As Byte, jmldigit0 As Byte
Dim j As Integer

jmldigit = 0
    With grid
        For i = 1 To .Rows - 1
            If InStr(1, .TextMatrix(i, bteColPrice), ".") > 0 Then _
                jmldigit0 = Len(Trim(.TextMatrix(i, bteColPrice))) - InStr(1, Trim(.TextMatrix(i, bteColPrice)), ".")
            If jmldigit0 > jmldigit Then jmldigit = jmldigit0
        Next i

        For i = 1 To .Rows - 1
            p0 = Trim(.TextMatrix(i, bteColPrice))
            If InStr(1, p0, ".") > 0 Then
                p1 = Len(p0) - InStr(1, p0, ".")
                For j = 1 To jmldigit - p1
                    p2 = p0 & ""
                    p0 = p2
                Next j
            End If
            .TextMatrix(i, bteColPrice) = p0
        Next i
    
    End With
End Sub

Sub Browse()
    
    Dim RS As New Recordset
    Dim rsnama As New Recordset
    Dim rsreason As New Recordset
    Dim i As Long
    Dim nama As String, reason As String, sqlnama As String, sqlreason As String, p As Double
    Dim tglAwal As String, tglAkhir As String
        
    sql = "select * from price_master where  item_code='" & cboitem.Text & "' and price_cls='" & Left(cbopricecls.Text, 2) & "' " & _
        "order by item_code, price_cls, trade_code, priority_cls, start_date, end_date"
    RS.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    i = 1
    If Not (RS.BOF And RS.EOF) Then
        cbopricecls.Text = Trim(RS("Price_Cls")) & " - " & Split(isiPrice, ",")(Val(Trim(RS("Price_Cls"))) - 1)
        With grid
            Do While Not RS.EOF
                .Rows = .Rows + 1
                
                sqlnama = "select trade_name from trade_master where trade_code='" & RS!Trade_Code & "' "
                Set rsnama = Db.Execute(sqlnama)
                If Not (rsnama.BOF And rsnama.EOF) Then
                    nama = Trim(rsnama(0))
                Else
                    nama = "Common"
                End If
                rsnama.Close
                
                sqlreason = "select * from reason_cls where reason_cls='" & RS!reason_cls & "' "
                Set rsreason = Db.Execute(sqlreason)
                If Not (rsreason.BOF And rsreason.EOF) Then
                    reason = Trim(RS("Reason_Cls")) & " - " & Trim(rsreason("description"))
                Else
                    reason = ""
                End If
                rsreason.Close
                
                tglAwal = Mid(RS("start_date"), 5, 2) & "/" & Right(RS("start_date"), 2) & "/" & Left(RS("start_date"), 4)
                tglAkhir = IIf(IsNull(RS("End_date")), "99/99/9999", Mid(RS("end_date"), 5, 2) & "/" & Right(RS("end_date"), 2) & "/" & Left(RS("end_date"), 4))
            
                .TextMatrix(i, bteColTrade) = Trim(RS("trade_code"))
                .TextMatrix(i, bteColName) = nama
                .TextMatrix(i, bteColPriority) = RS("Priority_cls")
                .TextMatrix(i, bteColCurrCode) = Trim(RS("currency_code"))
                If IsNull(RS("currency_code")) Then
                  .TextMatrix(i, bteColCurr) = ""
                Else
                  .TextMatrix(i, bteColCurr) = uf_GetCurrencyDescription(Trim(RS("Currency_code")))
                End If
                
                If Left(cbopricecls.Text, bteColName) = "01" Then
                    .TextMatrix(i, bteColPrice) = IIf(IsNull(RS("price")), 0, Format(Trim(RS("Price")), gs_formatPrice))
                Else
                    .TextMatrix(i, bteColPrice) = IIf(IsNull(RS("price")), 0, Format(Trim(RS("Price")), gs_formatPrice))
                End If
                .TextMatrix(i, bteColUnitCls) = IIf(IsNull(Trim(RS("unit_cls"))), "", Trim(RS("unit_cls")))
                If IsNull(RS("unit_cls")) Then
                  .TextMatrix(i, bteColUnit) = ""
                Else
                  .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(Trim(RS("Unit_Cls")))
                End If
                .TextMatrix(i, bteColDateStart) = Format(tglAwal, "dd MMM yyyy")
                .TextMatrix(i, bteColDateEnd) = Format(tglAkhir, "dd MMM yyyy")
                .TextMatrix(i, bteColReason) = reason
                .TextMatrix(i, bteColRemarks) = Trim(RS("remarks") & "")
                .TextMatrix(i, bteColLastUpdate) = IIf(IsNull(RS("Last_Update")), "", Trim(RS("Last_Update")))
                .TextMatrix(i, bteColLastuser) = IIf(IsNull(RS("Last_User")), "", Trim(RS("Last_User")))
                .Cell(flexcpBackColor, i, bteColSelect) = &HFFFFFF
                
                RS.MoveNext
                i = i + 1
            Loop
            formatprice
        End With
    Else
        kosonggrid
        Header
        filterunit
    End If
    RS.Close
    
    Set rsnama = Nothing
    Set rsreason = Nothing
    Set RS = Nothing
    
End Sub

Private Sub filterunit()
Dim sql1 As String
Dim rs1 As New Recordset
    
    sql1 = "select unit_cls from item_master where item_code='" & cboitem.Text & "' "
    Set rs1 = Db.Execute(sql1)
    
    If Not (rs1.BOF And rs1.EOF) Then
        cbounit.ListIndex = -1
        For i = 0 To cbounit.ListCount - 1
            If Trim(rs1(0)) = Left(cbounit.List(i), 2) Then
                cbounit.ListIndex = i
                Exit For
            End If
        Next
    End If
    rs1.Close
    Set rs1 = Nothing
End Sub


Sub cektgl()
Dim rs2 As New Recordset
Dim rs3 As New Recordset
Dim Tgl As Date
Dim TempDate

gavalid = False
ubahedate = False

If hapus Then
    sql = "select * from price_master where item_code='" & cboitem.Text & "' and " & _
          "price_cls='" & Left(cbopricecls.Text, 2) & "' and trade_code='" & tcode & "' and " & _
          "priority_cls='" & priority & "' and " & _
          "start_date<'" & SDate & "' order by start_date, end_date"
    If rs2.State <> adStateClosed Then rs2.Close
    rs2.Open sql, Db, adOpenKeyset, adLockOptimistic


    sql = "select * from price_master where item_code='" & cboitem.Text & "' and " & _
          "price_cls='" & Left(cbopricecls.Text, 2) & "' and trade_code='" & tcode & "' and " & _
          "priority_cls='" & priority & "' and " & _
          "start_date>'" & SDate & "' order by start_date, end_date"
    If rs3.State <> adStateClosed Then rs3.Close
    rs3.Open sql, Db, adOpenKeyset, adLockOptimistic

    If Not (rs2.BOF And rs2.EOF) Then
        rs2.MoveLast
      If Not (rs3.BOF And rs3.EOF) Then
        rs3.MoveFirst
        Tgl = Mid(rs3("start_date"), 5, 2) & "/" & Right(rs3("start_date"), 2) & "/" & Left(rs3("start_date"), 4)
        TempDate = Format(DateAdd("d", -1, CDate(Tgl)), "yyyymmdd")

        sql = "update price_master " & _
            "set Last_Update = getdate(), Last_User = '" & userLogin & "', end_date='" & TempDate & "' " & _
            "where item_code='" & rs2!Item_Code & "' " & _
            "and price_cls='" & rs2!price_cls & "' " & _
            "and trade_code='" & rs2!Trade_Code & "' " & _
            "and priority_cls='" & rs2!priority_cls & "' " & _
            "and start_date='" & rs2!Start_Date & "'"
        Db.Execute sql

      Else
        sql = "update price_master " & _
            "set Last_Update = getdate(), Last_User = '" & userLogin & "', end_date='99999999' " & _
            "where item_code='" & rs2!Item_Code & "' " & _
            "and price_cls='" & rs2!price_cls & "' " & _
            "and trade_code='" & rs2!Trade_Code & "' " & _
            "and priority_cls='" & rs2!priority_cls & "' " & _
            "and start_date='" & rs2!Start_Date & "'"
        Db.Execute sql
      End If
    End If
    Exit Sub
    
End If

If ubah = False Then
SDate = Format(dtsdate.Value, "yyyymmdd")
EDate = Format(mask.Text, "yyyymmdd")

    sql = "select * from price_master where item_code='" & cboitem.Text & "' and " & _
          "price_cls='" & Left(cbopricecls.Text, 2) & "' and trade_code='" & cboSupplier.Text & "' and " & _
          "priority_cls='" & cbopriority.Text & "' and " & _
          "start_date<'" & SDate & "' order by start_date,end_date"
    If rs2.State <> adStateClosed Then rs2.Close
    rs2.Open sql, Db, adOpenKeyset, adLockOptimistic

    sql = "select * from price_master where item_code='" & cboitem.Text & "' and " & _
          "price_cls='" & Left(cbopricecls.Text, 2) & "' and trade_code='" & cboSupplier.Text & "' and " & _
          "priority_cls='" & cbopriority.Text & "' and " & _
          "start_date>'" & SDate & "' order by start_date, end_date"
    If rs3.State <> adStateClosed Then rs3.Close
    rs3.Open sql, Db, adOpenKeyset, adLockOptimistic

      If Not (rs3.BOF And rs3.EOF) Then
        rs3.MoveFirst
        
        Tgl = Mid(rs3("start_date"), 5, 2) & "/" & Right(rs3("start_date"), 2) & "/" & Left(rs3("start_date"), 4)
        TempDate = Format(CDate(Tgl), "yyyymmdd")
        
        If EDate = "99/99/9999" Then
            ubahedate = True
            edateakhir = Format(DateAdd("d", -1, CDate(Tgl)), "yyyymmdd")
        Else
            If (EDate >= TempDate) Then
                LblErrMsg.Caption = DisplayMsg(8054) & "" & Format(CDate(Tgl), "dd MMM yyyy")
                gavalid = True
                dtedate.SetFocus
                mask.SetFocus
                Exit Sub
            End If
        End If
      End If


    If Not (rs2.BOF And rs2.EOF) Then
        rs2.MoveLast
        TempDate = Format(DateAdd("d", -1, CDate(dtsdate.Value)), "yyyymmdd")

        sql = "update price_master " & _
            "set Last_Update = getdate(), Last_User = '" & userLogin & "', end_date='" & TempDate & "' " & _
            "where item_code='" & rs2!Item_Code & "' " & _
            "and price_cls='" & rs2!price_cls & "' " & _
            "and trade_code='" & rs2!Trade_Code & "' " & _
            "and priority_cls='" & rs2!priority_cls & "' " & _
            "and start_date='" & rs2!Start_Date & "' "
        Db.Execute sql
    End If
    Exit Sub
Else

SDate = Format(dtsdate.Value, "yyyymmdd")
EDate = Format(mask.Text, "yyyymmdd")
    
    sql = "select * from price_master where item_code='" & cboitem.Text & "' and " & _
          "price_cls='" & Left(cbopricecls.Text, 2) & "' and trade_code='" & cboSupplier.Text & "' and " & _
          "priority_cls='" & cbopriority.Text & "' and " & _
          "start_date<'" & sdateawal & "' order by start_date,end_date"
    If rs2.State <> adStateClosed Then rs2.Close
    rs2.Open sql, Db, adOpenKeyset, adLockOptimistic

    sql = "select * from price_master where item_code='" & cboitem.Text & "' and " & _
          "price_cls='" & Left(cbopricecls.Text, 2) & "' and trade_code='" & cboSupplier.Text & "' and " & _
          "priority_cls='" & cbopriority.Text & "' and " & _
          "start_date>'" & sdateawal & "' order by start_date, end_date"
    If rs3.State <> adStateClosed Then rs3.Close
    rs3.Open sql, Db, adOpenKeyset, adLockOptimistic

      If Not (rs3.BOF And rs3.EOF) Then
        rs3.MoveFirst
        
        Tgl = Mid(rs3("start_date"), 5, 2) & "/" & Right(rs3("start_date"), 2) & "/" & Left(rs3("start_date"), 4)
        TempDate = Format(CDate(Tgl), "yyyymmdd")
        
        If EDate = "99/99/9999" Then
            ubahedate = True
            edateakhir = Format(DateAdd("d", -1, CDate(Tgl)), "yyyymmdd")
        Else
            If (EDate >= TempDate) Then
                LblErrMsg.Caption = DisplayMsg(8054) & "" & Format(CDate(Tgl), "dd MMM yyyy")
                gavalid = True
                dtedate.SetFocus
                mask.SetFocus
                Exit Sub
            End If
        End If
      End If
    
    If Not (rs2.BOF And rs2.EOF) Then
        rs2.MoveLast
        Tgl = Mid(rs2("start_date"), 5, 2) & "/" & Right(rs2("start_date"), 2) & "/" & Left(rs2("start_date"), 4)
        TempDate = Format(CDate(Tgl), "yyyymmdd")

        If (SDate <= TempDate) Then
            LblErrMsg.Caption = DisplayMsg(8055) & "" & Format(CDate(Tgl), "dd MMM yyyy")
            gavalid = True
            dtsdate.SetFocus
            Exit Sub
        Else
        
        TempDate = Format(DateAdd("d", -1, CDate(dtsdate.Value)), "yyyymmdd")
        sql = "update price_master set " & _
            "Last_Update = getdate(), Last_User = '" & userLogin & "', end_date='" & TempDate & "' " & _
            "where item_code='" & rs2!Item_Code & "' " & _
            "and price_cls='" & rs2!price_cls & "' " & _
            "and trade_code='" & rs2!Trade_Code & "' " & _
            "and priority_cls='" & rs2!priority_cls & "' " & _
            "and start_date='" & rs2!Start_Date & "' "
        Db.Execute sql
        
        End If
    End If
    
    Exit Sub
End If
End Sub

Private Sub cbocurr_Change()
If cbocurr.ListIndex >= 0 Then
txtDesc = cbocurr.List(cbocurr.ListIndex, 1)
Else
txtDesc = ""
End If
End Sub

Private Sub CboUnit_Change()
If cbounit.ListIndex >= 0 Then
txtUnit = cbounit.List(cbounit.ListIndex, 1)
Else
txtUnit = ""
End If
End Sub

Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = cboitem.Text
 frm_BrowseItem.Show 1
 cboitem.Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub cmdUpload_Click()
Me.MousePointer = vbHourglass
    FrmPriceUpload.Show 1
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    
   On Error GoTo errHandler
    
    
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    LblErrMsg = 1
    Kosong
    LblErrMsg = 2
    adtocbosupplier
    LblErrMsg = 3
    Header
    LblErrMsg = 4
    adtocboitem
    LblErrMsg = 5
    adtocboreason
    LblErrMsg = 6
    
'    cbopricecls.AddItem "01 - Purchase"
'    cbopricecls.AddItem "02 - Sales"
'    cbopricecls.AddItem "03 - Supply"
'    cbopricecls.AddItem "04 - Inventory"
'    cbopricecls.AddItem "05 - Service"
    
    
    
   
    
    Call adtocombo(cbopricecls, isiPrice, 0, 4, 90)
    LblErrMsg = 7
    Call up_FillCombo(cbounit, "unit_cls")
    LblErrMsg = 8
    Call up_FillCombo(cbocurr, "curr_cls")
    LblErrMsg = 9
    LblErrMsg = ""
    cbopriority.ColumnWidths = "30 pt"
    cbopriority.ListWidth = 30
    cbopriority.AddItem 0
    cbopriority.AddItem 1
    
    
errHandler:
    
    LblErrMsg.Caption = LblErrMsg.Caption
    
End Sub

Private Sub cboitem_Click()
    LblErrMsg = ""

    If cboitem.ListIndex <> -1 Then
        lblitem(0).Caption = cboitem.Column(1)
        lblitem(1).Caption = cboitem.Column(2)
        cbopricecls_Click
        filterunit
    Else
        lblitem(0) = ""
        lblitem(1) = ""
        Call up_FillCombo(cbounit, "unit_cls")
        kosonggrid
        Header
        cboitem.SetFocus
        LblErrMsg.Caption = DisplayMsg(4003)
        Exit Sub
    End If
End Sub

Private Sub cboitem_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then
    For i = 0 To cboitem.ListCount - 1
        If cboitem.Text = cboitem.List(i) Then
            cboitem.ListIndex = i
            Exit For
        End If
    Next
    cboitem_Click
  End If
End Sub

Private Sub cbopricecls_Click()
    If cbopricecls.ListIndex <> -1 Then
        If (cboitem.ListIndex <> -1) Then
            kosonggrid
            Header
            Browse
        Else
            kosonggrid
            Header
        End If
    Else
        kosonggrid
        Header
    End If
End Sub

Private Sub cbopricecls_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then cbopricecls_Click
End Sub

Private Sub cbosupplier_Click()
LblErrMsg = ""

    If cboSupplier.ListIndex <> -1 Then
        lbldesc.Text = cboSupplier.Column(1)
    Else
        lbldesc.Text = ""
        LblErrMsg.Caption = DisplayMsg(4013)
        cboSupplier.SetFocus
        Exit Sub
    End If
End Sub

Private Sub cbosupplier_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then
    For i = 0 To cboSupplier.ListCount - 1
        If cboSupplier.Text = cboSupplier.List(i) Then
            cboSupplier.ListIndex = i
            Exit For
        End If
    Next
   cbosupplier_Click
  End If
End Sub

Private Sub dtsdate_Change()
If mask.Text <> "99/99/9999" Then
   If CDate(dtsdate) > CDate(dtedate) Then
      LblErrMsg.Caption = DisplayMsg(4068)
      dtsdate.SetFocus
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
End If
End Sub

Private Sub dtsdate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then dtsdate_Change
End Sub

Private Sub dtedate_Change()
    mask.Text = Format(dtedate, "MM/dd/yyyy")
   If CDate(dtedate) < CDate(dtsdate) Then
      LblErrMsg.Caption = DisplayMsg(4066)
      mask.SetFocus
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If

End Sub

Private Sub dtedate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then dtedate_Change
End Sub

Private Sub mask_LostFocus()
    If IsDate(mask.Text) = True Then dtedate.Value = CDate(mask.Text)
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim TextGrid As String
Dim k As Boolean
Dim j As Integer

k = False
With grid
    TextGrid = grid.Text

    If TextGrid = "S" Then
        
        cboSupplier.Text = .TextMatrix(Row, bteColTrade)
        lbldesc.Text = .TextMatrix(Row, bteColName)
        cboSupplier.Enabled = False
        
        cbopriority.ListIndex = -1
        For i = 0 To 1
            If .TextMatrix(Row, bteColPriority) = cbopriority.List(i) Then
                cbopriority.ListIndex = i
                Exit For
            End If
        Next
        cbopriority.Enabled = False
        
        cbocurr.ListIndex = -1
        For i = 0 To 4
            If .TextMatrix(Row, bteColCurrCode) = cbocurr.List(i) Then
                cbocurr.ListIndex = i
                Exit For
            End If
        Next
        
        If Left(cbopricecls.Text, bteColName) = "01" Then
            txtprice.Text = Format(CDbl(.TextMatrix(Row, bteColPrice)), gs_formatPrice)
        Else
            txtprice.Text = Format(CDbl(.TextMatrix(Row, bteColPrice)), gs_formatPrice)
        End If
        
        cbounit.ListIndex = -1
        For i = 0 To cbounit.ListCount - 1
            If .TextMatrix(Row, bteColUnitCls) = cbounit.List(i) Then
                cbounit.ListIndex = i
                Exit For
            End If
        Next
        
        dtsdate.Value = Format(.TextMatrix(Row, bteColDateStart), "mm/dd/yyyy")
        sdateawal = Format(.TextMatrix(Row, bteColDateStart), "yyyymmdd")
        
        mask.Text = Format(.TextMatrix(Row, bteColDateEnd), "mm/dd/yyyy")
        If .TextMatrix(Row, bteColDateEnd) <> "99/99/9999" Then
            dtedate = Format(.TextMatrix(Row, bteColDateEnd), "mm/dd/yyyy")
        End If
        
        cboreason.ListIndex = -1
        For i = 0 To cboreason.ListCount - 1
            If .TextMatrix(Row, bteColReason) = (cboreason.List(i, 0) & " - " & cboreason.List(i, 1)) Then
                cboreason.ListIndex = i
                Exit For
            End If
        Next
        txtremarks.Text = .TextMatrix(Row, bteColRemarks)
                
        ubah = True
       Call kosongColGrid
    ElseIf TextGrid = "D" Then
       Call kosongColGrid("S")
    End If

    .TextMatrix(Row, Col) = TextGrid
    
    
    For j = 1 To .Rows - 1
        If .TextMatrix(j, bteColSelect) <> "" Then
            k = True
        End If
    Next j
    
    If k = False Then
        ubah = False
        cboSupplier.ListIndex = -1
        cboSupplier.Enabled = True
        lbldesc.Text = ""
        cbopriority.ListIndex = -1
        cbopriority.Enabled = True
        dtsdate.Value = Format(Now, "dd MMM yyyy")
        dtedate.Value = Format(Now, "dd MMM yyyy")
        mask.Text = "99/99/9999"
        cbocurr.ListIndex = -1
        txtprice.Text = ""
        cbounit.ListIndex = -1
        cboreason.ListIndex = -1
        txtremarks.Text = ""
    End If
    
End With

End Sub

Private Sub kosongColGrid(Optional Kolom As String)
    Dim i As Integer
    
    With grid
        .Col = bteColSelect
    
        If Kolom <> "" Then
           For i = 1 To .Rows - 1

              If .Text = Kolom Then .Text = ""
              If .TextMatrix(i, bteColSelect) <> "D" Then .TextMatrix(i, bteColSelect) = ""
           Next i
           kosonggrid
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
Dim sql1 As String, tanya
Dim RS As New Recordset

hapus = False
Select Case Index
Case 0:
        If hakUpdate(Me.Name) = 0 Then _
            LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub

          If cboitem.Text = "" Then
            cboitem.SetFocus
            LblErrMsg = DisplayMsg(1024)
            Exit Sub
          ElseIf cbopricecls.Text = "" Then
            cbopricecls.SetFocus
            LblErrMsg.Caption = DisplayMsg(1026)
            Exit Sub

          Else
          
          If cboitem.Text <> "" Then
            cboitem.MatchEntry = 1
            cboitem.Text = cboitem.Text
            If cboitem.MatchFound = False Then
                LblErrMsg = DisplayMsg(4003)
                cboitem.SetFocus
                cboitem.MatchEntry = 2
                Exit Sub
            End If
            cboitem.MatchEntry = 2
          End If
          
           
            With grid
                For i = 1 To .Rows - 1
                  If .TextMatrix(i, bteColSelect) = "D" Then
                    If IsEmpty(tanya) Then tanya = MsgBox("Do You Really Want To Delete This Data ?", vbQuestion & vbYesNo, "Confirmation")
                    If tanya = vbYes Then

                            sql1 = "delete from price_master where item_Code='" & cboitem.Text & "' and " & _
                                   "price_cls='" & Left(cbopricecls.Text, 2) & "' and trade_code='" & _
                                   .TextMatrix(i, 1) & "' and priority_cls='" & .TextMatrix(i, bteColPriority) & "' and " & _
                                   "start_date='" & Format(.TextMatrix(i, bteColDateStart), "yyyymmdd") & "'"
                            Db.Execute sql1
                        
                            hapus = True
                            SDate = Format(.TextMatrix(i, bteColDateStart), "yyyymmdd")
                            EDate = Format(.TextMatrix(i, bteColDateEnd), "yyyymmdd")
                            tcode = .TextMatrix(i, bteColTrade)
                            priority = .TextMatrix(i, bteColPriority)
                            
                            cektgl

                    Else
                        Exit For
                    End If
                  End If
                Next i
                
                If (hapus) Then kosonggrid: Header: Browse: filterunit: LblErrMsg = DisplayMsg(1201): Exit Sub
            
                  If cboSupplier.Text = "" Then
                    cboSupplier.SetFocus
                    LblErrMsg = DisplayMsg(1027)
                    Exit Sub
                  ElseIf cbopriority.Text = "" Then
                    cbopriority.SetFocus
                    LblErrMsg = DisplayMsg(1025)
                    Exit Sub
                  ElseIf cbocurr.Text = "" Then
                    cbocurr.SetFocus
                    LblErrMsg = DisplayMsg(1028)
                    Exit Sub
                  ElseIf txtprice.Text = "" Or IsNumeric(txtprice) = False Then
                    txtprice.SetFocus
                    LblErrMsg = DisplayMsg(8023)
                    Exit Sub
                   ElseIf CDbl(txtprice.Text) > gd_MaxPrice Then
                    txtprice.SetFocus
                    LblErrMsg = DisplayMsg(4048) & "" & gd_MaxPrice
                   Exit Sub
                  ElseIf cbounit.Text = "" Then
'                    cboUnit.SetFocus
                    LblErrMsg = DisplayMsg(1030)
                   Exit Sub
                  ElseIf cboreason.Text = "" Then
                    cboreason.SetFocus
                    LblErrMsg = DisplayMsg(1037)
                   Exit Sub
    
                  Else
                  
                    If cboSupplier.Text <> "" Then
                      cboSupplier.MatchEntry = 1
                      cboSupplier.Text = cboSupplier.Text
                      If cboSupplier.MatchFound = False Then
                          LblErrMsg = DisplayMsg(4013)
                          cboSupplier.SetFocus
                          cboSupplier.MatchEntry = 2
                          Exit Sub
                      End If
                      cboSupplier.MatchEntry = 2
                    End If
                  
                  If mask.Text <> "99/99/9999" Then
                       If IsDate(mask.Text) = False Then
                          LblErrMsg.Caption = DisplayMsg(4065) '"End Date is not valid"
                          mask.SetFocus
                          Exit Sub
                       End If
                       
                       If CDate(dtsdate) > CDate(dtedate) Then
                          LblErrMsg.Caption = DisplayMsg(4068)
                          dtsdate.SetFocus
                          Exit Sub
                       End If
                  End If
                    
                    If ubah = False Then
                        
                        sql = "select * from price_master where item_Code='" & cboitem.Text & "' and price_cls ='" & _
                            Left(cbopricecls.Text, 2) & "' and trade_code='" & cboSupplier.Text & _
                            "' and priority_cls='" & cbopriority.Text & "' and start_date='" & _
                            Format(dtsdate.Value, "yyyymmdd") & "' "
                        RS.Open sql, Db, adOpenDynamic, adLockOptimistic, adCmdText
                        
                        If Not (RS.EOF And RS.BOF) Then
                            LblErrMsg = DisplayMsg(1023): dtsdate.SetFocus: Exit Sub
                        Else
                            cektgl
                            If gavalid Then Exit Sub
                            RS.AddNew
                            RS("item_Code") = cboitem.Text
                            RS("price_cls") = Left(cbopricecls.Text, 2)
                        End If
                    
                    Else
                        
                        sql = "select * from price_master " & _
                            "where item_Code='" & cboitem.Text & "' and price_cls ='" & _
                            Left(cbopricecls.Text, 2) & "' and trade_code='" & cboSupplier.Text & _
                            "' and priority_cls='" & cbopriority.Text & "' and start_date='" & _
                            sdateawal & "' "
                        RS.Open sql, Db, adOpenDynamic, adLockOptimistic, adCmdText
                    
                    End If
                    
                    cektgl
                    If gavalid Then Exit Sub
                    
                    RS("Trade_code") = cboSupplier.Text
                    RS("Priority_cls") = cbopriority.Text
                    RS("Currency_code") = Left(cbocurr.Text, 2)
                    RS("Price") = txtprice.Text
                    RS("Unit_cls") = Left(cbounit.Text, 2)
                    RS("start_date") = Format(dtsdate.Value, "yyyymmdd")
                    
                    If mask.Text = "99/99/9999" Then
                       If ubahedate = True Then
                         RS("End_date") = edateakhir
                       Else
                         RS("end_date") = "99999999"
                       End If
                    Else
                        RS("end_date") = Format(mask.Text, "yyyymmdd")
                    End If
                    
                    RS("reason_cls") = cboreason.Text
                    RS("remarks") = txtremarks.Text
                    RS("Last_Update") = Now
                    RS("Last_User") = userLogin
                    
                    RS.update
                    RS.Close
                    
                    LblErrMsg = DisplayMsg(IIf((ubah = False), 1000, 1101))
                    
                    kosonggrid
                    Header
                    Browse
                    filterunit
                            
                End If
            End With
          End If
          
Case 1
    Kosong
    cboitem.SetFocus
End Select
Set RS = Nothing

End Sub

Private Sub command2_Click()
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

Private Sub txtPrice_LostFocus()
Dim z As Double
If Left(cbopricecls.Text, 2) = "01" Then
    txtprice.Text = Format(txtprice.Text, gs_formatPrice)
Else
    txtprice.Text = Format(txtprice.Text, gs_formatPrice)
End If
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
End If
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cmdReport_Click()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3

  
  Me.MousePointer = vbHourglass
  
  sql = "select tb1.* from ( " & _
        "select rtrim(prim.price_cls) price_cls, rtrim(im.makeritem_code) makeritem_code, " & _
        "rtrim(prim.item_code) item_code, rtrim(im.item_name) item_name, rtrim(prim.trade_code) trade_code, " & _
        "rtrim(tm.trade_name) trade_name, rtrim(prim.priority_cls) priority_cls, " & _
        "right(prim.start_date,2) + '-' + substring(prim.start_date,5,2) + '-' + left(prim.start_date,4) as start_date, " & _
        "right(prim.end_date,2) + '-' + substring(prim.end_date,5,2) + '-' + left(prim.end_date,4) as end_date, " & _
        "rtrim(prim.unit_cls) unit_cls, " & _
        "(select rtrim(description) from unit_cls uc where uc.unit_cls=prim.unit_cls) unit_desc, " & _
        "rtrim(prim.currency_code) currency_code, " & _
        "(select rtrim(description) from curr_cls where curr_cls=prim.currency_code) Curr_desc, " & _
        "prim.price, rtrim(rc.description) reason " & _
        "from price_master prim " & _
        "inner join item_master im on prim.item_code = im.item_code " & _
        "left join reason_cls rc on prim.reason_cls = rc.reason_cls " & _
        "inner join trade_master tm on prim.trade_code = tm.trade_code "
  sql = sql & _
        "Union All "
  sql = sql & _
        "select rtrim(prim.price_cls) price_cls, rtrim(im.makeritem_code) makeritem_code, " & _
        "rtrim(prim.item_code) item_code, rtrim(im.item_name) item_name, rtrim(prim.trade_code) trade_code, " & _
        "trade_name = 'Common', rtrim(prim.priority_cls) priority_cls, " & _
        "right(prim.start_date,2) + '-' + substring(prim.start_date,5,2) + '-' + left(prim.start_date,4) as start_date, " & _
        "right(prim.end_date,2) + '-' + substring(prim.end_date,5,2) + '-' + left(prim.end_date,4) as end_date, " & _
        "rtrim(prim.unit_cls) unit_cls, " & _
        "(select rtrim(description) from unit_cls uc where uc.unit_cls=prim.unit_cls) unit_desc, " & _
        "prim.currency_code, " & _
        "(select rtrim(description) from curr_cls where curr_cls=prim.currency_code) Curr_desc, " & _
        "prim.price, rtrim(rc.description) reason " & _
        "from price_master prim " & _
        "inner join item_master im on prim.item_code = im.item_code " & _
        "left join reason_cls rc on prim.reason_cls = rc.reason_cls " & _
        "where prim.trade_code = '000000' " & _
        ")tb1 "
  
  If Trim(cbopricecls.Text) <> "" Then
   sql = sql & _
         "where tb1.price_cls = '" & Left(Trim(cbopricecls.Text), 2) & "' "
  End If
        
   sql = sql & _
        "order by tb1.price_cls, tb1.makeritem_code, tb1.item_code, tb1.trade_code, tb1.priority_cls, tb1.start_date "
  
  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
  
  If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
  sqlprint = sql
  reportcode = "pricemaster"
  printorient = 2
  Set report = application.OpenReport(App.path & "\Reports\rptPriceMaster.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  
    '#####################################################################
    '# Price Digit and decimal
    report.FormulaFields(5).Text = "" & gi_decimalDigitPrice & ""
    report.FormulaFields(6).Text = "" & gi_decimalDigitPriceIDR & ""
    '#####################################################################
  Rpt.CRViewer1.ReportSource = report
  Rpt.CRViewer1.ViewReport
  Rpt.CRViewer1.Zoom 1
  
  Rpt.WindowState = 2
  Rpt.Show 1
  
  Me.MousePointer = vbDefault
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub
