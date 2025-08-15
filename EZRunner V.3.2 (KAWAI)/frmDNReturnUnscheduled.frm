VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDNReturnUnscheduled 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivery Note Return Unscheduled"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   Icon            =   "frmDNReturnUnscheduled.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,###"
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
      Height          =   345
      Left            =   13590
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   8190
      Width           =   1320
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,###"
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
      Height          =   345
      Left            =   10740
      TabIndex        =   49
      Top             =   8190
      Width           =   1380
   End
   Begin VB.TextBox TxtQty 
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
      Left            =   7710
      TabIndex        =   48
      Top             =   8160
      Width           =   705
   End
   Begin VB.TextBox TxtLotNo 
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
      Left            =   6780
      MaxLength       =   10
      TabIndex        =   47
      Top             =   8160
      Width           =   855
   End
   Begin VB.TextBox txtService 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,###"
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
      Height          =   345
      Left            =   12180
      TabIndex        =   46
      Top             =   8190
      Width           =   1380
   End
   Begin VB.TextBox tmpQty 
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
      Height          =   285
      Left            =   120
      TabIndex        =   45
      Top             =   150
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtRemark 
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
      Left            =   5730
      MaxLength       =   25
      TabIndex        =   11
      Top             =   8670
      Width           =   9225
   End
   Begin VB.TextBox txtReturnSeq_No 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,###"
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
      Height          =   285
      Left            =   3360
      TabIndex        =   41
      Top             =   240
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox txtRef 
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
      Left            =   1050
      MaxLength       =   25
      TabIndex        =   10
      Top             =   8640
      Width           =   3885
   End
   Begin VB.CommandButton cmdPopItem 
      Caption         =   "..."
      Height          =   315
      Left            =   2490
      TabIndex        =   7
      Top             =   8160
      Width           =   315
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   300
      TabIndex        =   34
      Top             =   9270
      Width           =   14700
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
         TabIndex        =   35
         Top             =   180
         Width           =   14265
      End
   End
   Begin VB.TextBox txtDesc 
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
      Left            =   2850
      TabIndex        =   32
      Tag             =   "1"
      Top             =   8190
      Width           =   2460
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1185
      Left            =   270
      TabIndex        =   17
      Top             =   990
      Width           =   14730
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0000FFFF&
         Caption         =   "Searc&h"
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
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   630
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtStart 
         Height          =   330
         Left            =   2010
         TabIndex        =   2
         Top             =   630
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
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
         Format          =   137756675
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtEnd 
         Height          =   330
         Left            =   4200
         TabIndex        =   3
         Top             =   630
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
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
         Format          =   137756675
         CurrentDate     =   37798
      End
      Begin VB.Label lblCust 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         Left            =   3750
         TabIndex        =   21
         Top             =   285
         Width           =   6015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
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
         Left            =   270
         TabIndex        =   20
         Top             =   285
         Width           =   1350
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3750
         X2              =   9750
         Y1              =   525
         Y2              =   525
      End
      Begin MSForms.ComboBox CboCust 
         Height          =   315
         Left            =   2010
         TabIndex        =   1
         Top             =   210
         Width           =   1635
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2884;556"
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
         Caption         =   "Return date from"
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
         TabIndex        =   19
         Top             =   705
         Width           =   1470
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
         Index           =   2
         Left            =   3870
         TabIndex        =   18
         Top             =   705
         Width           =   165
      End
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9960
      Width           =   1125
   End
   Begin VB.CommandButton cmdProcess 
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
      Left            =   13860
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9930
      Width           =   1125
   End
   Begin VB.CommandButton cmdProcess 
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
      Left            =   11415
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9930
      Width           =   1125
   End
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Cancel"
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
      Left            =   12630
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9930
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4845
      Left            =   300
      TabIndex        =   16
      Top             =   2250
      Width           =   14760
      _cx             =   26035
      _cy             =   8546
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
      Rows            =   1
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   275
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDNReturnUnscheduled.frx":0E42
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
      Begin MSComCtl2.DTPicker dtReturn2 
         Height          =   315
         Left            =   0
         TabIndex        =   36
         Top             =   300
         Visible         =   0   'False
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
         Format          =   137101315
         CurrentDate     =   37859
      End
   End
   Begin MSComCtl2.DTPicker dtReturn 
      Height          =   315
      Left            =   1410
      TabIndex        =   5
      Top             =   7290
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
      Format          =   137101315
      CurrentDate     =   37859
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13080
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   150
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin MSForms.ComboBox cboType 
      Height          =   315
      Left            =   9060
      TabIndex        =   44
      Top             =   8190
      Width           =   840
      VariousPropertyBits=   746604571
      MaxLength       =   4
      DisplayStyle    =   3
      Size            =   "1482;556"
      TextColumn      =   2
      MatchEntry      =   0
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label 
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
      Index           =   13
      Left            =   5010
      TabIndex        =   42
      Top             =   8700
      Width           =   675
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
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
      Index           =   13
      Left            =   8490
      TabIndex        =   40
      Top             =   8190
      Width           =   450
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref No."
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
      Index           =   9
      Left            =   420
      TabIndex        =   39
      Top             =   8670
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supply Cls"
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
      Left            =   9030
      TabIndex        =   38
      Top             =   7740
      Width           =   915
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service"
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
      Left            =   12270
      TabIndex        =   37
      Top             =   7740
      Width           =   645
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
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
      Index           =   12
      Left            =   300
      TabIndex        =   33
      Top             =   7350
      Width           =   1035
   End
   Begin MSForms.ComboBox cboProduct 
      Height          =   315
      Left            =   330
      TabIndex        =   6
      Top             =   8130
      Width           =   2145
      VariousPropertyBits=   746604571
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "3784;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2790
      X2              =   5310
      Y1              =   8460
      Y2              =   8460
   End
   Begin MSForms.ComboBox cboWhCode 
      Height          =   315
      Left            =   5460
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1275
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "2249;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   8460
      X2              =   9030
      Y1              =   8460
      Y2              =   8460
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
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
      Left            =   10020
      TabIndex        =   31
      Top             =   8250
      Width           =   480
   End
   Begin MSForms.ComboBox cbocurr 
      Height          =   315
      Left            =   9930
      TabIndex        =   9
      Top             =   8190
      Width           =   825
      VariousPropertyBits=   746604571
      MaxLength       =   4
      DisplayStyle    =   3
      Size            =   "1455;556"
      TextColumn      =   2
      MatchEntry      =   0
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label 
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
      Index           =   10
      Left            =   7860
      TabIndex        =   30
      Top             =   7740
      Width           =   300
   End
   Begin VB.Label Label 
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
      Index           =   8
      Left            =   8430
      TabIndex        =   29
      Top             =   7740
      Width           =   330
   End
   Begin VB.Label Label 
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
      Index           =   7
      Left            =   10050
      TabIndex        =   28
      Top             =   7740
      Width           =   390
   End
   Begin VB.Label Label 
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
      Index           =   6
      Left            =   10860
      TabIndex        =   27
      Top             =   7740
      Width           =   420
   End
   Begin VB.Label Label 
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
      Index           =   5
      Left            =   13650
      TabIndex        =   26
      Top             =   7740
      Width           =   660
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lot No."
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
      Left            =   6900
      TabIndex        =   25
      Top             =   7740
      Width           =   600
   End
   Begin VB.Label Label 
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
      Left            =   5640
      TabIndex        =   24
      Top             =   7740
      Width           =   795
   End
   Begin VB.Label Label 
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
      Index           =   1
      Left            =   3090
      TabIndex        =   23
      Top             =   7740
      Width           =   960
   End
   Begin VB.Label Label 
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
      Left            =   630
      TabIndex        =   22
      Top             =   7740
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   0
      Left            =   270
      Top             =   7680
      Width           =   14730
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   1155
      Index           =   0
      Left            =   270
      Top             =   8040
      Width           =   14730
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Note Return Unscheduled"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   14775
   End
End
Attribute VB_Name = "frmDNReturnUnscheduled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'RJ

' Update By Dudi Santosa, desember 4 2008
Option Explicit
Dim dbTrans As New ADODB.Connection
Dim ClsProc As New ClsProc
Dim i As Integer, HakU As Integer, posisi As Integer
Dim nilKosong As Boolean, ubahgrid As Boolean
Dim ErrorCol As Integer, KeyInput As Integer

Dim ColS As Integer
Dim ColDoNo As Integer
Dim ColProductCD As Integer, ColProductDesc As Integer, ColLotNo As Integer, ColCaseNo As Integer
Dim ColRQty As Integer, ColSpoolQty As Integer, ColUnit As Integer
Dim ColCurr As Integer, ColPrice As Integer, ColAmount As Integer
Dim ColHSeqNo As Integer, ColHDOSeqNo As Integer
Dim ColHWHBef As Integer
Dim ColHProdSeqNo As Integer, ColHProdResultSeqNo As Integer
Dim ColHUpload_Cls As Integer, ColChildSupDate As Integer
Dim ColHRowHeader As Integer, ColHUpdate As Integer
Dim KolItemCOde As Byte, Koldesc As Byte, KolLot As Byte, KolRetQty As Byte, KolWHCode As Byte, KolRetCls As Byte
Dim KolUnit As Byte, KolCurr As Byte, KolPrice As Byte, kolService As Byte, kolAmount As Byte, kolReff As Byte
Dim KolDNNo As Byte, KolRemark As Byte, kolRetDate As Byte, KolRetSeq_NO As Byte

Dim tempRow As Double, rowHeader As Double
Dim newCls As New clsMRP

Private Sub cboSize_Change()
End Sub

Private Sub cbocurr_Change()
CboProduct_Click
End Sub

Private Sub cbocurr_Click()
CboProduct_Click
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub cboCust_Click()
'cbocust_Change
End Sub

'-------------- Initial -------------------
Private Sub Form_Load()
nilKosong = True
    'CtrlMenu1.FormName = Me.Name
    
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

    HakU = hakUpdate(Me.Name)

    Load_Combo
    Call headerGrid
    dtReturn.Value = Now
'nilKosong = False

End Sub
Sub AdToCur()
Call up_FillCombo(cbocurr, "Curr_Cls")
    cbocurr.TextColumn = 2
End Sub
Sub Load_Combo()
    Call isiCbo(cboCust, "Trade_Master", "Trade_Code", "Trade_Name", 75, 200, "Trade_Code,Trade_Name", , , "(Trade_Cls = 2  or Trade_Cls = 4)", , 0)
    Call isiCbo(cboWhCode, "Warehouse_Master", "WH_Code", "WH_Name ", 60, 150, "WH_Code", , , "StockControl_Cls=1")
    Call isiCbo(CboProduct, "Item_Master", "Item_Code", "Item_name", 120, 160, "Item_Code", , 1)
    AdToCur
    
    
    
    CboType.AddItem
    CboType.columnCount = 1
    CboType.List(0, 0) = "DN Return"
    CboType.List(0, 1) = "D1"
    CboType.AddItem
    CboType.List(1, 0) = "Sales Return"
    CboType.List(1, 1) = "D2"
    
    cboWhCode.Enabled = True
    
    DtEnd = Now()
    DtStart = Now()
End Sub

Sub SetCol()
    ColS = 0
    kolRetDate = 1
    KolItemCOde = 2: Koldesc = 3: KolLot = 4: KolRetQty = 5: KolWHCode = 6
    KolRetCls = 7: KolUnit = 8: KolCurr = 9: KolPrice = 10: kolService = 11: kolReff = 12
    KolDNNo = 13: KolRemark = 14:
End Sub




'Sub IsiCurrency()
'    With cbocurr
'        .ColumnCount = 2
'        .clear
'
'        For i = 0 To 4
'        .AddItem ""
'        .List(i, 0) = Format(i + 1, "0#")
'        .List(i, 1) = Split(isiCurr, ",")(i)
'        Next i
'        .ListWidth = 60
'        .ColumnWidths = "20 pt;40 pt"
'
'    End With
'End Sub

'-----------------------------------------------------------

'--------------------------View Dt -------------------------
Private Sub headerGrid()
    ColS = 0
    kolRetDate = 1
    KolItemCOde = 2: Koldesc = 3: KolLot = 4: KolRetQty = 5: KolWHCode = 6
    KolRetCls = 7: KolUnit = 8: KolCurr = 9: KolPrice = 10: kolService = 11: kolAmount = 12: kolReff = 13
    KolRemark = 14: KolRetSeq_NO = 15

With grid
    .clear
    .ColS = 16: .Rows = 1
    .TextMatrix(0, ColS) = ""
    .TextMatrix(0, kolRetDate) = "Return Date"
    .TextMatrix(0, KolItemCOde) = "Product Code"
    .TextMatrix(0, Koldesc) = "Description"
    .TextMatrix(0, KolLot) = "Lot No "
    .TextMatrix(0, KolRetQty) = "Return Qty"
    .TextMatrix(0, KolWHCode) = "WH Code"
    .TextMatrix(0, KolRetCls) = "Return Cls "
    .TextMatrix(0, KolUnit) = "Unit"
    .TextMatrix(0, KolCurr) = "Curr"
    .TextMatrix(0, KolPrice) = "Price"
    .TextMatrix(0, kolService) = "Service"
    .TextMatrix(0, kolAmount) = "Amount"
    .TextMatrix(0, kolReff) = "Reff "
    '.TextMatrix(0, KolDNNo) = "DN No."
    .TextMatrix(0, KolRemark) = "Remarks"
        
    .ColWidth(ColS) = 300
    .ColWidth(kolRetDate) = 1250
    .ColWidth(KolItemCOde) = 2000
    .ColWidth(Koldesc) = 2500
    .ColWidth(KolLot) = 900
    .ColWidth(KolRetQty) = 1000
    .ColWidth(KolWHCode) = 1500
    .ColWidth(KolRetCls) = 1500
    .ColWidth(KolUnit) = 500
    .ColWidth(KolCurr) = 500
    .ColWidth(KolPrice) = 1250
    .ColWidth(kolService) = 1250
    .ColWidth(kolAmount) = 1250
    .ColWidth(kolReff) = 1350
    '.ColWidth(KolDNNo) = 1450
    .ColWidth(KolRemark) = 1450
    .ColHidden(KolRetSeq_NO) = True
'
'    .ColWidth(ColS) = flexAlignLeftCenter
'    .ColWidth(kolRetDate) = flexAlignLeftCenter
'    .ColWidth(KolItemCOde) = flexAlignLeftCenter
'    .ColWidth(KolDesc) = flexAlignLeftCenter
'    .ColWidth(KolLot) = flexAlignLeftCenter
'    .ColWidth(KolRetQty) = flexAlignLeftCenter
'    .ColWidth(KolWhCode) = flexAlignLeftCenter
'    .ColWidth(KolRetCls) = flexAlignLeftCenter
'    .ColWidth(KolUnit) = flexAlignLeftCenter
'    .ColWidth(kolCurr) = flexAlignLeftCenter
'    .ColWidth(KolPrice) = flexAlignLeftCenter
'    .ColWidth(KolService) = flexAlignLeftCenter
'    .ColWidth(kolAmount) = flexAlignLeftCenter
'    .ColWidth(kolReff) = flexAlignLeftCenter
'    .ColWidth(KolDNNo) = flexAlignLeftCenter
'    .ColWidth(KolRemark) = flexAlignLeftCenter
'     Call ClsProc.AlignHeader(Grid)
    .RowHeightMax = 250
    .EditMaxLength = 1
'    .Rows = .Rows + 2
End With
End Sub

Sub IsiGrid(Optional stFilter As Byte)
Dim sqlQ
    Dim rsGrid As New ADODB.Recordset
    Dim rsGridDet As New ADODB.Recordset
    dtReturn.Enabled = True
    Call KosongBawah
    With grid
        Call headerGrid
        sqlQ = "select ISNULL(Return_Date,'')as return_date ,Item_Code," & vbCrLf
        sqlQ = sqlQ & " (SELECT Item_Name FROM Item_Master WHERE Item_Code=A.Item_Code) as Descr" & vbCrLf
        sqlQ = sqlQ & ",ISNULL(Lot_No,'') as Lot_NO,isnull(Return_Qty,0)as return_qty" & vbCrLf
        sqlQ = sqlQ & " ,isnull((select description from Unit_Cls where unit_cls=a.unit_cls),'')as Unit" & vbCrLf
        sqlQ = sqlQ & ",isnull((select WH_Name From warehouse_master WHERE WH_COde=A.WH_COde),'')as wh_codes,a.wh_code,isnull(Return_Cls,'')as Return_Cls" & vbCrLf
        sqlQ = sqlQ & ",ISNULL((select Description FROM curr_Cls WHERE Curr_CLS=a.Curr_code),'') As Curr," & vbCrLf
        sqlQ = sqlQ & "isnull(Price,0)as price,isnull(Service,0)as service,ISNULL(Amount,0)Amount, ISNULL(Reference,1)Reference," & vbCrLf
        sqlQ = sqlQ & " ISNULL(DO_NO,'')DO_NO,isnull(Remarks,'')Remarks,ReturnSeq_No"
        sqlQ = sqlQ & " FROM delivery_Return A  WHERE Cust_Code='" & cboCust & "'"
        sqlQ = sqlQ & " AND Seq_no=0 AND Return_Date Between '" & Format(DtStart, "yyyy-mm-dd") & "' AND '" & Format(DtEnd, "yyyy-mm-dd") & "'"
        
        Set rsGrid = Db.Execute(sqlQ)
        If rsGrid.EOF Or rsGrid.BOF Then
            Set rsGrid = Nothing
            LblErrMsg.Caption = DisplayMsg(4006)
            grid.clear
            headerGrid
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        .ColComboList(KolRetCls) = "#D1;DN Return|#D2;Sales Return"
        Do While Not rsGrid.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, ColS) = ""
            .Cell(flexcpBackColor, .Rows - 1, ColS) = vbWhite
            .TextMatrix(.Rows - 1, kolRetDate) = Format(Trim(rsGrid!Return_Date), "yyyy MMM dd")
            .TextMatrix(.Rows - 1, KolItemCOde) = Trim(rsGrid!Item_Code)
            .TextMatrix(.Rows - 1, Koldesc) = Trim(rsGrid!Descr)
            .TextMatrix(.Rows - 1, KolLot) = Trim(rsGrid!Lot_no)
            .TextMatrix(.Rows - 1, KolRetQty) = Trim(rsGrid!Return_Qty)
            .TextMatrix(.Rows - 1, KolWHCode) = Trim(rsGrid!wh_code)
            .TextMatrix(.Rows - 1, KolRetCls) = Trim(rsGrid!Return_Cls)
            .TextMatrix(.Rows - 1, KolUnit) = Trim(rsGrid!unit)
            .TextMatrix(.Rows - 1, KolCurr) = Trim(rsGrid!Curr)
            .TextMatrix(.Rows - 1, KolPrice) = Trim(rsGrid!Price)
            .TextMatrix(.Rows - 1, kolService) = Trim(rsGrid!service)
            
            
            If .TextMatrix(.Rows - 1, KolCurr) = "IDR" Then
                .TextMatrix(.Rows - 1, KolCurr) = Format(Trim(rsGrid!Curr), gs_formatPriceIDR)
                .TextMatrix(.Rows - 1, KolPrice) = Format(Trim(rsGrid!Price), gs_formatPriceIDR)
                .TextMatrix(.Rows - 1, kolService) = Format(Trim(rsGrid!service), gs_formatPriceIDR)
            Else
                .TextMatrix(.Rows - 1, KolCurr) = Format(Trim(rsGrid!Curr), gs_formatPrice)
                .TextMatrix(.Rows - 1, KolPrice) = Format(Trim(rsGrid!Price), gs_formatPrice)
                .TextMatrix(.Rows - 1, kolService) = Format(Trim(rsGrid!service), gs_formatPrice)
            End If
            
            .TextMatrix(.Rows - 1, kolAmount) = Trim(rsGrid!Amount)
            .TextMatrix(.Rows - 1, kolReff) = Trim(rsGrid!Reference)
            '.TextMatrix(.Rows - 1, KolDNNo) = Trim(rsGrid!do_no)
            .TextMatrix(.Rows - 1, KolRemark) = Trim(rsGrid!Remarks)
            .TextMatrix(.Rows - 1, KolRetSeq_NO) = Trim(rsGrid!ReturnSeq_no)
            rowHeader = .Rows - 1
            rsGrid.MoveNext
            .Cell(flexcpBackColor, .Rows - 1, kolRetDate, .Rows - 1, .ColS - 1) = &HE0E0E0
        Loop
        Set rsGrid = Nothing
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid
If .TextMatrix(Row, ColS) = "d" Then .TextMatrix(Row, ColS) = "D"
    If Col = ColS Then RemoveES (Row)
    'If .Cell(flexcpBackColor, Row, 0) <> vbWhite Then Cancel = True
    
End With
End Sub

Private Sub grid_Click()
    nilKosong = True
    With grid
        LblErrMsg = ""
        
        If .Row > 0 Then
            If .Cell(flexcpBackColor, .Row, .Col) = vbWhite Then .FocusRect = flexFocusInset Else .FocusRect = flexFocusNone
            If .Cell(flexcpBackColor, .Row, ColChildSupDate) <> &HE0E0E0 Then
                If .Col = ColChildSupDate Then '
                    If Trim(.TextMatrix(.Row, kolRetDate)) = "" Then
                        'dtReturn2.Value = Format(.TextMatrix(.FindRow(rowHeader, , ColHRowHeader), ColChildSupDate), "dd MMM yyyy")
                    Else
                        
                        'dtReturn2.Value = Format(.TextMatrix(.Row, kolRetDate), "dd MMM yyyy")
                    End If
                 '   dtReturn2.Visible = True
                  '  dtReturn2.Left = .Cell(flexcpLeft, .Row, ColChildSupDate)
                   ' dtReturn2.top = .Cell(flexcpTop, .Row, ColChildSupDate)
                    'dtReturn2.Width = .CellWidth + 30
                    'dtReturn2.SetFocus
                    'tempRow = .Row
                End If
            End If
        End If
    End With
    nilKosong = False
End Sub

Private Sub Grid_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)

With grid
If .TextMatrix(Row, ColS) = "d" Then .TextMatrix(Row, ColS) = "D"
If KeyCode = 13 Then


    If UCase(.TextMatrix(Row, ColS)) = "S" Then
                CboProduct.Text = .TextMatrix(Row, KolItemCOde)
                txtDesc = .TextMatrix(Row, Koldesc)
                cboWhCode = .TextMatrix(Row, KolWHCode)
                lblUnit(13).Caption = .TextMatrix(Row, KolUnit)
                txtReturnSeq_No = .TextMatrix(Row, KolRetSeq_NO)
                txtRemark = .TextMatrix(Row, KolRemark)
                txtQty = .TextMatrix(Row, KolRetQty)
                txtprice = .TextMatrix(Row, KolPrice)
                TxtService = .TextMatrix(Row, kolService)
                txtamount = .TextMatrix(Row, kolAmount)
                cbocurr.Text = .TextMatrix(Row, KolCurr)  '  Get_Field("select Curr_Cls FROM CUrr_Cls WHERE Description='" & .TextMatrix(Row, kolCurr) & "'", 0)
                cboWhCode.Text = .TextMatrix(Row, KolWHCode)  'Get_Field("select wh_code From warehouse_master WHERE WH_Name='" & .TextMatrix(Row, KolWhCode) & "'", 0)
                CboType.Text = .TextMatrix(Row, KolRetCls)
                txtRef = .TextMatrix(Row, kolReff)
                'TxtDoNo = .TextMatrix(Row, KolDNNo)
                txtReturnSeq_No = .TextMatrix(Row, KolRetSeq_NO)
                TxtLotNo = .TextMatrix(Row, KolLot)
                dtReturn = .TextMatrix(Row, kolRetDate)
            End If

End If
End With
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
If grid.ColSel = 0 Then
    If KeyAscii = 100 Then KeyAscii = 90 'Grid.TextMatrix(Grid.RowSel, .ColSel) = ""
    If KeyAscii = 8 Then grid.TextMatrix(grid.RowSel, 0) = ""
End If
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 0 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii <> Asc("S") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
            KeyAscii = 0
        If KeyAscii = Asc(".") Then KeyAscii = 0
   End If
    If Col = ColS Then
        If (KeyAscii = 83 Or KeyAscii = 115 Or KeyAscii = 100 Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 68) Then
        If KeyAscii = 100 Then grid.TextMatrix(Row, ColS) = "D"
        If KeyAscii = 13 Then 'Jika teken enter..
            If txtReturnSeq_No <> grid.TextMatrix(Row, KolRetSeq_NO) Then
                Call Grid_KeyDownEdit(Row, Col, KeyAscii, 0)
            End If
        End If
        Exit Sub
        Else
        KeyAscii = 0
        End If
    End If
    
     
    
End Sub
Sub RemoveES(Baris)
With grid
For i = 1 To .Rows - 1
If i <> Baris Then
    .TextMatrix(i, 0) = ""
End If
Next
End With
End Sub
Public Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim rstime As Recordset
    Dim setRow As Integer, ClosedStock As String
    If nilKosong Then Exit Sub
    
    With grid
    
            If Col = ColS Then
            Call RemoveES(Row)
            If .TextMatrix(Row, Col) = "" Then Exit Sub

                
            If UCase(.TextMatrix(Row, ColS)) = "S" Then
                MousePointer = vbHourglass
                CboProduct.Text = .TextMatrix(Row, KolItemCOde)
                txtDesc = .TextMatrix(Row, Koldesc)
                cboWhCode = .TextMatrix(Row, KolWHCode)
                lblUnit(13).Caption = .TextMatrix(Row, KolUnit)
                txtReturnSeq_No = .TextMatrix(Row, KolRetSeq_NO)
                txtRemark = .TextMatrix(Row, KolRemark)
                txtQty = .TextMatrix(Row, KolRetQty)
                
                txtprice = .TextMatrix(Row, KolPrice)
                TxtService = .TextMatrix(Row, kolService)
                txtamount = .TextMatrix(Row, kolAmount)
                cbocurr.Text = .TextMatrix(Row, KolCurr)  '  Get_Field("select Curr_Cls FROM CUrr_Cls WHERE Description='" & .TextMatrix(Row, kolCurr) & "'", 0)
                cboWhCode = .TextMatrix(Row, KolWHCode)  ' Get_Field("select wh_code From warehouse_master WHERE WH_Name='" & .TextMatrix(Row, KolWhCode) & "'", 0)
                CboType.Text = .TextMatrix(Row, KolRetCls)
                txtRef = .TextMatrix(Row, kolReff)
                'eTxtDoNo = .TextMatrix(Row, KolDNNo)
                txtReturnSeq_No = .TextMatrix(Row, KolRetSeq_NO)
                TxtLotNo = .TextMatrix(Row, KolLot)
                dtReturn = .TextMatrix(Row, kolRetDate)
                                txtprice = .TextMatrix(Row, KolPrice)
                TxtService = .TextMatrix(Row, kolService)
                MousePointer = vbDefault

                
            
            End If
        
    End If
    End With
End Sub

Private Sub kosongColGrid(Row As Integer, Optional Kolom As String, Optional Kolom2 As String)
    With grid
        .Col = 0
        If Kolom <> "" Then
            If Kolom2 <> "" Then
                For i = 1 To .Rows - 1
                    If .Text = Kolom Or .Text = Kolom2 Then .Text = ""
                    If .TextMatrix(i, 0) <> "C" Then .TextMatrix(i, 0) = ""
                Next i
            Else
                For i = 1 To .Rows - 1
                   If .Text = Kolom Then .Text = ""
                   If .TextMatrix(i, 0) <> "D" Then .TextMatrix(i, 0) = ""
                Next i
            End If
           KosongBawah
        Else
           For i = 1 To .Rows - 1
              If .TextMatrix(i, 0) <> "" Then .TextMatrix(i, 0) = ""
           Next i
           .TextMatrix(Row, 0) = "S"
        End If
    End With
End Sub

'-----------------------------------------------------------

'------------------------- Process -------------------------
Function chkSave(Optional chkDetail As Byte) As Boolean
Dim rsCheck As New ADODB.Recordset

chkSave = False
    
    If HakU = 0 Then LblErrMsg = DisplayMsg(3008): Exit Function 'You don't have an access for Update
        
    If Trim(cboCust) = "" Then
        LblErrMsg = DisplayMsg(1033) 'Please Input Customer Code
        cboCust.SetFocus: Exit Function
    ElseIf cboCust.MatchFound = False Then
        LblErrMsg = DisplayMsg(4072) 'Customer code not found!
        cboCust.SetFocus: Exit Function
    ElseIf Format(DtStart, "yyyy-MM-dd") > Format(DtEnd, "yyyy-MM-dd") Then
        LblErrMsg = DisplayMsg(4025) 'Start Date must be lower than
    End If
    
chkSave = True
End Function

Private Sub cmdSearch_Click()
    Call Kosong
    CboType.Enabled = True
    If chkSave Then
        MousePointer = vbHourglass
        Call IsiGrid
        MousePointer = vbDefault
    End If
End Sub

Private Sub cmdprocess_Click(Index As Integer)
Dim tanya

Me.MousePointer = vbHourglass
Select Case Index
    Case 0: 'Save
        If chkSave(1) Then Call SaveData ' ProcessCmd

    Case 1:  'Cancel
        Call Kosong: Call IsiGrid
        
    Case 2:  'Clear
        Call Kosong(1): Call headerGrid: cboCust.SetFocus
End Select
Me.MousePointer = vbDefault
End Sub

Function ConvDate(Tgl)
ConvDate = IIf(IsDate(Tgl), Format(Tgl, "yyyy-MM-dd"), "1900-01-31")

End Function

Function CekTanggal()
CekTanggal = False
Dim pesandtAwal As String, pesandtAkhir As String


         
                pesandtAwal = up_ValidateDateRange(ConvDate(dtReturn), True)
                pesandtAkhir = up_ValidateDateRange(ConvDate(dtReturn), True)
                If pesandtAwal <> "" Or pesandtAkhir <> "" Then
                    CekTanggal = True
                End If

    

End Function
'----
'Fungsi menghapus data yang terdapat huruf D nya pada kolom 0
'-----
Function DeleteData()
Dim sDel As String
Dim LblInput As String

With grid
For i = 1 To .Rows - 1
    If UCase(.TextMatrix(i, 0)) = "D" Then
        LblInput = MsgBox("Do you really to delete  ?", _
         vbYesNo + vbQuestion, "Confirmation")
        If LblInput = vbYes Then
                sDel = "DELETE FROM  Delivery_Return WHERE ReturnSeq_NO=" & .TextMatrix(i, KolRetSeq_NO)
                
                Db.Execute (sDel)
                DeleteData = True
                Dim sCodewh As String
                
                'ambil code wh
                sCodewh = Get_Field("SELECT wh_code FROM Warehouse_Master WHERE WH_Name='" & .TextMatrix(i, KolWHCode) & "'", 0)
                
                'Update Master Stok
                Dim tampungBln As String
                Dim blnFix As Integer, thnFix As Integer
                tampungBln = newCls.blnAkhir()
                blnFix = Split(tampungBln, ",")(0)
                thnFix = Split(tampungBln, ",")(1)
                Call newCls.updateStock(sCodewh, .TextMatrix(i, KolItemCOde), .TextMatrix(i, KolRetQty), .TextMatrix(i, KolLot), Format(.TextMatrix(i, kolRetDate), "yyyy-mm-dd"), blnFix, thnFix, Db, "Supply", 0, 1)
                
                '-----Delete Part Supplier
                Db.Execute "DELETE FROM Part_Supply WHERE RecSeq_No=" & .TextMatrix(i, KolRetSeq_NO)
                
                
        End If
    End If
Next

End With
End Function

'---------------
'-Fungsi untuk  mengecek kelengkapan inputan data yang di inputkan
'----------------
Function TidakLengkap()


TidakLengkap = False
If cboCust = "" Then
    LblErrMsg = DisplayMsg(1033)
    TidakLengkap = True
    cboCust.SetFocus
Exit Function
End If

If CboProduct.Text = "" Then
    LblErrMsg = DisplayMsg(1009)
    TidakLengkap = True
    CboProduct.SetFocus
    Exit Function
ElseIf CboProduct.MatchFound = False Then
    LblErrMsg = DisplayMsg(4003)
    TidakLengkap = True
    CboProduct.SetFocus
    Exit Function
End If

If cboWhCode = "" Then
    LblErrMsg = DisplayMsg(1042)
    TidakLengkap = True
    cboWhCode.SetFocus
    Exit Function
End If

If txtQty = "" Then
    LblErrMsg = DisplayMsg(1012)
    TidakLengkap = True
    txtQty.SetFocus
    Exit Function
End If
If TxtLotNo.Text = "" Then
    LblErrMsg = DisplayMsg(1044)
    TidakLengkap = True
    TxtLotNo.SetFocus
    Exit Function
End If
If CboType.Text = "" Then
    LblErrMsg = DisplayMsg(8116)
    TidakLengkap = True
    CboType.SetFocus
    Exit Function
ElseIf CboType.MatchFound = False Then
    LblErrMsg = DisplayMsg(8070)
    TidakLengkap = True
    CboType.SetFocus
    Exit Function
End If

If cbocurr.Text = "" Then
    LblErrMsg = DisplayMsg(1011)
    TidakLengkap = True
    cbocurr.SetFocus
    Exit Function
ElseIf cbocurr.MatchFound = False Then
    LblErrMsg = DisplayMsg(4005)
    TidakLengkap = True
    cbocurr.SetFocus
    Exit Function
End If

If txtprice = "" Then
    LblErrMsg = DisplayMsg(1029)
    TidakLengkap = True
    txtprice.SetFocus
    Exit Function
End If
If CDbl(txtprice) > gd_MaxPrice Then
    LblErrMsg = DisplayMsg(4048) & " " & gd_MaxPrice
    TidakLengkap = True
    txtprice.SetFocus
    Exit Function
End If

If TxtService = "" Then
    LblErrMsg = DisplayMsg(8117)
    TidakLengkap = True
    TxtService.SetFocus
    Exit Function
End If

If CDbl(TxtService) > gd_MaxPrice Then
    LblErrMsg = DisplayMsg(8118) & " " & gd_MaxPrice
    TidakLengkap = True
    TxtService.SetFocus
    Exit Function
End If



End Function
Sub SaveData()
Dim ada
Dim JmlLama As Integer

'Menghapus data yang ada tanda GridnYe
If DeleteData() Then 'berarti melakukan proses hapus...
    
    IsiGrid 'tampilkan perubahan setelah menghapus
    
    LblErrMsg = DisplayMsg(1201)
    Exit Sub
End If
''----akan berkahir disini jika menghapus data




'--------Melakukan Proses INPUT ATAU UPdate Date


If TidakLengkap Then 'Mengecek kelengkapan data
Exit Sub
End If

'If CekTanggal Then
 '   lblErrMsg = DisplayMsg(1022)
  '  Exit Sub
'End If

Dim Rdata As New Recordset
Dim sq As String
If Rdata.State <> adStateClosed Then Rdata.Close
sq = "SELECT * FROM delivery_return WHERE ReturnSeq_No=" & IIf(txtReturnSeq_No = "", 0, txtReturnSeq_No)
Rdata.Open sq, Db, 1, 3
Dim tampungBln As String
Dim blnFix As Integer, thnFix As Integer
tampungBln = newCls.blnAkhir()
                    blnFix = Split(tampungBln, ",")(0)
                    thnFix = Split(tampungBln, ",")(1)

If Rdata.EOF Then
    'tambah baru
    Rdata.AddNew
    Rdata!Register_Date = Now()
    Call newCls.updateStock(cboWhCode, CboProduct, -txtQty, TxtLotNo, Format(dtReturn, "yyyy-mm-dd"), blnFix, thnFix, Db, "Supply", 0, 1)
Else
    
    
    'Untuk Update
    'Apus stok dulu baru di isi yang baru
    
     tmpQty = Get_Field("SELECT return_Qty FROM delivery_return WHERE returnSeq_no=" & txtReturnSeq_No & " AND Item_Code='" & CboProduct & "' AND WH_COde='" & cboWhCode & "' AND Lot_No='" & TxtLotNo & "'", 0) 'query nya diambil dengan kondisi item code,wh dan lot harus sama ...
    If tmpQty <> 0 Then
        'hapus terlebih dahulu
        Call newCls.updateStock(cboWhCode, CboProduct, tmpQty, TxtLotNo, Format(dtReturn, "yyyy-mm-dd"), blnFix, thnFix, Db, "Supply", 0, 1)
    End If
        
        
        'insert baru Stock
        Call newCls.updateStock(cboWhCode, CboProduct, txtQty, TxtLotNo, Format(dtReturn, "yyyy-mm-dd"), blnFix, thnFix, Db, "Supply", 0, 1)
End If

    Rdata!Return_Qty = txtQty
    Rdata!Cust_CodE = cboCust
    Rdata!do_no = ""
    Rdata!Seq_no = 0
    Rdata!DOSeq_No = 0
    Rdata!Item_Code = CboProduct
    Rdata!po_no = "PO"
    Rdata!Reference = txtRef
    Rdata!Price = IIf(txtprice = "", 0, txtprice)
    Rdata!service = TxtService
    Rdata!Lot_no = TxtLotNo
    Rdata!Unit_Cls = Get_Field("select Unit_cls FROM unit_cls WHERE Description='" & lblUnit(13).Caption & "'", 0)
    Rdata!curr_code = Get_Field("select Curr_Cls FROM Curr_cls WHERE Description='" & cbocurr.Text & "'", 0)
    Rdata!Amount = txtamount
    Rdata!Return_Cls = CboType.Text
    Rdata!wh_code = cboWhCode.Text
    Rdata!Last_Update = Now()
    Rdata!last_user = "SA"
    Rdata!Remarks = txtRemark
    Rdata!Return_Date = dtReturn
    Rdata.update
    Rdata.Close
    '-----------------------------------
    'Update Table PartSupply
     '-----------------------------------
          
           If IIf(txtReturnSeq_No = "", 0, txtReturnSeq_No) = 0 Then 'Mengecek dulu apakah data baru atau data lama dengan mengecek returnSeq_no
             'ambil maximal dari data yang baru di input,karena masih Baru di input
              txtReturnSeq_No = Get_Field("SELECT MAX(returnSeq_no) AS MaxRet FROM Delivery_Return", 0)
           
           End If
           
     
     'Mengupdate Part Supply
     UpdatePart_Supply
     
     
    
    '---------------
    IsiGrid ''Menampilkan data Hasil Perubahan
    '---------------
    
    
    
    LblErrMsg = DisplayMsg(8005)
    KosongBawah



End Sub


Sub UpdatePart_Supply()
Dim rupdate As New ADODB.Recordset
Dim sq As String
With grid
sq = "select * FROM Part_Supply WHERE RecSeq_No=" & txtReturnSeq_No
'Sq = Sq & " AND ChildITem_Code='" & cboProduct & "' AND FromWarehouse_Code='" & CboCust & "'"
'Sq = Sq & " AND towarehouse_code='" & cboWhCode & "'"

rupdate.Open sq, Db, 1, 3
If rupdate.EOF Then
    rupdate.AddNew
    rupdate!Register_Date = Now()
End If
    rupdate!RecSeq_no = txtReturnSeq_No
    rupdate!FromWarehouse_Code = cboCust
    rupdate!towarehouse_code = cboWhCode
    rupdate!childitem_code = CboProduct
    rupdate!supply_cls = CboType.Column(1)
    'rupdate!consumption_Qty = TxtQty
    rupdate!ChildRequirement_qty = txtQty
    rupdate!childunit_cls = Get_Field("select Unit_Cls FROM Unit_Cls WHERE Description='" & lblUnit(13) & "'", 0)
    rupdate!currency_code = Get_Field("select Curr_Cls FROM Curr_cls WHERE Description='" & cbocurr.Text & "'", 0)
    rupdate!Price = txtprice
    rupdate!service = TxtService
    rupdate!Amount = txtamount
    rupdate!Lot_no = TxtLotNo
    rupdate!do_no = ""
    rupdate!Remarks = txtRemark
    rupdate!SJNo = txtRef
    rupdate!childsupply_date = dtReturn
    rupdate!Last_Update = Now()
    rupdate!last_user = "SA"
    rupdate!from_address = "Addr"
    rupdate.update
    rupdate.Close


End With

End Sub


Private Sub cmdPopItem_Click()
    Me.MousePointer = vbHourglass
    frm_BrowseItem.getItemCode = CboProduct.Text
    frm_BrowseItem.Show 1
    CboProduct.Text = frm_BrowseItem.getItemCode
    Me.MousePointer = vbDefault
End Sub


'------------------------------------------------------


'------------------------- Validate -------------------
Private Sub CboCust_Change()


    Call headerGrid
    cboCust = cboCust
    If cboCust.MatchFound Then
        lblcust = cboCust.Column(1)
    Else
        lblcust = ""
    End If
    Call KosongBawah
    If nilKosong Then Exit Sub
End Sub

Private Sub dtStart_Change()
    LblErrMsg.Caption = ""
    If CDate(DtStart.Value) > CDate(DtEnd.Value) Then
       LblErrMsg.Caption = DisplayMsg(4025) & " " & Format(DtEnd, "dd MMM yyyy") '"Start Date must be lower than "
       Exit Sub
    End If
    Call CboCust_Change
End Sub

Private Sub dtEnd_Change()
    LblErrMsg.Caption = ""
    If CDate(DtEnd) < CDate(DtStart) Then
       LblErrMsg.Caption = DisplayMsg(4024) & " " & Format(DtStart, "dd MMM yyyy") '"End Date must be higher than "
       Exit Sub
    End If
    Call CboCust_Change
End Sub

Private Sub dtReturn2_Change()
If nilKosong Then Exit Sub

With grid
    If tempRow > 0 Then
        .TextMatrix(tempRow, ColChildSupDate) = Format(dtReturn2, "dd MMM yyyy")
        .TextMatrix(tempRow, ColHUpdate) = "1"
    End If
End With
End Sub

Private Sub dtReturn2_LostFocus()
    Call dtReturn2_Change
    dtReturn2.Visible = False
End Sub


Private Sub CboProduct_Change()
    LblErrMsg = ""
    txtDesc.Text = ""
    lblUnit(13).Caption = ""
End Sub
 
 Function Get_Field(sql, Field)
Dim Rdata As New ADODB.Recordset
Set Rdata = Db.Execute(sql)
Get_Field = ""
If Not Rdata.EOF Then
 Get_Field = IIf(IsNull(Rdata.Fields(Field)), "", Rdata.Fields(Field))
End If
End Function
Function FormatDate(Data)
If Data <> "" Or Not IsNull(Data) Then
Data = Right(Data, 2) & "/" & Left(Right(Data, 4), 2) & "/" & Left(Data, 4)
FormatDate = Format(Data, "MMM-dd-yyyy")
Else
FormatDate = ""
End If
End Function

Private Sub CboProduct_Click()
Dim ss As String
    LblErrMsg = ""
    If CboProduct.ListIndex <> -1 Then
        txtDesc.Text = CboProduct.Column(1)
        
        TxtService.locked = False: txtprice.locked = False 'Price dan service di Lock, karena prubahan price dan service hanya bisa di price master '202305 Pak Toha Minta lock nya dibuka
        
        'query ambil data price and service berdasar tipe price dan customer code
        ss = " select isnull(Price,0),currency_code,(select Description FROM Curr_Cls WHERE Curr_Cls=Currency_Code)Des  from price_master where  item_code='" & CboProduct & "'"
        ss = ss & " AND (convert(char(8)," & Format(dtReturn, "yyyymmdd") & " , 112) between  start_date and end_date)" ' and price_cls='01'"
        If cboCust <> "" Then
        ss = ss & " AND (Trade_Code='" & cboCust & "' OR Trade_Code='000000')"
        End If
        If cbocurr.ListIndex <> -1 Then
        ss = ss & " AND Currency_Code='" & cbocurr & "'"
        End If
        If getCount(ss & " AND Price_CLS='01'") <> 0 Then
                If cbocurr.Text = "IDR" Then
                    txtprice = Format(Get_Field(ss & " AND Price_CLS='01'", 0), gs_formatPriceIDR)
                Else
                    txtprice = Format(Get_Field(ss & " AND Price_CLS='01'", 0), gs_formatPrice)
                End If
            Else
            TxtService = "": txtprice = "":
            
        End If
        If getCount(ss & " AND Price_CLS='05'") <> 0 Then
                If cbocurr.Text = "IDR" Then
                    TxtService = Format(Get_Field(ss & " AND Price_CLS='05'", 0), gs_formatPriceIDR)
                Else
                    TxtService = Format(Get_Field(ss & " AND Price_CLS='05'", 0), gs_formatPrice)
                End If
            
            Else
            TxtService = "": txtprice = "":
        End If
        If getCount(ss & " AND Price_CLS='02'") <> 0 Then
                If cbocurr.Text = "IDR" Then
                    txtprice = Format(Get_Field(ss & " AND Price_CLS='02'", 0), gs_formatPriceIDR)
                Else
                    txtprice = Format(Get_Field(ss & " AND Price_CLS='02'", 0), gs_formatPrice)
                End If
            Else
            TxtService = "": txtprice = "":
            
        End If
        'Call isiCbo(cbocurr, "Curr_CLS", "Description", "Curr_Cls", 50, 20, "Curr_Cls", , 1)
        If getCount(ss) <> 0 Then
        'AdToCur
        cbocurr = Get_Field(ss, 1)
        Else
        TxtService = "": txtprice = "":
        End If
            lblUnit(13).Caption = Get_Field("select a.description FROM UNIT_CLS a  INNER JOIN Item_Master b ON a.Unit_cls=b.Unit_Cls WHERE b.ITem_Code='" & CboProduct & "'", 0)
            'ambil Default dari WH Code pada Setting Item  tersebut
            cboWhCode = Get_Field("SELECT WH_Code FROM Item_master WHERE Item_Code='" & CboProduct & "'", 0)
            'cbocurr.locked = False
            ''harga di lock
            TxtService.locked = False
            txtprice.locked = False
            cbocurr.locked = True
        
    Else
    TxtService = "": txtprice = "": cbocurr.Text = ""
    ''harga di lock
    TxtService.locked = False
    txtprice.locked = False
    cbocurr.locked = True
    'cbocurr.locked = False
    End If
End Sub
Function getCount(sql)
Dim R As New ADODB.Recordset
Set R = Nothing
R.Open sql, Db, adOpenDynamic, adLockBatchOptimistic
If R.EOF Then
getCount = 0
Else
getCount = 1
End If
R.Close
End Function
Private Sub CboProduct_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call CboProduct_Click
End Sub

Private Sub CboProduct_LostFocus()
    Call CboProduct_Click
End Sub

Private Sub CboWHCode_Change()
    cboWhCode = cboWhCode
End Sub

Private Sub Grid_Validate(Cancel As Boolean)
If grid.TextMatrix(grid.RowSel, 0) = "d" Then grid.TextMatrix(grid.RowSel, 0) = "D"
End Sub

Private Sub Grid_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid.TextMatrix(Row, ColS) = "d" Then grid.TextMatrix(grid.RowSel, 0) = "D"
End Sub

Private Sub txtPrice_Change()
    Ammount
End Sub

Private Sub txtPrice_LostFocus()
    Dim z As Double
    If txtprice <> "" Then
        z = CDbl(txtprice.Text)
        If z > (1E+16) Then txtprice = Left(z, 16)
    End If
      
            If cbocurr = "IDR" Then
                txtprice.Text = Format(txtprice.Text, gs_formatPriceIDR)
                
            Else
                txtprice.Text = Format(txtprice.Text, gs_formatPrice)
            End If

    
End Sub
Private Sub txtqty_Change()
    Ammount
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> Asc(".") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    If (txtQty.Text & Chr(KeyAscii)) > 1E+16 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtQty_LostFocus()
    Dim z As Double
    
    If txtQty <> "" Then
        z = CDbl(txtQty.Text)
        If z > (1E+16) Then txtQty = Left(z, 16)
    End If
End Sub

Private Sub TxtSpool_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> Asc(".") Then KeyAscii = 0
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> Asc(".") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    If (txtprice.Text & Chr(KeyAscii)) > 1E+16 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

'--------------------------------------------


'-------------- Out -----------------------
Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then Unload Me Else LblErrMsg.Caption = ErrMsg
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub
'----------------------------------------------------------------------

Sub Kosong(Optional stAwal As Byte)
nilKosong = True
    If stAwal = 1 Then
        cboCust = "": lblcust = ""
        DtStart = Format(DateValue(Year(Now) & "/" & Month(Now) & "/01"), "dd MMM YYYY")
        DtEnd = Format(Now, "dd MMM YYYY")
         
        Call KosongBawah
    End If
    TxtService = ""
    'TxtDoNo = ""
    txtRemark = ""
    txtRef = ""
    cboWhCode.Enabled = True
    cboWhCode = ""
    LblErrMsg = ""
    CboType = ""
    txtDesc = ""
    CboProduct.locked = False
    cmdPopItem.Enabled = True
nilKosong = False
End Sub

Sub KosongBawah()

   
    dtReturn = Format(Now, "dd MMM yyyy")
    CboProduct = "": txtDesc = ""
    TxtLotNo = ""
    txtQty = ""
    lblUnit(13).Caption = ""
    cbocurr = ""
    txtprice = ""
    txtamount = ""
    txtRemark = ""
    txtRef = ""
    'TxtDoNo = ""
    CboType = ""
    TxtService = ""
    txtReturnSeq_No = 0
    cboWhCode = ""
End Sub

Sub Ammount()


        
If Not IsNumeric(Trim(txtprice)) Or Trim(txtprice) = "" Or Not IsNumeric(Trim(txtQty.Text)) Or txtQty.Text = "" Or Not IsNumeric(Trim(TxtService)) Or Trim(TxtService) = "" Then
        txtamount.Text = 0
    Else
        txtamount.Text = CDbl(Trim(txtQty.Text)) * (CDbl(Trim(txtprice.Text) + CDbl(Trim(TxtService.Text))))
    End If
    If cbocurr.Text = "IDR" Then
            txtamount = Format(txtamount, gs_formatPriceIDR)
    Else
            txtamount = Format(txtamount, gs_formatPrice)
    End If
    
    
End Sub

Private Sub txtService_Change()
Ammount

End Sub

Private Sub txtService_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> Asc(".") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    If (txtprice.Text & Chr(KeyAscii)) > 1E+16 And KeyAscii <> vbKeyBack Then KeyAscii = 0

End Sub

Private Sub TxtService_LostFocus()
 Dim z As Double
    If TxtService <> "" Then
        z = CDbl(TxtService.Text)
        If z > (1E+16) Then TxtService = Left(z, 16)
    End If
    If cbocurr.Text = "IDR" Then
    TxtService.Text = Format(TxtService.Text, gs_formatPriceIDR)
    Else
    TxtService.Text = Format(TxtService.Text, gs_formatPrice)
    
    End If
End Sub
