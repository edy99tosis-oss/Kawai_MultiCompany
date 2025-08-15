VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmSupplyBySerialNo 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Supply By Serial No"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15195
   Icon            =   "FrmSupplyBySerialNo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   15195
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQty 
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
      Height          =   330
      Left            =   10920
      TabIndex        =   10
      Top             =   8640
      Width           =   1005
   End
   Begin EZRunnerv3.Anchor Anchor1 
      Left            =   2760
      Top             =   9840
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmd_PrintSuratJalan 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print Surat Jalan"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "FTTF*/"
      Top             =   9840
      Width           =   1605
   End
   Begin VB.CommandButton cmd_PrintDocBC 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print Doc BC"
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
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "FTTF*/"
      Top             =   9840
      Width           =   1605
   End
   Begin VB.TextBox txtDescription 
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
      MaxLength       =   100
      TabIndex        =   9
      Tag             =   "TTFF*/"
      Top             =   8640
      Width           =   2835
   End
   Begin VB.TextBox txtItemCode 
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
      Left            =   5640
      MaxLength       =   100
      TabIndex        =   8
      Tag             =   "TTFF*/"
      Top             =   8640
      Width           =   2115
   End
   Begin VB.TextBox txtSerialNo 
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
      Left            =   3360
      MaxLength       =   100
      TabIndex        =   7
      Tag             =   "TTFF*/"
      Top             =   8640
      Width           =   2115
   End
   Begin VB.TextBox txtBarcode 
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
      Left            =   360
      MaxLength       =   100
      TabIndex        =   6
      Tag             =   "TTFF*/"
      Top             =   8640
      Width           =   2835
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   2535
      Left            =   240
      TabIndex        =   22
      Tag             =   "TTTF*/"
      Top             =   720
      Width           =   14685
      Begin VB.OptionButton OptWithoutSerialNo 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Without Serial No"
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
         Left            =   1920
         TabIndex        =   43
         Top             =   215
         Width           =   1935
      End
      Begin VB.OptionButton OptSerialNo 
         BackColor       =   &H00FDDFE3&
         Caption         =   "With Serial No"
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
         Left            =   240
         TabIndex        =   42
         Top             =   200
         Width           =   1575
      End
      Begin VB.TextBox txtBCNo 
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
         Left            =   12360
         MaxLength       =   100
         TabIndex        =   39
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Search"
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
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   1920
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   5520
         TabIndex        =   34
         Tag             =   "TTFF*/"
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox lblFromWH 
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
         Height          =   240
         Index           =   0
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   24
         Tag             =   "TTFF*/"
         Top             =   650
         Width           =   4695
      End
      Begin VB.TextBox lblToWH 
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
         Height          =   240
         Index           =   2
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   23
         Tag             =   "TTFF*/"
         Top             =   1100
         Width           =   4695
      End
      Begin MSComCtl2.DTPicker DtpFrom 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Tag             =   "TTFF*/"
         Top             =   1560
         Width           =   1725
         _ExtentX        =   3043
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
      Begin MSComCtl2.DTPicker DtpTo 
         Height          =   315
         Left            =   4080
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   1560
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
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtpBCDate 
         Height          =   315
         Left            =   12360
         TabIndex        =   40
         Tag             =   "TTFF*/"
         Top             =   1560
         Width           =   1725
         _ExtentX        =   3043
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
      Begin VB.Label Label19 
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
         Index           =   7
         Left            =   11280
         TabIndex        =   41
         Tag             =   "TTFF*/"
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BC No"
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
         Left            =   11280
         TabIndex        =   38
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   540
      End
      Begin MSForms.ComboBox cboBCType 
         Height          =   315
         Left            =   12360
         TabIndex        =   37
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   1725
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "3043;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label19 
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
         Index           =   5
         Left            =   11280
         TabIndex        =   36
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   735
      End
      Begin MSForms.ComboBox cboDNo 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   2040
         Width           =   1725
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "3043;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboToWH 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   1725
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "3043;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboFromWH 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   1725
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "3043;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Transaction"
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
         Left            =   3960
         TabIndex        =   32
         Tag             =   "TTFF*/"
         Top             =   2040
         Width           =   1410
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN No"
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
         Left            =   240
         TabIndex        =   29
         Tag             =   "TTFF*/"
         Top             =   2040
         Width           =   540
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3720
         X2              =   8400
         Y1              =   1400
         Y2              =   1400
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To WareHouse"
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
         Left            =   240
         TabIndex        =   28
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From WareHouse"
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
         TabIndex        =   27
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3720
         TabIndex        =   26
         Tag             =   "TTTF*/"
         Top             =   1680
         Width           =   210
      End
      Begin VB.Label Label19 
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
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Tag             =   "TTFF*/"
         Top             =   1560
         Width           =   540
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3720
         X2              =   8415
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   240
      TabIndex        =   15
      Tag             =   "TTTF*/"
      Top             =   9120
      Width           =   14835
      Begin VB.Label lbl_pesan 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_pesan"
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
         Left            =   240
         TabIndex        =   17
         Tag             =   "TTTF*/"
         Top             =   240
         Width           =   14235
      End
   End
   Begin VB.CommandButton cmd_sub_menu 
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "TTFF*/"
      Top             =   9840
      Width           =   1125
   End
   Begin VB.CommandButton cmd_clear 
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "FTTF*/"
      Top             =   9840
      Width           =   1125
   End
   Begin VB.CommandButton cmdDelete 
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
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "FTTF*/"
      Top             =   9840
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13080
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "FTTF*/"
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4725
      Left            =   240
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "TTTF*/"
      Top             =   3360
      Width           =   14745
      _cx             =   26009
      _cy             =   8334
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
   Begin MSComDlg.CommonDialog cdg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblQty 
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
      Index           =   4
      Left            =   10920
      TabIndex        =   44
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   300
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   375
      Left            =   3840
      TabIndex        =   35
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   661
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Supply By Serial No"
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
      Left            =   120
      TabIndex        =   30
      Tag             =   "TTTF*/"
      Top             =   240
      Width           =   14565
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Description"
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
      Left            =   7920
      TabIndex        =   21
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   1425
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      TabIndex        =   20
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No"
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
      Left            =   3360
      TabIndex        =   19
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   780
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   1
      Left            =   240
      Tag             =   "TTTF*/"
      Top             =   8520
      Width           =   14775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Barcode No"
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
      Left            =   360
      TabIndex        =   18
      Tag             =   "TTFF*/"
      Top             =   8280
      Width           =   1470
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   240
      Tag             =   "TTTF*/"
      Top             =   8160
      Width           =   14775
   End
End
Attribute VB_Name = "FrmSupplyBySerialNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bteColSelect As Byte
Dim bteColSupplydate As Byte
Dim bteColBarcodeNo As Byte
Dim bteColSerialNo As Byte
Dim bteColItemCode As Byte
Dim bteColDescription As Byte
Dim bteColQty As Byte
Dim btecolDNNo As Byte
Dim tempIP As String
Dim Qty As Double
Dim validate As Boolean
Dim ls_PathExcel As String
Dim FSave As Boolean

Private Sub cboDNo_Change()
    lbl_pesan.Caption = ""
    up_FillBC
End Sub

Private Sub cmd_PrintDocBC_Click()
    up_DocBC
End Sub

Private Sub cmd_PrintSuratJalan_Click()
    up_PrintSuratJalan
End Sub

Private Sub DtpFrom_Change()
    Call up_FillComboDNNo
End Sub

Private Sub DtpTo_Change()
    Call up_FillComboDNNo
End Sub

Private Sub Form_Load()

    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

    Call up_Header
    Call up_FillComboWH
    Call up_FillComboDNNo
    Call up_Clear
    Call up_FillComboBC
    
With Anchor1
      .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
      .DoInit
End With
    
End Sub

Private Sub up_Header()

    bteColSelect = 0
    bteColSupplydate = 1
    bteColBarcodeNo = 2
    bteColSerialNo = 3
    bteColItemCode = 4
    bteColDescription = 5
    bteColQty = 6
    btecolDNNo = 7
    
    With grid
        .ColS = 8
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColSupplydate) = "Supply Date"
        .TextMatrix(0, bteColBarcodeNo) = "Barcode No"
        .TextMatrix(0, bteColSerialNo) = "Serial No"
        .TextMatrix(0, bteColItemCode) = "Item Code"
        .TextMatrix(0, bteColDescription) = "Description"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, btecolDNNo) = "DNNo"
    
         .ColWidth(bteColSelect) = 300
         .ColWidth(bteColSupplydate) = 1750
         .ColWidth(bteColBarcodeNo) = 3250
         .ColWidth(bteColSerialNo) = 3000
         .ColWidth(bteColItemCode) = 2000
         .ColWidth(bteColDescription) = 3000
         .ColWidth(bteColQty) = 1000
    
         .ColAlignment(bteColSelect) = flexAlignCenterCenter
         .ColAlignment(bteColSupplydate) = flexAlignCenterCenter
         .ColAlignment(bteColBarcodeNo) = flexAlignCenterCenter
         .ColAlignment(bteColSerialNo) = flexAlignCenterCenter
         .ColAlignment(bteColItemCode) = flexAlignCenterCenter
         .ColAlignment(bteColDescription) = flexAlignCenterCenter
         .ColAlignment(bteColQty) = flexAlignCenterCenter
              
        .ColHidden(btecolDNNo) = True
        
    End With
End Sub

Private Sub up_Clear()
        
    cboFromWH.Text = ""
    cboToWH.Text = ""
    cboDNo.Text = ""
    txtBarcode.Text = ""
    txtSerialNo.Text = ""
    txtItemCode.Text = ""
    txtDescription.Text = ""
    txtBCNo.Text = ""
    cboBCType.Text = ""
    lbl_pesan = ""
    
    lblFromWH(0) = ""
    lblToWH(2) = ""
    
    DTPFrom.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPTo.Value = Now()
    dtpBCDate.Value = Now()
    
    If OptWithoutSerialNo.Value = True Then
        lblQty(4).Visible = True
        txtQty.Visible = True
    Else
        lblQty(4).Visible = False
        txtQty.Visible = False
    End If
    
    up_Header
    
    validate = False
    
    FSave = False
    
End Sub

Private Sub up_FillComboWH()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_WH_Sel"
    
    Set RS = cmd.Execute

    With cboFromWH
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;180pt"
        .ListWidth = 230
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS("WH_Code") & "")
            .List(i, 1) = Trim(RS("WH_Name") & "")
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = 0
    End With
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_WH_Sel"
    
    Set RS = cmd.Execute

    With cboToWH
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;180pt"
        .ListWidth = 230
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS("WH_Code") & "")
            .List(i, 1) = Trim(RS("WH_Name") & "")
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = 0
    End With
    
End Sub

Private Sub up_FillComboDNNo()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_DNNo_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("FromwH", adVarChar, adParamInput, 15, RTrim(cboFromWH.Text))
    cmd.Parameters.append cmd.CreateParameter("ToWH", adVarChar, adParamInput, 15, RTrim(cboToWH))
    cmd.Parameters.append cmd.CreateParameter("DateFrom", adDate, adParamInput, , Format(DTPFrom.Value, "YYYY-MM-DD"))
    cmd.Parameters.append cmd.CreateParameter("DateTo", adDate, adParamInput, , Format(DTPTo.Value, "YYYY-MM-DD"))
    cmd.Parameters.append cmd.CreateParameter("Type", adVarChar, adParamInput, 1, "1")
    If OptSerialNo.Value = True Then
        cmd.Parameters.append cmd.CreateParameter("SerialNo", adVarChar, adParamInput, 1, "1")
    ElseIf OptWithoutSerialNo.Value = True Then
        cmd.Parameters.append cmd.CreateParameter("SerialNo", adVarChar, adParamInput, 1, "0")
    Else
        lbl_pesan = DisplayMsg(9017) & " Type Serial No "
        Exit Sub
    End If

    Set RS = cmd.Execute
    
    If RS.EOF = False Then

        With cboDNo
            .clear
            .columnCount = 1
            .ColumnWidths = "100pt"
            .ListWidth = 100
            .ListRows = 10
        
            i = 0
            
            Do While Not RS.EOF
                .AddItem
                .List(i, 0) = Trim(RS("SJ_No") & "")
                
                RS.MoveNext
                i = i + 1
            Loop
            
          .ListIndex = -1
        End With
    Else
        cboDNo.clear
    End If
End Sub

Private Sub up_FillBC()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_GetBC_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("SJNo", adVarChar, adParamInput, 25, RTrim(cboDNo.Text))
    
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
        cboBCType.Text = Trim(RS("BC_Type") & "")
        txtBCNo.Text = Trim(RS("BC40_No") & "")
        dtpBCDate.Value = IIf(IsNull(Trim(RS("BC40_Date"))) = True, Now, Trim(RS("BC40_Date")))
        'IIf(IsNull(Trim(RS("BC40_Date"))) = True, Now, Trim(RS("BC40_Date")))
    End If
End Sub

Private Sub up_GetSerialNo()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

MousePointer = vbHourglass
       
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_GetSerialNo"
    
    cmd.Parameters.append cmd.CreateParameter("BarcodeNo", adVarChar, adParamInput, 100, RTrim(txtBarcode.Text))
         
    Set RS = cmd.Execute
    
    If Trim(RS("Item_Code")) <> "NULL" Then
        If RS.EOF = False Then
            txtSerialNo.Text = Trim(RS("SerialNo"))
            txtItemCode.Text = IIf(IsNull(Trim(RS("Item_Code"))) = True, "", Trim(RS("Item_Code")))
            txtDescription.Text = IIf(IsNull(Trim(RS("Description"))) = True, "", Trim(RS("Description")))
        End If
        
        save_Load
        
        FSave = True
        
    Else
        lbl_pesan.Caption = "Invalid Barcode No !"
        
        txtBarcode.Text = ""
        
        wmp.URL = (App.path & "\Incorrect.mp3")

    End If
   
MousePointer = vbDefault
End Sub

Private Sub up_GetItemNo()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

MousePointer = vbHourglass
       
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_GetItemDesc"
    
    cmd.Parameters.append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 25, RTrim(txtItemCode.Text))
         
    Set RS = cmd.Execute
    
    'If Trim(RS("Item_Code")) <> "NULL" Then
        If RS.EOF = False Then
            txtDescription.Text = IIf(IsNull(Trim(RS("Item_Name"))) = True, "", Trim(RS("Item_Name")))
            
            If OptSerialNo.Value = True Then
                save_Load
            End If
        End If
        FSave = True
   ' End If
   
MousePointer = vbDefault
End Sub
Private Sub up_FillComboBC()
Dim ls_sql As String
Dim rs_combo As New ADODB.Recordset
Dim i As Long


cboBCType.columnCount = 1
cboBCType.clear

ls_sql = "select BC_Type from BC_master"
rs_combo.Open ls_sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
i = 0

Do While Not rs_combo.EOF
cboBCType.AddItem rs_combo("BC_Type")
rs_combo.MoveNext
Loop

cboBCType.ColumnWidths = "85"
cboBCType.ListWidth = 85
cboBCType.ListRows = 10

End Sub

Private Sub cek_SerialNo()
Dim sql As String
Dim RS As New Recordset

    sql = "select * from Supply_Scan_Detail WHERE Barcode_No='" & txtBarcode.Text & "'  AND Serial_No= '" & txtSerialNo.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lbl_pesan.Caption = DisplayMsg(71)
        
        wmp.URL = (App.path & "\Incorrect.mp3")
        Exit Sub
    End If
End Sub

Private Sub up_Save()
Dim strSQL As String
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

Qty = 1 'Set Qty
    If cboDNo.Text = "" Then
        Set cmd = New ADODB.Command
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 0
            cmd.ActiveConnection = Db
            cmd.CommandText = "sp_SupplyScanHeader_Ins"
            
            cmd.Parameters.append cmd.CreateParameter("FromWH", adVarChar, adParamInput, 15, RTrim(cboFromWH.Text))
            cmd.Parameters.append cmd.CreateParameter("ToWH", adVarChar, adParamInput, 15, RTrim(cboToWH.Text))
            cmd.Parameters.append cmd.CreateParameter("IPAddress", adVarChar, adParamInput, 50, tempIP)
            cmd.Parameters.append cmd.CreateParameter("BCType", adVarChar, adParamInput, 15, cboBCType.Text)
            cmd.Parameters.append cmd.CreateParameter("BC40No", adVarChar, adParamInput, 30, txtBCNo.Text)
            cmd.Parameters.append cmd.CreateParameter("BC40Date", adDBTime, adParamInput, , Format(dtpBCDate.Value, "dd-mmm-yy"))
            cmd.Parameters.append cmd.CreateParameter("RegisterUser", adVarChar, adParamInput, 50, userLogin)
                    
            Set RS = cmd.Execute
    End If
    
    If OptWithoutSerialNo.Value = True Then
        Qty = CDec(txtQty.Text)
    Else
        Qty = 1
    End If
    
    Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_SupplyScanDetail_Ins"
        
        cmd.Parameters.append cmd.CreateParameter("BarcodeNo", adVarChar, adParamInput, 100, RTrim(txtBarcode.Text))
        cmd.Parameters.append cmd.CreateParameter("SerialNo", adVarChar, adParamInput, 100, RTrim(txtSerialNo.Text))
        cmd.Parameters.append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 50, txtItemCode)
        Set prm = cmd.CreateParameter("Qty", adDecimal, adParamInput)
        prm.Precision = 38
        prm.NumericScale = 2
        prm.Value = Qty
        cmd.Parameters.append prm
        cmd.Parameters.append cmd.CreateParameter("DNNo", adVarChar, adParamInput, 50, RTrim(cboDNo.Text))
        cmd.Parameters.append cmd.CreateParameter("RegisterUser", adVarChar, adParamInput, 50, userLogin)
                
        Set RS = cmd.Execute
        
End Sub

Private Sub up_Delete()
Dim strSQL As String
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

Qty = 1

    Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_SupplyScan_Delete"
        
        cmd.Parameters.append cmd.CreateParameter("DNNo", adVarChar, adParamInput, 100, RTrim(cboDNo.Text))
        cmd.Parameters.append cmd.CreateParameter("BarcodeNo", adVarChar, adParamInput, 100, RTrim(txtBarcode.Text))
        cmd.Parameters.append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 25, RTrim(txtItemCode.Text))
                
        Set RS = cmd.Execute
        
End Sub

Private Sub up_GridLoad()
Dim sql As String
Dim cmd As ADODB.Command
Dim li_Row As Integer
     
up_Header
        
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_SupplyScan_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("FromwH", adVarChar, adParamInput, 15, RTrim(cboFromWH.Text))
    cmd.Parameters.append cmd.CreateParameter("ToWH", adVarChar, adParamInput, 15, RTrim(cboToWH))
    cmd.Parameters.append cmd.CreateParameter("DateFrom", adDBTime, adParamInput, , DTPFrom.Value)
    cmd.Parameters.append cmd.CreateParameter("DateTo", adDBTime, adParamInput, , DTPTo.Value)
    cmd.Parameters.append cmd.CreateParameter("DNNo", adVarChar, adParamInput, 25, RTrim(cboDNo.Text))
    cmd.Parameters.append cmd.CreateParameter("Type", adVarChar, adParamInput, 1, "1")
    
    Set RS = cmd.Execute
        
        i = 1
        With grid
            If RS.EOF = False Then
                While Not RS.EOF
                    .Rows = .Rows + 1
                    
                    .Cell(flexcpChecked, i, ColCheck) = flexUnchecked
                    .Cell(flexcpBackColor, i, ColCheck) = vbWhite
                    .TextMatrix(i, bteColSelect) = ""
                    .TextMatrix(i, bteColSupplydate) = Trim(RS("SupplyDate"))
                    .TextMatrix(i, bteColBarcodeNo) = IIf(IsNull(Trim(RS("Barcode_No"))) = True, "", Trim(RS("Barcode_No")))
                    .TextMatrix(i, bteColSerialNo) = IIf(IsNull(Trim(RS("Serial_No"))) = True, "", Trim(RS("Serial_No")))
                    .TextMatrix(i, bteColItemCode) = IIf(IsNull(Trim(RS("Item_Code"))) = True, "", Trim(RS("Item_Code")))
                    .TextMatrix(i, bteColDescription) = IIf(IsNull(Trim(RS("Description"))) = True, "", Trim(RS("Description")))
                    .TextMatrix(i, bteColQty) = IIf(IsNull(Trim(RS("Qty"))) = True, 0, Trim(RS("Qty")))
                    .TextMatrix(i, btecolDNNo) = Trim(RS("DNNo"))
                    
                    cboBCType.Text = IIf(IsNull(Trim(RS("BC_Type"))) = True, "", Trim(RS("BC_Type")))
                    txtBCNo.Text = IIf(IsNull(Trim(RS("BC40_No"))) = True, "", Trim(RS("BC40_No")))
                    dtpBCDate.Value = IIf(IsNull(Trim(RS("BC40_Date"))) = True, Now, Trim(RS("BC40_Date")))
                    
                    .Cell(flexcpAlignment, i, bteColSupplydate, i, bteColDescription) = flexAlignLeftCenter
                    i = i + 1
                RS.MoveNext
                Wend
            Else
'                lbl_pesan.Caption = DisplayMsg(8012)
            End If
        End With
End Sub

Private Sub get_IP()
    Dim WMI     As Object
    Dim qryWMI  As Object
    Dim Item    As Variant

    Set WMI = GetObject("winmgmts:\\.\root\cimv2")

    Set qryWMI = WMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration " & _
                               "WHERE IPEnabled = True")

    For Each Item In qryWMI
      getIP = Item.IPAddress(0)
    Next
    
    tempIP = getIP
    Set WMI = Nothing
    Set qryWMI = Nothing
End Sub

Private Sub up_Validate(Status As String)
    If Status = "Input" Then
        If cboFromWH.Text = "" Then
            lbl_pesan = DisplayMsg(9017) & " Warehouse From !": cboFromWH.SetFocus:
            validate = True
        ElseIf cboToWH.Text = "" Then
            lbl_pesan = DisplayMsg(9017) & " Warehouse To !": cboToWH.SetFocus:
            validate = True
'        ElseIf cboBCType.Text = "" Then
'             lbl_pesan = DisplayMsg(9017) & " BCType !": cboBCType.SetFocus:
'             validate = True
'        ElseIf txtBCNo.Text = "" Then
'             lbl_pesan = DisplayMsg("0001") & " BCNo !": txtBCNo.SetFocus:
'             validate = True
        End If
'        If OptWithoutSerialNo.Value = True Then
'            If cboBCType.Text = "" Then
'                 lbl_pesan = DisplayMsg(9017) & " BCType !": cboBCType.SetFocus:
'                 validate = True
'            ElseIf txtBCNo.Text = "" Then
'                 lbl_pesan = DisplayMsg("0001") & " BCNo !": txtBCNo.SetFocus:
'                 validate = True
'            End If
'        End If
    Else
         If cboFromWH.Text = "" Then
            lbl_pesan = DisplayMsg(9017) & " Warehouse From !": cboFromWH.SetFocus:
            validate = True
        ElseIf cboToWH.Text = "" Then
            lbl_pesan = DisplayMsg(9017) & " Warehouse To !": cboToWH.SetFocus:
            validate = True
        End If
    End If

End Sub

Sub up_PrintSuratJalan()
Dim xlapp As New Excel.application
Dim rsCek As New Recordset, Row As Long, tempi As String, tempcust As String
Dim bolcust As Boolean, bolinv As Boolean
Dim rsCompany As New Recordset, sql_plus As String, sqlP As String
Dim sqlControl As String, RsInvControl As New ADODB.Recordset
Dim selisih As Double
Dim nomor As Integer
Dim sql_sum As String
Dim cmd As ADODB.Command
Dim RS As ADODB.Recordset
Dim li_Row As Integer

Dim ls_no As String
Dim ls_nama_part As String
Dim ls_kode_part As String
Dim ls_qty_pengiriman As String
Dim ls_satuan As String
Dim ls_keterangan As String
Dim ls_c As String
Dim ls_d As String
    
ls_no = "A"
ls_nama_part = "B"
ls_kode_part = "E"
ls_qty_pengiriman = "F"
ls_satuan = "G"
ls_keterangan = "H"
ls_pallet = "I"
ls_totalqty = "J"
ls_c = "C"
ls_d = "D"
    
Me.MousePointer = vbHourglass

Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_SupplyScan_Report"
    
    cmd.Parameters.append cmd.CreateParameter("FromwH", adVarChar, adParamInput, 15, RTrim(cboFromWH.Text))
    cmd.Parameters.append cmd.CreateParameter("ToWH", adVarChar, adParamInput, 15, RTrim(cboToWH))
    cmd.Parameters.append cmd.CreateParameter("DateFrom", adDBTime, adParamInput, , DTPFrom.Value)
    cmd.Parameters.append cmd.CreateParameter("DateTo", adDBTime, adParamInput, , DTPTo.Value)
    cmd.Parameters.append cmd.CreateParameter("DNNo", adVarChar, adParamInput, 25, RTrim(cboDNo.Text))
    cmd.Parameters.append cmd.CreateParameter("Type", adVarChar, adParamInput, 1, "1")
    
    Set RS = cmd.Execute

If Not RS.EOF Then
Screen.MousePointer = vbHourglass
With xlapp

    sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax  From company_profile "
    If rsCompany.State <> adStateClosed Then rsCompany.Close
    rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
    If rsCompany.EOF Then Screen.MousePointer = vbDefault: Exit Sub

    .Workbooks.Add
    .Range("a4") = rsCompany!company_name '"Judul Company"
    .Range("A5") = rsCompany!address1 & " " & rsCompany!address2 & " " & rsCompany!City & " " & rsCompany!Province & " " & rsCompany!postal_code
    .Range("A6") = "Phone (0264)351323-6. Fax:(0264)351327"
    .Range("A8") = "No"
    .Range("A9") = "Cust PO No"
    .Range("A10") = "BC Type"
    .Range("A11") = "BC Number"
    .Range("A12") = "QTY"
    .Range("F8") = "Delivery To"
    .Range("G8") = ": " '& CboDelivery.Text
    .Range("F9") = "Kawasan Industri Kota Bukit Indah"
    .Range("F10") = "Blok A-III, No. 23, Dangdeur, Kab Purwakarta, Jawa Barat"
    .Range("F12") = "Model"
    .Range("A14") = "Kami Kirimkan barang-barang tersebut dibawah ini dengan kendaraan: .......................................... No: ........................"
    .Range("H1") = "Cikampek," & " " & Format(Now, "dd mmm YYYY")
    .Range("I4") = "Surat Jalan"
    .Range("C8") = ": " & Trim(RS!DNNo)
    .Range("C9") = ": " '& txtPoNo.Text
    .Range("C10") = ": " '& Trim(rsCek!BC_Type)
    .Range("C11") = ": " '& Trim(rsCek!BC40_No)
    
    .Range("I4:J4").Merge
    .Range("B16:D16").Merge
    .Range("A4").Font.Size = 14
    .Range("I4").Font.Size = 18
    .Range("I4").Font.Bold = True
    
    .ActiveSheet.Cells(1, 1).columnWidth = 5
    .ActiveSheet.Cells(1, 2).columnWidth = 8.15
    .ActiveSheet.Cells(1, 3).columnWidth = 20
    .ActiveSheet.Cells(1, 4).columnWidth = 18
    .ActiveSheet.Cells(1, 5).columnWidth = 12
    .ActiveSheet.Cells(1, 6).columnWidth = 15
    .ActiveSheet.Cells(1, 7).columnWidth = 6.71
    .ActiveSheet.Cells(1, 8).columnWidth = 30
    '.Columns("H:H")ColumnWidth = 30
    .Range(ls_no & 6, ls_totalqty & 6).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Row = 16

Dim jumlah As Double
    Do While Not RS.EOF
        If Row = 16 Then
            .Range(ls_no & Row) = "No"
            .Range(ls_nama_part & Row) = "Nama Part"
            .Range(ls_kode_part & Row) = "Kode Part"
            .Range(ls_qty_pengiriman & Row) = "QTY Pengiriman"
            .Range(ls_satuan & Row) = "Satuan"
            .Range(ls_keterangan & Row) = "Keterangan"
            .Range(ls_pallet & Row) = "Pallet No"
            .Range(ls_totalqty & Row) = "Total Qty"

            Row = Row + 1
        End If
        nomor = nomor + 1
        Row = Row
        
        jumlah = jumlah + RS!Qty
        .Range(ls_no & Row) = nomor
        .Range(ls_nama_part & Row) = Trim(RS!Name_Part)
        .Range(ls_kode_part & Row) = "'" + Trim(RS!Kode_Part)
        .Range(ls_qty_pengiriman & Row) = Format(RS!Qty)
        .Range(ls_satuan & Row) = (RS!Description)
        .Range(ls_keterangan & Row) = Format(RS!Remarks)
        
        Row = Row + 1
        .Range(ls_no & Row - 1).horizontalAlignment = xlCenter
        .Range(ls_kode_part & Row - 1).horizontalAlignment = xlLeft
        .Range(ls_qty_pengiriman & Row - 1).horizontalAlignment = xlCenter
        .Range(ls_satuan & Row - 1).horizontalAlignment = xlLeft
        .Range(ls_kode_part & Row - 1).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range(ls_nama_part & Row - 1, ls_d & Row - 1).Merge
        
        RS.MoveNext
    Loop
    
    .Range("C12") = ": " & jumlah
'    'Border
    .Range(ls_no & 16, ls_totalqty & Row - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(ls_no & 16, ls_totalqty & Row - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range(ls_no & 16, ls_totalqty & Row - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range(ls_no & 16, ls_totalqty & Row - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range(ls_no & 16, ls_totalqty & Row - 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Range(ls_no & 16, ls_totalqty & Row - 1).Borders(xlInsideVertical).LineStyle = xlContinuous
        
    .Range(ls_nama_part & 16, ls_c & 16 & Row - 1).Borders(xlEdgeRight).LineStyle = xlNone
    
    .Range(ls_no & Row + 2) = "* Please return this original latter to PT.Kawai Indonesia Plat-3"
    .Range(ls_no & Row + 2).Font.Italic = True
    
    .Range(ls_no & Row + 5) = "Delivered by,"
    .Range(ls_no & Row + 10, ls_nama_part & Row + 10).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(ls_no & Row + 10) = "WH.Member"
    
    .Range(ls_d & Row + 5) = "Approved by,"
    .Range(ls_d & Row + 10).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(ls_d & Row + 10) = "WH.Function Head"
    
    .Range(ls_qty_pengiriman & Row + 5) = "Checked by,"
    .Range(ls_qty_pengiriman & Row + 10).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(ls_qty_pengiriman & Row + 10) = "Security"
    
    .Range(ls_keterangan & Row + 5) = "Received by,"
    .Range(ls_keterangan & Row + 10).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(ls_keterangan & Row + 10) = "Customer"
    
    .Range("A1", "C1").Font.Bold = True
    .Range("A16:J16").Font.Bold = True
    .Range("a4").Font.Bold = True
    .Range("A1:H1").Columns.Font.Name = "Arial"
    .Range("A1:H1").Columns.Font.Size = "10"
    .Range("H1", "H4").horizontalAlignment = xlRight
    .Range("A16", "J16").horizontalAlignment = xlCenter
    .ActiveSheet.PageSetup.Orientation = xlLandscape
    .Range("A1:H1").Columns.AutoFit
    .Range("A1").Select
    .WindowState = xlMaximized
    .Visible = True
End With

Else
    lbl_pesan = DisplayMsg(4006)
End If

Screen.MousePointer = vbDefault
Me.MousePointer = vbDefault
End Sub

Private Sub up_DocBC()
Dim xlapp As New Excel.application
Dim rsCek As New Recordset, RowNo As Long, RowDesc As Integer, RowQty As Integer, RowRemark As Integer, tempi As String, tempcust As String
Dim bolcust As Boolean, bolinv As Boolean
Dim rsCompany As New Recordset, sql_plus As String, sqlP As String
Dim sqlControl As String, RsInvControl As New ADODB.Recordset
Dim selisih As Double
Dim nomor As Integer
Dim sql_sum As String
Dim cmd As ADODB.Command
Dim RS As ADODB.Recordset
Dim li_Row As Integer

Dim ls_a As String
Dim ls_b As String
Dim ls_c As String
Dim ls_d As String
Dim ls_e As String
Dim ls_f As String
Dim ls_g As String
Dim ls_h As String
Dim ls_i As String
Dim ls_j As String
    
ls_a = "A"
ls_b = "B"
ls_c = "C"
ls_d = "D"
ls_e = "E"
ls_f = "F"
ls_g = "G"
ls_h = "H"
ls_i = "I"
ls_j = "J"
    
Me.MousePointer = vbHourglass

Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_SupplyScan_Report"
    
    cmd.Parameters.append cmd.CreateParameter("FromwH", adVarChar, adParamInput, 15, RTrim(cboFromWH.Text))
    cmd.Parameters.append cmd.CreateParameter("ToWH", adVarChar, adParamInput, 15, RTrim(cboToWH))
    cmd.Parameters.append cmd.CreateParameter("DateFrom", adDBTime, adParamInput, , DTPFrom.Value)
    cmd.Parameters.append cmd.CreateParameter("DateTo", adDBTime, adParamInput, , DTPTo.Value)
    cmd.Parameters.append cmd.CreateParameter("DNNo", adVarChar, adParamInput, 25, RTrim(cboDNo.Text))
    cmd.Parameters.append cmd.CreateParameter("Type", adVarChar, adParamInput, 1, "2")
    
    Set RS = cmd.Execute

If Not RS.EOF Then
Screen.MousePointer = vbHourglass
With xlapp

    sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax, RTRIM(No_Izin)No_Izin From company_profile "
    If rsCompany.State <> adStateClosed Then rsCompany.Close
    rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
    If rsCompany.EOF Then Screen.MousePointer = vbDefault: Exit Sub

    .Workbooks.Add
    .Range("a2") = "PPB-KB"
    .Range("A3") = "F.2 - R.2"
    .Range("C2") = "PEMBERITAHUAN PEMINDAHAN BARANG"
    .Range("C3") = "DALAM SATU KAWASAN BERIKAT"
    .Range("A2 : J2").Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range("A3 : J3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    .Range("A2 : B2").Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range("A2 : B2").Borders(xlEdgeRight).LineStyle = xlContinuous
    
    .Range("A3 : B3").Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range("A3 : B3").Borders(xlEdgeRight).LineStyle = xlContinuous
    
    .Range("B2 : J2").Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range("B2 : J2").Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range("B3 : J3").Borders(xlEdgeRight).LineStyle = xlContinuous
    
    .Range("A4") = "Nomor"
    .Range("A5") = "Tanggal"
    .Range("C4") = ": " & txtBCNo.Text
    .Range("C5") = ": " & Format(Trim(RS!SJ_DATE), "dd mmm YYYY")
    .Range("A5 : J5").Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range("A4 : A5").Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range("J4 : J5").Borders(xlEdgeRight).LineStyle = xlContinuous
    
   
    .Range("A6") = "Identitas Pengusaha Kawasan Berikat / PDKB"
    .Range("A7") = "Nama Perusahaan"
    .Range("E7") = ": " & Trim(rsCompany!company_name)
    .Range("A8") = "Nomor Izin"
    .Range("E8") = ": " & Trim(rsCompany!No_Izin)
    .Range("A9") = "Lokasi"
    .Range("E9") = ": " & Trim(rsCompany!City)
    .Range("A9 : J9").Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range("A6 : A9").Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range("J6 : J9").Borders(xlEdgeRight).LineStyle = xlContinuous
    
    .Range("A10") = "A. Asal Lokasi Barang dan Tujuan Pemindahan Barang :"
    .Range("A11 : J11").Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range("A10 : A11").Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range("J10 : J11").Borders(xlEdgeRight).LineStyle = xlContinuous
    
    .Range("A12") = "Lokasi Asal Barang :"
    .Range("F12") = "Lokasi Tujuan Barang :"
    .Range("A12 : J12").Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range("A12").Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range("E12").Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range("J12").Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range("A13 : A15").Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range("E13 : E15").Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range("J13 : J15").Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range("A15 : J15").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    .Range("A13") = Trim(rsCompany!company_name)
    .Range("F13") = Trim(rsCompany!company_name)
    .Range("A14") = Trim(rsCompany!address1)
    .Range("F14") = "Kawasan Industri Kota Bukit Indah"
    .Range("A15") = Trim(rsCompany!address2)
    .Range("F15") = "Blok A-III. No 23, Dangdeur, Kab Purwakarta, Jawabarat"
    
    .Range("A16") = "B. Uraian Barang Yang Dipindahkan :"
    .Range("A16 : A17").Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range("J16 : J17").Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range("A17 : J17").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    .Range("A18") = "No"
    .Range("B18") = "'- Nama Barang"
    .Range("B19") = "'- Kode Barang"
    .Range("B20") = "'- Kode HS"
    .Range("G18") = "'- Jumlah"
    .Range("G19") = "'- Satuan"
    .Range("H18") = "'- Dokumen Pemasukan"
    .Range("H19") = "'- Nomor"
    .Range("H20") = "'- Tanggal"
    .Range("A18 : A20").Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range("A18 : A20").Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range("F18 : F20").Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range("G18 : G20").Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range("J18 : J20").Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range("A20 : J20").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    .Range("A2").Font.Size = 18
    .Range("A3").Font.Size = 18
    .Range("A2:J2").Font.Bold = True
    .Range("A3:J3").Font.Bold = True
    .Range("E7").Font.Bold = True
    .Range("A13").Font.Bold = True
    .Range("F13").Font.Bold = True
    
    .Range("A2:B2").Merge
    .Range("A3:B3").Merge
    .Range("A12:E12").Merge
    .Range("F12:J12").Merge
    
    .Range("C2:J2").Merge
    .Range("C3:J3").Merge
    .Range("C2").Font.Size = 14
    .Range("C3").Font.Size = 14
    .Range("C2:J2").horizontalAlignment = xlCenter
    .Range("C3:J3").horizontalAlignment = xlCenter
    .Range("A2:B2").horizontalAlignment = xlCenter
    .Range("A3:B3").horizontalAlignment = xlCenter
    .Range("C2:J2").verticalAlignment = xlCenter
    .Range("C3:J3").verticalAlignment = xlCenter
    .Range("A2:B2").verticalAlignment = xlCenter
    .Range("A3:B3").verticalAlignment = xlCenter
    
    .Range("A12:E12").verticalAlignment = xlCenter
    .Range("A12:E12").horizontalAlignment = xlCenter
    
    .Range("F12:J12").verticalAlignment = xlCenter
    .Range("F12:J12").horizontalAlignment = xlCenter
    
    .ActiveSheet.Cells(1, 1).columnWidth = 3.5
    .ActiveSheet.Cells(1, 2).columnWidth = 14
    .ActiveSheet.Cells(1, 10).columnWidth = 12

    RowNo = 21
    RowDesc = 22

Dim jumlah As Double
    Do While Not RS.EOF
        nomor = nomor + 1

        .Range(ls_a & RowNo) = nomor
        .Range(ls_a & RowNo & ":" & ls_a & RowNo + 3).Merge
        .Range(ls_a & RowNo & ":" & ls_a & RowNo + 3).horizontalAlignment = xlCenter
        .Range(ls_a & RowNo & ":" & ls_a & RowNo + 3).verticalAlignment = xlCenter
        
        .Range(ls_a & RowNo & ":" & ls_a & RowNo + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range(ls_a & RowNo & ":" & ls_a & RowNo + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range(ls_a & RowNo & ":" & ls_a & RowNo + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range(ls_f & RowNo & ":" & ls_f & RowNo + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range(ls_g & RowNo & ":" & ls_g & RowNo + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range(ls_g & RowNo & ":" & ls_g & RowNo + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range(ls_j & RowNo & ":" & ls_j & RowNo + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        
        .Range(ls_h & RowNo) = "NO DO : " & Trim(RS!DNNo)
        
        RowNo = RowNo + 1
        
        .Range(ls_b & RowDesc) = "Nama Barang :"
        .Range(ls_c & RowDesc) = Trim(RS!Name_Part)
        .Range(ls_g & RowNo - 1) = Format(RS!Qty)
        .Range(ls_g & RowNo - 1 & ":" & ls_g & RowNo).Merge
        .Range(ls_g & RowNo + 1) = Trim(RS!Description)
        .Range(ls_g & RowNo - 1 & ":" & ls_g & RowNo + 2).horizontalAlignment = xlCenter
        .Range(ls_g & RowNo - 1 & ":" & ls_g & RowNo + 2).verticalAlignment = xlCenter
        
        .Range(ls_h & RowNo) = "Tanggal : " & Format(RS!SJ_DATE, "DD-MMM-YYYY")
        
        
        RowNo = RowNo + 1
        RowDesc = RowDesc + 1
        
        .Range(ls_b & RowDesc) = "Kode Barang :"
        .Range(ls_c & RowDesc) = "'" + Trim(RS!Kode_Part)
        
'        .Range(ls_h & RowNo) = "NO.PENDAFTARAN :"
        
        RowNo = RowNo + 1
        RowDesc = RowDesc + 1
        
        .Range(ls_b & RowDesc) = "HS Code : "
        .Range(ls_c & RowDesc) = "'" & Trim(RS!HS_Code)
'        .Range(ls_h & RowNo) = "Tanggal : "
        .Range(ls_b & RowDesc & ":" & ls_j & RowDesc).Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        
        RowNo = RowNo + 1
        RowDesc = RowDesc + 2

        RS.MoveNext
    Loop
    
'    RowNo = RowNo + 1
    .Range(ls_a & RowNo) = "Lembar Persetujuan Pejabat Bea dan Cukai"
    .Range(ls_g & RowNo) = "Penanggung Jawab"
    
    RowNo = RowNo + 1
    .Range(ls_a & RowNo) = "Nomor Agenda Persetujuan"
    .Range(ls_g & RowNo) = "Pengusaha KB / PDKB"
    
    RowNo = RowNo + 2
    .Range(ls_a & RowNo) = "Tanggal Persetujuan"
    .Range(ls_c & RowNo) = ": " & "" & Format(Now, "dd mmm YYYY")
    
    RowNo = RowNo + 5
    .Range(ls_a & RowNo) = "Nama  : "
    
    RowNo = RowNo + 1
    .Range(ls_a & RowNo) = "Jabatan : "
    
    RowNo = RowNo + 2
    .Range(ls_a & RowNo) = "Catatan :"
    
    RowNo = RowNo + 2
    .Range(ls_a & RowNo) = "Selesai dipindahkan pada tanggal :"
    
    RowNo = RowNo + 1
    .Range(ls_a & RowNo) = "Pukul :"
    
    RowNo = RowNo + 2
    .Range(ls_a & RowNo) = "Nama :"
    
    RowNo = 25
'    .Range(ls_a & RowNo & ":" & ls_a & RowNo + 17).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    .Range(ls_f & RowNo & ":" & ls_f & RowNo + 10).Borders(xlEdgeRight).LineStyle = xlContinuous
'    .Range(ls_j & RowNo & ":" & ls_j & RowNo + 17).Borders(xlEdgeRight).LineStyle = xlContinuous
'    .Range(ls_a & RowNo + 10 & ":" & ls_j & RowNo + 10).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    .Range(ls_a & RowNo + 17 & ":" & ls_j & RowNo + 17).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    .ActiveSheet.PageSetup.Orientation = xlPortrait
    .Range("A1:H1").Columns.AutoFit
    .Range("A1").Select
    .WindowState = xlMaximized
    .Visible = True
    
    lbl_pesan = DisplayMsg(9008)
    
End With

Else
    lbl_pesan = DisplayMsg(4006)
End If

Screen.MousePointer = vbDefault
Me.MousePointer = vbDefault
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        up_Clear
        cboDNo.Enabled = False
    Else
        cboDNo.Enabled = True
    End If
End Sub

Private Sub cmd_clear_Click()
    up_Clear
    Check1.Value = 0
End Sub

Private Sub cmd_sub_menu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub


Private Sub save_Load()
Dim sql As String
Dim RS As New Recordset
Dim RsGetSJ As New Recordset

up_Validate ("Input")

    If validate = True Then
        validate = False
        Exit Sub
    End If

If Check1.Value = 0 And cboDNo.Text = "" Then
    lbl_pesan = "Please Checked New Transaction"
    Check1.SetFocus
    Exit Sub
End If

MousePointer = vbHourglass
    
    If OptSerialNo.Value = True Then
        If cboDNo.Text <> "" Then
'            sql = " SELECT * FROM Supply_Scan_Detail WHERE Barcode_No = '" & txtBarcode.Text & "' AND Serial_No = '" & txtSerialNo.Text & "' " & vbCrLf & _
'                  " AND Item_Code= '" & txtItemCode.Text & "' --AND SJ_No = '" & cboDNo.Text & "' "

'            sql = " SELECT Barcode_No, * from Supply_Scan_Header SSH " & vbCrLf & _
'                  " LEFT JOIN Supply_Scan_Detail SSD ON SSD.SJ_No = SSH.SJ_No " & vbCrLf & _
'                  " WHERE From_WH = '" & Trim(cboFromWH.Text) & "' AND To_WH = '" & Trim(cboToWH.Text) & "' " & vbCrLf & _
'                  " AND Barcode_No = '" & Trim(txtBarcode.Text) & "' "
'
'            Set RS = Db.Execute(sql)

            sql = "EXEC sp_SupplyBySerialNo_Validate '1', '" & Trim(cboFromWH.Text) & "', '" & Trim(cboToWH.Text) & "', '" & Trim(txtBarcode.Text) & "'"
            
            If RS.State = adStateOpen Then RS.Close
            RS.CursorLocation = adUseClient
            RS.Open sql, Db, adOpenKeyset, adLockOptimistic
            
            
            If RS.EOF = False Then
'                    lbl_pesan.Caption = DisplayMsg("0071")
'                    MousePointer = vbDefault
'
'                    wmp.URL = (App.path & "\Incorrect.mp3")
'                Exit Sub
                If IIf(IsNull(Trim(RS("Receipt_Status"))) = True, "0", Trim(RS("Receipt_Status"))) <> 1 And IIf(IsNull(Trim(RS("Scan_Cls"))) = True, "0", Trim(RS("Scan_Cls"))) <> 1 Then
                    lbl_pesan.Caption = DisplayMsg("0071")
                    MousePointer = vbDefault
                   
                    wmp.URL = (App.path & "\Incorrect.mp3")
                    Exit Sub
                End If
                RS.MoveNext
                
            End If
        End If
        
        'Validasi Serial No yang akan di supply ulang
        If OptSerialNo.Value = True Then
            
'            sql = " SELECT Barcode_No, * from Supply_Scan_Header SSH " & vbCrLf & _
'                  " LEFT JOIN Supply_Scan_Detail SSD ON SSD.SJ_No = SSH.SJ_No " & vbCrLf & _
'                  " WHERE From_WH = '" & Trim(cboFromWH.Text) & "' AND To_WH = '" & Trim(cboToWH.Text) & "' " & vbCrLf & _
'                  " AND Barcode_No = '" & Trim(txtBarcode.Text) & "' "
'
'            Set RS = Db.Execute(sql)
            
            sql = "EXEC sp_SupplyBySerialNo_Validate '1', '" & Trim(cboFromWH.Text) & "', '" & Trim(cboToWH.Text) & "', '" & Trim(txtBarcode.Text) & "'"
            
            If RS.State = adStateOpen Then RS.Close
            RS.CursorLocation = adUseClient
            RS.Open sql, Db, adOpenKeyset, adLockOptimistic
            
            Do While Not RS.EOF
                If IIf(IsNull(Trim(RS("Receipt_Status"))) = True, "0", Trim(RS("Receipt_Status"))) <> 1 And IIf(IsNull(Trim(RS("Scan_Cls"))) = True, "0", Trim(RS("Scan_Cls"))) <> 1 Then
                 lbl_pesan.Caption = DisplayMsg("0071")
                 MousePointer = vbDefault
                    
                 wmp.URL = (App.path & "\Incorrect.mp3")
                 Exit Sub
                End If
                RS.MoveNext
            Loop
        End If
     
    End If
       
'        End If
    'End If
    
    get_IP
    up_Save
    up_FillComboDNNo
        
    sql = " SELECT TOP 1 SH.SJ_No FROM " & vbCrLf & _
            " Supply_Scan_Header SH " & vbCrLf & _
            " LEFT JOIN Supply_Scan_Detail SD ON SH.SJ_No = SD.SJ_No " & vbCrLf & _
            " WHERE From_WH = '" & cboFromWH.Text & "' AND To_WH = '" & cboToWH.Text & "'  " & vbCrLf & _
            " AND Barcode_No= '" & txtBarcode.Text & "' ORDER BY SH.Register_Date DESC "
    
    Set RsGetSJ = Db.Execute(sql)
    
    If RsGetSJ.EOF = False Then
        cboDNo.Text = Trim(RsGetSJ("SJ_No") & "")
    End If
    
    up_GridLoad
    
    If OptSerialNo.Value = True Then
        wmp.URL = (App.path & "\Correct.mp3")
    End If
        
    txtBarcode.Text = ""
    txtSerialNo.Text = ""
    txtItemCode.Text = ""
    txtDescription.Text = ""
    
    cboDNo.Enabled = True
    
    
    
    lbl_pesan.Caption = DisplayMsg(1000)
    
MousePointer = vbDefault
End Sub

Private Sub cmdDelete_Click()
Dim sql As String
Dim RsDel As Recordset
Dim RsValDel As Recordset
Dim RStatus As String
    
    If OptSerialNo.Value = True Then
        If txtBarcode.Text <> "" Then
            If IsEmpty(tanya) Then tanya = MsgBox("Do you really want to delete this data ?", vbQuestion & vbYesNo, "Confirmation")
                If tanya = vbYes Then
                    
                    sql = "SELECT Receipt_Status FROM Supply_Scan_Header WHERE SJ_No='" & cboDNo.Text & "'"
                    Set RsValDel = Db.Execute(sql)
        
                    If Trim(RsValDel("Receipt_Status") & "") = 1 Then
                        lbl_pesan.Caption = "Item Cannot Delete, Data Already Supply"
                        Exit Sub
                    End If
                    
                    up_Delete
                    up_GridLoad
                    
                    If i = 1 Then
                        up_FillComboDNNo
                        Check1.Value = 0
                    End If
                    
                    lbl_pesan.Caption = DisplayMsg(1201)
                    
                    txtBarcode.Text = ""
                    txtSerialNo.Text = ""
                    txtItemCode.Text = ""
                    txtDescription.Text = ""
                
                Else
                    lbl_pesan.Caption = "Delete Record Canceled"
                    Exit Sub
                End If
        Else
            lbl_pesan.Caption = "Please Select Data To Delete !"
        End If
    Else
        If IsEmpty(tanya) Then tanya = MsgBox("Do you really want to delete this data ?", vbQuestion & vbYesNo, "Confirmation")
                If tanya = vbYes Then
                    
                    sql = "SELECT Receipt_Status FROM Supply_Scan_Header WHERE SJ_No='" & cboDNo.Text & "'"
                    Set RsValDel = Db.Execute(sql)
        
                    If Trim(RsValDel("Receipt_Status") & "") = 1 Then
                        lbl_pesan.Caption = "Item Cannot Delete, Data Already Supply"
                        Exit Sub
                    End If
                    
                    up_Delete
                    up_GridLoad
                    
                    If i = 1 Then
                        up_FillComboDNNo
                        Check1.Value = 0
                    End If
                    
                    lbl_pesan.Caption = DisplayMsg(1201)
                    
                    txtBarcode.Text = ""
                    txtSerialNo.Text = ""
                    txtItemCode.Text = ""
                    txtDescription.Text = ""
                
                Else
                    lbl_pesan.Caption = "Delete Record Canceled"
                    Exit Sub
                End If
    End If

End Sub

Private Sub cmdSearch_Click()
lbl_pesan = ""
    up_Validate ("Search")
    
    If validate = True Then
        validate = False
        Exit Sub
    End If
    
    If cboDNo.Text = "" Then
        lbl_pesan = "Please Select DNNo"
        cboDNo.SetFocus
        Exit Sub
    End If
    
    Call up_GridLoad
End Sub

Private Sub cboFromWH_Change()
    lblFromWH(0) = ""
    up_FillComboDNNo
    lbl_pesan = ""
End Sub

Private Sub cboFromWH_Click()
cboFromWH = cboFromWH
    If cboFromWH.MatchFound Then
        lblFromWH(0) = cboFromWH.Column(1)
        lbl_pesan = ""
    Else
        lblFromWH(0) = ""
        lbl_pesan = DisplayMsg(4018)
    End If
End Sub

Private Sub cboToWH_Change()
    lblToWH(2) = ""
    lbl_pesan = ""
    up_FillComboDNNo
End Sub

Private Sub cboToWH_Click()
cboToWH = cboToWH
    If cboToWH.MatchFound Then
        lblToWH(2) = cboToWH.Column(1)
        lbl_pesan = ""
    Else
        lblToWH(2) = ""
        lbl_pesan = DisplayMsg(4018)
    End If
End Sub

Private Sub OptSerialNo_Click()
lbl_pesan.Caption = ""
up_Clear
txtSerialNo.Enabled = True
txtBarcode.Enabled = True
lblQty(4).Visible = False
txtQty.Visible = False
End Sub

Private Sub OptWithoutSerialNo_Click()
lbl_pesan.Caption = ""
up_Clear
txtSerialNo.Enabled = False
txtBarcode.Enabled = False
lblQty(4).Visible = True
txtQty.Visible = True
End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)

lbl_pesan.Caption = ""

'If up_validateNew = False Then Exit Sub

If OptSerialNo.Value = True Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If (KeyAscii = vbKeyReturn) Then
       
        up_GetSerialNo
        
        If FSave = True Then
            txtBarcode.Text = ""
            txtSerialNo.Text = ""
            txtItemCode.Text = ""
            txtDescription.Text = ""
            
            FSave = False
        End If
    End If
End If
    
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim cek As Integer

    If Col = bteColSelect Then
        If grid.Cell(flexcpChecked, Row, Col) = flexChecked Then
            cek = 1
        Else
            cek = 2
        End If
        
        For i = 1 To grid.Rows - 1
            grid.Cell(flexcpChecked, i, 0) = flexUnchecked
            txtBarcode.Text = ""
            txtSerialNo.Text = ""
            txtDescription.Text = ""
            txtItemCode.Text = ""
            'cboDNo.Text = ""
        Next i
        
        grid.Cell(flexcpChecked, Row, Col) = cek
    End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> bteColSelect Then
        Cancel = True
    Else
        For i = 1 To grid.Rows - 1
            If grid.Cell(flexcpChecked, i, Col) = flexChecked Then
                grid.Cell(flexcpChecked, i, 0) = flexChecked
                txtBarcode.Text = grid.TextMatrix(grid.RowSel, bteColBarcodeNo)
                txtSerialNo.Text = grid.TextMatrix(grid.RowSel, bteColSerialNo)
                txtItemCode.Text = grid.TextMatrix(grid.RowSel, bteColItemCode)
                txtDescription.Text = grid.TextMatrix(grid.RowSel, bteColDescription)
                txtQty.Text = grid.TextMatrix(grid.RowSel, bteColQty)
                cboDNo.Text = grid.TextMatrix(grid.RowSel, btecolDNNo)
                lbl_pesan.Caption = ""
            Else
                grid.Cell(flexcpChecked, i, 0) = flexUnchecked
                
            End If
        Next i
    End If
End Sub

Public Function up_validateNew() As Boolean
    If Check1.Value = 0 Then
        lbl_pesan = DisplayMsg(9017) & " New Transaction !": Check1.SetFocus:
        up_validateNew = False
        Exit Function
    End If
    
    up_validateNew = True
End Function

'Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
''LblErrMsg = ""
''If OptWithoutSerialNo.Value = True Then
''    If KeyAscii = 13 Then SendKeys vbTab
''End If
'End Sub

Private Sub txtItemCode_LostFocus()
lbl_pesan.Caption = ""

If OptWithoutSerialNo.Value = True And txtItemCode.Text <> "" Then
        up_GetItemNo
        
        up_Validate ("Input")
        
        If validate = True Then
            validate = False
            Exit Sub
        End If
    
        
        
        If OptSerialNo.Value = True Then
            If FSave = True Then
                txtBarcode.Text = ""
                txtSerialNo.Text = ""
                txtItemCode.Text = ""
                txtDescription.Text = ""
                
                FSave = False
            End If
        Else
            txtQty.SetFocus
        End If
        
End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        up_Validate ("Input")
        
        If validate = True Then
            validate = False
            Exit Sub
        End If
        
        save_Load
        txtQty.Text = ""
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
End Sub
