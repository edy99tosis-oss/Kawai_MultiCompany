VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_part_supplyAuto 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Parts (Material) Supply Request [Automatic]"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "frm_part_supplyAuto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtParent 
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
      Left            =   1200
      MaxLength       =   12
      TabIndex        =   45
      Top             =   10680
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Left            =   360
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   10680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cmd_delete 
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10170
      Width           =   1125
   End
   Begin VB.CommandButton cmd_submit 
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10170
      Width           =   1125
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
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   10170
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   10170
      Width           =   1125
   End
   Begin VB.CommandButton cmd_cancel 
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10170
      Width           =   1125
   End
   Begin VB.CommandButton cmd_preview 
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10170
      Width           =   1125
   End
   Begin VB.CommandButton cmd_daily 
      BackColor       =   &H0080FFFF&
      Caption         =   "To Production Schedule"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10170
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton cmd_navigate 
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10170
      Width           =   1140
   End
   Begin VB.CommandButton cmd_navigate 
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
      Left            =   2685
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10170
      Width           =   1140
   End
   Begin VB.CommandButton cmd_navigate 
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
      Left            =   3915
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10170
      Width           =   1140
   End
   Begin VB.CommandButton cmd_navigate 
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10170
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   240
      TabIndex        =   20
      Top             =   9480
      Width           =   14655
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
         Left            =   90
         TabIndex        =   21
         Top             =   240
         Width           =   14235
      End
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   330
      Left            =   1770
      TabIndex        =   0
      Top             =   1110
      Width           =   1500
      _ExtentX        =   2646
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
      Format          =   134086659
      CurrentDate     =   37867
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   1560
      Width           =   14550
      Begin VB.CommandButton cmd_update 
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
         Index           =   0
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label LblWomIn 
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
         Left            =   2160
         TabIndex        =   43
         Top             =   300
         Width           =   60
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   8970
         X2              =   11400
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Material Cls"
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
         Left            =   11775
         TabIndex        =   37
         Top             =   735
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSForms.ComboBox cbo_material 
         Height          =   330
         Index           =   0
         Left            =   12930
         TabIndex        =   6
         Top             =   660
         Visible         =   0   'False
         Width           =   1455
         VariousPropertyBits=   746604575
         BackColor       =   14737632
         MaxLength       =   6
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2566;582"
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_supplyNo 
         Height          =   330
         Index           =   0
         Left            =   2130
         TabIndex        =   2
         Top             =   225
         Visible         =   0   'False
         Width           =   1995
         VariousPropertyBits=   746604571
         MaxLength       =   13
         DisplayStyle    =   3
         Size            =   "3519;582"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Request No."
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
         Left            =   225
         TabIndex        =   36
         Top             =   300
         Width           =   1680
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machine Name"
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
         Left            =   11520
         TabIndex        =   35
         Top             =   1140
         Width           =   1245
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   12930
         X2              =   14310
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label lbl_machine 
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
         Left            =   12915
         TabIndex        =   34
         Top             =   1155
         Width           =   1470
      End
      Begin MSForms.ComboBox cbo_MachineNo 
         Height          =   330
         Index           =   0
         Left            =   5280
         TabIndex        =   8
         Top             =   1080
         Width           =   1995
         VariousPropertyBits=   746604575
         BackColor       =   14737632
         MaxLength       =   6
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3519;582"
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machine No."
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
         Left            =   4005
         TabIndex        =   33
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location Name"
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
         Left            =   7560
         TabIndex        =   32
         Top             =   1140
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   4005
         TabIndex        =   31
         Top             =   735
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Warehouse CD"
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
         Left            =   225
         TabIndex        =   29
         Top             =   735
         Width           =   1785
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Location CD"
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
         Left            =   225
         TabIndex        =   28
         Top             =   1155
         Width           =   1305
      End
      Begin VB.Label Label5 
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
         Left            =   8610
         TabIndex        =   27
         Top             =   735
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSForms.ComboBox cbo_warehouse 
         Height          =   330
         Index           =   0
         Left            =   2130
         TabIndex        =   4
         Top             =   675
         Width           =   1500
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_location 
         Height          =   330
         Index           =   0
         Left            =   2160
         TabIndex        =   7
         Top             =   1095
         Width           =   1500
         VariousPropertyBits=   746604575
         BackColor       =   14737632
         MaxLength       =   6
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2646;582"
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_supply 
         Height          =   330
         Index           =   0
         Left            =   9690
         TabIndex        =   5
         Top             =   667
         Visible         =   0   'False
         Width           =   780
         VariousPropertyBits=   746604575
         BackColor       =   14737632
         MaxLength       =   2
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1376;582"
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lbl_warehouse 
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
         Left            =   5265
         TabIndex        =   26
         Top             =   735
         Width           =   3000
      End
      Begin VB.Label lbl_location 
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
         Left            =   8970
         TabIndex        =   25
         Top             =   1155
         Width           =   2445
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   5265
         X2              =   8325
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line3 
         Index           =   0
         Visible         =   0   'False
         X1              =   10560
         X2              =   11655
         Y1              =   975
         Y2              =   975
      End
      Begin VB.Label lbl_supply 
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
         Left            =   10560
         TabIndex        =   24
         Top             =   735
         Visible         =   0   'False
         Width           =   60
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13020
      TabIndex        =   42
      Top             =   240
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   6075
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   3240
      Width           =   14685
      _cx             =   25903
      _cy             =   10716
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
      GridColorFixed  =   8421504
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
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
      Begin MSForms.ComboBox cboRepItem 
         Height          =   315
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "3201;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Label lbl_record 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_record"
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
      Left            =   6840
      TabIndex        =   41
      Top             =   10680
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lbl_supply 
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
      Index           =   2
      Left            =   5310
      TabIndex        =   40
      Top             =   1185
      Width           =   60
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   5310
      X2              =   6405
      Y1              =   1395
      Y2              =   1395
   End
   Begin MSForms.ComboBox cbo_supply 
      Height          =   330
      Index           =   2
      Left            =   4440
      TabIndex        =   1
      Top             =   1110
      Width           =   780
      VariousPropertyBits=   746604571
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "1376;582"
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
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
      Index           =   2
      Left            =   3360
      TabIndex        =   39
      Top             =   1185
      Width           =   915
   End
   Begin VB.Label Label6 
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
      Left            =   465
      TabIndex        =   38
      Top             =   1185
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Parts (Material) Supply Request [Automatic]"
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
      Index           =   0
      Left            =   300
      TabIndex        =   22
      Top             =   375
      Width           =   14520
   End
End
Attribute VB_Name = "frm_part_supplyAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim kali As Integer
Dim rs_part_supply As New ADODB.Recordset
Dim rs_warehouse As New ADODB.Recordset
Dim rs_trade_master As New ADODB.Recordset
Dim rs_item As New ADODB.Recordset
Dim l_update_stock As Double
Dim l_tambah_stock As Double
Dim l_update_lotNo As String
Dim l_item_code_update As String, l_supply_cls As String, l_stock_warehouse As String
Dim stockcontrol_cls   As String, l_stock_location As String, l_seqNo As Double
Dim cboEmpty As Boolean
Dim LocationHasNoLine As Boolean
Dim ld_formulaQty As Double
Dim tempRowPartSupplyDetail As Integer, Idx As Integer

Public parentItemCode As String
Public is_groupRequest As String
Public il_selectedRecord As Long
Public is_status As String
Public is_DailySeqNo As String
Public ib_fromProd As Boolean
Public is_group As String

Dim bteColMatCode As Byte
Dim bteColPartNumber As Byte
Dim bteColDesc As Byte
Dim bteColWHCode As Byte
Dim bteColQty As Byte
Dim bteColStock As Byte
Dim bteColReqQty As Byte
Dim bteColUnit As Byte
Dim bteColPackSize As Byte
Dim bteColLotNo As Byte
Dim bteColRemarks As Byte
Dim bteColSeqNo As Byte
Dim bteColParentItem As Byte
Dim bteColUnitCode As Byte
Dim bteColDetailStatus As Byte
Dim bteColQtyBef As Byte
Dim bteColStdPack As Byte
Dim bteColChangeItem As Byte

Const bteColReqCls As Byte = 16 'Link ke Part Supply Request
Const bteColGroupCls As Byte = 17 'Link ke Part Supply Request

Sub Header(Index As Integer)

    bteColMatCode = 0
    bteColPartNumber = 1
    bteColDesc = 2
    bteColWHCode = 3
    bteColQty = 4
    bteColStock = 5
    bteColReqQty = 6
    bteColUnit = 7
    bteColPackSize = 8
    bteColLotNo = 9
    bteColChangeItem = 10
    bteColRemarks = 11
    bteColParentItem = 12
    bteColSeqNo = 13 'Seq No Link ke Part Supply Request
    bteColUnitCode = 14
    bteColDetailStatus = 15
    bteColQtyBef = 16
    bteColStdPack = 17
        
    With Grid1(Index)
        .ColS = 18
        .Rows = 1
        .clear
        
        .TextMatrix(0, bteColMatCode) = "Material Code"
        .TextMatrix(0, bteColPartNumber) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColWHCode) = "Warehouse CD"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColStock) = "Current Stock"
        .TextMatrix(0, bteColReqQty) = "Request Qty"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColPackSize) = "P. Size"
        .TextMatrix(0, bteColLotNo) = "Lot No."
        .TextMatrix(0, bteColChangeItem) = "Change Material Code"
        .TextMatrix(0, bteColRemarks) = "Remarks"
        .TextMatrix(0, bteColSeqNo) = "SeqNo"
        .TextMatrix(0, bteColParentItem) = "ParentItemCode"
        .TextMatrix(0, bteColUnitCode) = "UnitCode"
        '#If DetailStatus=1 then data from current Supply Req No.
        '#If DetailStatus=0 then data from other Supply Req No.
        .TextMatrix(0, bteColDetailStatus) = "DetailStatus"
        .TextMatrix(0, bteColQtyBef) = "QtyBeforeUpdate"
        .TextMatrix(0, bteColStdPack) = "shdPackingSize"
'        .TextMatrix(0, bteColChangeItem) = "Change Material Code"
        
        .ColWidth(bteColMatCode) = 1500
        .ColWidth(bteColPartNumber) = 1500
        .ColWidth(bteColDesc) = 3500
        .ColWidth(bteColWHCode) = 1500
        .ColWidth(bteColQty) = 1300
        .ColWidth(bteColStock) = 1300
        .ColWidth(bteColReqQty) = 1300
        .ColWidth(bteColUnit) = 700
        .ColWidth(bteColPackSize) = 700
        .ColWidth(bteColLotNo) = 1000
        .ColWidth(bteColRemarks) = 2000
        .ColWidth(bteColSeqNo) = 1000
        .ColWidth(bteColChangeItem) = 2000
        
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColParentItem) = True
        .ColHidden(bteColUnitCode) = True
        .ColHidden(bteColDetailStatus) = True
        .ColHidden(bteColQtyBef) = True
        .ColHidden(bteColStdPack) = True
        .ColHidden(bteColPackSize) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignLeftCenter
        .ColAlignment(bteColMatCode) = flexAlignLeftCenter
        .ColAlignment(bteColPartNumber) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColWHCode) = flexAlignLeftCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColStock) = flexAlignRightCenter
        .ColAlignment(bteColReqQty) = flexAlignRightCenter
        .ColAlignment(bteColUnit) = flexAlignLeftCenter
        .ColAlignment(bteColPackSize) = flexAlignRightCenter
        .ColAlignment(bteColLotNo) = flexAlignLeftCenter
        .ColAlignment(bteColRemarks) = flexAlignLeftCenter
        .ColAlignment(bteColSeqNo) = flexAlignLeftCenter
        .ColAlignment(bteColChangeItem) = flexAlignLeftCenter
        
        .EditMaxLength = 1
    End With

End Sub
Private Sub cbo_warehouse_Change(Index As Integer)
If Index = 0 Then
Call uf_ValidateComboData(cbo_warehouse(0), "4018", lbl_pesan, lbl_warehouse(0))
End If
End Sub
Private Sub cbo_location_Change(Index As Integer)
If Index = 0 Then
Call uf_ValidateComboData(cbo_location(0), "4014", lbl_pesan, lbl_location(0))
End If
End Sub

'Private Sub CboItemCode_Change()
'
'    If cboitemcode.MatchFound Then
'        'LblPart.Text = CboItemCode.Column(1)
'        lbldesc.Text = cboitemcode.Column(2)
'        lblWHCode.Text = cboitemcode.Column(3)
'        lblUnit(1).Text = cboitemcode.Column(4)
'        txtqty.Text = Format(0, gs_formatQty)
'    Else
'        'LblPart.Text = ""
'        lbldesc.Text = ""
'        lblWHCode.Text = ""
'        lblUnit(1).Text = ""
'        txtqty.Text = Format(0, gs_formatQty)
'    End If
'
'
'End Sub

Private Sub Cmd_delete_Click()
Dim lrs_check As New ADODB.Recordset
Dim ls_supReqNoResin As String

'#Loop to get SupplyReqNo Resin
If cbo_supplyNo(0).ListCount > 0 Then
    For i = 0 To cbo_supplyNo(0).ListCount - 1
        ls_supReqNoResin = ls_supReqNoResin + "'" + Trim(cbo_supplyNo(0).List(i, 0)) + "',"
    Next
    
    If Len(Trim(ls_supReqNoResin)) > 0 Then _
    ls_supReqNoResin = Left(ls_supReqNoResin, Len(Trim(ls_supReqNoResin)) - 1)
End If

If Trim(ls_supReqNoResin) = "" Then Exit Sub

i = MsgBox("Do you really want to delete this data ?", vbYesNo + vbQuestion, "Warning")
If i = vbYes Then
   
    '#Check if data request already used by supply
    If lrs_check.State <> adStateClosed Then lrs_check.Close
    lrs_check.Open " select * from part_supply where supplyRec_No in (" & ls_supReqNoResin & ")", Db, adOpenKeyset, adLockOptimistic
        
    If lrs_check.EOF = False Then
        lbl_pesan.Caption = DisplayMsg(1204) '"Data already used by another application!"
        Exit Sub
    End If
    
    '#Delete data request
    Call UpdateRequestNo("delete")
    Call cmd_clear_Click
    Call frm_ProdResultAutoRequest.cmdSearch_Click
    lbl_pesan = DisplayMsg(1201) '#Delete data success
End If
End Sub

Private Sub cmd_navigate_Click(Index As Integer)
Select Case Index
Case 1:
    navigate ("first")
Case 2:
    navigate ("previous")
Case 3:
    navigate ("next")
Case 4:
    navigate ("last")
End Select
End Sub

Private Sub navigate(ls_direction As String)

Dim li_increment As Integer

For i = 0 To 0
    If Trim(ls_direction) = "next" Or Trim(ls_direction) = "previous" Then
        '#Init li_increment
        li_increment = IIf(Trim(ls_direction) = "next", 1, -1)
        
        '#Navigate
        If cbo_supplyNo(i).ListIndex + li_increment <= cbo_supplyNo(i).ListCount - 1 And _
           cbo_supplyNo(i).ListIndex + li_increment >= 0 Then
            cbo_supplyNo(i) = cbo_supplyNo(i).List(cbo_supplyNo(i).ListIndex + li_increment, 0)
        End If
            
    Else
        
        '#Init li_increment
        li_increment = IIf(Trim(ls_direction) = "first", 0, cbo_supplyNo(i).ListCount - 1)
        
        '#Navigate
        cbo_supplyNo(i) = cbo_supplyNo(i).List(li_increment, 0)
    
    End If
    
    '#Call data
    Call cmd_update_Click((i))
Next

End Sub

Private Sub GetFormulaQty(StrParent As String, strItemCode As String, dblQty As Double)
    
    Dim adoRs As New ADODB.Recordset
    Dim booRecurring As Boolean
    
    On Error GoTo errHandler
    
    sql = "Select c.Item_Code, c.Production_Cls, a.Qty, Item_Child = (Select Distinct Parent_ItemCode From BOM_Master Where Parent_ItemCode = a.Item_Code) " & _
        "From BOM_Master a " & _
        "Inner Join Item_Master b On a.Parent_ItemCode = b.Item_Code " & _
        "Inner Join Item_Master c On a.Item_Code = c.Item_Code " & _
        "Where b.Control_Cls = '01' And c.Control_Cls = '01' And c.SupplyIssue_Cls = '01' And a.Parent_ItemCode = '" & StrParent & "'"
    
    adoRs.Open sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
    While Not adoRs.EOF
        
        booRecurring = adoRs.Fields("Production_Cls") <> "01"
        If booRecurring Then booRecurring = Not IsNull(adoRs.Fields("Item_Child"))
        
        If booRecurring Then
            GetFormulaQty Trim(adoRs.Fields("Item_Code")), strItemCode, dblQty * adoRs.Fields("Qty")
            'GetFormulaQty Trim(adoRs.Fields("Item_Code")), adoRs.Fields("Item_Child"), dblQty * adoRs.Fields("Qty")
        Else
            If Trim(adoRs.Fields("Item_Code")) = strItemCode Then
                ld_formulaQty = ld_formulaQty + (dblQty * adoRs.Fields("Qty"))
            End If
        End If
        adoRs.MoveNext
    Wend
    adoRs.Close
    
ErrExit:
    Set adoRs = Nothing
    Exit Sub
errHandler:
    lbl_pesan.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
    
End Sub

Private Sub setting_grid(Index As Integer)
    
    Dim rs_join As New ADODB.Recordset
    Dim ls_sqlJoin As String
    Dim ld_pakingSize As Double

    On Error GoTo errHandler
    Me.MousePointer = vbHourglass
'    kali = kali + 1
'    MsgBox kali
    Call Header(Index)

    If Trim(is_group) = "" Then is_group = "0"

'    ls_sqlJoin = "Select psm.Request_Cls, psm.FromWarehouse_Code, " & _
'        "ActualQty = " & _
'        "Case When psd.SupplyRec_No = '" & Trim(cbo_supplyNo(Index)) & "' Then ChildRequirement_Qty Else " & _
'        "Isnull((Select Isnull(ChildRequirement_Qty, 0) From PartSupplyRequest_Detail p Left Join PartSupplyRequest_Master q On q.SupplyRec_No = p.SupplyRec_No " & _
'        "Where p.ChildItem_Code = psd.ChildItem_Code And p.ChildLot_No = psd.ChildLot_No And p.SupplyRec_No = '" & Trim(cbo_supplyNo(Index)) & "' And q.FromWarehouse_Code = '" & Trim(cbo_warehouse(Index)) & "'), 0) End, " & _
'        "DetailStatus = Case When psd.SupplyRec_No = '" & Trim(cbo_supplyNo(Index)) & "' Then '1' Else '0' End, psd.*, im.Item_Name, psd.ChildUnit_Cls Unit_Cls, " & _
'        "Unit_Desc = (Select Description From Unit_Cls Where Unit_Cls = psd.ChildUnit_Cls), " & _
'        "im.Number_Box As PackingSize, dp.Lot_No As DLot_No, dp.Remark As DRemarks, dp.Qty, im.makeritem_code " & _
'        "From PartSupplyRequest_Detail psd " & _
'        "Left Join PartSupplyRequest_Master psm on psd.SupplyRec_No = psm.SupplyRec_No " & _
'        "Inner Join Item_Master im on psd.ChildItem_Code = im.Item_Code " & _
'        "Inner Join Daily_Production dp On psd.DailySeq_No = dp.Seq_No " & _
'        "Where psm.Request_Cls = '" & Trim(is_group) & "' " & _
'        "Order By psm.FromWarehouse_Code, ChildItem_Code, ChildLot_No"
    
    ' *************************************
    ' Perubahan Proses Perhitungan untuk QtyBOM
    ' *************************************
    ls_sqlJoin = " Select psm.Request_Cls, psm.FromWarehouse_Code, " & vbCrLf & _
                              "  ActualQty = Case When psd.SupplyRec_No = '" & Trim(cbo_supplyNo(Index)) & "'  " & vbCrLf & _
                              "     Then ChildRequirement_Qty Else Isnull((Select Isnull(ChildRequirement_Qty, 0)  " & vbCrLf & _
                              "         From PartSupplyRequest_Detail p Left Join PartSupplyRequest_Master q On q.SupplyRec_No = p.SupplyRec_No  " & vbCrLf & _
                              "             Where p.ChildItem_Code = psd.ChildItem_Code And p.ChildLot_No = psd.ChildLot_No  " & vbCrLf & _
                              "                     And p.SupplyRec_No = '" & Trim(cbo_supplyNo(Index)) & "' And q.FromWarehouse_Code = '" & Trim(cbo_warehouse(Index)) & "'), 0) End, " & vbCrLf & _
                              "  RequestQty = (Select isnull(Sum(QtyBOM),0)  " & vbCrLf & _
                              "         From TempRequest Tq  " & vbCrLf & _
                              "             Where TQ.Parent_ItemCode=psd.ParentItem_Code and Tq.Item_Code = psd.ChildItem_Code and SeqNo=psd.dailySeq_No " & vbCrLf & _
                              "                     And Tq.LotNo = psd.ChildLot_No group By Parent_ItemCode,SeqNo,LotNo,Item_Code), " & vbCrLf & _
                              "  DetailStatus = Case When psd.SupplyRec_No = '" & Trim(cbo_supplyNo(Index)) & "' Then '1' Else '0' End,  " & vbCrLf
    
    ls_sqlJoin = ls_sqlJoin + "  psd.*, im.Item_Name, psd.ChildUnit_Cls Unit_Cls, Unit_Desc = (Select Description From Unit_Cls Where Unit_Cls = psd.ChildUnit_Cls),  " & vbCrLf & _
                              "  im.Number_Box As PackingSize, dp.Lot_No As DLot_No, dp.Remark As DRemarks, dp.Qty, im.makeritem_code, ISNULL(psd.ReplacementItem_Code,'')ChangeItem_Code   " & vbCrLf & _
                              "     From PartSupplyRequest_Detail psd  " & vbCrLf & _
                              "     Left Join PartSupplyRequest_Master psm on psd.SupplyRec_No = psm.SupplyRec_No  " & vbCrLf & _
                              "     Inner Join Item_Master im on psd.ChildItem_Code = im.Item_Code  " & vbCrLf & _
                              "     Inner Join Daily_Production dp On psd.DailySeq_No = dp.Seq_No  " & vbCrLf & _
                              "         Where psm.Request_Cls = '" & Trim(is_group) & "' Order By psm.FromWarehouse_Code, ChildItem_Code, ChildLot_No "
    
    ' *************************
    
    '20230118
    'If rs_join.State <> adStateClosed Then rs_join.Close
    'rs_join.Open ls_sqlJoin, Db, adOpenDynamic
            
    If rs_join.State <> adStateClosed Then rs_join.Close
    rs_join.Open ls_sqlJoin, Db, adOpenKeyset, adLockOptimistic
            
    tempRowPartSupplyDetail = rs_join.RecordCount
   
    If CDec(txtTemp) <> 0 Then
        If CDec(txtTemp) <> tempRowPartSupplyDetail Then
            cmd_submit.Enabled = False
            lbl_pesan = DisplayMsg(9014)
        End If
    End If
        
   
   Do While Not rs_join.EOF ''While rs_join.EOF = False
        
        '#Init Formula Qty
        ld_formulaQty = 0
        'Perubahan Yudi
       'GetFormulaQty Trim(rs_join!parentItem_code), Trim(rs_join!childitem_code), rs_join!Qty
        
        '#Init packingSize
        If rs_join!PackingSize <> 0 Then
          ld_pakingSize = uf_Ceiling(rs_join!ActualQty / rs_join!PackingSize)
        Else
          ld_pakingSize = 0
        End If
        
        With Grid1(Index)
            .AddItem ""
            .Cell(flexcpBackColor, .Rows - 1, bteColReqQty) = vbWhite
             .Cell(flexcpBackColor, .Rows - 1, bteColChangeItem) = vbWhite
            
            .TextMatrix(.Rows - 1, bteColMatCode) = Trim(rs_join!childitem_code)
            .TextMatrix(.Rows - 1, bteColPartNumber) = Trim(rs_join!MakerItem_Code)
            .TextMatrix(.Rows - 1, bteColDesc) = uf_GetItemDescription(Trim(rs_join!childitem_code))
            .TextMatrix(.Rows - 1, bteColWHCode) = Trim(rs_join!FromWarehouse_Code)
            .TextMatrix(.Rows - 1, bteColQty) = Format(Trim(rs_join!RequestQty), gs_formatQtyBOM) ' Change Source of QtyBOM
            .TextMatrix(.Rows - 1, bteColStock) = Format(GetCurrentStock(rs_join!childitem_code, cbo_warehouse(0)), gs_formatQtyBOM)
            .TextMatrix(.Rows - 1, bteColReqQty) = Format(Trim(rs_join!ActualQty), gs_formatQtyBOM)
            .TextMatrix(.Rows - 1, bteColUnit) = Trim(rs_join!Unit_Desc)
            .TextMatrix(.Rows - 1, bteColPackSize) = Format(ld_pakingSize, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColLotNo) = Trim(rs_join!DLot_No)
            .TextMatrix(.Rows - 1, bteColRemarks) = Trim(rs_join!DRemarks)
            .TextMatrix(.Rows - 1, bteColSeqNo) = Trim(rs_join!Seq_no)
            .TextMatrix(.Rows - 1, bteColParentItem) = Trim(rs_join!parentItem_code)
            .TextMatrix(.Rows - 1, bteColUnitCode) = Trim(rs_join!Unit_cls)
            .TextMatrix(.Rows - 1, bteColDetailStatus) = Trim(rs_join!DetailStatus)
            .TextMatrix(.Rows - 1, bteColChangeItem) = Trim(rs_join!ChangeItem_Code)
            .TextMatrix(.Rows - 1, bteColQtyBef) = Format(Trim(rs_join!ActualQty), gs_formatQtyBOM)
            .TextMatrix(.Rows - 1, bteColStdPack) = Trim(rs_join!PackingSize & "")
        End With
        rs_join.MoveNext
        
    Loop ''Wend
        
    rs_join.Close

'    For i = 1 To Grid1(Index).Rows - 1
''        Grid1(Index).Cell(flexcpBackColor, i, bteColSelect) = vbWhite
'        Grid1(Index).Cell(flexcpBackColor, i, bteColReqQty) = vbWhite
'    Next

ErrExit:
    Set rs_join = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    lbl_pesan.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit

End Sub

Private Sub cbo_location_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
If Index = 0 Then
   Call uf_ValidateComboData(cbo_location(0), "4014", lbl_pesan, lbl_location(0))
End If
End Sub

Private Sub cbo_location_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
If Index = 0 Then
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
Else
    KeyAscii = 0
End If
End Sub

Private Sub cbo_MachineNo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
KeyCode = 0
End Sub

Private Sub cbo_MachineNo_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub cbo_material_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
KeyCode = 0
End Sub

Private Sub cbo_material_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub cbo_supply_Change(Index As Integer)

If ib_fromProd = False Then
    cbo_supply(0) = cbo_supply(2)
Else
    cbo_supply(2) = cbo_supply(0)
End If

lbl_pesan = validCombo
If Trim(lbl_pesan) <> "" Then
    Exit Sub
End If

If cbo_supply(Index).ListIndex >= 0 Then
    lbl_supply(Index).Caption = cbo_supply(Index).List(cbo_supply(Index).ListIndex, 1)
Else
    lbl_supply(Index).Caption = ""
End If

End Sub

Private Sub cbo_supply_Click(Index As Integer)
If cbo_supply(Index).ListCount <= 0 Then Exit Sub
If cbo_supply(Index).ListIndex <= 0 Then Exit Sub
lbl_supply(Index).Caption = cbo_supply(Index).List(cbo_supply(Index).ListIndex, 1)
End Sub

Private Sub cbo_supply_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyBack Then lbl_supply(Index) = ""

If KeyCode = 13 Then

    lbl_pesan = validCombo
    If Trim(lbl_pesan) <> "" Then
        If Index > 1 Then Index = 1
        clearGrid (Index)
        Exit Sub
    End If
End If
End Sub

Private Sub cbo_supply_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbo_supplyNo_Click(Index As Integer)
Call clearGrid(Index)
End Sub

Private Sub cbo_supplyNo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
KeyCode = 0
End Sub

Private Sub cbo_warehouse_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
If Index = 0 Then
    Call uf_ValidateComboData(cbo_warehouse(0), "4018", lbl_pesan, lbl_warehouse(0))
End If
End Sub

Sub clearGrid(Index As Integer)
Call clearFrameTop(Index)
Grid1(Index).clear
Grid1(Index).Rows = 1
Call Header(Index)
End Sub

Private Sub cbo_warehouse_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
If Index = 0 Then
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
Else
    KeyAscii = 0
End If
End Sub

Public Sub navigateButton(Status As Boolean)
    cmd_sub_menu.Enabled = Status
    cmd_daily.Enabled = Status
    cmd_preview.Enabled = Status
    cmd_clear.Enabled = Status
    Cmd_delete.Enabled = Status
End Sub

Private Sub cmd_Cancel_Click()

'#Enabled Button Preview etc
Call navigateButton(True)

lbl_pesan = ""
Call setting_grid(0)

'#Delete Current SupplyReqNo in Supply Request Master
If is_status = "insert" Then
    Call UpdateRequestNo("delete")
    Call frm_ProdResultAutoRequest.cmdSearch_Click
    Call cmd_clear_Click
End If

End Sub

Private Sub cmd_clear_Click()
cbo_supplyNo(0) = ""
Call clearFrameTop(0)
Call Header(0)
End Sub

Sub clearFrameTop(Index As Integer)

If Index > 1 Then Index = 1

DTPicker3.Value = Format(Date, "dd MMM yyyy")

lbl_warehouse(Index).Caption = ""
lbl_location(Index).Caption = ""
lbl_pesan.Caption = ""
lbl_supply(Index).Caption = ""
lbl_machine(Index).Caption = ""

cbo_location(Index) = ""
cbo_warehouse(Index) = ""
cbo_MachineNo(Index) = ""
cbo_supply(Index) = ""

cbo_material(Index) = ""

'Call setting_grid(Index)
End Sub


Private Sub cmd_daily_Click()
    With frm_ProdResultAutoRequest
        .fromProd = False
        .Command1(1).Enabled = False
        .cmdSubMenu.Caption = "&Back"
        .Show
        Call .initGroup
        Me.Hide
    End With
End Sub

Private Sub cmd_sub_menu_Click()

Dim n As Integer

If cmd_sub_menu.Caption = "&Back" Then
    '#Delete Current SupplyReqNo in Supply Request Master
    If is_status = "insert" Then
        cmd_submit.Enabled = True
        Call UpdateRequestNo("delete")
        Call cmd_clear_Click
    End If
    
    '#Back to frmProdResultInquiry
    With frm_ProdResultAutoRequest
        
        .fromProd = False
        .is_groupRequest = is_group 'Request
        .il_selectedRecord = il_selectedRecord
        
        If is_group <> "" Then Call frm_ProdResultAutoRequest.cmdSearch_Click
        
'        Request jika di Back, grid di-refresh saja, data yang sudah di check sebelumnya di clear saja.
'        For i = 1 To .Grid.Rows - 1
'            For n = 0 To UBound(Split(is_DailySeqNo, ","))
'                If .Grid.TextMatrix(i, bteColSeqNo) = Mid$(Split(is_DailySeqNo, ",")(n), 1, Len(Split(is_DailySeqNo, ",")(n))) Then
'                    .Grid.Cell(flexcpChecked, i, bteColSelect) = flexChecked
'                    .is_request = Trim(.Grid.TextMatrix(i, bteColReqCls))
'                    .is_groupRequest = Trim(.Grid.TextMatrix(i, bteColGroupCls))
'                End If
'            Next
'        Next
        
        .Show
        Me.Hide
    
    End With

Else
    '#Back to Main Menu
    Grid1(0).Rows = 1
    
    Unload frm_part_supplyAuto
    Unload frm_ProdResultAutoRequest
    frmMainMenu.Show

End If
End Sub

Private Function checkParentItemCode(serialNo As String, lotno As String) As String
 '# Check if Supply Data had been used for another receipt data or not
Dim recCheck As Integer
Dim sqlCheck As String
Dim rsCheck As New ADODB.Recordset

sqlCheck = " select * from daily_production where serial_no ='" & serialNo & "' and lot_no='" & lotno & "' "
                              
If rsCheck.State <> adStateClosed Then rsCheck.Close
rsCheck.Open sqlCheck, Db, adOpenKeyset, adLockOptimistic

If rsCheck.EOF = False Then '#Parent Item Code Exist
    checkParentItemCode = Trim(rsCheck!Item_Code)
Else '#Parent Item Code doesn't Exist
    checkParentItemCode = ""
End If

End Function

Private Sub UpdateRequestNo(Status As String)

Dim ls_supReqNoResin As String
Dim ls_supReqNoOther As String

'#Loop to get SupplyReqNo Resin
If cbo_supplyNo(0).ListCount > 0 Then
    For i = 0 To cbo_supplyNo(0).ListCount - 1
        ls_supReqNoResin = ls_supReqNoResin + "'" + Trim(cbo_supplyNo(0).List(i, 0)) + "',"
    Next
    
    If Len(Trim(ls_supReqNoResin)) > 0 Then _
    ls_supReqNoResin = Left(ls_supReqNoResin, Len(Trim(ls_supReqNoResin)) - 1)
    
End If

Db.BeginTrans

If Status = "update" Then
       
    '#Update Suppply_Cls and supply_date in Supply Request Master
    Dim ls_sqlMasterUpdate As String
    '#Resin
        If Trim(ls_supReqNoResin) <> "" Then
            ls_sqlMasterUpdate = " Update partSupplyRequest_master set " & _
                               " childSupply_date='" & Format(DTPicker3, "yyyy-MM-dd") & "', " & _
                               " supply_Cls='" & Trim(cbo_supply(0)) & "' " & _
                               " where supplyRec_No in (" & Trim(ls_supReqNoResin) & ") "

            Db.Execute ls_sqlMasterUpdate

        End If
    

Else '#Delete

    '#Update status in daily
    If is_status = "insert" Then
        '#If cancel insert
        If Trim(is_DailySeqNo) <> "" Then
            Db.Execute "update  daily_production  set request_cls=null where seq_No in(" & Trim(is_DailySeqNo) & ")"
        End If
    Else
        '#If delete after insert
        Db.Execute "update  daily_production  set request_cls=null where request_cls = (select distinct request_cls from daily_production where seq_No in(" & Trim(is_DailySeqNo) & "))"
    End If
       
    
    '#Delete in Supply Request Master
    Dim ls_sqlMasterUpdate2 As String
    '#Resin
        If Trim(ls_supReqNoResin) <> "" Then
            ls_sqlMasterUpdate2 = " delete from partSupplyRequest_master with (updlock) " & _
                                " where supplyRec_No in (" & Trim(ls_supReqNoResin) & ") "
            
            Db.Execute ls_sqlMasterUpdate2
            cbo_supplyNo(0).clear
        End If

End If

Db.CommitTrans
End Sub

Private Sub Cmd_Submit_Click()

Dim j As Integer
Dim sqlupd As String

If hakUpdate(Me.Name) = 0 Then _
lbl_pesan = DisplayMsg(3008): Exit Sub

lbl_pesan = validCombo
If lbl_pesan <> "" Then Exit Sub

    '#If No data in grid
    If Grid1(0).Rows <= 1 Then Exit Sub
    
    '#Enable Button preview etc
    Call navigateButton(True)
    
    Me.MousePointer = vbHourglass
    
    '#Update data Supply CLs and Request date in Supply Request Master
    Call UpdateRequestNo("update")
    
    If cbo_location(0).Enabled = True Then
        sqlupd = "update partsupplyrequest_master " & _
            "set FromWarehouse_Code = '" & Trim(cbo_warehouse(0)) & "', " & _
            "ToWarehouse_code='" & Trim(cbo_location(0)) & "' " & _
            "where SupplyRec_No='" & Trim(cbo_supplyNo(0)) & "'"
        Db.Execute sqlupd
    End If
    sqlupd = ""
    
    Db.BeginTrans
    For i = 0 To 0
    
        For j = 1 To Grid1(i).Rows - 1
            
            If Trim(Grid1(i).TextMatrix(j, bteColReqQty)) = "" Then Grid1(i).TextMatrix(j, bteColReqQty) = "0"
            If Trim(Grid1(i).TextMatrix(j, bteColQtyBef)) = "" Then Grid1(i).TextMatrix(j, bteColQtyBef) = "0"
            
            '#If data didn't change skip update
            If CDbl(Grid1(i).TextMatrix(j, bteColReqQty)) = CDbl(Grid1(i).TextMatrix(j, bteColQtyBef)) And Grid1(i).TextMatrix(j, bteColChangeItem) = "" Then GoTo NextData
            
            '#Update & Delete
            If (CDbl(Grid1(i).TextMatrix(j, bteColQtyBef)) > 0 And (Grid1(i).TextMatrix(j, bteColDetailStatus)) = "0") Or _
               (CDbl(Grid1(i).TextMatrix(j, bteColQtyBef)) >= 0 And (Grid1(i).TextMatrix(j, bteColDetailStatus)) = "1") Or _
               (Grid1(i).TextMatrix(j, bteColChangeItem) <> "") Then
                
                '#If Actual Qty is Set To Zero and data from another
                '#Supply Request no then
                '#Delete
                If CDbl(Grid1(i).TextMatrix(j, bteColReqQty)) = 0 _
                    And (Grid1(i).TextMatrix(j, bteColDetailStatus)) = "0" Then
                            
                    '#Delete data in Supply request Detail
                    sqlupd = " delete partSupplyRequest_Detail with (updlock) " & _
                                " where supplyRec_No='" & Trim(cbo_supplyNo(i)) & "' and " & _
                                " childLot_no='" & Trim(Grid1(i).TextMatrix(j, bteColLotNo)) & "' and " & _
                                " ChildItem_code='" & Trim(Grid1(i).TextMatrix(j, bteColMatCode)) & "' "
                                    
                    Db.Execute sqlupd
                            
                Else '#Update
                    '# Data from Current Supply Req No.
                    '# or (data from Another Supply Req No. And Actual
                    '# Qty not Zero)
                    
    
                    sqlupd = " update partSupplyRequest_Detail with (updlock) set " & _
                                " childRequirement_Qty='" & CDbl(Grid1(i).TextMatrix(j, bteColReqQty)) & "', " & _
                                " ReplacementItem_Code ='" & Grid1(i).TextMatrix(j, bteColChangeItem) & "', Last_Update = getdate(), " & _
                                " Last_User ='" & userLogin & "', " & _
                                " Remarks = '" & Trim(Grid1(i).TextMatrix(j, 9)) & "' " & _
                                " where supplyRec_No='" & Trim(cbo_supplyNo(i)) & "' and " & _
                                " childLot_no='" & Trim(Grid1(i).TextMatrix(j, bteColLotNo)) & "' and " & _
                                " ChildItem_code='" & Trim(Grid1(i).TextMatrix(j, bteColMatCode)) & "'"
                                
                    Db.Execute sqlupd
                End If
                
            Else '#Insert
                Dim ls_insert As String
                Dim ls_SeqNoDetail As String
                            
                '#Check Last Detail Seqno for current Supply Request No
                Dim ls_sqlCheck As String
                Dim lrs_check As New ADODB.Recordset
                
                ls_sqlCheck = " select max(seq_no)+1 seq_no from partsupplyRequest_detail psd " & _
                              " where supplyRec_no='" & Trim(cbo_supplyNo(i)) & "'"
                                              
                If lrs_check.State <> adStateClosed Then lrs_check.Close
                lrs_check.Open ls_sqlCheck, Db, adOpenKeyset, adLockOptimistic
                            
                '#Init ls_SeqNoDetail
                If lrs_check.EOF = False Then
                    ls_SeqNoDetail = Trim(lrs_check!Seq_no)
                Else
                    ls_SeqNoDetail = "0"
                End If
                
                ls_insert = "'" & Trim(cbo_supplyNo(i)) & "','" & _
                            Trim(ls_SeqNoDetail) & "','" & _
                            Trim(Grid1(i).TextMatrix(j, bteColMatCode)) & "','" & _
                            Trim(Grid1(i).TextMatrix(j, bteColLotNo)) & "'," & _
                            CDbl(Grid1(i).TextMatrix(j, bteColReqQty)) & ",'" & _
                            Trim(Grid1(i).TextMatrix(j, bteColUnitCode)) & "','" & _
                            Trim(Grid1(i).TextMatrix(j, bteColParentItem)) & "','" & _
                            Now & "','" & _
                            userLogin & "','" & _
                            Trim(Grid1(i).TextMatrix(j, bteColRemarks)) & "'"
                            
                sqlupd = " insert into partSupplyRequest_Detail with (updlock)  " & _
                            " (supplyRec_No,Seq_No,ChildItem_code,ChildLot_no," & _
                            " ChildRequirement_qty,ChildUnit_Cls,parentItem_code, Last_Update,Last_User, remarks) " & _
                            " values(" & Trim(ls_insert) & ")"
                            
                Db.Execute sqlupd
    
            End If
            
            '#Init QtyBeforeUpdate after Changes
            Grid1(i).TextMatrix(j, bteColQtyBef) = Trim(Grid1(i).TextMatrix(j, bteColReqQty))
            
NextData:
        Next
        
        '#Refresh Grid
        Call setting_grid((i))
    Next

Me.MousePointer = vbDefault
If is_status = "insert" Then
    lbl_pesan.Caption = DisplayMsg(1000) ' "Insert data success !"
    is_status = "update"
Else
    lbl_pesan.Caption = DisplayMsg(1101) ' "Update data success !"
End If

Db.CommitTrans

End Sub

Public Sub cmd_update_Click(Index As Integer)

If Index = 1 Then Exit Sub

    '#Show Master Request Data
    Dim rsUpd As New ADODB.Recordset
    Dim ls_sql As String
    
    '==================================================================================
    '20070419 Herfin Tambah fungsi u/ check supaya kalau sudah ada data supply, tidak boleh rubah
    'Location code
    ls_sql = " select * from part_supply where supplyrec_no='" & Trim(cbo_supplyNo(0)) & "'"
    If rsUpd.State <> adStateClosed Then rsUpd.Close
    rsUpd.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
    If rsUpd.EOF = True Then
        cbo_location(0).Enabled = True
    Else
        cbo_location(0).Enabled = False
    End If
    '==================================================================================
    
    ls_sql = " select wmFrom.wh_name FromWHname,wmTo.wh_name ToWHname,ml.line_name, " & _
            " prm.* from partSupplyRequest_Master prm " & _
            " left join " & _
            " ( " & _
            "   select wh_code,wh_name,stockControl_cls from warehouse_master  " & _
            "   union all  " & _
            "   select distinct(manufacture_line.manufacture_code)wh_code,trade_name wh_name,stockControl_Cls='01' " & _
            "   from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code  " & _
            " )wmFrom on prm.fromwarehouse_code=wmFrom.wh_code " & _
            " left join " & _
            " ( " & _
            "   select wh_code,wh_name,stockControl_cls from warehouse_master  " & _
            "   union all  " & _
            "   select distinct(manufacture_line.manufacture_code)wh_code,trade_name wh_name,stockControl_Cls='01' " & _
            "   from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code  " & _
            " )wmTo on prm.towarehouse_code=wmTo.wh_code " & _
            " left join manufacture_line ml on prm.machine_no =ml.line_code " & _
            " where supplyRec_No='" & Trim(cbo_supplyNo(Index)) & "'"
    If rsUpd.State <> adStateClosed Then rsUpd.Close
    rsUpd.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
    
    If rsUpd.EOF Then
        rsUpd.Close
        lbl_pesan = DisplayMsg(8093) '"Data with this Supply No. not found !"
        Exit Sub
    End If
    
    rsUpd.MoveFirst
    cbo_warehouse(Index) = Trim(rsUpd!FromWarehouse_Code)
    lbl_warehouse(Index) = Trim(rsUpd!FromWHname)
    cbo_location(Index) = Trim(rsUpd!towarehouse_code)
    lbl_location(Index) = Trim(rsUpd!ToWhName)
    cbo_MachineNo(Index) = Trim(rsUpd!Machine_no)
    lbl_machine(Index) = Trim(rsUpd!line_name)
    
    DTPicker3 = Format(rsUpd!childsupply_date, "dd MMM yyyy")
    cbo_supply(Index) = Trim(rsUpd!supply_cls)
    'cbo_material(Index) = IIf(Trim(rsUpd!material_cls) = "1", "Resin", "Others")
    cbo_supplyNo(Index) = Trim(rsUpd!supplyRec_No)
    rsUpd.Close
    
    '#Show Detail Request Data
    Call setting_grid(Index)

End Sub

Private Sub cmdRepItem_Click()
    Me.MousePointer = vbHourglass
    frm_BrowseItem_WIP.getItemCode = cboRepItem.Text
    frm_BrowseItem_WIP.getParent = txtParent.Text
    frm_BrowseItem_WIP.Show 1
    cboRepItem.Text = frm_BrowseItem_WIP.getItemCode
    Me.MousePointer = vbDefault
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    lbl_pesan.Caption = ErrMsg
End If
End Sub

Function validCombo() As String

Dim j As Integer

'# cek combo Supply Cls
If Trim(cbo_supply(2)) = "" Then
    validCombo = DisplayMsg(4052) '"Please input Supply Clasification !"
    cboEmpty = True: Exit Function
End If

j = 0
For i = 0 To cbo_supply(2).ListCount - 1
    If UCase(Trim(cbo_supply(2))) = UCase(Trim(cbo_supply(2).List(i, 0))) Then
        cbo_supply(2).Text = cbo_supply(2).List(i, 0)
        lbl_pesan.Caption = ""
        j = 1
        Exit For
    End If
Next
      
If j = 0 Then
    lbl_supply(2) = ""
    validCombo = DisplayMsg(4052) '"Invalid supply clasification !"
    Exit Function
End If

 If uf_ValidateComboData(cbo_warehouse(0), "4018", lbl_pesan, lbl_warehouse(0)) = False Then
    validCombo = DisplayMsg(4018)
    Exit Function
 End If
 If uf_ValidateComboData(cbo_location(0), "4014", lbl_pesan, lbl_location(0)) = False Then
    validCombo = DisplayMsg(4014)
    Exit Function
 End If
End Function

Function validasi() As Boolean

lbl_pesan = validCombo
If lbl_pesan <> "" Then Exit Function

End Function



Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
ib_fromProd = False
'Frame1(0).Caption = getMaterialDesc("01")
CtrlMenu1.FormName = Me.Name
'Me.Caption = Me.Caption & " (Menu ID : " & frmcode("frm_partsupplyAuto") & ")"

DTPicker3.Value = Format(Date, "dd MMM yyyy")

'#Form Setting
Call setting

'#Grid Setting
Call Header(0)
'Call Header(1)
cbo_supplyNo(0) = ""
'cbo_supplyNo(1) = ""
Call clearFrameTop(0)
'Call clearFrameTop(1)
Call cmd_clear_Click

'LblWomIn = Trim(cbo_supplyNo(0)) & "-" & ls_FG
lbl_pesan.Caption = ""
is_status = "insert"
End Sub

Private Sub setting()
    cbo_supply(2).clear
    cbo_supply(2).columnCount = 2
    cbo_supply(2).TextColumn = 1
    
    cbo_supply(2).AddItem ""
    cbo_supply(2).List(0, 0) = "S1"
    cbo_supply(2).List(0, 1) = "Supply"
    cbo_supply(2).AddItem ""
    cbo_supply(2).List(1, 0) = "S"
    cbo_supply(2).List(1, 1) = "Consumption"
    cbo_supply(2).AddItem ""
    cbo_supply(2).List(2, 0) = "L"
    cbo_supply(2).List(2, 1) = "Loss"
    cbo_supply(2).AddItem ""
    cbo_supply(2).List(3, 0) = "RJ"
    cbo_supply(2).List(3, 1) = "Reject"
    cbo_supply(2).ColumnWidths = "25 pt; 75 pt"
    cbo_supply(2).ListWidth = 100
    
    Dim SqlW As String
    SqlW = " (select wh_code,wh_name,stockControl_cls from warehouse_master " & _
        " union all  select distinct(manufacture_line.manufacture_code)wh_code,trade_name wh_name,stockControl_Cls='01' from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code )tbJ order by wh_code"
    cbo_location(0).locked = False
    Call up_FillCombo(cbo_location(0), SqlW, "*", "", False)
    cbo_location(0).ColumnWidths = "50 pt; 175 pt"
    cbo_location(0).ListWidth = 225
    cbo_location(0).locked = True
    Call up_FillCombo(cbo_warehouse(0), SqlW, "*", "", False)
    cbo_warehouse(0).ColumnWidths = "50 pt; 175 pt"
    cbo_warehouse(0).ListWidth = 225
    
    addToCboItemCode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid1_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
Dim cek As Integer

'# Check Valid Qty Numeric Input
If IsNumeric(Trim(Grid1(Index).TextMatrix(Row, Col))) = False Then Grid1(Index).TextMatrix(Row, Col) = 0
If Trim(Grid1(Index).TextMatrix(Row, Col)) = "" Then
    Exit Sub
Else
    Grid1(Index).TextMatrix(Row, Col) = Format(Grid1(Index).TextMatrix(Row, Col), gs_formatQtyBOM)
End If
If CDbl(Trim(Grid1(Index).TextMatrix(Row, Col))) > gd_MaxQty Then
    Grid1(Index).TextMatrix(Row, Col) = gd_MaxQty
    lbl_pesan = DisplayMsg(4045) & " " & gd_MaxQty '"Quantity must be equal or less than 9,999,999.99 !"
End If

'#Set Packing Size Qty
If Trim(Grid1(Index).TextMatrix(Row, bteColStdPack)) <> "0" Then
    Grid1(Index).TextMatrix(Row, bteColPackSize) = Format(uf_Ceiling(CDbl(Trim(Grid1(Index).TextMatrix(Row, bteColReqQty))) / CDbl(Trim(Grid1(Index).TextMatrix(Row, bteColStdPack)))), gs_formatQtyBOM)
Else
    Grid1(Index).TextMatrix(Row, bteColPackSize) = "0"
End If
    
End Sub

Private Sub Grid1_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Select Case Grid1(Index).Col
    Case bteColReqQty: Grid1(Index).EditMaxLength = 15
    Case bteColChangeItem: Grid1(Index).EditMaxLength = 25
    Case Else: Cancel = True
End Select

End Sub

Private Sub Grid1_Click(Index As Integer)
If Grid1(Index).Row <> 0 Then
    If Grid1(Index).Col = bteColReqQty Then
       Grid1(Index).FocusRect = flexFocusInset
    Else
       Grid1(Index).FocusRect = flexFocusNone
    End If
End If

With Grid1
    If Grid1(Index).Row <> 0 Then
        If Grid1(Index).Col = bteColChangeItem Then
            cboRepItem.Left = Grid1(Index).CellLeft
            cboRepItem.top = Grid1(Index).CellTop
            cboRepItem.Width = Grid1(Index).CellWidth
            cboRepItem.Height = Grid1(Index).CellHeight
            cboRepItem.Visible = True
            cboRepItem.SetFocus
            cboRepItem.DropDown
        End If
    End If
    
    Idx = Grid1(Index).Row
End With
    
End Sub

Private Sub cmd_preview_Click()

    Dim requestno As String
    
    Me.MousePointer = vbHourglass
    requestno = ""
    For i = 0 To cbo_supplyNo(0).ListCount - 1
        If requestno = "" Then
            requestno = "'" & cbo_supplyNo(0).List(0) & "'"
        Else
            requestno = requestno & ",'" & cbo_supplyNo(0).List(i) & "'"
        End If
    Next
    
    If requestno = "" Then
        lbl_pesan = DisplayMsg(8093) '"Please input Supply Request No. first !"
        cbo_supplyNo(0).SetFocus
        Me.MousePointer = vbDefault: Exit Sub
    End If
    
    Call reportrequestauto(requestno)
    Me.MousePointer = vbDefault
End Sub

Private Sub Grid1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 And Grid1(Index).Row + 1 <> Grid1(Index).Rows Then _
    Grid1(Index).Row = Grid1(Index).Row + 1
    
If Grid1(Index).Col = bteColReqQty Then  '# Qty Column
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 And KeyAscii <> 8 Then KeyAscii = 0
End If

End Sub

Private Sub Grid1_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

If KeyAscii = 13 And Grid1(Index).Row + 1 <> Grid1(Index).Rows Then _
    Grid1(Index).Row = Grid1(Index).Row + 1
    
If Col = bteColReqQty Then '# Qty Column
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 And KeyAscii <> 8 Then KeyAscii = 0
End If

End Sub

Private Function GetCurrentStock(StrItem As String, Optional strWH As String) As Double
    
    Dim adoRs As New ADODB.Recordset
    
    GetCurrentStock = 0
    
    sql = "select isnull(sum(lm_inventory), 0) lm_inventory, " & _
        "isnull(sum(tm_current), 0) tm_current, " & _
        "isnull(sum(nm_current), 0) nm_current " & _
        "from stock_master " & _
        "where item_code = '" & StrItem & "' "
    If Not IsNull(strWH) Then sql = sql & "and warehouse_code = '" & strWH & "' "
    
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        Select Case DateDiff("m", uf_GetLastClosing("fulldate"), DTPicker3.Value)
        Case 0: GetCurrentStock = adoRs.Fields("lm_inventory")
        Case 1: GetCurrentStock = adoRs.Fields("tm_current")
        Case 2: GetCurrentStock = adoRs.Fields("nm_current")
        End Select
    End If
    adoRs.Close
    
    
End Function

Sub addToCboItemCode()
Dim sqlChangeItem As String
Dim RsChangeItem As New Recordset
Dim i As Long
    
    sqlChangeItem = "select RTRIM(item_code)item_code, RTRIM(item_name)item_name, unit_cls, (select description from unit_cls uc where uc.unit_cls=item_master.unit_cls)  Unit_Desc, WH_Code " & _
        "from item_master where FinishGoodPart_Cls = '02'  and Item_Code not in (select Item_Code from BOM_Master where Parent_ItemCode ='" & parentItemCode & "') " & _
        "and use_endday > convert(char(8), getdate(), 112) order by item_name"
    Set RsChangeItem = Db.Execute(sqlChangeItem)
    
    With cboRepItem
        .clear
        .columnCount = 3
        .ColumnWidths = "80pt;80pt;30pt"
        .ListWidth = 210
        .ListRows = 10
        
        i = 0
        Do While Not RsChangeItem.EOF
            .AddItem
            .List(i, 0) = Trim(RsChangeItem("item_code"))
            .List(i, 1) = Trim(RsChangeItem("item_Name"))
            RsChangeItem.MoveNext
            i = i + 1
        Loop
    End With
End Sub

Private Sub cboRepItem_Change()
    If cboRepItem.MatchFound Or cboRepItem.Text <> "" Then
        With Grid1(0)
            .TextMatrix(Idx, bteColChangeItem) = cboRepItem.Text
        End With
    Else
         With Grid1(0)
            .TextMatrix(Idx, bteColChangeItem) = ""
        End With
    End If
End Sub

Private Sub cboRepItem_Click()
    If cboRepItem.MatchFound Or cboRepItem.Text <> "" Then
         With Grid1(0)
            .TextMatrix(Idx, bteColChangeItem) = cboRepItem.Text
        End With
    Else
         With Grid1(0)
            .TextMatrix(Idx, bteColChangeItem) = ""
        End With
    End If
End Sub

Private Sub cboRepItem_LostFocus()
    If cboRepItem.MatchFound Or cboRepItem.Text <> "" Then
        With Grid1(0)
            .TextMatrix(Idx, bteColChangeItem) = cboRepItem.Text
        End With
    Else
         With Grid1(0)
            .TextMatrix(Idx, bteColChangeItem) = ""
        End With
    End If
    
    cboRepItem.Visible = False
End Sub

