VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_part_supplyIn 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Parts (Material) Supply [Automatic]"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frm_part_supplyIn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtIndex 
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
      Height          =   225
      Left            =   6300
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   1
      Top             =   950
      Width           =   465
   End
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
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   12990
      TabIndex        =   34
      Top             =   270
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   767
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
      Left            =   11445
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9780
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
      Left            =   12630
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9780
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1560
      Left            =   315
      TabIndex        =   20
      Top             =   1320
      Width           =   14550
      Begin VB.CommandButton Command1 
         Caption         =   "Set"
         Height          =   405
         Left            =   13500
         TabIndex        =   39
         Top             =   210
         Width           =   525
      End
      Begin VB.TextBox TxtUnitQty 
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
         Height          =   225
         Left            =   14040
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.TextBox TxtUnitSet 
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
         Height          =   345
         Left            =   11970
         TabIndex        =   10
         Top             =   240
         Width           =   1425
      End
      Begin VB.TextBox lbl_location 
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
         Left            =   9060
         TabIndex        =   36
         Text            =   "lbl_location"
         Top             =   698
         Width           =   2460
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   330
         Left            =   9060
         TabIndex        =   9
         Top             =   1065
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
         Format          =   141230083
         CurrentDate     =   37867
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   330
         Left            =   2340
         TabIndex        =   6
         Top             =   1087
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
         CustomFormat    =   "MMM yyyy"
         Format          =   141230083
         UpDown          =   -1  'True
         CurrentDate     =   37867
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Set"
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
         Left            =   11040
         TabIndex        =   37
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label2 
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
         Left            =   225
         TabIndex        =   33
         Top             =   1155
         Width           =   540
      End
      Begin VB.Line Line3 
         X1              =   6525
         X2              =   7515
         Y1              =   1365
         Y2              =   1365
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
         Left            =   10740
         TabIndex        =   32
         Top             =   1140
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSForms.ComboBox cbo_material 
         Height          =   330
         Left            =   11970
         TabIndex        =   11
         Top             =   1065
         Visible         =   0   'False
         Width           =   1455
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2566;582"
         ShowDropButtonWhen=   2
         Value           =   "cbo_material"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- Name"
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
         Left            =   11655
         TabIndex        =   31
         Top             =   720
         Width           =   765
      End
      Begin VB.Line Line4 
         X1              =   12645
         X2              =   14310
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label lbl_machine 
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_machine"
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
         Left            =   12645
         TabIndex        =   30
         Top             =   735
         Width           =   1020
      End
      Begin MSForms.ComboBox cbo_MachineNo 
         Height          =   330
         Left            =   5640
         TabIndex        =   7
         Top             =   675
         Width           =   1995
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "3519;582"
         ShowDropButtonWhen=   2
         Value           =   "cbo_MachineNo"
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
         Left            =   4380
         TabIndex        =   29
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label8 
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
         Left            =   7800
         TabIndex        =   28
         Top             =   720
         Width           =   495
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
         Left            =   4380
         TabIndex        =   27
         Top             =   315
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
         Left            =   225
         TabIndex        =   25
         Top             =   315
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
         Left            =   225
         TabIndex        =   24
         Top             =   735
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
         Left            =   4380
         TabIndex        =   23
         Top             =   1155
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Date"
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
         TabIndex        =   3
         Top             =   1140
         Width           =   1050
      End
      Begin MSForms.ComboBox cbo_warehouse 
         Height          =   330
         Left            =   2340
         TabIndex        =   4
         Top             =   225
         Width           =   1500
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cbo_warehouse"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_location 
         Height          =   330
         Left            =   2340
         TabIndex        =   5
         Top             =   667
         Width           =   1500
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cbo_location"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_supply 
         Height          =   330
         Left            =   5640
         TabIndex        =   8
         Top             =   1065
         Width           =   780
         VariousPropertyBits=   746604571
         MaxLength       =   2
         DisplayStyle    =   3
         Size            =   "1376;582"
         ShowDropButtonWhen=   2
         Value           =   "cbo_supply"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lbl_warehouse 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_warehouse"
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
         Left            =   5655
         TabIndex        =   22
         Top             =   315
         Width           =   3210
      End
      Begin VB.Line Line1 
         X1              =   5655
         X2              =   8715
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line2 
         X1              =   9060
         X2              =   11610
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label lbl_supply 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_supply"
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
         Left            =   6540
         TabIndex        =   21
         Top             =   1125
         Width           =   900
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
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9780
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
      Left            =   13815
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9780
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   330
      TabIndex        =   16
      Top             =   9090
      Width           =   14595
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
         TabIndex        =   17
         Top             =   240
         Width           =   14235
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   5940
      Left            =   300
      TabIndex        =   26
      Top             =   2940
      Width           =   14565
      _cx             =   25691
      _cy             =   10477
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
   Begin MSForms.ComboBox CboWominNo 
      Height          =   330
      Left            =   1920
      TabIndex        =   42
      Top             =   900
      Width           =   3585
      VariousPropertyBits=   746604571
      MaxLength       =   50
      DisplayStyle    =   3
      Size            =   "6324;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "cbo_supplyNo"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Index"
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
      Left            =   5640
      TabIndex        =   41
      Top             =   960
      Width           =   495
   End
   Begin MSForms.ComboBox cboIndex 
      Height          =   315
      Left            =   6240
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   900
      Width           =   855
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "1508;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supply Req No."
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
      Left            =   540
      TabIndex        =   35
      Top             =   990
      Width           =   1320
   End
   Begin MSForms.ComboBox cbo_supplyNo 
      Height          =   330
      Left            =   1980
      TabIndex        =   0
      Top             =   900
      Visible         =   0   'False
      Width           =   2145
      VariousPropertyBits=   746604571
      MaxLength       =   13
      DisplayStyle    =   3
      Size            =   "3784;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "cbo_supplyNo"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      Left            =   1950
      TabIndex        =   19
      Top             =   9960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Parts (Material) Supply [Automatic]"
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
      Left            =   360
      TabIndex        =   18
      Top             =   225
      Width           =   14490
   End
End
Attribute VB_Name = "frm_part_supplyIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_part_supply As New ADODB.Recordset
Dim rs_warehouse As New ADODB.Recordset
Dim rs_trade_master As New ADODB.Recordset
Dim rs_item As New ADODB.Recordset

Dim idt_dateUp As Date
Dim Status As String
Dim cboEmpty As Boolean
Dim InvYear As String
Dim InvMonth As String
Dim LocationHasNoLine As Boolean
Dim RecordHasAdjustment As Boolean
Dim RecordHasChild As Boolean
Dim listError As Boolean

Dim bteColSelect As Byte
Dim bteColMatCode As Byte
Dim bteColDesc As Byte
Dim bteColQtyReq As Byte
Dim bteColUnitDesc As Byte
Dim bteColQty As Byte
'Dim bteColAlreadySupplyQty As Byte
'Dim bteColRemaining As Byte
Dim bteRemaining As Byte
Dim bteColUnitCls As Byte
Dim bteColSerialNo As Byte
Dim bteColLotNo As Byte
Dim bteColMatLotNo As Byte
Dim bteColRemark As Byte
Dim bteColReqSeqNo As Byte
Dim bteColSuplQty As Byte
Dim bteColRemQty As Byte
Dim bteColStock As Byte
Dim bteColshdQty As Byte
Dim bteColshdMatLotNo As Byte
Dim bteColshdUnit As Byte
Dim bteColshdItemControl As Byte
Dim bteColshdFromControl As Byte
Dim bteColshdSupplyDate As Byte
Dim bteColshdParentItem As Byte
Dim bteColshdUnitItem As Byte

Dim boolAdd As Boolean

Sub Header()

bteColSelect = 0
bteColMatCode = 1
bteColDesc = 2

bteColStock = 3

bteColQtyReq = 4

bteColSuplQty = 5
bteColRemQty = 6

bteColUnitDesc = 7
bteColQty = 8

bteColUnitCls = 9
bteColSerialNo = 10
bteColLotNo = 11
bteColMatLotNo = 12
bteColRemark = 13
bteColReqSeqNo = 14
bteColshdQty = 15
bteColshdMatLotNo = 16
bteColshdUnit = 17
bteColshdItemControl = 18
bteColshdFromControl = 19
bteColshdSupplyDate = 20
bteColshdParentItem = 21
bteColshdUnitItem = 22

With Grid1
    
.ColS = 23
.Rows = 1
.clear

.TextMatrix(0, bteColSelect) = "S"
.TextMatrix(0, bteColMatCode) = "Material Code"
.TextMatrix(0, bteColDesc) = "Description"
.TextMatrix(0, bteColQtyReq) = "Req Qty"

.TextMatrix(0, bteColSuplQty) = "Supply Qty"
.TextMatrix(0, bteColRemQty) = "Remain Qty"
.TextMatrix(0, bteColStock) = "Stock Qty"

.TextMatrix(0, bteColUnitDesc) = "Unit" 'Request Unit Desc
.TextMatrix(0, bteColQty) = "Qty"

.TextMatrix(0, bteColUnitCls) = "Unit" '#Request Unit Cls
.TextMatrix(0, bteColSerialNo) = "Serial No."
.TextMatrix(0, bteColLotNo) = "Lot No."
.TextMatrix(0, bteColMatLotNo) = "Material Lot No."
.TextMatrix(0, bteColRemark) = "Remarks"
.TextMatrix(0, bteColReqSeqNo) = "ReqSeqNo" '#Request Seq Cls
.TextMatrix(0, bteColshdQty) = "shdQty"
.TextMatrix(0, bteColshdMatLotNo) = "shdMaterialLotNo"
.TextMatrix(0, bteColshdUnit) = "shdUnit" '#Request Unit Cls
.TextMatrix(0, bteColshdItemControl) = "shdItemControl"
.TextMatrix(0, bteColshdFromControl) = "shdFromControl"
.TextMatrix(0, bteColshdSupplyDate) = "shdSupplyDate"
.TextMatrix(0, bteColshdParentItem) = "shdParentItemCode"
.TextMatrix(0, bteColshdUnitItem) = "shdUnitItemMaster" '#Item Master Unit Cls

.ColWidth(bteColSelect) = 250
.ColWidth(bteColMatCode) = 1500
.ColWidth(bteColDesc) = 3000
.ColWidth(bteColQtyReq) = 1000

.ColWidth(bteColSuplQty) = 1000
.ColWidth(bteColRemQty) = 1000
.ColWidth(bteColStock) = 1000

.ColWidth(bteColUnitDesc) = 500
.ColWidth(bteColQty) = 1500

.ColWidth(bteColUnitCls) = 500
.ColWidth(bteColSerialNo) = 1000
.ColWidth(bteColLotNo) = 1500
.ColWidth(bteColMatLotNo) = 1500
.ColWidth(bteColRemark) = 2000
.ColWidth(bteColReqSeqNo) = 1500
.ColWidth(bteColshdQty) = 1500
.ColWidth(bteColshdParentItem) = 1500

.ColHidden(bteColSelect) = True
.ColHidden(bteColSerialNo) = True '#Serial No.

.ColHidden(bteColMatLotNo) = True '#Material Lot No.
.ColHidden(bteColUnitDesc) = True
.ColHidden(bteColReqSeqNo) = True
.ColHidden(bteColshdQty) = True
.ColHidden(bteColshdMatLotNo) = True
.ColHidden(bteColshdUnit) = True
.ColHidden(bteColshdItemControl) = True
.ColHidden(bteColshdFromControl) = True
.ColHidden(bteColshdSupplyDate) = True
.ColHidden(bteColshdParentItem) = True
.ColHidden(bteColshdUnitItem) = True

.Cell(flexcpAlignment, 0, 0, 0, bteColLotNo) = flexAlignLeftCenter
.ColAlignment(bteColMatCode) = flexAlignLeftCenter
.ColAlignment(bteColDesc) = flexAlignLeftCenter
.ColAlignment(bteColQtyReq) = flexAlignRightCenter
.ColAlignment(bteColUnitDesc) = flexAlignLeftCenter
.ColAlignment(bteColQty) = flexAlignRightCenter

.ColAlignment(bteColUnitCls) = flexAlignLeftCenter
.ColAlignment(bteColSerialNo) = flexAlignLeftCenter
.ColAlignment(bteColLotNo) = flexAlignLeftCenter
.ColAlignment(bteColMatLotNo) = flexAlignLeftCenter
.ColAlignment(bteColRemark) = flexAlignLeftCenter
.ColAlignment(bteColReqSeqNo) = flexAlignLeftCenter

.EditMaxLength = 1

End With
End Sub

Private Sub comboDataChange(Status As Boolean)
cbo_supplyNo.DataChanged = Status
CboWominNo.DataChanged = Status
cbo_warehouse.DataChanged = Status
cbo_location.DataChanged = Status
cbo_MachineNo.DataChanged = Status
cbo_supply.DataChanged = Status
cbo_material.DataChanged = Status
End Sub

Private Sub setting_grid()
Dim sql_join As String

With Grid1

Call Header

RecordHasAdjustment = False
RecordHasChild = False


Dim rs_join As New ADODB.Recordset

'sql_join = " select  IM.stockControl_Cls as itemControl, PSM.FromWarehouse_code,isnull(WM.StockControl_cls ,'01')as FromControl,  PS.ChildSupply_date," & _
'            " PSD.*,IM.unit_cls,IM.Item_name,PS.consumption_Qty as psQty, " & _
'            " (select description from unit_cls uc where uc.unit_cls= IM.unit_cls) unit_desc, " & _
'            " (select description from unit_cls uc where uc.unit_cls= PSD.childunit_cls) requestUnit_desc " & _
'            " from partSupplyRequest_detail PSD " & _
'            " left join item_master IM on PSD.childitem_code=IM.item_code " & _
'            " left join partSupplyRequest_Master PSM on PSD.supplyRec_No=PSM.supplyRec_No " & _
'            " left join warehouse_Master wm on PSM.fromWarehouse_Code=WM.wh_code " & _
'            " left join part_supply PS on PSD.supplyRec_No=PS.supplyRec_No and  PSD.seq_No=PS.RecSeq_No " & _
'            " where PSD.supplyRec_No= '" & Trim(cbo_supplyNo) & "'"
 
If boolAdd = False Then

    sql_join = " select  IM.stockControl_Cls as itemControl,  " & vbCrLf & _
                          "     PSM.FromWarehouse_code,isnull(WM.StockControl_cls ,'01')as FromControl,   " & vbCrLf & _
                          "     PS.ChildSupply_date, PSD.*,IM.unit_cls,IM.Item_name,PS.consumption_Qty as psQty,   " & vbCrLf & _
                          "     (select description from unit_cls uc where uc.unit_cls= IM.unit_cls) unit_desc,   " & vbCrLf & _
                          "     (select description from unit_cls uc where uc.unit_cls= PSD.childunit_cls) requestUnit_desc,   " & vbCrLf & _
                          "     (Select NM_Current from Stock_Master SM where PSD.childitem_code=SM.Item_Code and PSM.FromWarehouse_code=SM.Warehouse_Code) CurrStock, " & vbCrLf & _
                          "     (Select isnull(Sum(isnull(consumption_Qty,0)),0) From Part_supply PS1 Where PS1.SupplyRec_No=PSD.supplyRec_No and PSD.Childitem_Code=PS1.ChildItem_Code) Supply_Qty, ISNULL(PSD.ReplacementItem_Code,'')ChangeItem_Code " & vbCrLf & _
                          " from partSupplyRequest_detail PSD  " & vbCrLf & _
                          " left join item_master IM on PSD.childitem_code=IM.item_code   " & vbCrLf & _
                          " left join partSupplyRequest_Master PSM on PSD.supplyRec_No=PSM.supplyRec_No   " & vbCrLf & _
                          " left join warehouse_Master wm on PSM.fromWarehouse_Code=WM.wh_code   "
    
    sql_join = sql_join + " left join part_supply PS on PSD.supplyRec_No=PS.supplyRec_No and  " & vbCrLf & _
                          " PSD.seq_No=PS.RecSeq_No  where PSD.supplyRec_No= '" & Trim(cbo_supplyNo) & "' and SJNo='" & Trim(txtIndex) & "'"

Else
'Request Pak toha ambil stock currentnya dari NM Current 8 Feb 2017

    sql_join = " select  IM.stockControl_Cls as itemControl,  " & vbCrLf & _
                          "     PSM.FromWarehouse_code,isnull(WM.StockControl_cls ,'01')as FromControl,   " & vbCrLf & _
                          "     Getdate() ChildSupply_date, PSD.*,IM.unit_cls,IM.Item_name,   " & vbCrLf & _
                          "     (Select sum(consumption_Qty) From Part_supply PS Where PS.supplyRec_No=PSD.supplyRec_No And PS.ChildItem_Code=PSD.ChildItem_Code and PSD.seq_No=PS.RecSeq_No) as psQty,   " & vbCrLf & _
                          "     (select description from unit_cls uc where uc.unit_cls= IM.unit_cls) unit_desc,   " & vbCrLf & _
                          "     (select description from unit_cls uc where uc.unit_cls= PSD.childunit_cls) requestUnit_desc,   " & vbCrLf & _
                          "     (Select NM_Current  from Stock_Master SM where PSD.childitem_code=SM.Item_Code and PSM.FromWarehouse_code=SM.Warehouse_Code) CurrStock, " & vbCrLf & _
                          "     (Select isnull(Sum(isnull(consumption_Qty,0)),0) From Part_supply PS1 Where PS1.SupplyRec_No=PSD.supplyRec_No and PSD.Childitem_Code=PS1.ChildItem_Code) Supply_Qty, ISNULL(PSD.ReplacementItem_Code,'')ChangeItem_Code " & vbCrLf & _
                          " from partSupplyRequest_detail PSD  " & vbCrLf & _
                          " left join item_master IM on PSD.childitem_code=IM.item_code   " & vbCrLf & _
                          " left join partSupplyRequest_Master PSM on PSD.supplyRec_No=PSM.supplyRec_No   " & vbCrLf & _
                          " left join warehouse_Master wm on PSM.fromWarehouse_Code=WM.wh_code   "
    
    sql_join = sql_join + "   where PSD.supplyRec_No= '" & Trim(cbo_supplyNo) & "' "

End If

If rs_join.State <> adStateClosed Then rs_join.Close
rs_join.Open sql_join, Db, adOpenForwardOnly, adLockReadOnly
If rs_join.EOF = False Or rs_join.BOF = False Then

    While rs_join.EOF = False
    
'         RecordHasAdjustment = FIFOcheckDownLevelAdjustment("", Trim(cbo_supplyNo))
'         RecordHasChild = FIFOcheckDownLevelReceipt("", Trim(cbo_supplyNo))
        
        With Grid1
            .AddItem ""
            .Cell(flexcpBackColor, .Rows - 1, bteColSelect) = vbWhite
            .Cell(flexcpBackColor, .Rows - 1, bteColQty) = vbWhite
            
            If Trim(rs_join!ChangeItem_Code) <> "" Then
                .TextMatrix(.Rows - 1, bteColMatCode) = Trim(rs_join!ChangeItem_Code)
                .TextMatrix(.Rows - 1, bteColDesc) = uf_GetItemDescription(Trim(rs_join!ChangeItem_Code))
            Else
                .TextMatrix(.Rows - 1, bteColMatCode) = Trim(rs_join!childitem_code)
                .TextMatrix(.Rows - 1, bteColDesc) = uf_GetItemDescription(Trim(rs_join!childitem_code))
            End If
            
            '.TextMatrix(.Rows - 1, bteColMatCode) = Trim(rs_join!childitem_code)
'            .TextMatrix(.Rows - 1, bteColDesc) = uf_GetItemDescription(Trim(rs_join!childitem_code))
            .TextMatrix(.Rows - 1, bteColQtyReq) = Format(Trim(rs_join!ChildRequirement_qty), gs_formatQtyBOM)
            .TextMatrix(.Rows - 1, bteColUnitDesc) = Format(Trim(rs_join!requestUnit_desc), gs_formatQtyBOM)
            
            .TextMatrix(.Rows - 1, bteColQty) = Format(Trim(rs_join!psQty), gs_formatQtyBOM)
            .TextMatrix(.Rows - 1, bteColStock) = Format(Trim(rs_join!CurrStock), gs_formatQtyBOM)
            .TextMatrix(.Rows - 1, bteColSuplQty) = Format(Trim(rs_join!Supply_Qty), gs_formatQtyBOM)
            
            .TextMatrix(.Rows - 1, bteColRemQty) = Format(Trim(rs_join!ChildRequirement_qty) - Trim(rs_join!Supply_Qty & ""), gs_formatQtyBOM)
                    
            If boolAdd = False Then
                .TextMatrix(.Rows - 1, bteColQty) = Format(Trim(rs_join!psQty), gs_formatQtyBOM)
            Else
                .TextMatrix(.Rows - 1, bteColQty) = Format(CDbl(.TextMatrix(.Rows - 1, bteColRemQty)), gs_formatQtyBOM)
            End If
            
            If .TextMatrix(.Rows - 1, bteColQty) >= 0 Then
                .Cell(flexcpBackColor, .Rows - 1, bteColQty) = vbYellow
            Else
                .Cell(flexcpBackColor, .Rows - 1, bteColQty) = vbWhite
            End If
            
            
            If Trim(.TextMatrix(.Rows - 1, bteColQty)) = "" Then
                If gb_AllowSetDefaultSupplyQty_MaterialSupplyAutomatic = True Then
                    .TextMatrix(.Rows - 1, bteColQty) = Format(Trim(rs_join!ChildRequirement_qty), gs_formatQtyBOM)
                    .Cell(flexcpBackColor, .Rows - 1, bteColQty) = vbYellow
                End If
            End If
            
            .TextMatrix(.Rows - 1, bteColUnitCls) = Trim(rs_join!requestUnit_desc)
            .TextMatrix(.Rows - 1, bteColSerialNo) = ""
            .TextMatrix(.Rows - 1, bteColLotNo) = Trim(rs_join!ChildLot_no)
            .TextMatrix(.Rows - 1, bteColMatLotNo) = Trim("")
            .TextMatrix(.Rows - 1, bteColRemark) = Trim(rs_join!Remarks)
            .TextMatrix(.Rows - 1, bteColReqSeqNo) = Trim(rs_join!Seq_no)
            .TextMatrix(.Rows - 1, bteColshdQty) = Format(Trim(rs_join!psQty), gs_formatQtyBOM)
            .TextMatrix(.Rows - 1, bteColshdMatLotNo) = Trim("")
            .TextMatrix(.Rows - 1, bteColshdUnit) = Trim(rs_join!childunit_cls)
            .TextMatrix(.Rows - 1, bteColshdItemControl) = Trim(rs_join!itemControl)
            .TextMatrix(.Rows - 1, bteColshdFromControl) = Trim(rs_join!FromControl)
            .TextMatrix(.Rows - 1, bteColshdSupplyDate) = Format(rs_join!childsupply_date, "yyyy-MM-dd")
            .TextMatrix(.Rows - 1, bteColshdParentItem) = Trim(rs_join!parentItem_code)
            .TextMatrix(.Rows - 1, bteColshdUnitItem) = Trim(rs_join!Unit_cls)
        End With
        rs_join.MoveNext
    Wend
    
End If
Status = "insertnew"
rs_join.Close

End With

End Sub

Private Sub cbo_location_Click()

Call set_line

If cbo_location.DataChanged = True Then CboWominNo = "": cbo_supplyNo = "": Call setCboSupplyNo
cbo_location.DataChanged = True

If cbo_location.ListIndex <> -1 Then
    lbl_location = cbo_location.List(cbo_location.ListIndex, 1)
Else
    lbl_location = ""
End If

lbl_pesan = validCombo
If Trim(lbl_pesan) <> "" Then
    If cboEmpty = True Then lbl_pesan = ""
    clearGrid
    Exit Sub
End If

lbl_pesan = ""

End Sub


Private Sub cbo_location_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

If KeyCode = vbKeyBack Then lbl_location = "": clearGrid

If KeyCode = 13 Then
    lbl_pesan = ""
    Call set_line
    
    If cbo_location.DataChanged = True Then cbo_supplyNo = "": Call setCboSupplyNo
    cbo_location.DataChanged = True
    
    lbl_pesan = validCombo
    If Trim(lbl_pesan) <> "" Then
        If cboEmpty = True Then lbl_pesan = ""
        clearGrid
        Exit Sub
    End If
End If

End Sub

Private Sub cbo_location_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


Private Sub cbo_MachineNo_Click()

If cbo_MachineNo.DataChanged = True Then cbo_supplyNo = "": Call setCboSupplyNo
cbo_MachineNo.DataChanged = True

If cbo_MachineNo.ListIndex <> -1 Then
    lbl_machine.Caption = cbo_MachineNo.List(cbo_MachineNo.ListIndex, 1)
Else
    lbl_machine.Caption = ""
End If

lbl_pesan = validCombo
If Trim(lbl_pesan) <> "" Then
    If cboEmpty = True Then lbl_pesan = ""
    clearGrid
    Exit Sub
End If

clearGrid
lbl_pesan = ""

End Sub

Private Sub cbo_MachineNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyBack Then lbl_location = "": clearGrid

If KeyCode = 13 Then
    
    If cbo_MachineNo.DataChanged = True Then cbo_supplyNo = "": Call setCboSupplyNo
    cbo_MachineNo.DataChanged = True
    
    If cbo_MachineNo.ListIndex <> -1 Then
    lbl_machine.Caption = cbo_MachineNo.List(cbo_MachineNo.ListIndex, 1)
    Else
        lbl_machine.Caption = ""
    End If
    
    lbl_pesan = validCombo
    If Trim(lbl_pesan) <> "" Then
        If cboEmpty = True Then lbl_pesan = ""
        clearGrid
        Exit Sub
    End If
End If

clearGrid

End Sub

Private Sub cbo_MachineNo_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbo_material_Click()
If cbo_material.DataChanged = True Then cbo_supplyNo = "": Call setCboSupplyNo
cbo_material.DataChanged = True
Call clearGrid
End Sub

Private Sub cbo_material_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub cbo_supply_Change()

If cbo_supply.DataChanged = True Then cbo_supplyNo = "": 'Call setCboSupplyNo
cbo_supply.DataChanged = True

lbl_pesan = validCombo
If Trim(lbl_pesan) <> "" Then
    If cboEmpty = True Then lbl_pesan = ""
    clearGrid
    Exit Sub
End If
Call clearGrid
End Sub

Private Sub cbo_supply_Click()
lbl_supply.Caption = cbo_supply.List(cbo_supply.ListIndex, 1)
End Sub

Private Sub cbo_supply_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyBack Then lbl_supply = ""

If KeyCode = 13 Then
    If cbo_supply.DataChanged = True Then cbo_supplyNo = "": Call setCboSupplyNo
    cbo_supply.DataChanged = True
    
    lbl_pesan = validCombo
    If Trim(lbl_pesan) <> "" Then
        If cboEmpty = True Then lbl_pesan = ""
        clearGrid
        Exit Sub
    End If
End If
End Sub

Private Sub cbo_supply_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


Private Sub cbo_supplyNo_Change()
Call clearGrid
End Sub

Private Sub GetIndex()
Dim rs_Index As New ADODB.Recordset
Dim strSQL As String

strSQL = "Select cast(isnull(SJNo,0) as int) From Part_Supply Where SupplyRec_No='" & Trim(cbo_supplyNo) & "' Group By cast(isnull(SJNo,0) as int)"
rs_Index.Open strSQL, Db, adOpenForwardOnly, adLockReadOnly

cboIndex.clear

Do While Not rs_Index.EOF
    If rs_Index(0) <> 0 Then cboIndex.AddItem rs_Index(0)
    rs_Index.MoveNext
Loop
rs_Index.Close

strSQL = "Select isnull(Max(cast(isnull(SJNo,0) as int)),0) + 1 From Part_Supply Where SupplyRec_No='" & Trim(cbo_supplyNo) & "'"
rs_Index.Open strSQL, Db, adOpenForwardOnly, adLockReadOnly

txtIndex = rs_Index(0)
rs_Index.Close

boolAdd = True

End Sub

Private Sub cbo_supplyNo_Click()
If cbo_supplyNo <> "" Then Call GetIndex

End Sub

Private Sub cbo_supplyNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then _
Call clearGrid
End Sub

Private Sub cbo_update_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub cbo_warehouse_Click()

lbl_warehouse.Caption = cbo_warehouse.List(cbo_warehouse.ListIndex, 1)

If cbo_warehouse.DataChanged = True Then cbo_supplyNo = "": Call setCboSupplyNo
cbo_warehouse.DataChanged = True

lbl_pesan = validCombo
If Trim(lbl_pesan) <> "" Then
    If cboEmpty = True Then lbl_pesan = ""
    clearGrid
    Exit Sub
End If

Call Header
lbl_pesan = ""

End Sub

Private Sub setCboSupplyNo()
    

'------------------ cbo supplyNo-------------------

cbo_supplyNo.clear
cbo_supplyNo.columnCount = 1
cbo_supplyNo.TextColumn = 1

CboWominNo.clear
CboWominNo.columnCount = 2
CboWominNo.TextColumn = 1

Dim rsSup As New ADODB.Recordset

' For Kawai, Do Not Show Complete WomIn No or Request Number

sql = " Select * From (Select psm.*, DP.Qty-(Select isnull(sum(Qty),0) From Part_Receipt PR Where PR.DailySeq_No=DP.seq_No and PR.Item_Code=DP.Item_Code And REceipt_Cls='P1') QtyRemaining,  " & vbCrLf & _
        " (Select Sum(isnull(ChildRequirement_Qty,0)) RequestQty From PartSupplyRequest_DEtail Where SupplyRec_No=PSM.SupplyRec_No Group By SupplyRec_No) -" & vbCrLf & _
        " (Select Sum(isnull(Consumption_Qty,0)) SupQty From Part_Supply Where SupplyRec_No=PSM.SupplyRec_No Group By SupplyRec_No) WomIn_Rem " & vbCrLf & _
        "   From PartSupplyRequest_Master PSM" & _
        " inner join Daily_Production DP on PSM.REquest_Cls=DP.Request_Cls  Where "

sql = "Select PSM.* From PartSupplyRequest_Master PSM " & vbCrLf & _
      " inner join Daily_Production DP on PSM.REquest_Cls=DP.Request_Cls  Where " & vbCrLf

If Trim(cbo_warehouse) <> "" Then sql = sql + " fromWarehouse_code = '" & Trim(cbo_warehouse) & "' and" & vbCrLf
If Trim(cbo_location) <> "" Then sql = sql + " toWarehouse_code = '" & Trim(cbo_location) & "' and" & vbCrLf
If Trim(cbo_supply) <> "" Then sql = sql + " supply_cls = '" & Trim(cbo_supply) & "' and" & vbCrLf

If Trim(Format(date1, "yyyy-MM-dd")) <> "" Then _
    sql = sql + "  month(childSupply_date) = '" & Trim(Format(date1, "MM")) & "' and year(childSupply_date) = '" & Trim(Format(date1, "yyyy")) & "' and" & vbCrLf
If Trim(cbo_MachineNo) <> "" Then sql = sql + " machine_No = '" & Trim(cbo_MachineNo) & "' and" & vbCrLf
If Trim(cbo_material) <> "" Then sql = sql + " material_cls= '" & IIf(Trim(cbo_material) = "Resin", 1, 0) & "' and" & vbCrLf
       
If Right(Trim(sql), 3) = "and" Then sql = Left(Trim(sql), Len(Trim(sql)) - 3)
If Right(Trim(sql), bteColQty) = "where" Then sql = Left(Trim(sql), Len(Trim(sql)) - 5)

sql = sql & " 'A'='A' --) WomIn --Where (WomIn_Rem is null) or (WomIn_Rem>0) " & vbCrLf & _
            " Order By SupplyRec_No"
            
rsSup.Open sql, Db

i = 0

If rsSup.EOF = False Or rsSup.BOF = False Then
    rsSup.MoveFirst
    While rsSup.EOF = False
        cbo_supplyNo.AddItem ""
        cbo_supplyNo.List(i, 0) = Trim(rsSup!supplyRec_No)
        
        CboWominNo.AddItem ""
        CboWominNo.List(i, 1) = Trim(rsSup!supplyRec_No) & ""
        CboWominNo.List(i, 0) = Trim(rsSup!WomIn_No) & ""
        
        rsSup.MoveNext
        i = i + 1
    Wend
    cbo_supplyNo.ColumnWidths = "100 pt"
    cbo_supplyNo.ListWidth = 100
    
    CboWominNo.ColumnWidths = "200 pt,100 pt"
    CboWominNo.ListWidth = 200
    
End If

rsSup.Close

End Sub


Private Sub cbo_warehouse_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

If KeyCode = vbKeyBack Then lbl_warehouse = "": clearGrid
If KeyCode = 13 Then

lbl_pesan = ""

If cbo_warehouse.DataChanged = True Then cbo_supplyNo = "": Call setCboSupplyNo
cbo_warehouse.DataChanged = True

lbl_pesan = validCombo
If Trim(lbl_pesan) <> "" Then
    If cboEmpty = True Then lbl_pesan = ""
    clearGrid
    Exit Sub
End If

End If

End Sub

Sub clearGrid()
Grid1.clear
Grid1.Rows = 1
Call Header
End Sub

Private Sub cbo_warehouse_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


Private Sub cboIndex_Click()
txtIndex = cboIndex
boolAdd = False
Call clearGrid

cmd_update.SetFocus

End Sub

Private Sub CboWominNo_Click()
    If CboWominNo.ListIndex < 0 Then
        CboWominNo = ""
        cbo_supplyNo = ""
    Else
        cbo_supplyNo = CboWominNo.Column(1)
    End If
End Sub

Private Sub cmd_Cancel_Click()

lbl_pesan = ""
Call setting_grid
End Sub

Private Sub cmd_clear_Click()
comboDataChange (False)
DTPicker3.Value = Format(Date, "dd MMM yyyy")

lbl_warehouse.Caption = ""
lbl_location = ""
lbl_pesan.Caption = ""
lbl_supply.Caption = ""
lbl_machine.Caption = ""

cbo_location = ""
cbo_warehouse = ""
cbo_MachineNo = ""
cbo_supply = ""
cbo_supplyNo = ""
cbo_material = ""

Call setting_grid
comboDataChange (True)

Call setCboSupplyNo
End Sub

Private Sub cmd_sub_menu_Click()
frmMainMenu.Show
Unload Me
End Sub

Private Sub Cmd_Submit_Click()

Dim maxDateFroms As String
Dim maxDateTos As String
Dim l_update_stock As Double
Dim gd_SupplyQtyAfterUpdate As Double
Dim gd_SupplyQtyBeforeUpdate As Double
Dim gs_SupplySeqNo As String, strSQL As String
Dim sql_del As String, KeySeqNo As Double
Dim rsseqno As New ADODB.Recordset

If hakUpdate(Me.Name) = 0 Then _
lbl_pesan = DisplayMsg(3008): Exit Sub

Dim ls_SupplySeqNo As String
Dim ll_updatedRecord As Long
Dim rsSTHUpdU As New ADODB.Recordset
Dim sqlSTHUpdU As String
Dim rsSTHUpdS As New ADODB.Recordset
Dim sqlSTHUpdS As String
Dim i As Integer
Dim ls_ClosingMonth As String
Dim ls_ClosingYear As String
Dim lcls_clsProc As New ClsProc
Dim IntRedColor As Integer

On Error GoTo errHandler


For i = 1 To Grid1.Rows - 1
    If Grid1.Cell(flexcpBackColor, i, bteColQty) = vbRed Then
        lbl_pesan = "[0000]-Invalid Qty Entry (Red Cell)"
        Exit Sub
    End If
    IntRedColor = IntRedColor + 1
Next
'#Init Last Closing
ls_ClosingMonth = uf_GetLastClosing("month")
ls_ClosingYear = uf_GetLastClosing("year")

ll_updatedRecord = 0
listError = False

'# Check Valid Date Range of Inventory Control
lbl_pesan = up_ValidateDateRange(DTPicker3.Value, True)
If lbl_pesan.Caption <> "" Then Call setting_grid: Exit Sub

If Month(date1) <> Month(DTPicker3) Then
    If MsgBox("Request Period : " & Format(date1, "MMM yyyy") & ", Supply Date : " & Format(DTPicker3, "dd-MMM-yyyy") & "." & vbNewLine & _
        "Are you sure want to process?", vbExclamation + vbYesNo, "Confirm") = vbNo Then
            Call setting_grid
            Exit Sub
    End If
End If

Me.MousePointer = vbHourglass

    Dim rsSbm As New ADODB.Recordset
    Dim sqlSbm As String
    
    For i = 1 To Grid1.Rows - 1
        '#Init Control Cls
        ItemControlCls = Grid1.TextMatrix(i, bteColshdItemControl)
        FromControlCls = Grid1.TextMatrix(i, bteColshdFromControl)
        
        '# if SupplyQty or MaterialLotNo has changed
        If (Trim(Grid1.TextMatrix(i, bteColQty)) <> Trim(Grid1.TextMatrix(i, bteColshdQty))) Or (Format(DTPicker3, "yyyy-MM-dd") <> Trim(Grid1.TextMatrix(i, bteColshdSupplyDate)) Or boolAdd = True) Then
            
            '# Updated record Counter
             If (Trim(Grid1.TextMatrix(i, bteColQty)) <> Trim(Grid1.TextMatrix(i, bteColshdQty))) Then ll_updatedRecord = ll_updatedRecord + 1
            
            sqlSbm = "select * from part_supply where recSeq_No='" & Trim(Grid1.TextMatrix(i, bteColReqSeqNo)) & "' and SupplyRec_No='" & Trim(cbo_supplyNo) & "'" & vbLf & _
                " And SJNo='" & Trim(txtIndex) & "'"
            
            If rsSbm.State <> adStateClosed Then rsSbm.Close
            rsSbm.Open sqlSbm, Db, adOpenKeyset, adLockOptimistic
            
            If rsSbm.EOF = False Then '# Update or Delete data in part Supply
                If Val(Grid1.TextMatrix(i, bteColQty)) = 0 Then '# if Qty is Blank then delete data in part Supply
                    Grid1.TextMatrix(i, bteColQty) = ""
                    Dim deleteErr As Boolean

                    '# refresh Indicate the error data
                    Call gridColor("peach", i)
                    
                    '#Check if item influence the stock or not
                    If ItemControlCls = "01" Then
                        
                        '# Check Data Last Update from part_supply & part_receipt
                        maxDateFroms = uf_GetSupplyLastUpdate(Trim(cbo_warehouse), Trim(Grid1.TextMatrix(i, bteColMatCode)), Trim(Grid1.TextMatrix(i, bteColLotNo)))
                        maxDateTos = uf_GetSupplyLastUpdate(Trim(cbo_location), Trim(Grid1.TextMatrix(i, bteColMatCode)), Trim(Grid1.TextMatrix(i, bteColLotNo)))
                
                        Db.BeginTrans
                                                
                        '#Init Delete Stock
                        l_update_stock = lcls_clsProc.nilConvertUnit(CDbl(Grid1.TextMatrix(i, bteColshdQty)), Trim(Grid1.TextMatrix(i, bteColshdUnit)), Trim(Grid1.TextMatrix(i, bteColshdUnitItem)))
                        
                        '# Delete data from stock Master
                        
                        gd_SupplyQtyAfterUpdate = 0
                        gd_SupplyQtyBeforeUpdate = l_update_stock
                    
                         gs_SupplySeqNo = Trim(rsSbm!Seq_no)
                         Call up_UpdateStockMaster(Format(DTPicker3.Value, "yyyy-MM-dd"), _
                                    ls_ClosingMonth, ls_ClosingYear, Trim(cbo_warehouse), _
                                    Trim(cbo_location), Trim(Grid1.TextMatrix(i, bteColMatCode)), 0 - l_update_stock, _
                                    Trim(cbo_supply), Trim(ToControlCls), "", _
                                    "D", (maxDateFroms), (maxDateTos), False, False)
                                                
                         '# Delete data from part Supply
                         sql = "delete from part_supply where recSeq_No='" & Trim(Grid1.TextMatrix(i, bteColReqSeqNo)) & "' and SupplyRec_No='" & Trim(cbo_supplyNo) & "' " & _
                                " And SJNo='" & Trim(txtIndex) & "'"
                                
                         Db.Execute sql
                         
                        '# Erase data from stockMaster ( base on From WareHouse code)
                        Call up_EraseBlankDataInStockMaster(Trim(cbo_warehouse), Trim(Grid1.TextMatrix(i, bteColMatCode)), Trim(Grid1.TextMatrix(i, bteColLotNo)))
                        
                        '# Erase data from stockMaster ( base on To WareHouse code)
                        Call up_EraseBlankDataInStockMaster(Trim(cbo_location), Trim(Grid1.TextMatrix(i, bteColMatCode)), Trim(Grid1.TextMatrix(i, bteColLotNo)))
                                                                                            
                        Db.CommitTrans
                        
                    End If
                    Grid1.TextMatrix(i, bteColshdQty) = ""
                Else '# if Qty is not Blank then update data in part Supply
                                                                                                                                            
                    '#Init UpdateQty
                    Dim ld_UpdateQty As Double '#Qty that already converted to its base Unit
                    
                    ld_UpdateQty = lcls_clsProc.nilConvertUnit(CDbl(Grid1.TextMatrix(i, bteColQty)), Trim(Grid1.TextMatrix(i, bteColshdUnit)), Trim(Grid1.TextMatrix(i, bteColshdUnitItem)))
                    gd_SupplyQtyAfterUpdate = ld_UpdateQty
                    gd_SupplyQtyBeforeUpdate = lcls_clsProc.nilConvertUnit(CDbl(Grid1.TextMatrix(i, bteColshdQty)), Trim(Grid1.TextMatrix(i, bteColshdUnit)), Trim(Grid1.TextMatrix(i, bteColshdUnitItem)))
                    
                    '# refresh Indicate the error data
                    Call gridColor("peach", i)
                    
                    Db.BeginTrans
                    '#Update Data in part Supply
                    sql_del = "update part_supply with (updlock) " & _
                                "set childrequirement_qty = " & ld_UpdateQty & ", " & _
                                "Consumption_qty = " & CDbl(Trim(Grid1.TextMatrix(i, bteColQty))) & ", " & _
                                "from_address = '', " & _
                                "Last_Update = getdate(), " & _
                                "Last_User = '" & userLogin & "' " & _
                                "where RecSeq_no='" & Trim(Grid1.TextMatrix(i, bteColReqSeqNo)) & "' and supplyRec_No='" & Trim(cbo_supplyNo) & "' " & _
                                " And SJNo='" & txtIndex & "' "
                    
                    Db.Execute (sql_del)
                    
                    sql_del = "select * from part_supply " & _
                                "  where RecSeq_no='" & Trim(Grid1.TextMatrix(i, bteColReqSeqNo)) & "' " & _
                                "  and supplyRec_No='" & Trim(cbo_supplyNo) & "' " & _
                                " And SJNo='" & txtIndex & "' "
                                
                    
                    '#Check Part Supply Seq_No for allocate process
                    Dim lrs_update As New ADODB.Recordset
                    If lrs_update.State <> adStateClosed Then lrs_update.Close
                    lrs_update.Open sql_del, Db, adOpenKeyset, adLockOptimistic
                    ls_SupplySeqNo = Trim(lrs_update!Seq_no)
                    lrs_update.Close
                    
                    '#Init Update Stock Qty
                    l_update_stock = lcls_clsProc.nilConvertUnit(CDbl(Grid1.TextMatrix(i, bteColshdQty)) - CDbl(Grid1.TextMatrix(i, bteColQty)), Trim(Grid1.TextMatrix(i, bteColshdUnit)), Trim(Grid1.TextMatrix(i, bteColshdUnitItem)))
                    
                    '#Check if item influence the stock or not
                    If ItemControlCls = "01" Then
                        '# Update data in stock Master
                        gs_SupplySeqNo = ls_SupplySeqNo
                        Call up_UpdateStockMaster(Format(DTPicker3.Value, "yyyy-MM-dd"), _
                            ls_ClosingMonth, ls_ClosingYear, Trim(cbo_warehouse), _
                            Trim(cbo_location), Trim(Grid1.TextMatrix(i, bteColMatCode)), 0 - l_update_stock, _
                            Trim(cbo_supply), Trim(ToControlCls), "", _
                            "U", "", "", False, False)
                    End If
                    
                    Db.CommitTrans
                                                                        
                    '# Adjust Qty
                    Grid1.TextMatrix(i, bteColshdQty) = Trim(Grid1.TextMatrix(i, bteColQty))
                   
               End If
               
        ElseIf CDbl(Trim(Grid1.TextMatrix(i, bteColQty))) <> 0 Then 'And Grid1.Cell(flexcpBackColor, i, bteColQty) = vbYellow Then  '# insert data to part Supply
            
            Call gridColor("peach", i) '# refresh Indicate the error data
                        
                Dim InsertQty As Double
                InsertQty = lcls_clsProc.nilConvertUnit(CDbl(Grid1.TextMatrix(i, bteColQty)), Trim(Grid1.TextMatrix(i, bteColshdUnit)), Trim(Grid1.TextMatrix(i, bteColshdUnitItem)))
                
                
                Set rsseqno = Db.Execute("Select isnull(max(seq_No),0)+1 From Part_Supply")
                KeySeqNo = rsseqno(0)
                rsseqno.Close
                
                Db.BeginTrans
                
            '# insert data to part Supply
                rsSbm.AddNew
                'rsSbm!Seq_No = KeySeqNo
                rsSbm!supplyRec_No = Trim(cbo_supplyNo)
                rsSbm!RecSeq_no = Trim(Grid1.TextMatrix(i, bteColReqSeqNo))
                rsSbm!FromWarehouse_Code = Trim(cbo_warehouse)
                rsSbm!towarehouse_code = Trim(cbo_location)
                rsSbm!childsupply_date = Format(DTPicker3, "yyyy-MM-dd")
                rsSbm!childitem_code = Trim(Grid1.TextMatrix(i, bteColMatCode))
                rsSbm!Lot_no = Trim(Grid1.TextMatrix(i, bteColLotNo)) '#Tadinya ChildLot No
                rsSbm!supply_cls = Trim(cbo_supply)
                rsSbm!ChildRequirement_qty = InsertQty
                rsSbm!consumption_Qty = Trim(Grid1.TextMatrix(i, bteColQty))
                rsSbm!childunit_cls = Trim(Grid1.TextMatrix(i, bteColshdUnit))
                rsSbm!Remarks = Trim(Grid1.TextMatrix(i, bteColRemark))
                rsSbm!do_no = ""
                rsSbm!SJNo = Trim(txtIndex)
                rsSbm!from_address = ""
                rsSbm!parentItem_code = Trim(Grid1.TextMatrix(i, bteColshdParentItem))
                rsSbm!Last_Update = Now
                rsSbm!last_user = userLogin
                rsSbm.update
                
                '#Init Supply Seq No for allocate Process
                ls_SupplySeqNo = Trim(rsSbm!Seq_no)
                                
                '#Adjustment for Qty
                Grid1.TextMatrix(i, bteColshdQty) = Trim(Grid1.TextMatrix(i, bteColQty))
                
                '#Adjustment for Supply Date
                Grid1.TextMatrix(i, bteColshdSupplyDate) = Format(DTPicker3, "yyyy-MM-dd")
            
                '#if Item influence the stock then insert data into stockMaster
                If ItemControlCls = "01" Then
                    '# Insert data into stock Master
                    gs_SupplySeqNo = ls_SupplySeqNo
                    gd_SupplyQtyAfterUpdate = InsertQty
                    gd_SupplyQtyBeforeUpdate = 0
                    Call up_UpdateStockMaster(Format(DTPicker3.Value, "yyyy-MM-dd"), _
                    ls_ClosingMonth, ls_ClosingYear, Trim(cbo_warehouse), _
                    Trim(cbo_location), Trim(Grid1.TextMatrix(i, bteColMatCode)), InsertQty, _
                    Trim(cbo_supply), Trim(ToControlCls), "", _
                    "I", "", "", False, False)
                End If
                Db.CommitTrans
skipInsert:
            End If
        End If
        Grid1.Cell(flexcpBackColor, i, bteColQty) = vbWhite
    Next
    
Me.MousePointer = vbDefault
'Call clearGrid
If listError = False Then
    lbl_pesan = IIf(ll_updatedRecord <> 0, DisplayMsg(1101), "")
Else
    lbl_pesan = DisplayMsg(8094) '"Some data are not avaiable in stock !"
End If

Call cmd_clear_Click

Exit Sub
ErrExit:
    Db.RollbackTrans
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    lbl_pesan.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit


End Sub
Private Sub gridColor(color As String, Index As Integer)
If color = "peach" Then
    Grid1.Cell(flexcpBackColor, Index, 0, Index, Grid1.ColS - 1) = &H80000018  '# refresh Indicate the error data
Else
    Grid1.Cell(flexcpBackColor, Index, 0, Index, Grid1.ColS - 1) = vbRed  '# Indicate the error data
End If
Grid1.Cell(flexcpBackColor, Index, bteColQty, Index, bteColQty) = vbWhite
End Sub

Private Sub cmd_update_Click()

Dim strSQL As String

comboDataChange (False)
'# if Search

Dim rsUpd As New ADODB.Recordset
' Supply with Kanban System ( 1 Request --> more than 1 Supply / Partial Supply ) - 20090421

'StrSql = " Select PSM.*,DP.Qty QtySet From PartSupplyRequest_Master PSM " & _
'        " inner join Daily_Production DP on PSM.REquest_Cls=DP.Request_Cls " & _
'        " where supplyRec_No='" & Trim(cbo_supplyNo) & "'"
        
    strSQL = "  Select PSM.*,DP.Qty QtySet, " & vbCrLf & _
                      "     (Select isnull(sum(Qty),0) From Part_Receipt PR Where PR.DailySeq_No=DP.seq_No  " & vbCrLf & _
                      "         and PR.Item_Code=DP.Item_Code And REceipt_Cls='P1') QtySupply " & vbCrLf & _
                      "  From PartSupplyRequest_Master PSM " & vbCrLf & _
                      "   inner join Daily_Production DP on PSM.REquest_Cls=DP.Request_Cls  " & vbCrLf & _
                      "  where supplyRec_No='" & Trim(cbo_supplyNo) & "' "
        
rsUpd.Open strSQL, Db, adOpenForwardOnly, adLockReadOnly

If rsUpd.EOF Then
    rsUpd.Close
    Call clearGrid
    lbl_pesan = DisplayMsg(8093) '"Data with this Supply No. not found !"
    Exit Sub
End If

rsUpd.MoveFirst

cbo_warehouse.DataChanged = False
cbo_warehouse = Trim(rsUpd!FromWarehouse_Code)
cbo_warehouse.DataChanged = True
cbo_location.DataChanged = False
cbo_location = Trim(rsUpd!towarehouse_code)
cbo_location.DataChanged = True
cbo_MachineNo.DataChanged = False
cbo_MachineNo = Trim(rsUpd!Machine_no)
cbo_MachineNo.DataChanged = True
'TxtUnitSet = rsUpd!QtySet - rsUpd!QtySupply
TxtUnitQty = rsUpd!QtySet

Dim rsSup As New ADODB.Recordset
rsSup.Open "select * from part_Supply where supplyRec_No='" & Trim(cbo_supplyNo) & _
            "' and SJNo='" & Trim(txtIndex) & "'", Db, adOpenKeyset, adLockOptimistic
            
If rsSup.EOF Then

Else
    DTPicker3 = Format(rsSup!childsupply_date, "dd MMM yyyy")
    date1 = Format(Right(Trim(cbo_supplyNo), 7), "MMM yyyy")
End If
rsSup.Close

cbo_supply.DataChanged = False
cbo_supply = Trim(rsUpd!supply_cls)
cbo_supply.DataChanged = True
cbo_material.DataChanged = False

cbo_material.DataChanged = True
Call setCboSupplyNo
cbo_supplyNo = Trim(rsUpd!supplyRec_No)
CboWominNo.ListIndex = cbo_supplyNo.ListIndex
rsUpd.Close

Call setting_grid

comboDataChange (True)
End Sub

Private Sub Command1_Click()
Dim TmpSet As Double
Dim lngRow As Integer

lbl_pesan = ""
On Error GoTo handler

If TxtUnitSet.Text <> "" And TxtUnitQty.Text <> "" Then
    If CDbl(TxtUnitSet) > CDbl(TxtUnitQty) Then
        lbl_pesan = "[0000]-Supply Qty greater than Request !"
        Exit Sub
    End If

TmpSet = CDbl(TxtUnitSet) / CDbl(TxtUnitQty)

End If

lngRow = 1
Do
    If lngRow > Grid1.Rows - 1 Then Exit Do
    Grid1.TextMatrix(lngRow, bteColQty) = Format(CDbl(Grid1.TextMatrix(lngRow, bteColQtyReq)) * TmpSet, gs_formatQtyBOM)
    
    If CDbl(Grid1.TextMatrix(lngRow, bteColQty)) > CDbl(Grid1.TextMatrix(lngRow, bteColRemQty)) Then
        Grid1.Cell(flexcpBackColor, lngRow, bteColQty) = vbRed
    Else
        Grid1.Cell(flexcpBackColor, lngRow, bteColQty) = vbYellow
    End If
    lngRow = lngRow + 1
Loop

Exit Sub


handler:
 lbl_pesan.Caption = err.Description
 err.clear

End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    lbl_pesan.Caption = ErrMsg
End If
End Sub



Private Sub date1_Change()
Dim sql As String, RsMonth As New ADODB.Recordset, BMonth, BYear, BTgl As String

If Format(date1.Value, "MM") < Format(idt_dateUp, "MM") And Val(Format(date1.Value, "MM")) = 1 And Val(Format(idt_dateUp, "MM")) = 12 Then _
            date1.Year = date1.Year + 1: GoTo pass
    If Format(date1.Value, "MM") > Format(idt_dateUp, "MM") And Val(Format(date1.Value, "MM")) = 12 And Val(Format(idt_dateUp, "MM")) = 1 Then _
            date1.Year = date1.Year - 1
pass:
    idt_dateUp = Format(date1.Value, "dd MMM yyyy")


Call clearGrid
Call setCboSupplyNo
cbo_supplyNo = ""
End Sub

Private Sub DTPicker3_Change()

lbl_pesan = up_ValidateDateRange(Format(DTPicker3, "yyyy-MM-dd"), True)
If Trim(lbl_pesan) <> "" Then
    clearGrid
    Exit Sub
End If

End Sub

Function validCombo() As String

Dim j As Integer

'# cek combo warehouse
If Trim(cbo_warehouse) = "" Then
    validCombo = DisplayMsg(1042) ' "Please input Warehouse Code !"
    cboEmpty = True: Exit Function
End If
cboEmpty = False

j = 0
For i = 0 To cbo_warehouse.ListCount - 1
    If UCase(Trim(cbo_warehouse)) = UCase(Trim(cbo_warehouse.List(i, 0))) Then
        cbo_warehouse.Text = cbo_warehouse.List(i, 0)
        lbl_warehouse.Caption = cbo_warehouse.List(i, 1)
        lbl_pesan.Caption = ""
        j = 1
        Exit For
    End If
Next

If j = 0 Then
    lbl_warehouse.Caption = "": validCombo = DisplayMsg(4018) ' "Invalid warehouse code !"
    Exit Function
End If

'# cek combo Location
If Trim(cbo_location) = "" Then
    validCombo = DisplayMsg(1031) ' "Please input Location Code !"
    cboEmpty = True: Exit Function
End If
cboEmpty = False

j = 0
For i = 0 To cbo_location.ListCount - 1
    If UCase(Trim(cbo_location)) = UCase(Trim(cbo_location.List(i, 0))) Then
        cbo_location.Text = cbo_location.List(i, 0)
        lbl_location = cbo_location.List(i, 1)
        ToControlCls = cbo_location.List(i, 2)
        lbl_pesan.Caption = ""
        j = 1
        Exit For
    End If
Next

If j = 0 Then
    lbl_location = "":  validCombo = DisplayMsg(4014) '"Invalid location code !"
    Exit Function
End If
 
'# cek combo Machine No.
If Trim(cbo_MachineNo) = "" Then
    validCombo = DisplayMsg(4079) '"Please input Machine No. !"
    cboEmpty = True: Exit Function
End If
cboEmpty = False

j = 0
For i = 0 To cbo_MachineNo.ListCount - 1
    If UCase(Trim(cbo_MachineNo)) = UCase(Trim(cbo_MachineNo.List(i, 0))) Then
        cbo_MachineNo.Text = cbo_MachineNo.List(i, 0)
        lbl_machine.Caption = cbo_MachineNo.List(i, 1)
        lbl_pesan.Caption = ""
        j = 1
        Exit For
    End If
Next
      
If j = 0 Then
    lbl_machine.Caption = ""
    validCombo = DisplayMsg(4079) '" Invalid Machine No. ! "
    Exit Function
End If

'# cek combo Supply Cls
If Trim(cbo_supply) = "" Then
    validCombo = DisplayMsg(4052) ' "Please input Supply Clasification !"
    cboEmpty = True: Exit Function
End If
cboEmpty = False

j = 0
For i = 0 To cbo_supply.ListCount - 1
    If UCase(Trim(cbo_supply)) = UCase(Trim(cbo_supply.List(i, 0))) Then
        cbo_supply.Text = cbo_supply.List(i, 0)
        lbl_pesan.Caption = ""
        j = 1
        Exit For
    End If
Next
      
If j = 0 Then
    lbl_supply = "": validCombo = DisplayMsg(4056) '"Invalid supply clasification !"
    Exit Function
End If

'# cek combo Material Cls
If Trim(cbo_material) = "" Then
    validCombo = DisplayMsg(8095) ' "Please input Material Clasification !"
    cboEmpty = True:  Exit Function
End If
cboEmpty = False
 
j = 0
For i = 0 To cbo_material.ListCount - 1
    If UCase(Trim(cbo_material)) = UCase(Trim(cbo_material.List(i, 0))) Then
        cbo_material.Text = cbo_material.List(i, 0)
        lbl_pesan.Caption = ""
        j = 1
        Exit For
    End If
Next
      
If j = 0 Then
    validCombo = DisplayMsg(8095) '"Invalid Material Clasification !"
    Exit Function
End If

End Function


Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

Call koneksi
DTPicker3.Value = Format(Date, "dd MMM yyyy")

lbl_warehouse.Caption = ""
lbl_location = ""
lbl_machine.Caption = ""
lbl_pesan.Caption = ""

cbo_location = ""
cbo_warehouse = ""
cbo_MachineNo = ""
cbo_material = ""
cbo_supplyNo = ""
CboWominNo = ""

date1.Value = Format(Now, "MMM yyyy")
idt_dateUp = Date

Call setting
Call setting_grid
lbl_supply.Caption = ""
cbo_supply = ""
CboWominNo = ""

boolAdd = True

lbl_pesan.Caption = ""
Call comboDataChange(True)
End Sub


Private Sub set_line()

Dim rs_line As New ADODB.Recordset

cbo_MachineNo.clear
cbo_MachineNo = ""
lbl_machine = ""

If rs_line.State <> adStateClosed Then rs_line.Close
rs_line.Open " select * from manufacture_line ", Db ' where manufacture_code='" & Trim(cbo_location.Text) & "' ", Db

If Not (rs_line.EOF And rs_line.BOF) Then
 
    rs_line.MoveFirst
    i = 0

    While Not rs_line.EOF
        cbo_MachineNo.AddItem ""
        cbo_MachineNo.List(i, 0) = rs_line!line_code
        cbo_MachineNo.List(i, 1) = rs_line!line_name

        i = i + 1
        rs_line.MoveNext
    Wend
    LocationHasNoLine = False
Else
    LocationHasNoLine = True
End If
        
End Sub

Private Sub setting()
Dim SqlW As String

'------------------ cbo warehouse-------------------

cbo_warehouse.clear
cbo_warehouse.columnCount = 3
cbo_warehouse.TextColumn = 1

i = 0

If rs_warehouse.EOF = False Or rs_warehouse.BOF = False Then
    rs_warehouse.MoveFirst
    While rs_warehouse.EOF = False
        cbo_warehouse.AddItem ""
        cbo_warehouse.List(i, 0) = Trim(rs_warehouse!wh_code)
        cbo_warehouse.List(i, 1) = rs_warehouse!WH_Name
        cbo_warehouse.List(i, 2) = rs_warehouse!stockcontrol_cls
        rs_warehouse.MoveNext
        i = i + 1
    Wend
    cbo_warehouse.ColumnWidths = "50 pt; 175 pt; 0 pt"
    cbo_warehouse.ListWidth = 225
End If
rs_warehouse.Close


'------------------ cbo Location-------------------
cbo_location.clear
cbo_location.columnCount = 3
cbo_location.TextColumn = 1

Dim rsFac As New ADODB.Recordset
SqlW = "select wh_code,wh_name,stockControl_cls from warehouse_master " & _
         " union all " & _
         " select distinct(manufacture_line.manufacture_code)wh_code,trade_name wh_name,stockControl_Cls='01' from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code "
'rsFac.Open "select distinct(manufacture_line.manufacture_code),trade_name from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code", Db
rsFac.Open SqlW, Db

If rsFac.EOF = False Then
    rsFac.MoveFirst
    i = 0
    
    If rsFac.EOF = False Or rsFac.BOF = False Then
        rsFac.MoveFirst
        While rsFac.EOF = False
            cbo_location.AddItem ""
'            cbo_location.List(i, 0) = Trim(rsFac!Manufacture_Code)
'            cbo_location.List(i, 1) = Trim(rsFac!trade_Name)
            cbo_location.List(i, 0) = Trim(rsFac!wh_code)
            cbo_location.List(i, 1) = rsFac!WH_Name
            cbo_location.List(i, 2) = rsFac!stockcontrol_cls
            rsFac.MoveNext
            i = i + 1
        Wend
        cbo_location.ColumnWidths = "50 pt;175 pt;0 pt"
        cbo_location.ListWidth = 225
    End If

End If
rsFac.Close

'------------------ cbo supply-------------------
cbo_supply.clear
cbo_supply.columnCount = 2
cbo_supply.TextColumn = 1

cbo_supply.AddItem ""
cbo_supply.List(0, 0) = "S1"
cbo_supply.List(0, 1) = "Supply"
cbo_supply.AddItem ""
cbo_supply.List(1, 0) = "S"
cbo_supply.List(1, 1) = "Consumption"
cbo_supply.AddItem ""
cbo_supply.List(2, 0) = "L"
cbo_supply.List(2, 1) = "Loss"
cbo_supply.AddItem ""
cbo_supply.List(3, 0) = "RJ"
cbo_supply.List(3, 1) = "Reject"
cbo_supply.ColumnWidths = "25 pt; 75 pt"
cbo_supply.ListWidth = 100

cbo_supply = "S1"


'------------------ cbo material-------------------

cbo_material.clear
cbo_material.columnCount = 2
cbo_material.TextColumn = 1

        cbo_material.AddItem ""
        cbo_material.List(0, 0) = uf_GetMaterialDescription("01") '"Resin"
        cbo_material.List(0, 1) = "1"
        cbo_material.AddItem ""
        cbo_material.List(1, 0) = "Others"
        cbo_material.List(1, 1) = "0"
        
cbo_material.ColumnWidths = "100 pt; 0 pt"
cbo_material.ListWidth = 100


'------------------ cbo supply No-------------------
Call setCboSupplyNo


'------------------ Last Data Inventory Closing-------------------
Dim sqlControl As String, RsInvControl As New ADODB.Recordset
sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year desc ,inventory_month desc"

If RsInvControl.State <> adStateClosed Then RsInvControl.Close
RsInvControl.Open sqlControl, Db, adOpenKeyset, adLockOptimistic

If RsInvControl.EOF = True And RsInvControl.BOF = True Then
    'lbl_pesan = DisplayMsg(4022) '"Inventory Stock hasn't been closed !"
    Exit Sub
End If
RsInvControl.MoveFirst
InvYear = Trim(RsInvControl!Inventory_Year)
InvMonth = Trim(RsInvControl!Inventory_Month)

End Sub

Private Sub koneksi()
Dim SqlW As String
rs_part_supply.Open "select top 5 * from part_supply", Db, adOpenKeyset, adLockOptimistic
SqlW = " select * from (select wh_code,wh_name,stockControl_cls from warehouse_master " & _
         " union all " & _
         " select distinct(manufacture_line.manufacture_code)wh_code,trade_name wh_name,stockControl_Cls='01' from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code)tbJ order by wh_code "
rs_warehouse.Open SqlW, Db, adOpenKeyset, adLockOptimistic
rs_trade_master.Open "select * from trade_master where trade_cls='1'", Db, adOpenKeyset, adLockOptimistic
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'db2.Close
If rs_part_supply.State <> adStateClosed Then rs_part_supply.Close
If rs_warehouse.State <> adStateClosed Then rs_warehouse.Close
If rs_trade_master.State <> adStateClosed Then rs_trade_master.Close
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'# Check Valid Qty Numeric Input
If IsNumeric(Trim(Grid1.TextMatrix(Row, Col))) = False And Trim(Grid1.TextMatrix(Row, Col)) <> "" Then Grid1.TextMatrix(Row, Col) = 0
If Trim(Grid1.TextMatrix(Row, Col)) = "" Then
    Exit Sub
Else
    Grid1.TextMatrix(Row, Col) = Format(Grid1.TextMatrix(Row, Col), gs_formatQtyBOM)
End If
If CDbl(Trim(Grid1.TextMatrix(Row, Col))) > gd_MaxQty Then
    Grid1.TextMatrix(Row, Col) = gd_MaxQty
    lbl_pesan = DisplayMsg(4045) & " " & gd_MaxQty
End If

If CDbl(Trim(Grid1.TextMatrix(Row, Col))) > CDbl(Trim(Grid1.TextMatrix(Row, bteColRemQty))) Then
    Grid1.CellBackColor = vbRed
Else
    Grid1.CellBackColor = vbWhite
End If

End Sub

Private Sub Grid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Select Case Grid1.Col
    Case bteColSelect: Grid1.EditMaxLength = 1
    Case bteColQty: Grid1.EditMaxLength = 15
    'Case 9: Grid1.EditMaxLength = 7
    Case Else: Cancel = True
End Select
End Sub

Private Sub Grid1_Click()
If Grid1.Row <> bteColSelect Then
    'If Grid1.Col = 9 Or Grid1.Col = bteColQty Then Grid1.FocusRect = flexFocusInset Else Grid1.FocusRect = flexFocusNone
    If Grid1.Col = bteColQty Then Grid1.FocusRect = flexFocusInset Else Grid1.FocusRect = flexFocusNone
End If
End Sub

Private Sub Grid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 13 And Grid1.Row + 1 <> Grid1.Rows Then Grid1.Row = Grid1.Row + 1
If Col = bteColQty Then '# Qty Column
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> vbKeyReturn Then KeyAscii = 0
End If
End Sub

Private Sub lbl_location_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub TxtUnitSet_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> vbKeyReturn Then KeyAscii = 0

End Sub
