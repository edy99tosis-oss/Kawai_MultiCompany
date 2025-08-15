VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_item_master2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FDDFE3&
   Caption         =   "Item Master"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   ForeColor       =   &H00C0E0FF&
   Icon            =   "Frm_item_master2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin EZRunnerv3.CtrlMenu txtmenu 
      Height          =   420
      Left            =   13035
      TabIndex        =   174
      TabStop         =   0   'False
      Top             =   345
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Inquiry"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   9705
      Width           =   1140
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7110
      Left            =   360
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   1800
      Width           =   14475
      _ExtentX        =   25532
      _ExtentY        =   12541
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16637923
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "    General    "
      TabPicture(0)   =   "Frm_item_master2.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Manufacture"
      TabPicture(1)   =   "Frm_item_master2.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame9"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         Caption         =   "Material Dimension"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3300
         Left            =   -74760
         TabIndex        =   152
         Top             =   3000
         Width           =   5250
         Begin VB.TextBox txt_gross 
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
            Left            =   1800
            MaxLength       =   13
            TabIndex        =   46
            Text            =   "txt_gross"
            Top             =   2460
            Width           =   1575
         End
         Begin VB.TextBox txt_width 
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
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   43
            Text            =   "txt_width"
            Top             =   1425
            Width           =   1575
         End
         Begin VB.TextBox txt_thickness 
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
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   42
            Text            =   "txt_thicknes"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txt_weight 
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
            Left            =   1800
            MaxLength       =   13
            TabIndex        =   45
            Text            =   "txt_weight"
            Top             =   2115
            Width           =   1575
         End
         Begin VB.TextBox txt_length 
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
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   44
            Text            =   "txt_length"
            Top             =   1770
            Width           =   1575
         End
         Begin VB.Label Label66 
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Weight"
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
            Left            =   150
            TabIndex        =   181
            Top             =   2513
            Width           =   1785
         End
         Begin VB.Label Label62 
            BackStyle       =   0  'Transparent
            Caption         =   "Bag Packing Style"
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
            Left            =   150
            TabIndex        =   178
            Top             =   705
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSForms.ComboBox cmbbox_mPacking 
            Height          =   315
            Left            =   1800
            TabIndex        =   41
            Top             =   690
            Visible         =   0   'False
            Width           =   780
            VariousPropertyBits=   746604571
            MaxLength       =   15
            DisplayStyle    =   3
            Size            =   "1376;556"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_mPacking"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lbl_mPacking 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_mPacking"
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
            Left            =   2655
            TabIndex        =   177
            Top             =   735
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.Line Line26 
            Visible         =   0   'False
            X1              =   2610
            X2              =   5010
            Y1              =   990
            Y2              =   990
         End
         Begin VB.Label Label16 
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
            Height          =   225
            Left            =   150
            TabIndex        =   158
            Top             =   330
            Width           =   1785
         End
         Begin MSForms.ComboBox cmbbox_material 
            Height          =   330
            Left            =   1800
            TabIndex        =   40
            Top             =   330
            Width           =   765
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1349;582"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_material"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lbl_material 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_material"
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
            Left            =   2640
            TabIndex        =   157
            Top             =   360
            Width           =   2355
         End
         Begin VB.Line Line11 
            X1              =   2640
            X2              =   5010
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Thickness"
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
            Left            =   150
            TabIndex        =   156
            Top             =   1133
            Width           =   1785
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
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
            Left            =   150
            TabIndex        =   155
            Top             =   1478
            Width           =   1785
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Net Weight"
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
            Left            =   150
            TabIndex        =   154
            Top             =   2168
            Width           =   1785
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Length"
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
            Left            =   150
            TabIndex        =   153
            Top             =   1823
            Width           =   1785
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Factory  Info"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   -74820
         TabIndex        =   148
         Top             =   585
         Width           =   5250
         Begin VB.TextBox txt_factory 
            BackColor       =   &H8000000F&
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   172
            TabStop         =   0   'False
            Text            =   "txt_factory"
            Top             =   720
            Width           =   3105
         End
         Begin MSForms.ComboBox cmbbox_line 
            Height          =   330
            Left            =   1890
            TabIndex        =   39
            Top             =   1170
            Width           =   1575
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2778;582"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_line"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lbl_factory 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_factory"
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
            Left            =   2790
            TabIndex        =   159
            Top             =   1980
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Line Line5 
            X1              =   1920
            X2              =   5040
            Y1              =   990
            Y2              =   990
         End
         Begin VB.Line Line2 
            X1              =   1920
            X2              =   5040
            Y1              =   1815
            Y2              =   1815
         End
         Begin MSForms.ComboBox cmbox_manufacture 
            Height          =   330
            Left            =   1890
            TabIndex        =   38
            Top             =   330
            Width           =   1575
            VariousPropertyBits=   746604571
            MaxLength       =   6
            DisplayStyle    =   3
            Size            =   "2778;582"
            ShowDropButtonWhen=   2
            Value           =   "cmbox_manufacture"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lbl_line 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_line"
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
            Left            =   1920
            TabIndex        =   151
            Top             =   1560
            Width           =   3105
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Line Code"
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
            Left            =   120
            TabIndex        =   150
            Top             =   1155
            Width           =   1785
         End
         Begin VB.Label Label8 
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
            Height          =   225
            Left            =   120
            TabIndex        =   149
            Top             =   360
            Width           =   1785
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   6120
         TabIndex        =   121
         Top             =   480
         Width           =   8265
         Begin VB.TextBox txtSafetyStock2 
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
            Left            =   2460
            MaxLength       =   6
            TabIndex        =   189
            Text            =   "txt_safety"
            Top             =   1980
            Width           =   1620
         End
         Begin VB.TextBox TxtMinOrder 
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
            Left            =   6345
            MaxLength       =   13
            TabIndex        =   36
            Top             =   2040
            Width           =   1695
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   330
            Left            =   2460
            TabIndex        =   29
            Top             =   5130
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
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
            Format          =   "MM/dd/yyyy"
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txt_number_of_box 
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
            Left            =   2460
            MaxLength       =   13
            TabIndex        =   23
            Text            =   "txt_number"
            Top             =   3000
            Width           =   1620
         End
         Begin VB.TextBox txt_ne 
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
            Left            =   2460
            MaxLength       =   9
            TabIndex        =   16
            Text            =   "txt_ne"
            Top             =   210
            Width           =   1620
         End
         Begin VB.TextBox txt_accounting_code 
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
            Left            =   2460
            MaxLength       =   7
            TabIndex        =   24
            Text            =   "txt_accounting_code"
            Top             =   3360
            Width           =   1620
         End
         Begin VB.TextBox txt_min_stock 
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
            Left            =   2460
            MaxLength       =   13
            TabIndex        =   22
            Text            =   "txt_min_st"
            Top             =   2670
            Width           =   1620
         End
         Begin VB.TextBox txt_max_stock 
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
            Left            =   2460
            MaxLength       =   13
            TabIndex        =   21
            Text            =   "txt_max_st"
            Top             =   2340
            Width           =   1620
         End
         Begin VB.TextBox txt_order_point_qty 
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
            Left            =   6345
            MaxLength       =   13
            TabIndex        =   35
            Text            =   "txt_order_"
            Top             =   1710
            Width           =   1695
         End
         Begin VB.TextBox txt_allowance_day 
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
            Left            =   2475
            MaxLength       =   2
            TabIndex        =   28
            Text            =   "tx"
            Top             =   4785
            Width           =   765
         End
         Begin VB.TextBox txt_delivery_read_time 
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
            Left            =   2460
            MaxLength       =   2
            TabIndex        =   30
            Text            =   "tx"
            Top             =   5850
            Width           =   765
         End
         Begin VB.TextBox txt_standart_stock 
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
            Left            =   2460
            MaxLength       =   13
            TabIndex        =   19
            Text            =   "txt_standa"
            Top             =   1275
            Width           =   1620
         End
         Begin VB.TextBox txt_safety_stock 
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
            Left            =   2460
            MaxLength       =   13
            TabIndex        =   20
            Text            =   "txt_safety"
            Top             =   1620
            Width           =   1620
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   2475
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   4770
            Width           =   1545
            _ExtentX        =   2725
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
            Format          =   117374979
            CurrentDate     =   37818
         End
         Begin VB.Label Label74 
            BackStyle       =   0  'Transparent
            Caption         =   "Color Cls"
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
            Left            =   4515
            TabIndex        =   207
            Tag             =   "TTFF*/"
            Top             =   5640
            Width           =   1785
         End
         Begin VB.Line Line31 
            X1              =   6360
            X2              =   8180
            Y1              =   6240
            Y2              =   6240
         End
         Begin VB.Label lbl_Color 
            Caption         =   "lbl_Color"
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
            Left            =   6360
            TabIndex        =   206
            Tag             =   "TTFF*/"
            Top             =   6000
            Width           =   1800
         End
         Begin MSForms.ComboBox cboColor 
            Height          =   345
            Left            =   6360
            TabIndex        =   205
            Tag             =   "TTFF*/"
            Top             =   5640
            Width           =   765
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1349;609"
            ShowDropButtonWhen=   2
            Value           =   "cmb_POType"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Line Line30 
            X1              =   6360
            X2              =   8180
            Y1              =   5530
            Y2              =   5530
         End
         Begin VB.Label lbl_Destination 
            Caption         =   "lbl_Destination"
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
            Left            =   6360
            TabIndex        =   204
            Tag             =   "TTFF*/"
            Top             =   5310
            Width           =   1800
         End
         Begin MSForms.ComboBox cboDestination 
            Height          =   345
            Left            =   6360
            TabIndex        =   203
            Tag             =   "TTFF*/"
            Top             =   4930
            Width           =   765
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1349;609"
            ShowDropButtonWhen=   2
            Value           =   "cmb_POType"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label72 
            BackStyle       =   0  'Transparent
            Caption         =   "Destination Cls"
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
            Left            =   4515
            TabIndex        =   202
            Tag             =   "TTFF*/"
            Top             =   4960
            Width           =   1785
         End
         Begin VB.Line Line35 
            X1              =   6360
            X2              =   8180
            Y1              =   4850
            Y2              =   4850
         End
         Begin VB.Label lbl_POType 
            Caption         =   "lbl_POType"
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
            Left            =   6360
            TabIndex        =   198
            Tag             =   "TTFF*/"
            Top             =   4605
            Width           =   1800
         End
         Begin VB.Label Label82 
            BackStyle       =   0  'Transparent
            Caption         =   "PO Type"
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
            Left            =   4515
            TabIndex        =   197
            Tag             =   "TTFF*/"
            Top             =   4290
            Width           =   1785
         End
         Begin MSForms.ComboBox cbo_POType 
            Height          =   345
            Left            =   6360
            TabIndex        =   196
            Tag             =   "TTFF*/"
            Top             =   4200
            Width           =   765
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1349;609"
            ShowDropButtonWhen=   2
            Value           =   "cmb_POType"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblModel 
            Caption         =   "lbl_Model"
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
            Left            =   7200
            TabIndex        =   195
            Top             =   3840
            Width           =   885
         End
         Begin VB.Line Line28 
            X1              =   7200
            X2              =   8055
            Y1              =   4080
            Y2              =   4080
         End
         Begin MSForms.ComboBox cboModel 
            Height          =   315
            Left            =   6360
            TabIndex        =   194
            Top             =   3840
            Width           =   750
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1323;556"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_unit"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label71 
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
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
            Left            =   4515
            TabIndex        =   193
            Top             =   3870
            Width           =   1155
         End
         Begin MSForms.ComboBox cboTypeAccs 
            Height          =   315
            Left            =   6360
            TabIndex        =   192
            Top             =   3435
            Width           =   765
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1349;556"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_packing"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label70 
            BackStyle       =   0  'Transparent
            Caption         =   "Type Accs"
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
            Left            =   4515
            TabIndex        =   191
            Top             =   3480
            Width           =   1155
         End
         Begin VB.Label Label69 
            BackStyle       =   0  'Transparent
            Caption         =   "Safety Stock (%)"
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
            Left            =   150
            TabIndex        =   190
            Top             =   2010
            Width           =   1995
         End
         Begin VB.Label Label68 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum Order Qty"
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
            Left            =   4485
            TabIndex        =   188
            Top             =   2100
            Width           =   1755
         End
         Begin MSForms.ComboBox cbo_PackingCode 
            Height          =   315
            Left            =   6345
            TabIndex        =   32
            Top             =   600
            Visible         =   0   'False
            Width           =   1695
            VariousPropertyBits=   746604571
            MaxLength       =   15
            DisplayStyle    =   3
            Size            =   "2990;556"
            ShowDropButtonWhen=   2
            Value           =   "cbo_PackingCode"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label64 
            BackStyle       =   0  'Transparent
            Caption         =   "Packing Item Code"
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
            Left            =   4470
            TabIndex        =   176
            Top             =   645
            Visible         =   0   'False
            Width           =   1710
         End
         Begin VB.Label lbl_packing2 
            Caption         =   "lbl_packing"
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
            Left            =   6345
            TabIndex        =   169
            Top             =   3045
            Width           =   1425
         End
         Begin VB.Line Line25 
            X1              =   6360
            X2              =   8055
            Y1              =   3285
            Y2              =   3285
         End
         Begin MSForms.ComboBox cmbbox_packing2 
            Height          =   315
            Left            =   6360
            TabIndex        =   37
            Top             =   2640
            Width           =   765
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1349;556"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_packing"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "Packing Style Part / Material"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   4515
            TabIndex        =   168
            Top             =   2655
            Width           =   1455
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Qty/box (Parts/Material)"
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
            Left            =   150
            TabIndex        =   162
            Top             =   3030
            Width           =   2655
         End
         Begin MSForms.ComboBox cmb_explosion2 
            Height          =   315
            Left            =   1650
            TabIndex        =   25
            Top             =   3705
            Width           =   765
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1349;556"
            ShowDropButtonWhen=   2
            Value           =   "cmb_explosion2"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lbl_unit 
            Caption         =   "lbl_unit"
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
            Left            =   7215
            TabIndex        =   161
            Top             =   1380
            Width           =   885
         End
         Begin VB.Line Line19 
            X1              =   7200
            X2              =   8055
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Line Line20 
            X1              =   7200
            X2              =   8055
            Y1              =   1215
            Y2              =   1215
         End
         Begin MSForms.ComboBox cmbbox_unit 
            Height          =   315
            Left            =   6345
            TabIndex        =   34
            Top             =   1335
            Width           =   750
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1323;556"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_unit"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmb_stock_control2 
            Height          =   315
            Left            =   1650
            TabIndex        =   27
            Top             =   4410
            Width           =   765
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1349;556"
            ShowDropButtonWhen=   2
            Value           =   "cmd_stock_control2"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Label4"
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
            Left            =   2550
            TabIndex        =   160
            Top             =   5535
            Width           =   555
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Qty/Case (Finish goods)"
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
            Left            =   150
            TabIndex        =   147
            Top             =   270
            Width           =   2235
         End
         Begin MSForms.ComboBox cmb_make 
            Height          =   315
            Left            =   6345
            TabIndex        =   31
            Top             =   210
            Width           =   750
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1323;556"
            ShowDropButtonWhen=   2
            Value           =   "cmb_make"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "Make or Buy Cls"
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
            Left            =   4440
            TabIndex        =   146
            Top             =   255
            Width           =   1695
         End
         Begin MSForms.ComboBox cmbbox_control 
            Height          =   315
            Left            =   6345
            TabIndex        =   33
            Top             =   945
            Width           =   750
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1323;556"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_cotrol"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "Control"
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
            Left            =   4485
            TabIndex        =   145
            Top             =   990
            Width           =   1455
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "Accounting Code"
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
            Left            =   150
            TabIndex        =   144
            Top             =   3390
            Width           =   1395
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Max Stock"
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
            Left            =   150
            TabIndex        =   143
            Top             =   2370
            Width           =   1395
         End
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            Caption         =   "Min Stock"
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
            Left            =   150
            TabIndex        =   142
            Top             =   2715
            Width           =   1395
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "Allowance Day"
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
            Left            =   150
            TabIndex        =   141
            Top             =   4815
            Width           =   1395
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Lead Time"
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
            Left            =   150
            TabIndex        =   140
            Top             =   5880
            Width           =   1785
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "Order Point Qty"
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
            Left            =   4485
            TabIndex        =   139
            Top             =   1740
            Width           =   1635
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Cls"
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
            Left            =   4485
            TabIndex        =   138
            Top             =   1380
            Width           =   1155
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Standard Stock"
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
            Left            =   150
            TabIndex        =   137
            Top             =   1305
            Width           =   1395
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Safety Stock"
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
            Left            =   150
            TabIndex        =   136
            Top             =   1650
            Width           =   1395
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "Explosion Cls"
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
            Left            =   150
            TabIndex        =   135
            Top             =   3750
            Width           =   1395
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Person in Charge"
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
            Left            =   150
            TabIndex        =   134
            Top             =   4080
            Width           =   1485
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Control "
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
            Left            =   150
            TabIndex        =   133
            Top             =   4455
            Width           =   1575
         End
         Begin MSForms.ComboBox cmbbox_purchase_person 
            Height          =   315
            Left            =   1650
            TabIndex        =   26
            Top             =   4050
            Width           =   765
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1349;556"
            ListWidth       =   7761
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_purchase_person"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "Use End Date"
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
            Left            =   150
            TabIndex        =   132
            Top             =   5160
            Width           =   1395
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Update"
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
            Left            =   150
            TabIndex        =   131
            Top             =   5520
            Width           =   1575
         End
         Begin VB.Label Label1 
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
            Height          =   225
            Left            =   150
            TabIndex        =   130
            Top             =   975
            Width           =   1395
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Packing Style"
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
            Left            =   150
            TabIndex        =   129
            Top             =   630
            Width           =   1305
         End
         Begin MSForms.ComboBox cmbbox_packing 
            Height          =   315
            Left            =   1650
            TabIndex        =   17
            Top             =   585
            Width           =   765
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1349;556"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_packing"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbbox_group 
            Height          =   315
            Left            =   1650
            TabIndex        =   18
            Top             =   930
            Width           =   765
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1349;556"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_group"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Line Line16 
            X1              =   2475
            X2              =   4050
            Y1              =   870
            Y2              =   870
         End
         Begin VB.Line Line17 
            X1              =   2475
            X2              =   4050
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line18 
            X1              =   7200
            X2              =   8055
            Y1              =   525
            Y2              =   525
         End
         Begin VB.Line Line21 
            X1              =   2475
            X2              =   4320
            Y1              =   3960
            Y2              =   3960
         End
         Begin VB.Line Line22 
            X1              =   2475
            X2              =   4320
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Line Line23 
            X1              =   2475
            X2              =   4320
            Y1              =   4650
            Y2              =   4650
         End
         Begin VB.Label lbl_packing 
            Caption         =   "lbl_packing"
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
            Left            =   2475
            TabIndex        =   128
            Top             =   630
            Width           =   1575
         End
         Begin VB.Label lbl_group 
            Caption         =   "lbl_group"
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
            Left            =   2475
            TabIndex        =   127
            Top             =   960
            Width           =   1530
         End
         Begin VB.Label lbl_make 
            Caption         =   "lbl_make"
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
            Left            =   7215
            TabIndex        =   126
            Top             =   270
            Width           =   885
         End
         Begin VB.Label lbl_control 
            Caption         =   "lbl_control"
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
            Left            =   7215
            TabIndex        =   125
            Top             =   990
            Width           =   885
         End
         Begin VB.Label lbl_explosion 
            Caption         =   "lbl_explosion"
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
            Left            =   2475
            TabIndex        =   124
            Top             =   3750
            Width           =   1860
         End
         Begin VB.Label lbl_purchase 
            Caption         =   "lbl_purchase"
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
            Left            =   2475
            TabIndex        =   123
            Top             =   4050
            Width           =   1860
         End
         Begin VB.Label lbl_stock_control 
            Caption         =   "lbl_stock_control"
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
            Left            =   2475
            TabIndex        =   122
            Top             =   4410
            Width           =   1860
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Item Classification"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2550
         Left            =   150
         TabIndex        =   112
         Top             =   4380
         Width           =   5865
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "Production Cls"
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
            Left            =   180
            TabIndex        =   167
            Top             =   1710
            Width           =   1395
         End
         Begin MSForms.ComboBox cmb_prod 
            Height          =   315
            Left            =   2010
            TabIndex        =   15
            Top             =   1665
            Width           =   795
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1402;556"
            ShowDropButtonWhen=   2
            Value           =   "cmb_prod"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Line Line24 
            X1              =   2910
            X2              =   4185
            Y1              =   1935
            Y2              =   1935
         End
         Begin VB.Label lbl_prod 
            Caption         =   "lbl_prod"
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
            Left            =   2940
            TabIndex        =   166
            Top             =   1695
            Width           =   1530
         End
         Begin MSForms.ComboBox cmb_part2 
            Height          =   315
            Left            =   2010
            TabIndex        =   11
            Top             =   270
            Width           =   795
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1402;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            Value           =   "cmb_part2"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmb_provision2 
            Height          =   315
            Left            =   2010
            TabIndex        =   14
            Top             =   1305
            Width           =   795
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1402;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            Value           =   "cmb_provision2"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmb_supply2 
            Height          =   315
            Left            =   2010
            TabIndex        =   13
            Top             =   945
            Width           =   795
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1402;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            Value           =   "cmb_supply2"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmb_reserve2 
            Height          =   315
            Left            =   2010
            TabIndex        =   12
            Top             =   585
            Width           =   795
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1402;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            Value           =   "cmb_reserve2"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Line Line8 
            X1              =   2910
            X2              =   4140
            Y1              =   885
            Y2              =   885
         End
         Begin VB.Label lbl_provotion 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_provotion"
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
            Left            =   2940
            TabIndex        =   120
            Top             =   1365
            Width           =   2775
         End
         Begin VB.Line Line10 
            X1              =   2910
            X2              =   4185
            Y1              =   1635
            Y2              =   1635
         End
         Begin VB.Label lbl_supply 
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
            Height          =   225
            Left            =   2940
            TabIndex        =   119
            Top             =   1005
            Width           =   2775
         End
         Begin VB.Line Line9 
            X1              =   2910
            X2              =   4185
            Y1              =   1275
            Y2              =   1275
         End
         Begin VB.Label lbl_reserve 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_reserve"
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
            Left            =   2940
            TabIndex        =   118
            Top             =   615
            Width           =   2775
         End
         Begin VB.Label lbl_part 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_part"
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
            Left            =   2940
            TabIndex        =   117
            Top             =   285
            Width           =   2775
         End
         Begin VB.Line Line7 
            X1              =   2910
            X2              =   4140
            Y1              =   555
            Y2              =   555
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Part Cls"
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
            Left            =   180
            TabIndex        =   116
            Top             =   315
            Width           =   1785
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Provision Cls"
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
            Left            =   180
            TabIndex        =   115
            Top             =   1365
            Width           =   1785
         End
         Begin VB.Label Label14 
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
            Height          =   225
            Left            =   180
            TabIndex        =   114
            Top             =   1005
            Width           =   1785
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Reserve Cls"
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
            Left            =   180
            TabIndex        =   113
            Top             =   630
            Width           =   1785
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Material Clasification"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5760
         Left            =   -69420
         TabIndex        =   90
         Top             =   585
         Width           =   8760
         Begin VB.TextBox txt_heat_qty 
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
            Left            =   7290
            MaxLength       =   9
            TabIndex        =   55
            Text            =   "txt_heat_"
            Top             =   2655
            Width           =   1290
         End
         Begin VB.TextBox txt_surface_qty 
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
            Left            =   7290
            MaxLength       =   9
            TabIndex        =   53
            Text            =   "txt_surfa"
            Top             =   2205
            Width           =   1290
         End
         Begin VB.TextBox txt_lot_coef 
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
            Left            =   6150
            MaxLength       =   9
            TabIndex        =   64
            Text            =   "txt_lot_c"
            Top             =   3780
            Width           =   1560
         End
         Begin VB.TextBox txt_number_producible 
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
            Left            =   2160
            MaxLength       =   12
            TabIndex        =   49
            Text            =   "txt_number_producible"
            Top             =   1080
            Width           =   1560
         End
         Begin VB.TextBox txt_scrap_weight 
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
            Left            =   2160
            MaxLength       =   12
            TabIndex        =   50
            Text            =   "txt_scrap_weight"
            Top             =   1425
            Width           =   1560
         End
         Begin VB.TextBox txt_pitch 
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
            Left            =   2160
            MaxLength       =   12
            TabIndex        =   48
            Text            =   "txt_pitch"
            Top             =   720
            Width           =   1560
         End
         Begin VB.TextBox txt_prt 
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
            Left            =   6150
            MaxLength       =   2
            TabIndex        =   65
            Text            =   "tx"
            Top             =   4095
            Width           =   1560
         End
         Begin VB.TextBox txt_lot 
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
            Left            =   6150
            MaxLength       =   9
            TabIndex        =   63
            Text            =   "txt_lot"
            Top             =   3450
            Width           =   1560
         End
         Begin VB.TextBox txt_yp 
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
            Left            =   6150
            MaxLength       =   6
            TabIndex        =   66
            Text            =   "txt_yp"
            Top             =   4425
            Width           =   1560
         End
         Begin VB.TextBox txt_sw 
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
            Left            =   2160
            MaxLength       =   12
            TabIndex        =   57
            Text            =   "txt_sw"
            Top             =   3375
            Width           =   1560
         End
         Begin VB.TextBox txt_sample 
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
            Left            =   2160
            MaxLength       =   12
            TabIndex        =   56
            Text            =   "txt_sample"
            Top             =   3045
            Width           =   1560
         End
         Begin VB.TextBox txt_mc 
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
            Left            =   2160
            MaxLength       =   12
            TabIndex        =   60
            Text            =   "txt_mc"
            Top             =   4395
            Width           =   1560
         End
         Begin VB.TextBox txt_ew 
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
            Left            =   2160
            MaxLength       =   12
            TabIndex        =   58
            Text            =   "txt_ew"
            Top             =   3675
            Width           =   1560
         End
         Begin VB.TextBox txt_min_lot 
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
            Left            =   6150
            MaxLength       =   9
            TabIndex        =   62
            Text            =   "txt_min"
            Top             =   3090
            Width           =   1560
         End
         Begin VB.TextBox txt_pc 
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
            Left            =   2160
            MaxLength       =   12
            TabIndex        =   61
            Text            =   "txt_pc"
            Top             =   4770
            Width           =   1560
         End
         Begin VB.TextBox txt_np 
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
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   59
            Text            =   "tx"
            Top             =   4035
            Width           =   765
         End
         Begin VB.Label Label73 
            BackStyle       =   0  'Transparent
            Caption         =   "Clasification Part"
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
            Left            =   120
            TabIndex        =   201
            Top             =   5210
            Width           =   1755
         End
         Begin MSForms.ComboBox cboClasificationPart 
            Height          =   375
            Left            =   2160
            TabIndex        =   200
            Top             =   5160
            Width           =   765
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1349;661"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_heat"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblClasificationPart 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_heat"
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
            Left            =   3045
            TabIndex        =   199
            Top             =   5220
            Width           =   2235
         End
         Begin VB.Line Line29 
            X1              =   3045
            X2              =   5310
            Y1              =   5460
            Y2              =   5460
         End
         Begin VB.Label Label65 
            Caption         =   "x"
            Height          =   915
            Left            =   5430
            TabIndex        =   180
            Top             =   2100
            Width           =   1785
         End
         Begin VB.Label Label63 
            Caption         =   "x"
            Height          =   915
            Left            =   90
            TabIndex        =   179
            Top             =   3030
            Width           =   1875
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "Heat Order point Qty"
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
            Left            =   5445
            TabIndex        =   165
            Top             =   2685
            Width           =   1800
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "Srfc Order Point Qty"
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
            Left            =   5445
            TabIndex        =   164
            Top             =   2265
            Width           =   1800
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Lot Coefficient"
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
            Left            =   4110
            TabIndex        =   163
            Top             =   3780
            Width           =   1395
         End
         Begin VB.Line Line15 
            X1              =   3015
            X2              =   5325
            Y1              =   585
            Y2              =   585
         End
         Begin VB.Label lbl_sheet 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_sheet"
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
            Left            =   3045
            TabIndex        =   111
            Top             =   360
            Width           =   4875
         End
         Begin VB.Line Line14 
            X1              =   3045
            X2              =   5310
            Y1              =   2910
            Y2              =   2910
         End
         Begin VB.Label lbl_heat 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_heat"
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
            Left            =   3045
            TabIndex        =   110
            Top             =   2670
            Width           =   4875
         End
         Begin VB.Line Line13 
            X1              =   3045
            X2              =   5310
            Y1              =   2535
            Y2              =   2535
         End
         Begin VB.Label lbl_surface 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_surface"
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
            Left            =   3045
            TabIndex        =   109
            Top             =   2265
            Width           =   4875
         End
         Begin VB.Line Line12 
            X1              =   3045
            X2              =   5310
            Y1              =   2100
            Y2              =   2100
         End
         Begin VB.Label lbl_drawing 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_drawing"
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
            Left            =   3045
            TabIndex        =   108
            Top             =   1860
            Width           =   4875
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Drawing Material Cls"
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
            Left            =   120
            TabIndex        =   107
            Top             =   1860
            Width           =   2115
         End
         Begin MSForms.ComboBox cmbbox_drawing 
            Height          =   375
            Left            =   2160
            TabIndex        =   51
            Top             =   1785
            Width           =   765
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1349;661"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_drawing"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Surface Treatment Cls"
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
            Left            =   120
            TabIndex        =   106
            Top             =   2265
            Width           =   2085
         End
         Begin MSForms.ComboBox cmbbox_surface 
            Height          =   375
            Left            =   2160
            TabIndex        =   52
            Top             =   2190
            Width           =   765
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1349;661"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_surface"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Heat Treatment Cls"
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
            Left            =   120
            TabIndex        =   105
            Top             =   2685
            Width           =   2085
         End
         Begin MSForms.ComboBox cmbbox_heat 
            Height          =   375
            Left            =   2160
            TabIndex        =   54
            Top             =   2610
            Width           =   765
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1349;661"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_heat"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbbox_sheet 
            Height          =   375
            Left            =   2160
            TabIndex        =   47
            Top             =   270
            Width           =   765
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   3
            Size            =   "1349;661"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_sheet"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Scrap Weight"
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
            Left            =   120
            TabIndex        =   104
            Top             =   1455
            Width           =   1785
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Number Producible"
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
            Left            =   120
            TabIndex        =   103
            Top             =   1110
            Width           =   1785
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Pitch"
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
            Left            =   120
            TabIndex        =   102
            Top             =   750
            Width           =   1785
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Sheet / Coil Cls"
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
            Left            =   120
            TabIndex        =   101
            Top             =   360
            Width           =   1785
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Lot Qty"
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
            Left            =   4110
            TabIndex        =   100
            Top             =   3450
            Width           =   1395
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Lead Time"
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
            Left            =   4110
            TabIndex        =   99
            Top             =   4095
            Width           =   1845
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Yield Percentages"
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
            Left            =   4110
            TabIndex        =   98
            Top             =   4425
            Width           =   1395
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Sample"
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
            Left            =   120
            TabIndex        =   97
            Top             =   3075
            Width           =   1395
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "SW Qty"
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
            Left            =   120
            TabIndex        =   96
            Top             =   3375
            Width           =   1395
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "EW Qty"
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
            Left            =   120
            TabIndex        =   95
            Top             =   3705
            Width           =   1395
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Number of Process"
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
            Left            =   120
            TabIndex        =   94
            Top             =   4035
            Width           =   1845
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Material Coefficient"
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
            Left            =   120
            TabIndex        =   93
            Top             =   4425
            Width           =   1755
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Process Coefficient"
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
            Left            =   120
            TabIndex        =   92
            Top             =   4785
            Width           =   1755
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Min Lot"
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
            Left            =   4110
            TabIndex        =   91
            Top             =   3090
            Width           =   1395
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Item Info"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   135
         TabIndex        =   86
         Tag             =   " '#Loop to get SupplyReqNo Resin"
         Top             =   450
         Width           =   5910
         Begin VB.TextBox txt_drawingCode 
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
            Left            =   2010
            MaxLength       =   15
            TabIndex        =   5
            Text            =   "txt_drawingCode"
            Top             =   1020
            Width           =   3660
         End
         Begin VB.TextBox txt_maker_item_code 
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
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   4
            Text            =   "txt_maker_item_code"
            Top             =   660
            Width           =   3660
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Drawing Code"
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
            Left            =   210
            TabIndex        =   175
            Top             =   1050
            Width           =   1785
         End
         Begin VB.Label lbl_finish 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_finish"
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
            Left            =   2940
            TabIndex        =   89
            Top             =   270
            Width           =   1845
         End
         Begin VB.Line Line6 
            X1              =   2940
            X2              =   5670
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Finish Good Part Cls"
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
            Index           =   2
            Left            =   210
            TabIndex        =   88
            Top             =   300
            Width           =   1785
         End
         Begin MSForms.ComboBox cmb_finish_good 
            Height          =   360
            Left            =   2010
            TabIndex        =   3
            Top             =   240
            Width           =   795
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1402;635"
            ShowDropButtonWhen=   2
            Value           =   "cmb_finish_good"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label12 
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
            Height          =   225
            Left            =   210
            TabIndex        =   87
            Top             =   690
            Width           =   1785
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Order  && Delivery"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   165
         TabIndex        =   84
         Top             =   1950
         Width           =   5865
         Begin VB.TextBox Text1 
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
            Left            =   2010
            MaxLength       =   6
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   705
            Width           =   1035
         End
         Begin VB.TextBox txt_supplier 
            BackColor       =   &H8000000F&
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
            Left            =   3375
            Locked          =   -1  'True
            TabIndex        =   171
            TabStop         =   0   'False
            Text            =   "txt_supplier"
            Top             =   1170
            Width           =   2295
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Warehouse Code"
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
            Index           =   1
            Left            =   180
            TabIndex        =   187
            Top             =   345
            Width           =   1785
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Place"
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
            Left            =   180
            TabIndex        =   186
            Top             =   1553
            Width           =   1785
         End
         Begin MSForms.ComboBox cmbbox_warehouse 
            Height          =   330
            Left            =   2010
            TabIndex        =   6
            Top             =   315
            Width           =   1290
            VariousPropertyBits=   746604571
            MaxLength       =   6
            DisplayStyle    =   3
            Size            =   "2275;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_warehouse"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmb_delivery 
            Height          =   330
            Left            =   2010
            TabIndex        =   9
            Top             =   1500
            Width           =   1290
            VariousPropertyBits=   746604571
            MaxLength       =   6
            DisplayStyle    =   3
            Size            =   "2275;582"
            ShowDropButtonWhen=   2
            Value           =   "cmb_delivery"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbox_suplier 
            Height          =   330
            Left            =   2010
            TabIndex        =   8
            Top             =   1105
            Width           =   1290
            VariousPropertyBits=   746604571
            MaxLength       =   6
            DisplayStyle    =   3
            Size            =   "2275;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            Value           =   "cmbox_suplier"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Code"
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
            Left            =   180
            TabIndex        =   185
            Top             =   1158
            Width           =   1785
         End
         Begin VB.Label Label3 
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
            Height          =   225
            Index           =   3
            Left            =   180
            TabIndex        =   184
            Top             =   763
            Width           =   1515
         End
         Begin MSForms.ComboBox cmb_hs 
            Height          =   330
            Left            =   2010
            TabIndex        =   10
            Top             =   1890
            Width           =   2085
            VariousPropertyBits=   746604571
            MaxLength       =   15
            DisplayStyle    =   3
            Size            =   "3678;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            Value           =   "cmb_hs"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label67 
            Caption         =   "HS Code"
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
            Left            =   180
            TabIndex        =   183
            Top             =   1950
            Width           =   1605
         End
         Begin VB.Line Line27 
            Visible         =   0   'False
            X1              =   3780
            X2              =   5640
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label lbl_hs 
            Caption         =   "lbl_hs"
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
            Left            =   3780
            TabIndex        =   182
            Top             =   1920
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label lbl_delivery_place 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_delivery_place"
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
            Left            =   3375
            TabIndex        =   170
            Top             =   1530
            Width           =   2295
         End
         Begin VB.Line Line4 
            X1              =   3375
            X2              =   5655
            Y1              =   1770
            Y2              =   1770
         End
         Begin VB.Line Line3 
            X1              =   3375
            X2              =   5640
            Y1              =   1410
            Y2              =   1410
         End
         Begin VB.Line Line1 
            X1              =   3375
            X2              =   5655
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lbl_warehouse 
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
            Height          =   225
            Left            =   3375
            TabIndex        =   85
            Top             =   360
            Width           =   2295
         End
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   11303
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   9705
      Width           =   1140
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FDDFE3&
      Height          =   855
      Left            =   360
      TabIndex        =   80
      Top             =   840
      Width           =   14580
      Begin VB.CommandButton cmd_Browser 
         Caption         =   "..."
         Height          =   300
         Left            =   3360
         TabIndex        =   1
         Top             =   337
         Width           =   300
      End
      Begin VB.TextBox txt_item_code 
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
         Left            =   1530
         MaxLength       =   25
         TabIndex        =   0
         Text            =   "txt_item_code"
         Top             =   345
         Width           =   1725
      End
      Begin VB.TextBox txt_item_name 
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
         Left            =   4995
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txt_item_master"
         Top             =   345
         Width           =   5550
      End
      Begin VB.Label lbl_record 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Left            =   11475
         TabIndex        =   173
         Top             =   367
         Width           =   2895
      End
      Begin VB.Label Label2 
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
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   82
         Top             =   375
         Width           =   1785
      End
      Begin VB.Label Label50 
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
         Height          =   225
         Left            =   3900
         TabIndex        =   81
         Top             =   375
         Width           =   1785
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   323
      TabIndex        =   78
      Top             =   9015
      Width           =   14580
      Begin VB.Label Lbl_pesan 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lbl_pesan"
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
         Left            =   120
         TabIndex        =   79
         Top             =   180
         Width           =   14295
      End
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
      Left            =   13763
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   9705
      Width           =   1140
   End
   Begin VB.CommandButton cmd_Clear 
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
      Left            =   12533
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   9705
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4253
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   9705
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Left            =   5483
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   9705
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Left            =   6713
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   9705
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Left            =   7943
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   9705
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Index           =   0
      Left            =   323
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   9705
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Master"
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
      Height          =   330
      Index           =   0
      Left            =   6375
      TabIndex        =   77
      Top             =   360
      Width           =   1380
   End
End
Attribute VB_Name = "frm_item_master2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txt_box, combo, Label
Dim jumlah_titik As Integer, l_txt_item_code As String
Dim l_txt_item_code2 As String, l_cmb_stock_control As String

Dim f_loss As Boolean
Public k_pertama As Boolean
Public Status As String
Public rs_item_master As New ADODB.Recordset
Dim rs_delivery As New ADODB.Recordset
Dim rs_hs_master As New ADODB.Recordset
Dim rs_trade_master As New ADODB.Recordset
Dim rs_material_cls As New ADODB.Recordset
Dim rs_manufacture_line2 As New ADODB.Recordset
Dim rs_manufacture_line As New ADODB.Recordset
Dim rs_sheetcoil_cls As New ADODB.Recordset
Dim rs_drawingmaterial_cls As New ADODB.Recordset
Dim rs_surfacetreatment_cls As New ADODB.Recordset
Dim rs_heattreatment_cls As New ADODB.Recordset
Dim rs_group_cls As New ADODB.Recordset
Dim rs_packingstyle_cls As New ADODB.Recordset
Dim rs_control_cls As New ADODB.Recordset
Dim rs_personincharge_cls As New ADODB.Recordset
Dim rs_bom_master As New ADODB.Recordset
Dim rs_line As New ADODB.Recordset
Dim rs_clasification As New ADODB.Recordset
Dim rs_destination As New ADODB.Recordset
Dim rs_color As New ADODB.Recordset

Private Sub koneksi()
rs_bom_master.Open " select * from bom_master", Db, adOpenKeyset, adLockOptimistic
rs_item_master.Open " select * from item_master", Db, adOpenKeyset, adLockOptimistic
rs_hs_master.Open " select * from HS_Master", Db, adOpenKeyset, adLockOptimistic
rs_trade_master.Open " select * from trade_master", Db, adOpenKeyset, adLockOptimistic
rs_material_cls.Open " select * from material_cls", Db, adOpenKeyset, adLockOptimistic
rs_manufacture_line2.Open "select distinct(manufacture_line.manufacture_code),trade_name from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code", Db, adOpenKeyset, adLockOptimistic
rs_manufacture_line.Open "select manufacture_line.manufacture_code,trade_name from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code", Db, adOpenKeyset, adLockOptimistic
rs_sheetcoil_cls.Open "select * from sheetcoil_cls", Db, adOpenKeyset, adLockOptimistic
rs_drawingmaterial_cls.Open "select * from drawingmaterial_cls", Db, adOpenKeyset, adLockOptimistic
rs_heattreatment_cls.Open "select * from heattreatment_cls", Db, adOpenKeyset, adLockOptimistic
rs_surfacetreatment_cls.Open "select * from surfacetreatment_cls", Db, adOpenKeyset, adLockOptimistic
rs_group_cls.Open "select * from group_cls", Db, adOpenKeyset, adLockOptimistic
rs_packingstyle_cls.Open "select * from packingstyle_cls", Db, adOpenKeyset, adLockOptimistic
rs_control_cls.Open "select * from control_cls", Db, adOpenKeyset, adLockOptimistic
rs_personincharge_cls.Open "select * from personincharge_cls", Db, adOpenKeyset, adLockOptimistic
rs_line.Open "select * from manufacture_line", Db, adOpenKeyset, adLockOptimistic
rs_clasification.Open "select * from ClasificationPart_Cls", Db, adOpenKeyset, adLockOptimistic
rs_destination.Open "select * from Destination_Cls", Db, adOpenKeyset, adLockOptimistic
rs_color.Open "select * from Color_Cls", Db, adOpenKeyset, adLockOptimistic
End Sub

Private Sub cbo_POType_Click()
lbl_POType.Caption = cbo_POType.List(cbo_POType.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cbo_POType_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_POType.Caption = ""
If KeyCode = vbKeyBack Then lbl_POType.Caption = ""
End Sub

Private Sub cboClasificationPart_Click()
lblClasificationPart.Caption = cboClasificationPart.List(cboClasificationPart.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cboColor_Click()
lbl_Color.Caption = cboColor.List(cboColor.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cboColor_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_Color.Caption = ""
If KeyCode = vbKeyBack Then lbl_Color.Caption = ""
End Sub

Private Sub cboDestination_Click()
lbl_Destination.Caption = cboDestination.List(cboDestination.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cboDestination_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_Destination.Caption = ""
If KeyCode = vbKeyBack Then lbl_Destination.Caption = ""
End Sub

Private Sub cboModel_Click()
lblModel.Caption = cboModel.List(cboModel.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

'Private Sub cboModel_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'If KeyCode = vbKeyDelete Then lblModel.Caption = ""
'If KeyCode = vbKeyBack Then lblModel.Caption = ""
'
'If KeyCode = 13 Then
'    If rs_model_cls.EOF = False Or rs_model_cls.BOF = False Then
'        rs_model_cls.MoveFirst
'        rs_model_cls.Find "model_cls='" & Trim(cboModel.Text) & "'"
'        If rs_model_cls.EOF = False Then
'            lblModel.Caption = Trim(rs_model_cls!Description)
'            lbl_pesan.Caption = ""
'        Else
'            lblModel.Caption = ""
'            lbl_pesan.Caption = DisplayMsg("0020")
'            cboModel.SetFocus
'        End If
'    End If
'End If
'End Sub

Private Sub cmb_delivery_click()
lbl_delivery_place.Caption = cmb_delivery.List(cmb_delivery.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmb_delivery_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_delivery_place.Caption = ""
If KeyCode = vbKeyBack Then lbl_delivery_place.Caption = ""

If KeyCode = 13 Then
    If cmb_delivery.ListCount > 0 Then
        For i = 0 To cmb_delivery.ListCount - 1
            If UCase(Trim(cmb_delivery.Text)) = UCase(Trim(cmb_delivery.List(i, 0))) Then
                  cmb_delivery = Trim(cmb_delivery.List(i, 0))
                  lbl_delivery_place.Caption = Trim(cmb_delivery.List(i, 1))
                  lbl_pesan.Caption = ""
                Exit For
            Else
              lbl_delivery_place.Caption = ""
              lbl_pesan.Caption = DisplayMsg("0014")
              cmb_delivery.SetFocus
            End If
        Next
    Else
        If Trim(cmb_delivery.Text) <> "" Then
            lbl_delivery_place.Caption = ""
            lbl_pesan.Caption = DisplayMsg("0014")
            cmb_delivery.SetFocus
        End If
    End If
End If
End Sub

Private Sub cmb_delivery_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmb_explosion2_Click()
lbl_explosion.Caption = cmb_explosion2.List(cmb_explosion2.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmb_explosion2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_explosion.Caption = ""
If KeyCode = vbKeyBack Then lbl_explosion.Caption = "" ' KeyCode = 0
If KeyCode = 13 Then
    If cmb_explosion2.ListCount < 1 Then Exit Sub
    For i = 0 To cmb_explosion2.ListCount - 1
        If Trim(cmb_explosion2.Text) = cmb_explosion2.List(i, 0) Then
            lbl_explosion.Caption = cmb_explosion2.List(i, 1): lbl_pesan.Caption = "": Exit Sub
        End If
    Next
    lbl_pesan.Caption = DisplayMsg("0015")
    cmb_explosion2.SetFocus
    SSTab1.Tab = 0
    lbl_explosion.Caption = ""
End If
End Sub

Private Sub cmb_finish_good_Click()
lbl_finish.Caption = cmb_finish_good.List(cmb_finish_good.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmb_finish_good_GotFocus()
SSTab1.Tab = 0
End Sub

Private Sub cmb_finish_good_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyBack Then lbl_finish.Caption = "" 'KeyCode = 0
If KeyCode = vbKeyDelete Then lbl_finish.Caption = ""
If KeyCode = 13 Then
    If cmb_finish_good.ListCount < 1 Then Exit Sub
    For i = 0 To cmb_finish_good.ListCount - 1
        If Trim(cmb_finish_good.Text) = cmb_finish_good.List(i, 0) Then
            cmb_finish_good = cmb_finish_good.List(i, 0)
            lbl_finish.Caption = cmb_finish_good.List(i, 1)
            lbl_pesan.Caption = "": Exit Sub
        End If
    Next
    lbl_pesan.Caption = DisplayMsg("8057")
     lbl_finish.Caption = ""
    cmb_finish_good.SetFocus
    SSTab1.Tab = 0
End If
End Sub

Private Sub cmb_make_Click()
Lbl_Make.Caption = cmb_make.List(cmb_make.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmb_make_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then Lbl_Make.Caption = ""
If KeyCode = vbKeyBack Then Lbl_Make.Caption = ""  'KeyCode = 0
If KeyCode = 13 Then
    If cmb_make.ListCount < 1 Then Exit Sub
    For i = 0 To cmb_make.ListCount - 1
        If Trim(cmb_make.Text) = cmb_make.List(i, 0) Then
            Lbl_Make.Caption = cmb_make.List(i, 1): lbl_pesan.Caption = "": Exit Sub
        End If
    Next
    lbl_pesan.Caption = DisplayMsg("8096")
    cmb_make.SetFocus
    SSTab1.Tab = 0
    Lbl_Make.Caption = ""
End If
End Sub


Private Sub cmb_part2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_part.Caption = ""
If KeyCode = vbKeyBack Then lbl_part.Caption = ""  'KeyCode = 0
If KeyCode = 13 Then
    If cmb_part2.ListCount < 1 Then Exit Sub
    For i = 0 To cmb_part2.ListCount - 1
        If Trim(cmb_part2.Text) = cmb_part2.List(i, 0) Then
            lbl_part.Caption = cmb_part2.List(i, 1): lbl_pesan.Caption = "": Exit Sub
        End If
    Next
    lbl_pesan.Caption = DisplayMsg("0016")
    cmb_part2.SetFocus
    SSTab1.Tab = 0
    lbl_part.Caption = ""
End If
End Sub

Private Sub cmb_prod_Change()
If cmb_prod.ListIndex > -1 Then
lbl_prod.Caption = cmb_prod.List(cmb_prod.ListIndex, 1)
lbl_pesan.Caption = ""
End If
End Sub

Private Sub cmb_prod_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_group.Caption = ""
If KeyCode = vbKeyBack Then lbl_group.Caption = ""

If KeyCode = 13 Then
    If cmb_prod.ListCount < 1 Then Exit Sub
    For i = 0 To cmb_prod.ListCount - 1
        If Trim(cmb_prod.Text) = cmb_prod.List(i, 0) Then
            lbl_prod.Caption = cmb_prod.List(i, 1): lbl_pesan.Caption = "": Exit Sub
        End If
    Next
    lbl_pesan.Caption = DisplayMsg("0017")
    cmb_prod.SetFocus
    SSTab1.Tab = 0
    lbl_prod.Caption = ""
End If
End Sub

Private Sub cmb_prod_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmb_provision2_Click()
lbl_provotion.Caption = cmb_provision2.List(cmb_provision2.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmb_provision2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_provotion.Caption = ""
If KeyCode = vbKeyBack Then lbl_provotion.Caption = "" 'KeyCode = 0
If KeyCode = 13 Then
    If cmb_provision2.ListCount < 1 Then Exit Sub
    For i = 0 To cmb_provision2.ListCount - 1
        If Trim(cmb_provision2.Text) = cmb_provision2.List(i, 0) Then
            lbl_provotion.Caption = cmb_provision2.List(i, 1): lbl_pesan.Caption = "": Exit Sub
        End If
    Next
    lbl_pesan.Caption = DisplayMsg("0018")
    cmb_provision2.SetFocus
    SSTab1.Tab = 0
    lbl_provotion.Caption = ""
End If
End Sub

Private Sub cmb_reserve2_Click()
lbl_reserve.Caption = cmb_reserve2.List(cmb_reserve2.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmb_reserve2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_reserve.Caption = ""
If KeyCode = vbKeyBack Then lbl_reserve.Caption = "" 'KeyCode = 0
If KeyCode = 13 Then
    If cmb_reserve2.ListCount < 1 Then Exit Sub
    For i = 0 To cmb_reserve2.ListCount - 1
        If Trim(cmb_reserve2.Text) = cmb_reserve2.List(i, 0) Then
            lbl_reserve.Caption = cmb_reserve2.List(i, 1): lbl_pesan.Caption = "": Exit Sub
        End If
    Next
    lbl_pesan.Caption = DisplayMsg("0023")
    cmb_reserve2.SetFocus
    SSTab1.Tab = 0
    lbl_reserve.Caption = ""
End If
End Sub

Private Sub cmb_stock_control2_Click()
If Trim(cmb_stock_control2.List(cmb_stock_control2.ListIndex, 1)) <> "" Then
    lbl_stock_control.Caption = Trim(cmb_stock_control2.List(cmb_stock_control2.ListIndex, 1))
Else
    lbl_stock_control.Caption = ""
End If
lbl_pesan.Caption = ""
End Sub

Private Sub cmb_stock_control2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_stock_control.Caption = ""
If KeyCode = vbKeyBack Then lbl_stock_control.Caption = ""
If KeyCode = 13 Then
    If cmb_stock_control2.ListCount < 1 Then Exit Sub
    For i = 0 To cmb_stock_control2.ListCount - 1
        If Trim(cmb_stock_control2.Text) = cmb_stock_control2.List(i, 0) Then
            lbl_stock_control.Caption = cmb_stock_control2.List(i, 1): lbl_pesan.Caption = "": Exit Sub
        End If
    Next
    lbl_pesan.Caption = DisplayMsg("0019")
    cmb_stock_control2.SetFocus
    SSTab1.Tab = 0
    lbl_stock_control.Caption = ""
End If
End Sub

Private Sub cmb_supply2_Click()
lbl_supply.Caption = cmb_supply2.List(cmb_supply2.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmb_supply2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_supply.Caption = ""
If KeyCode = vbKeyBack Then lbl_supply.Caption = ""  'KeyCode = 0
If KeyCode = 13 Then
    If cmb_supply2.ListCount < 1 Then Exit Sub
    For i = 0 To cmb_supply2.ListCount - 1
        If Trim(cmb_supply2.Text) = cmb_supply2.List(i, 0) Then
            lbl_supply.Caption = cmb_supply2.List(i, 1): lbl_pesan.Caption = "": Exit Sub
        End If
    Next
    lbl_pesan.Caption = DisplayMsg("4056")
    cmb_supply2.SetFocus
    SSTab1.Tab = 0
    lbl_supply.Caption = ""
End If
End Sub

Private Sub cmbbox_control_Click()
 lbl_control.Caption = cmbbox_control.List(cmbbox_control.ListIndex, 1)
 lbl_pesan.Caption = ""
End Sub

Private Sub cmbbox_control_Keydown(KeyAscii As MSForms.ReturnInteger, Shift As Integer)
If KeyAscii = vbKeyDelete Then lbl_control.Caption = ""
If KeyAscii = vbKeyBack Then lbl_control.Caption = ""

If KeyAscii = 13 Then
    If rs_control_cls.EOF = False Or rs_control_cls.BOF = False Then
        rs_control_cls.MoveFirst
        rs_control_cls.Find "control_cls='" & Trim(cmbbox_control.Text) & "'"
        If rs_control_cls.EOF = False Then
            lbl_control.Caption = Trim(rs_control_cls!Description)
            lbl_pesan.Caption = ""
        Else
            lbl_control.Caption = ""
            lbl_pesan.Caption = DisplayMsg("0020")
            cmbbox_control.SetFocus
        End If
    End If
End If
End Sub

Private Sub cmbbox_drawing_click()
  lbl_drawing.Caption = cmbbox_drawing.List(cmbbox_drawing.ListIndex, 1)
  lbl_pesan.Caption = ""
End Sub

Private Sub cmbbox_drawing_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_drawing.Caption = ""
If KeyCode = vbKeyBack Then lbl_drawing.Caption = ""

If KeyCode = 13 Then
    If rs_drawingmaterial_cls.EOF = False Or rs_drawingmaterial_cls.BOF = False Then
        rs_drawingmaterial_cls.MoveFirst
        rs_drawingmaterial_cls.Find "drawingmaterial_cls='" & Trim(cmbbox_drawing.Text) & "'"
        If rs_drawingmaterial_cls.EOF = False Then
            lbl_drawing.Caption = Trim(rs_drawingmaterial_cls!Description)
            lbl_pesan.Caption = ""
        Else
            lbl_drawing.Caption = ""
            lbl_pesan.Caption = DisplayMsg("0021")
            cmbbox_drawing.SetFocus
        End If
    End If
End If
End Sub

Private Sub cmbbox_drawing_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmbbox_group_Click()
lbl_group.Caption = cmbbox_group.List(cmbbox_group.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmbbox_group_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_group.Caption = ""
If KeyCode = vbKeyBack Then lbl_group.Caption = ""

If KeyCode = 13 Then
    If rs_group_cls.EOF = False Or rs_group_cls.BOF = False Then
        rs_group_cls.MoveFirst
        rs_group_cls.Find "group_cls='" & Trim(cmbbox_group.Text) & "'"
        If rs_group_cls.EOF = False Then
            lbl_group.Caption = Trim(rs_group_cls!Description)
            lbl_pesan.Caption = ""
        Else
            lbl_group.Caption = ""
            lbl_pesan.Caption = DisplayMsg("8056")
            cmbbox_group.SetFocus
        End If
    End If
End If
End Sub

Private Sub cmbbox_group_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmbbox_heat_Click()
  lbl_heat.Caption = cmbbox_heat.List(cmbbox_heat.ListIndex, 1)
  lbl_pesan.Caption = ""
End Sub

Private Sub cmbbox_heat_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_heat.Caption = ""
If KeyCode = vbKeyBack Then lbl_heat.Caption = ""

If KeyCode = 13 Then
    If rs_heattreatment_cls.EOF = False Or rs_heattreatment_cls.BOF = False Then
        rs_heattreatment_cls.MoveFirst
        rs_heattreatment_cls.Find "heattreatment_cls='" & Trim(cmbbox_heat.Text) & "'"
        If rs_heattreatment_cls.EOF = False Then
            lbl_heat.Caption = Trim(rs_heattreatment_cls!Description)
            lbl_pesan.Caption = ""
        Else
            lbl_heat.Caption = ""
            lbl_pesan.Caption = DisplayMsg("0022")
            cmbbox_heat.SetFocus
        End If
    End If
End If
End Sub

Private Sub cmbbox_heat_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmbbox_line_Click()
lbl_line.Caption = cmbbox_line.List(cmbbox_line.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmbbox_line_GotFocus()
SSTab1.Tab = 1
End Sub

Private Sub cmbbox_line_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_line.Caption = ""
If KeyCode = vbKeyBack Then lbl_line.Caption = ""

If KeyCode = 13 Then
    If rs_line.EOF = False Or rs_line.BOF = False Then
        rs_line.MoveFirst
        rs_line.Find "line_code ='" & Trim(cmbbox_line.Text) & "'"
        If rs_line.EOF = False Then
            cmbbox_line = Trim(rs_line!line_code)
            lbl_line.Caption = Trim(rs_line!line_name)
            lbl_pesan.Caption = ""
        Else
            lbl_line.Caption = ""
            lbl_pesan.Caption = DisplayMsg("8009")
            cmbbox_line.SetFocus
        End If
    End If
End If
End Sub

Private Sub cmbbox_line_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmbbox_material_Click()
lbl_material.Caption = cmbbox_material.List(cmbbox_material.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub


Private Sub cmbbox_material_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_material.Caption = ""
If KeyCode = vbKeyBack Then lbl_material.Caption = ""

If KeyCode = 13 Then
    If rs_material_cls.EOF = False Or rs_material_cls.BOF = False Then
        rs_material_cls.MoveFirst
        rs_material_cls.Find "material_cls='" & Trim(cmbbox_material.Text) & "'"
        If rs_material_cls.EOF = False Then
            cmbbox_material.Text = Trim(rs_material_cls!Material_Cls)
            lbl_material.Caption = Trim(rs_material_cls!Description)
            lbl_pesan.Caption = ""
        Else
            lbl_material.Caption = ""
            lbl_pesan.Caption = DisplayMsg("8095")
            cmbbox_material.SetFocus
        End If
    End If
End If
End Sub

Private Sub cmbbox_material_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmbbox_packing_Click()
lbl_packing.Caption = cmbbox_packing.List(cmbbox_packing.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmbbox_packing_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

If KeyCode = vbKeyDelete Then lbl_packing.Caption = ""
If KeyCode = vbKeyBack Then lbl_packing.Caption = ""

If KeyCode = 13 Then
    If rs_packingstyle_cls.EOF = False Or rs_packingstyle_cls.BOF = False Then
        rs_packingstyle_cls.MoveFirst
        rs_packingstyle_cls.Find "packingstyle_cls='" & Trim(cmbbox_packing.Text) & "'"
        If rs_packingstyle_cls.EOF = False Then
            lbl_packing.Caption = Trim(rs_packingstyle_cls!Description)
            lbl_pesan.Caption = ""
        Else
            lbl_packing.Caption = ""
            lbl_pesan.Caption = DisplayMsg("0024")
            cmbbox_packing.SetFocus
        End If
    End If
End If

End Sub

Private Sub cmbbox_packing_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmbbox_packing2_Click()
lbl_packing2.Caption = cmbbox_packing2.List(cmbbox_packing2.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmbbox_packing2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_packing2.Caption = ""
If KeyCode = vbKeyBack Then lbl_packing2.Caption = ""

If KeyCode = 13 Then
    If rs_packingstyle_cls.EOF = False Or rs_packingstyle_cls.BOF = False Then
        rs_packingstyle_cls.MoveFirst
        rs_packingstyle_cls.Find "packingstyle_cls='" & Trim(cmbbox_packing2.Text) & "'"
        If rs_packingstyle_cls.EOF = False Then
            lbl_packing2.Caption = Trim(rs_packingstyle_cls!Description)
            lbl_pesan.Caption = ""
        Else
            lbl_packing2.Caption = ""
            lbl_pesan.Caption = DisplayMsg("0025")
            cmbbox_packing2.SetFocus
        End If
    End If
End If
End Sub

Private Sub cmbbox_packing2_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmbbox_purchase_person_Click()
lbl_purchase.Caption = cmbbox_purchase_person.List(cmbbox_purchase_person.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmbbox_purchase_person_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_purchase.Caption = ""
If KeyCode = vbKeyBack Then lbl_purchase.Caption = ""

If KeyCode = 13 Then
    If rs_personincharge_cls.EOF = False Or rs_personincharge_cls.BOF = False Then
        rs_personincharge_cls.MoveFirst
        rs_personincharge_cls.Find "personincharge_cls='" & Trim(cmbbox_purchase_person.Text) & "'"
        If rs_personincharge_cls.EOF = False Then
            lbl_purchase.Caption = Trim(rs_personincharge_cls!Description)
            lbl_pesan.Caption = ""
        Else
            lbl_purchase.Caption = ""
            lbl_pesan.Caption = DisplayMsg("0026")
            cmbbox_purchase_person.SetFocus
        End If
    End If
End If
End Sub

Private Sub cmbbox_purchase_person_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmbbox_sheet_Click()
lbl_sheet.Caption = cmbbox_sheet.List(cmbbox_sheet.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmbbox_sheet_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_sheet.Caption = ""

If KeyCode = 13 Then
    If rs_sheetcoil_cls.EOF = False Or rs_sheetcoil_cls.BOF = False Then
        rs_sheetcoil_cls.MoveFirst
        rs_sheetcoil_cls.Find "sheetcoil_cls='" & Trim(cmbbox_sheet.Text) & "'"
        If rs_sheetcoil_cls.EOF = False Then
            lbl_sheet.Caption = Trim(rs_sheetcoil_cls!Description)
            lbl_pesan.Caption = ""
        Else
            lbl_sheet.Caption = ""
            lbl_pesan.Caption = DisplayMsg("0027")
            cmbbox_sheet.SetFocus
        End If
    End If
End If
End Sub

Private Sub cmbbox_sheet_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmbbox_surface_click()
lbl_surface.Caption = cmbbox_surface.List(cmbbox_surface.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmbbox_surface_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_surface.Caption = ""
If KeyCode = vbKeyBack Then lbl_surface.Caption = ""

If KeyCode = 13 Then
    If rs_surfacetreatment_cls.EOF = False Or rs_surfacetreatment_cls.BOF = False Then
        rs_surfacetreatment_cls.MoveFirst
        rs_surfacetreatment_cls.Find "surfacetreatment_cls='" & Trim(cmbbox_surface.Text) & "'"
        If rs_surfacetreatment_cls.EOF = False Then
            lbl_surface.Caption = Trim(rs_surfacetreatment_cls!Description)
            lbl_pesan.Caption = ""
        Else
            lbl_surface.Caption = ""
            lbl_pesan.Caption = DisplayMsg("0028")
            cmbbox_surface.SetFocus
        End If
    End If
End If
End Sub

Private Sub cmbbox_surface_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmbbox_unit_Click()
lbl_unit.Caption = cmbbox_unit.List(cmbbox_unit.ListIndex, 1)
lbl_pesan.Caption = ""
End Sub

Private Sub cmbbox_unit_GotFocus()
SSTab1.Tab = 0
End Sub

Private Sub cmbbox_unit_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyBack Then lbl_unit.Caption = ""
If KeyCode = vbKeyDelete Then lbl_unit.Caption = ""

If KeyCode = 13 Then
    If cmbbox_unit.ListCount < 1 Then Exit Sub
    For i = 0 To cmbbox_unit.ListCount - 1
        If Trim(cmbbox_unit.Text) = cmbbox_unit.List(i, 0) Then
            lbl_unit.Caption = cmbbox_unit.List(i, 1): lbl_pesan.Caption = "": Exit Sub
        End If
    Next
    lbl_pesan.Caption = DisplayMsg("1010")
    cmbbox_unit.SetFocus
    SSTab1.Tab = 0
    lbl_unit.Caption = ""
End If
End Sub

Private Sub cmb_hs_Change()
lbl_pesan.Caption = ""
End Sub

Private Sub cmb_hs_Click()
lbl_hs.Caption = cmb_hs.List(cmb_hs.ListIndex, 1)
End Sub

Private Sub cmb_hs_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_hs.Caption = ""
If KeyCode = vbKeyBack Then lbl_hs.Caption = ""

If KeyCode = 13 Then
    If cmb_hs.ListCount > 0 Then
        For i = 0 To cmb_hs.ListCount - 1
            If UCase(Trim(cmb_hs.Text)) = UCase(Trim(cmb_hs.List(i, 0))) Then
                  cmb_hs.Text = Trim(cmb_hs.List(i, 0))
                  lbl_hs.Caption = Trim(cmb_hs.List(i, 1))
                  lbl_pesan.Caption = ""
                Exit For
            Else
              lbl_hs.Caption = ""
              lbl_pesan.Caption = DisplayMsg(8060)
              cmb_hs.SetFocus
            End If
        Next
    End If
End If
End Sub

Private Sub cmb_hs_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub


Private Sub cmbbox_warehouse_Change()
    lbl_pesan.Caption = ""
    lbl_warehouse.Caption = ""
    If cmbbox_warehouse.MatchFound Then
        lbl_warehouse.Caption = cmbbox_warehouse.List(cmbbox_warehouse.ListIndex, 1)
    End If
End Sub

Private Sub cmbbox_warehouse_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub set_line()
cmbbox_line.clear
lbl_line.Caption = ""

If rs_manufacture_line2.EOF = False Or rs_manufacture_line2.BOF = False Then
    If rs_line.State <> adStateClosed Then rs_line.Close
    rs_line.Open " select * from manufacture_line where manufacture_code='" & Trim(cmbox_manufacture.Text) & "' "
    
    If Not (rs_line.EOF And rs_line.BOF) Then
     
        rs_line.MoveFirst
        i = 1
        cmbbox_line.AddItem ""
        cmbbox_line.List(0, 0) = ""
        cmbbox_line.List(0, 1) = ""
        While Not rs_line.EOF
            cmbbox_line.AddItem ""
            cmbbox_line.List(i, 0) = rs_line!line_code
            cmbbox_line.List(i, 1) = rs_line!line_name
            i = i + 1
            rs_line.MoveNext
        Wend
    End If
        
End If
End Sub

Private Sub set_Clasification()
cboClasificationPart.clear
lblClasificationPart.Caption = ""

    If rs_clasification.State <> adStateClosed Then rs_clasification.Close
    rs_clasification.Open " select * from ClasificationPart_Cls where ClasificationPart_Cls='" & Trim(cboClasificationPart.Text) & "' "
    
    If Not (rs_clasification.EOF And rs_clasification.BOF) Then
     
        rs_clasification.MoveFirst
        i = 1
        cboClasificationPart.AddItem ""
        cboClasificationPart.List(0, 0) = ""
        cboClasificationPart.List(0, 1) = ""
        While Not rs_clasification.EOF
            cboClasificationPart.AddItem ""
            cboClasificationPart.List(i, 0) = rs_clasification!ClasificationPart_Cls
            cboClasificationPart.List(i, 1) = rs_clasification!Description
            i = i + 1
            rs_clasification.MoveNext
        Wend
    End If
End Sub

Private Sub set_Destination()
cboDestination.clear
lbl_Destination.Caption = ""

    If rs_destination.State <> adStateClosed Then rs_destination.Close
    rs_destination.Open " select * from Destination_Cls where Destination_Cls='" & Trim(cboDestination.Text) & "' "
    
    If Not (rs_destination.EOF And rs_destination.BOF) Then
     
        rs_destination.MoveFirst
        i = 1
        cboDestination.AddItem ""
        cboDestination.List(0, 0) = ""
        cboDestination.List(0, 1) = ""
        While Not rs_destination.EOF
            cboDestination.AddItem ""
            cboDestination.List(i, 0) = rs_destination!Destination_Cls
            cboDestination.List(i, 1) = rs_destination!Description
            i = i + 1
            rs_destination.MoveNext
        Wend
    End If
End Sub

Private Sub set_Color()
cboColor.clear
lbl_Color.Caption = ""

    If rs_color.State <> adStateClosed Then rs_color.Close
    rs_color.Open " select * from Color_Cls where Color_Cls='" & Trim(cboColor.Text) & "' "
    
    If Not (rs_color.EOF And rs_color.BOF) Then
     
        rs_color.MoveFirst
        i = 1
        cboColor.AddItem ""
        cboColor.List(0, 0) = ""
        cboColor.List(0, 1) = ""
        While Not rs_color.EOF
            cboColor.AddItem ""
            cboColor.List(i, 0) = rs_color!Color_Cls
            cboColor.List(i, 1) = rs_color!Description
            i = i + 1
            rs_color.MoveNext
        Wend
    End If
End Sub

Private Sub cmbox_manufacture_Click()

Call set_line

If cmbox_manufacture.ListIndex < 0 Then Exit Sub
txt_factory = cmbox_manufacture.List(cmbox_manufacture.ListIndex, 1)
lbl_pesan.Caption = ""

End Sub

Private Sub text_clear()
For Each txt_box In frm_item_master2.Controls
   If TypeOf txt_box Is TextBox Then txt_box.Text = ""
Next
End Sub

Private Sub combo_clear()
For Each combo In frm_item_master2.Controls
    If TypeOf combo Is ComboBox Then combo.Text = ""
Next
For Each combo In frm_item_master2.Controls
    If TypeOf combo Is MSForms.ComboBox Then combo.Text = ""
Next
End Sub

Private Sub label_clear()

lbl_line.Caption = ""
lbl_warehouse.Caption = ""
lbl_hs.Caption = ""
lbl_mPacking = ""
txt_supplier = ""
lbl_delivery_place.Caption = ""

txt_factory = ""
lbl_unit.Caption = ""

lbl_finish.Caption = ""
lbl_part.Caption = ""
lbl_provotion.Caption = ""
lbl_reserve.Caption = ""
lbl_supply.Caption = ""
lbl_prod.Caption = ""

lbl_material.Caption = ""
lbl_drawing.Caption = ""
lbl_surface.Caption = ""
lbl_heat.Caption = ""
lbl_sheet.Caption = ""

lbl_packing.Caption = ""
lbl_packing2.Caption = ""
lbl_purchase.Caption = ""
Lbl_Make.Caption = ""
lbl_stock_control.Caption = ""
lbl_control.Caption = ""
lbl_group.Caption = ""
lbl_explosion.Caption = ""
lblModel.Caption = ""
lbl_POType.Caption = ""
lblClasificationPart.Caption = ""
lbl_Destination.Caption = ""
lbl_Color.Caption = ""

End Sub

Private Sub setting()

Dim rs_warehouse_master As New ADODB.Recordset

sql = "select wh_code,wh_name,stockControl_cls from warehouse_master " & _
    "Union All " & _
    "select distinct(manufacture_line.manufacture_code)wh_code,trade_name wh_name,stockControl_Cls='01' " & _
    "from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code"
rs_warehouse_master.Open sql, Db, adOpenKeyset, adLockOptimistic

'=====================Setting Combo Material Packing =======================
cmbbox_mPacking.clear
cmbbox_mPacking.columnCount = 2
cmbbox_mPacking.TextColumn = 1
i = 1

If rs_packingstyle_cls.EOF = False Or rs_packingstyle_cls.BOF = False Then
    rs_packingstyle_cls.MoveFirst
    cmbbox_mPacking.AddItem ""
    cmbbox_mPacking.List(0, 0) = ""
    cmbbox_mPacking.List(0, 1) = ""
    While rs_packingstyle_cls.EOF = False
        cmbbox_mPacking.AddItem ""
        cmbbox_mPacking.List(i, 0) = rs_packingstyle_cls!PackingStyle_Cls
        cmbbox_mPacking.List(i, 1) = rs_packingstyle_cls!Description
        rs_packingstyle_cls.MoveNext
        i = i + 1
    Wend
    
    cmbbox_mPacking.ColumnWidths = "20 pt; 120 pt"
    cmbbox_mPacking.ListWidth = 140
    cmbbox_mPacking.ListIndex = 0
    cmbbox_mPacking.Text = cmbbox_mPacking.List(0, 0)
End If

'=====================Setting Combo Finish=========================
cmb_finish_good.clear
cmb_finish_good.columnCount = 2

cmb_finish_good.AddItem
cmb_finish_good.List(0, 0) = "01"
cmb_finish_good.List(0, 1) = "Finish Goods"
cmb_finish_good.AddItem
cmb_finish_good.List(1, 0) = "02"
cmb_finish_good.List(1, 1) = "Parts/WIP/Material"

cmb_finish_good.ListWidth = 110
cmb_finish_good.ColumnWidths = "20 pt ; 90 pt "
cmb_finish_good.ListIndex = 0
cmb_finish_good.Text = cmb_finish_good.List(0, 0)
'==============================================================
'=====================Setting Combo Part===========================
cmb_part2.clear
cmb_part2.columnCount = 2

cmb_part2.AddItem
cmb_part2.List(0, 0) = "01"
cmb_part2.List(0, 1) = ""
cmb_part2.AddItem
cmb_part2.List(1, 0) = "02"
cmb_part2.List(1, 1) = ""
cmb_part2.AddItem
cmb_part2.List(2, 0) = "03"
cmb_part2.List(2, 1) = ""
cmb_part2.AddItem
cmb_part2.List(3, 0) = "04"
cmb_part2.List(3, 1) = ""

cmb_part2.ListWidth = 70
cmb_part2.ColumnWidths = "20 pt ; 50 pt "
cmb_part2.ListIndex = 0
cmb_part2.Text = cmb_part2.List(0, 0)
'==============================================================
'=====================Setting Combo production cls=======================
cmb_prod.clear
cmb_prod.columnCount = 2

cmb_prod.AddItem
cmb_prod.List(0, 0) = "01"
cmb_prod.List(0, 1) = "Yes"
cmb_prod.AddItem
cmb_prod.List(1, 0) = "02"
cmb_prod.List(1, 1) = "No"

cmb_prod.ListWidth = 70
cmb_prod.ColumnWidths = "20 pt ; 50 pt "
cmb_prod.ListIndex = 0
cmb_prod.Text = cmb_prod.List(0, 0)
'==============================================================

'=====================Setting Combo Reserve=======================
cmb_reserve2.clear
cmb_reserve2.columnCount = 2

cmb_reserve2.AddItem
cmb_reserve2.List(0, 0) = "01"
cmb_reserve2.List(0, 1) = "Yes"
cmb_reserve2.AddItem
cmb_reserve2.List(1, 0) = "02"
cmb_reserve2.List(1, 1) = "No"

cmb_reserve2.ListWidth = 70
cmb_reserve2.ColumnWidths = "20 pt ; 50 pt "
cmb_reserve2.ListIndex = 0
cmb_reserve2.Text = cmb_reserve2.List(0, 0)
'==============================================================
'=====================Setting Combo Supply===========================
cmb_supply2.clear
cmb_supply2.columnCount = 2

cmb_supply2.AddItem
cmb_supply2.List(0, 0) = "01"
cmb_supply2.List(0, 1) = "Yes"
cmb_supply2.AddItem
cmb_supply2.List(1, 0) = "02"
cmb_supply2.List(1, 1) = "No"

cmb_supply2.ListWidth = 70
cmb_supply2.ColumnWidths = "20 pt ; 50 pt "
cmb_supply2.ListIndex = 0
cmb_supply2.Text = cmb_supply2.List(0, 0)
'==============================================================
'=====================Setting Combo Provision========================
cmb_provision2.clear
cmb_provision2.columnCount = 2

cmb_provision2.AddItem
cmb_provision2.List(0, 0) = "01"
cmb_provision2.List(0, 1) = "Yes"
cmb_provision2.AddItem
cmb_provision2.List(1, 0) = "02"
cmb_provision2.List(1, 1) = "No"

cmb_provision2.ListWidth = 70
cmb_provision2.ColumnWidths = "20 pt ; 50 pt "
cmb_provision2.ListIndex = 0
cmb_provision2.Text = cmb_provision2.List(0, 0)
'==============================================================
'=====================Setting Combo Make / Buy======================
cmb_make.clear
cmb_make.columnCount = 2

cmb_make.AddItem
cmb_make.List(0, 0) = "01"
cmb_make.List(0, 1) = "Make"
cmb_make.AddItem
cmb_make.List(1, 0) = "02"
cmb_make.List(1, 1) = "Buy"

cmb_make.ListWidth = 70
cmb_make.ColumnWidths = "20 pt ; 50 pt "
cmb_make.ListIndex = 0
cmb_make.Text = cmb_make.List(0, 0)
'==============================================================
'=====================Setting Combo explosion===================
cmb_explosion2.clear
cmb_explosion2.columnCount = 2

cmb_explosion2.AddItem
cmb_explosion2.List(0, 0) = "01"
cmb_explosion2.List(0, 1) = "All"
cmb_explosion2.AddItem
cmb_explosion2.List(1, 0) = "02"
cmb_explosion2.List(1, 1) = "1 Level"

cmb_explosion2.ListWidth = 90
cmb_explosion2.ColumnWidths = "20 pt ; 70 pt "
cmb_explosion2.ListIndex = 0
cmb_explosion2.Text = cmb_explosion2.List(0, 0)
'==============================================================

'=====================Setting Combo Stock Control 2======================
cmb_stock_control2.clear
cmb_stock_control2.columnCount = 2
cmb_stock_control2.TextColumn = 1
cmb_stock_control2.AddItem ""
cmb_stock_control2.List(0, 0) = "01"
cmb_stock_control2.List(0, 1) = "Yes"
cmb_stock_control2.AddItem ""
cmb_stock_control2.List(1, 0) = "02"
cmb_stock_control2.List(1, 1) = "No"
cmb_stock_control2.ColumnWidths = "20 pt; 30 pt"
cmb_stock_control2.ListWidth = 50
cmb_stock_control2.ListIndex = 0
'==============================================================
'=====================Setting Combo line=================
cmbbox_line.clear
cmbbox_line.columnCount = 2
cmbbox_line.TextColumn = 1
cmbbox_line.ColumnWidths = "50 pt; 120 pt"
cmbbox_line.ListWidth = 170
'==============================================================

'=====================Setting Combo Drawing Material=================
cmbbox_drawing.clear
cmbbox_drawing.columnCount = 2
cmbbox_drawing.TextColumn = 1
i = 1

If rs_drawingmaterial_cls.EOF = False Or rs_drawingmaterial_cls.BOF = False Then
    rs_drawingmaterial_cls.MoveFirst
        cmbbox_drawing.AddItem ""
        cmbbox_drawing.List(0, 0) = ""
        cmbbox_drawing.List(0, 1) = ""
    While rs_drawingmaterial_cls.EOF = False
        cmbbox_drawing.AddItem ""
        cmbbox_drawing.List(i, 0) = rs_drawingmaterial_cls!drawingmaterial_cls
        cmbbox_drawing.List(i, 1) = rs_drawingmaterial_cls!Description
        rs_drawingmaterial_cls.MoveNext
        i = i + 1
    Wend
    
    cmbbox_drawing.ColumnWidths = "20 pt; 120 pt"
    cmbbox_drawing.ListWidth = 140
    cmbbox_drawing.ListIndex = 0
    cmbbox_drawing.Text = cmbbox_drawing.List(0, 0)
End If
'==============================================================
'=====================Setting Combo Surface Treatment===============
cmbbox_surface.clear
cmbbox_surface.columnCount = 2
cmbbox_surface.TextColumn = 1
i = 1

If rs_surfacetreatment_cls.EOF = False Or rs_surfacetreatment_cls.BOF = False Then
    rs_surfacetreatment_cls.MoveFirst
    cmbbox_surface.AddItem ""
    cmbbox_surface.List(0, 0) = ""
    cmbbox_surface.List(0, 1) = ""
    While rs_surfacetreatment_cls.EOF = False
       cmbbox_surface.AddItem ""
        cmbbox_surface.List(i, 0) = rs_surfacetreatment_cls!surfacetreatment_cls
        cmbbox_surface.List(i, 1) = rs_surfacetreatment_cls!Description
        rs_surfacetreatment_cls.MoveNext
        i = i + 1
    Wend
    
    cmbbox_surface.ColumnWidths = "20 pt; 120 pt"
    cmbbox_surface.ListWidth = 140
    cmbbox_surface.ListIndex = 0
    cmbbox_surface = cmbbox_surface.List(0, 0)
End If
'==============================================================
'=====================Setting Combo heat Treatment===============
cmbbox_heat.clear
cmbbox_heat.columnCount = 2
cmbbox_heat.TextColumn = 1
i = 1

If rs_heattreatment_cls.EOF = False Or rs_heattreatment_cls.BOF = False Then
    rs_heattreatment_cls.MoveFirst
    cmbbox_heat.AddItem ""
    cmbbox_heat.List(0, 0) = ""
    cmbbox_heat.List(0, 1) = ""
    
    While rs_heattreatment_cls.EOF = False
        cmbbox_heat.AddItem ""
        cmbbox_heat.List(i, 0) = rs_heattreatment_cls!heattreatment_cls
        cmbbox_heat.List(i, 1) = rs_heattreatment_cls!Description
        rs_heattreatment_cls.MoveNext
        i = i + 1
    Wend
    
    cmbbox_heat.ColumnWidths = "20 pt; 120 pt"
    cmbbox_heat.ListWidth = 140
    cmbbox_heat.ListIndex = 0
    cmbbox_heat = cmbbox_heat.List(0, 0)
End If
'==============================================================
'=====================Setting Combo WareHouse=====================
cmbbox_warehouse.clear
cmbbox_warehouse.columnCount = 2
cmbbox_warehouse.TextColumn = 1
i = 0

If rs_warehouse_master.EOF = False Or rs_warehouse_master.BOF = False Then
    rs_warehouse_master.MoveFirst
    While rs_warehouse_master.EOF = False
        cmbbox_warehouse.AddItem ""
        cmbbox_warehouse.List(i, 0) = Trim(rs_warehouse_master!wh_code)
        cmbbox_warehouse.List(i, 1) = Trim(rs_warehouse_master!WH_Name)
        rs_warehouse_master.MoveNext
        i = i + 1
    Wend
    
    cmbbox_warehouse.ColumnWidths = "50 pt; 120 pt"
    cmbbox_warehouse.ListWidth = 170
    cmbbox_warehouse.ListIndex = 0
    cmbbox_warehouse.Text = cmbbox_warehouse.List(0, 0)
End If
'==============================================================
'=====================Setting Combo Sheet Coil=====================
cmbbox_sheet.clear
cmbbox_sheet.columnCount = 2
cmbbox_sheet.TextColumn = 1
i = 1

If rs_sheetcoil_cls.EOF = False Or rs_sheetcoil_cls.BOF = False Then
    rs_sheetcoil_cls.MoveFirst
    cmbbox_sheet.AddItem ""
    cmbbox_sheet.List(0, 0) = ""
    cmbbox_sheet.List(0, 1) = ""

    While rs_sheetcoil_cls.EOF = False
           cmbbox_sheet.AddItem ""
        cmbbox_sheet.List(i, 0) = rs_sheetcoil_cls!sheetcoil_cls
        cmbbox_sheet.List(i, 1) = rs_sheetcoil_cls!Description
         rs_sheetcoil_cls.MoveNext
        i = i + 1
    Wend
    
    cmbbox_sheet.ColumnWidths = "20 pt; 120 pt"
    cmbbox_sheet.ListWidth = 140
    cmbbox_sheet.ListIndex = 0
    cmbbox_sheet.Text = cmbbox_sheet.List(0, 0)
End If
'==============================================================
'=====================Setting Combo Manufacture=====================
cmbox_manufacture.clear
cmbox_manufacture.columnCount = 2
cmbox_manufacture.TextColumn = 1
i = 1

If rs_manufacture_line2.EOF = False Or rs_manufacture_line2.BOF = False Then
    rs_manufacture_line2.MoveFirst
    cmbox_manufacture.AddItem ""
        cmbox_manufacture.List(0, 0) = ""
        cmbox_manufacture.List(0, 1) = ""
    While rs_manufacture_line2.EOF = False
        cmbox_manufacture.AddItem ""
        cmbox_manufacture.List(i, 0) = Trim(rs_manufacture_line2!Manufacture_Code)
        cmbox_manufacture.List(i, 1) = Trim(rs_manufacture_line2!trade_name)
        rs_manufacture_line2.MoveNext
        i = i + 1
    Wend
    
    cmbox_manufacture.ColumnWidths = "50 pt; 350 pt"
    cmbox_manufacture.ListWidth = 400
    cmbox_manufacture.ListIndex = 0
    cmbox_manufacture.Text = cmbox_manufacture.List(0, 0)
End If
'==============================================================


'=====================Setting Combo Suplier=========================
cmbox_suplier.clear
cmbox_suplier.columnCount = 2
cmbox_suplier.TextColumn = 1
i = 0

If rs_trade_master.EOF = False Or rs_trade_master.BOF = False Then
    rs_trade_master.MoveFirst
    While rs_trade_master.EOF = False
        cmbox_suplier.AddItem ""
        cmbox_suplier.List(i, 0) = Trim(rs_trade_master!Trade_Code)
        cmbox_suplier.List(i, 1) = Trim(rs_trade_master!trade_name)
        rs_trade_master.MoveNext
        i = i + 1
    Wend
    
    cmbox_suplier.ColumnWidths = "50 pt; 350 pt"
    cmbox_suplier.ListWidth = 400
    cmbox_suplier.ListIndex = 0
    cmbox_suplier.Text = cmbox_suplier.List(0, 0)
End If
'=================================================================

'=====================Setting Combo HS Code=====================
cmb_hs.clear
cmb_hs.columnCount = 2
cmb_hs.TextColumn = 1
i = 0

If rs_hs_master.EOF = False Or rs_hs_master.BOF = False Then
    rs_hs_master.MoveFirst
    While rs_hs_master.EOF = False
        cmb_hs.AddItem ""
        cmb_hs.List(i, 0) = rs_hs_master!HS_Code
        cmb_hs.List(i, 1) = rs_hs_master!tax
        rs_hs_master.MoveNext
        i = i + 1
    Wend
    
    cmb_hs.ColumnWidths = "50 pt; 120 pt"
    cmb_hs.ListWidth = 170
    cmb_hs.ListIndex = 0
    cmb_hs.Text = cmb_hs.List(0, 0)
End If
'==============================================================

'=====================Setting Combo Material=======================
cmbbox_material.clear
cmbbox_material.columnCount = 2
cmbbox_material.TextColumn = 1
i = 1

If rs_material_cls.EOF = False Or rs_material_cls.BOF = False Then
    rs_material_cls.MoveFirst
         cmbbox_material.AddItem ""
        cmbbox_material.List(0, 0) = ""
        cmbbox_material.List(0, 1) = ""
    While rs_material_cls.EOF = False
       cmbbox_material.AddItem ""
        cmbbox_material.List(i, 0) = rs_material_cls!Material_Cls
        cmbbox_material.List(i, 1) = rs_material_cls!Description
        rs_material_cls.MoveNext
        i = i + 1
    Wend
    
    cmbbox_material.ColumnWidths = "20 pt; 120 pt"
    cmbbox_material.ListWidth = 140
    cmbbox_material.ListIndex = 0
    cmbbox_material.Text = cmbbox_material.List(0, 0)
End If
'==============================================================

'=====================Setting Combo Packing =======================
cmbbox_packing.clear
cmbbox_packing.columnCount = 2
cmbbox_packing.TextColumn = 1
i = 1

If rs_packingstyle_cls.EOF = False Or rs_packingstyle_cls.BOF = False Then
    rs_packingstyle_cls.MoveFirst
    cmbbox_packing.AddItem ""
    cmbbox_packing.List(0, 0) = ""
    cmbbox_packing.List(0, 1) = ""
    While rs_packingstyle_cls.EOF = False
        cmbbox_packing.AddItem ""
        cmbbox_packing.List(i, 0) = rs_packingstyle_cls!PackingStyle_Cls
        cmbbox_packing.List(i, 1) = rs_packingstyle_cls!Description
        rs_packingstyle_cls.MoveNext
        i = i + 1
    Wend
    
    cmbbox_packing.ColumnWidths = "20 pt; 120 pt"
    cmbbox_packing.ListWidth = 140
    cmbbox_packing.ListIndex = 0
    cmbbox_packing.Text = cmbbox_packing.List(0, 0)
End If
'==============================================================

'=====================Setting Combo Packing2 =======================
cmbbox_packing2.clear
cmbbox_packing2.columnCount = 2
cmbbox_packing2.TextColumn = 1
i = 1

If rs_packingstyle_cls.EOF = False Or rs_packingstyle_cls.BOF = False Then
    rs_packingstyle_cls.MoveFirst
    cmbbox_packing2.AddItem ""
    cmbbox_packing2.List(0, 0) = ""
    cmbbox_packing2.List(0, 1) = ""
    While rs_packingstyle_cls.EOF = False
        cmbbox_packing2.AddItem ""
        cmbbox_packing2.List(i, 0) = rs_packingstyle_cls!PackingStyle_Cls
        cmbbox_packing2.List(i, 1) = rs_packingstyle_cls!Description
        rs_packingstyle_cls.MoveNext
        i = i + 1
    Wend
    
    cmbbox_packing2.ColumnWidths = "20 pt; 120 pt"
    cmbbox_packing2.ListWidth = 140
    cmbbox_packing2.ListIndex = 0
    cmbbox_packing2.Text = cmbbox_packing2.List(0, 0)
End If
'==============================================================
'=====================Setting Combo Group=======================
cmbbox_group.clear
cmbbox_group.columnCount = 2
cmbbox_group.TextColumn = 1
i = 1

If rs_group_cls.EOF = False Or rs_group_cls.BOF = False Then
    rs_group_cls.MoveFirst
        cmbbox_group.AddItem ""
        cmbbox_group.List(0, 0) = ""
        cmbbox_group.List(0, 1) = ""
    While rs_group_cls.EOF = False
     cmbbox_group.AddItem ""
        cmbbox_group.List(i, 0) = rs_group_cls!group_cls
        cmbbox_group.List(i, 1) = rs_group_cls!Description
        rs_group_cls.MoveNext
        i = i + 1
    Wend
    
    cmbbox_group.ColumnWidths = "20 pt; 120 pt"
    cmbbox_group.ListWidth = 140
    cmbbox_group.ListIndex = 0
    cmbbox_group.Text = cmbbox_group.List(0, 0)
End If
'==============================================================
'=====================Setting Combo Control=======================
cmbbox_control.clear
cmbbox_control.columnCount = 2
cmbbox_control.TextColumn = 1
i = 0

If rs_control_cls.EOF = False Or rs_control_cls.BOF = False Then
    rs_control_cls.MoveFirst
    While rs_control_cls.EOF = False
        cmbbox_control.AddItem ""
        cmbbox_control.List(i, 0) = rs_control_cls!control_cls
        cmbbox_control.List(i, 1) = rs_control_cls!Description
        rs_control_cls.MoveNext
        i = i + 1
    Wend
    
    cmbbox_control.ColumnWidths = "20 pt; 50 pt"
    cmbbox_control.ListWidth = 70
    cmbbox_control.ListIndex = 0
    cmbbox_control.Text = cmbbox_control.List(0, 0)
End If
'==============================================================
'=====================Setting Combo Purchase Person==================
cmbbox_purchase_person.clear
cmbbox_purchase_person.columnCount = 2
cmbbox_purchase_person.TextColumn = 1
i = 1

If rs_personincharge_cls.EOF = False Or rs_personincharge_cls.BOF = False Then
    rs_personincharge_cls.MoveFirst
     cmbbox_purchase_person.AddItem ""
    cmbbox_purchase_person.List(0, 0) = ""
    cmbbox_purchase_person.List(0, 1) = ""
    While rs_personincharge_cls.EOF = False
    cmbbox_purchase_person.AddItem ""
        cmbbox_purchase_person.List(i, 0) = rs_personincharge_cls!personincharge_cls
        cmbbox_purchase_person.List(i, 1) = rs_personincharge_cls!Description
        rs_personincharge_cls.MoveNext
        i = i + 1
    Wend
    
    cmbbox_purchase_person.ColumnWidths = "20 pt; 140 pt"
    cmbbox_purchase_person.ListWidth = 160
    cmbbox_purchase_person.ListIndex = 0
    cmbbox_purchase_person.Text = cmbbox_purchase_person.List(0, 0)
End If
'==============================================================
'=====================Setting Combo unit2=================
cmbbox_unit.clear
cmbbox_unit.columnCount = 2
cmbbox_unit.TextColumn = 1
i = 1
Call up_FillCombo(cmbbox_unit, "unit_cls")

cmbbox_unit.ColumnWidths = "20 pt; 50 pt"
cmbbox_unit.ListWidth = 70
cmbbox_unit.ListIndex = 0


'==============================================================

cboTypeAccs.clear
cboTypeAccs.columnCount = 2

cboTypeAccs.AddItem
cboTypeAccs.List(0, 0) = "01"
cboTypeAccs.List(0, 1) = "Set"
cboTypeAccs.AddItem
cboTypeAccs.List(1, 0) = "02"
cboTypeAccs.List(1, 1) = "FG"

cboTypeAccs.ColumnWidths = "20 pt; 50 pt"

'==============================================================
'=====================Setting Combo Model=================
cboModel.clear
cboModel.columnCount = 2
cboModel.TextColumn = 1
i = 1
Call up_FillCombo(cboModel, "Model_Cls")

cboModel.ColumnWidths = "20 pt; 70 pt"
cboModel.ListWidth = 90
cboModel.ListIndex = 0
'==============================================================

'==============================================================
'=====================Setting Combo PO Type=================
cbo_POType.clear
cbo_POType.columnCount = 2
cbo_POType.TextColumn = 1
i = 1
Call up_FillCombo(cbo_POType, "POType_Cls")

cbo_POType.ColumnWidths = "20 pt; 70 pt"
cbo_POType.ListWidth = 90
cbo_POType.ListIndex = 0
'==============================================================


Set rs_warehouse_master = Nothing

'==============================================================
'=====================Setting Clasification Part=====================
cboClasificationPart.clear
cboClasificationPart.columnCount = 2
cboClasificationPart.TextColumn = 1
i = 0

If rs_clasification.EOF = False Or rs_clasification.BOF = False Then
    rs_clasification.MoveFirst
    While rs_clasification.EOF = False
        cboClasificationPart.AddItem ""
        cboClasificationPart.List(i, 0) = Trim(rs_clasification!ClasificationPart_Cls)
        cboClasificationPart.List(i, 1) = Trim(rs_clasification!Description)
        rs_clasification.MoveNext
        i = i + 1
    Wend
    
    cboClasificationPart.ColumnWidths = "50 pt; 120 pt"
    cboClasificationPart.ListWidth = 170
    cboClasificationPart.ListIndex = 0
    cboClasificationPart.Text = cboClasificationPart.List(0, 0)
End If

'==============================================================
'=====================Setting Clasification Destination========
cboDestination.clear
cboDestination.columnCount = 2
cboDestination.TextColumn = 1
i = 1
Call up_FillCombo(cboDestination, "Destination_Cls")

cboDestination.ColumnWidths = "20 pt; 70 pt"
cboDestination.ListWidth = 90
cboDestination.ListIndex = 0
'==============================================================

'==============================================================
'=====================Setting Combo Color=================
cboColor.clear
cboColor.columnCount = 2
cboColor.TextColumn = 1
i = 1
Call up_FillCombo(cboColor, "Color_Cls")

cboColor.ColumnWidths = "20 pt; 70 pt"
cboColor.ListWidth = 90
cboColor.ListIndex = 0
'==============================================================

End Sub

Function validasi2() As Boolean
Dim j As Integer
If Trim(txt_item_code.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg(1009) '"Please insert Product Code !"
    txt_item_code.SetFocus: validasi2 = False: Exit Function
End If

If Trim(txt_item_name.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg(1006) '"Please insert description !"
    txt_item_name.SetFocus: validasi2 = False: Exit Function
End If

If Trim(cmb_finish_good.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("0029")
    cmb_finish_good.SetFocus
    SSTab1.Tab = 0
    validasi2 = False
    Exit Function
Else
    j = 0
    If cmb_finish_good.ListCount < 1 Then Exit Function
       For i = 0 To cmb_finish_good.ListCount - 1
           If Trim(cmb_finish_good.Text) = cmb_finish_good.List(i, 0) Then
               j = 1
           End If
       Next
       If j = 0 Then
            lbl_pesan.Caption = DisplayMsg("0029")
            cmb_finish_good.SetFocus
            SSTab1.Tab = 0
            lbl_finish.Caption = ""
            Exit Function
        End If
End If

If Trim(txt_maker_item_code.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("0030")
    txt_maker_item_code.SetFocus
    validasi2 = False
    Exit Function
End If

If Trim(cmbbox_warehouse.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("0031")
    cmbbox_warehouse.SetFocus
    SSTab1.Tab = 0
    validasi2 = False
    Exit Function
End If
'======================================================================================================================================
If cmbbox_warehouse.MatchFound = False Then
    lbl_warehouse.Caption = ""
    lbl_pesan.Caption = DisplayMsg("4023")
    cmbbox_warehouse.SetFocus
    SSTab1.Tab = 0
    Exit Function
End If
'======================================================================================================================================
If Trim(cmbox_suplier.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("1054")
    cmbox_suplier.SetFocus
    SSTab1.Tab = 0
    validasi2 = False
    Exit Function
End If
'======================================================================================================================================
    If rs_trade_master.EOF = False Or rs_trade_master.BOF = False Then
        rs_trade_master.MoveFirst
        rs_trade_master.Find "trade_code='" & Trim(cmbox_suplier.Text) & "'"
        If rs_trade_master.EOF = True Then
no2:
            'lbl_suplier.Caption = ""
            txt_supplier = ""
            lbl_pesan.Caption = DisplayMsg("0032")
            cmbox_suplier.SetFocus
            SSTab1.Tab = 0
            Exit Function
        End If
    Else
        GoTo no2
    End If
'======================================================================================================================================
    j = 0
    If cmb_delivery.ListCount > 0 Then
        For i = 0 To cmb_delivery.ListCount - 1
            If Trim(cmb_delivery.Text) = Trim(cmb_delivery.List(i, 0)) Then
                  j = 1
            End If
        Next
        If j < 1 Then
              lbl_delivery_place.Caption = ""
              lbl_pesan.Caption = DisplayMsg("0033")
              cmb_delivery.SetFocus
              SSTab1.Tab = 0
              Exit Function
        End If
    Else
        If Trim(cmb_delivery.Text) <> "" Then
            lbl_delivery_place.Caption = ""
            lbl_pesan.Caption = DisplayMsg("0033")
            cmb_delivery.SetFocus
            SSTab1.Tab = 0
            Exit Function
        End If
    End If
'===============================================================================
If rs_hs_master.EOF = False Or rs_hs_master.BOF = False Then
        rs_hs_master.MoveFirst
        rs_hs_master.Find "hs_code='" & Trim(cmb_hs.Text) & "'"
        If rs_hs_master.EOF = True Then
NO:
            lbl_hs.Caption = ""
            lbl_pesan.Caption = DisplayMsg(8060)
            cmb_hs.SetFocus
            SSTab1.Tab = 0
            Exit Function
        End If
    Else
        GoTo NO
    End If
    
'======================================================================================================================================
If Trim(cmb_part2.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("0016")
    cmb_part2.SetFocus
    SSTab1.Tab = 0
    validasi2 = False
    Exit Function
Else
    j = 0
    If cmb_part2.ListCount < 1 Then Exit Function
       For i = 0 To cmb_part2.ListCount - 1
           If Trim(cmb_part2.Text) = cmb_part2.List(i, 0) Then
               j = 1
           End If
       Next
       If j = 0 Then
            lbl_pesan.Caption = DisplayMsg("0016")
            cmb_part2.SetFocus
            SSTab1.Tab = 0
            lbl_part.Caption = ""
            Exit Function
        End If
End If

If Trim(cmb_reserve2.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("0023")
    cmb_reserve2.SetFocus
    SSTab1.Tab = 0
    validasi2 = False
    Exit Function
Else
    j = 0
    If cmb_reserve2.ListCount < 1 Then Exit Function
       For i = 0 To cmb_reserve2.ListCount - 1
           If Trim(cmb_reserve2.Text) = cmb_reserve2.List(i, 0) Then
               j = 1
           End If
       Next
       If j = 0 Then
            lbl_pesan.Caption = DisplayMsg("0023")
            cmb_reserve2.SetFocus
            SSTab1.Tab = 0
            lbl_reserve.Caption = ""
            Exit Function
        End If
End If
If Trim(cmb_supply2.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("4052")
    cmb_supply2.SetFocus
    SSTab1.Tab = 0
    validasi2 = False
    Exit Function
Else
    j = 0
    If cmb_supply2.ListCount < 1 Then Exit Function
       For i = 0 To cmb_supply2.ListCount - 1
           If Trim(cmb_supply2.Text) = cmb_supply2.List(i, 0) Then
               j = 1
           End If
       Next
       If j = 0 Then
            lbl_pesan.Caption = DisplayMsg("4056")
            cmb_supply2.SetFocus
            SSTab1.Tab = 0
            lbl_supply.Caption = ""
            Exit Function
        End If
End If

If Trim(cmb_provision2.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("0018")
    cmb_provision2.SetFocus
    SSTab1.Tab = 0
    validasi2 = False
    Exit Function
Else
    j = 0
    If cmb_provision2.ListCount < 1 Then Exit Function
       For i = 0 To cmb_provision2.ListCount - 1
           If Trim(cmb_provision2.Text) = cmb_provision2.List(i, 0) Then
               j = 1
           End If
       Next
       If j = 0 Then
            lbl_pesan.Caption = DisplayMsg("0018")
            cmb_provision2.SetFocus
            SSTab1.Tab = 0
            lbl_provotion.Caption = ""
            Exit Function
        End If
End If

'===========================================================================================================================================

If IsNumeric(txt_ne) = False Then
    txt_ne.SetFocus
    lbl_pesan = DisplayMsg(8024)
    Exit Function
End If
If CDbl(txt_ne) > gd_MaxBox Then
    txt_ne.SetFocus
    lbl_pesan = DisplayMsg(8025) & " " & gd_MaxBox
    Exit Function
End If

'===========================================================================================================================================
    If Trim(cmbbox_packing.Text) <> "" Then
        If rs_packingstyle_cls.EOF = False Or rs_packingstyle_cls.BOF = False Then
            rs_packingstyle_cls.MoveFirst
            rs_packingstyle_cls.Find "packingstyle_cls='" & Trim(cmbbox_packing.Text) & "'"
            If rs_packingstyle_cls.EOF = True Then
                lbl_pesan.Caption = DisplayMsg("0024")
                cmbbox_packing.SetFocus
                SSTab1.Tab = 0
                Exit Function
            End If
        Else
             lbl_pesan.Caption = DisplayMsg("0024")
             cmbbox_packing.SetFocus
             SSTab1.Tab = 0
             Exit Function
        End If
    End If
'===========================================================================================================================================

   '#Must input PackingItemCode for Item with packing Style "99"/Bag
    If Trim(cbo_PackingCode.Text) = "" And Trim(cmbbox_packing.Text) = "99" Then
        lbl_pesan.Caption = DisplayMsg("0034")
        cbo_PackingCode.SetFocus
        SSTab1.Tab = 0
        validasi2 = False
        Exit Function
    Else
        j = 0
        If Trim(cbo_PackingCode.Text) = "" Then GoTo proceed
        If cbo_PackingCode.ListCount < 1 Then GoTo invalid
        
           For i = 0 To cbo_PackingCode.ListCount - 1
               If Trim(cbo_PackingCode.Text) = cbo_PackingCode.List(i, 0) Then
                   j = 1
               End If
           Next
           If j = 0 Then
invalid:
                lbl_pesan.Caption = DisplayMsg("0035")
                cbo_PackingCode.SetFocus
                SSTab1.Tab = 0
                Exit Function
           End If
                       
            '#cek qty/Case

            If IsNumeric(txt_ne) = False Or CDbl(txt_ne) = 0 Then
                txt_ne.SetFocus
                lbl_pesan = DisplayMsg(8024)
                Exit Function
            End If
            If CDbl(txt_ne) > gd_MaxBox Then
                txt_ne.SetFocus
                lbl_pesan = DisplayMsg(8025) & " " & gd_MaxBox
                Exit Function
            End If
            If Trim(lbl_pesan) <> "" Then: txt_ne.SetFocus: Exit Function
    End If
proceed:
'===========================================================================================================================================
    If Trim(cmbbox_group.Text) <> "" Then
        If rs_group_cls.EOF = False Or rs_group_cls.BOF = False Then
            rs_group_cls.MoveFirst
            rs_group_cls.Find "group_cls='" & Trim(cmbbox_group.Text) & "'"
            If rs_group_cls.EOF = True Then
                lbl_pesan.Caption = DisplayMsg("8056")
                lbl_group = ""
                cmbbox_group.SetFocus
                SSTab1.Tab = 0
                Exit Function
            End If
        Else
            lbl_pesan.Caption = DisplayMsg("8056")
            lbl_group = ""
            cmbbox_group.SetFocus
            SSTab1.Tab = 0
            Exit Function
        End If
    End If
'===========================================================================================================================================
If Trim(cmb_prod.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("0017")
    cmb_prod.SetFocus
    SSTab1.Tab = 0
    validasi2 = False
    Exit Function
Else
    j = 0
    If cmb_prod.ListCount < 1 Then Exit Function
       For i = 0 To cmb_prod.ListCount - 1
           If Trim(cmb_prod.Text) = cmb_prod.List(i, 0) Then
               j = 1
           End If
           
       Next
       If j = 0 Then
            lbl_pesan.Caption = DisplayMsg("0017")
            cmb_prod.SetFocus
            SSTab1.Tab = 0
            lbl_prod.Caption = ""
            Exit Function
        End If
End If
'===========================================================================================================================================
If Trim(cmb_explosion2.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("0015")
    cmb_explosion2.SetFocus
    SSTab1.Tab = 0
    validasi2 = False
    Exit Function
Else
    j = 0
    If cmb_explosion2.ListCount < 1 Then Exit Function
       For i = 0 To cmb_explosion2.ListCount - 1
           If Trim(cmb_explosion2.Text) = cmb_explosion2.List(i, 0) Then
               j = 1
           End If
       Next
       If j = 0 Then
            lbl_pesan.Caption = DisplayMsg("0015")
            cmb_explosion2.SetFocus
            SSTab1.Tab = 0
            lbl_explosion.Caption = ""
            Exit Function
        End If
End If

'===========================================================================================================================================
    If Trim(cmbbox_purchase_person.Text) <> "" Then
        If rs_personincharge_cls.EOF = False Or rs_personincharge_cls.BOF = False Then
            rs_personincharge_cls.MoveFirst
            rs_personincharge_cls.Find "personincharge_cls='" & Trim(cmbbox_purchase_person.Text) & "'"
            If rs_personincharge_cls.EOF = True Then
                lbl_purchase.Caption = ""
                lbl_pesan.Caption = DisplayMsg("0026")
                SSTab1.Tab = 0
                cmbbox_purchase_person.SetFocus
                Exit Function
            End If
        Else
            lbl_purchase.Caption = ""
            lbl_pesan.Caption = DisplayMsg("0026")
            SSTab1.Tab = 0
            cmbbox_purchase_person.SetFocus
            Exit Function
        End If
    End If
'===========================================================================================================================================
If Trim(cmb_stock_control2.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("0019")
    cmb_stock_control2.SetFocus
    SSTab1.Tab = 0
    validasi2 = False
    Exit Function
Else
    j = 0
    If cmb_stock_control2.ListCount < 1 Then Exit Function
       For i = 0 To cmb_stock_control2.ListCount - 1
           If Trim(cmb_stock_control2.Text) = cmb_stock_control2.List(i, 0) Then
               j = 1
           End If
       Next
       If j = 0 Then
            lbl_pesan.Caption = DisplayMsg("0019")
            cmb_stock_control2.SetFocus
            SSTab1.Tab = 0
            lbl_stock_control.Caption = ""
            Exit Function
        End If
End If

'===========================================================================================================================================
MaskEdBox1.PromptInclude = False
If Len(Trim(MaskEdBox1.Text)) <> 8 Then
    lbl_pesan.Caption = DisplayMsg("1021")
    validasi2 = False
    MaskEdBox1.SetFocus
    Exit Function
Else
    MaskEdBox1.PromptInclude = True
    If IsDate(Trim(MaskEdBox1.Text)) = False And Trim(MaskEdBox1.Text) <> "99/99/9999" Then
        lbl_pesan.Caption = DisplayMsg("1021")
        validasi2 = False
        MaskEdBox1.SetFocus
        Exit Function
    End If
End If

If Trim(cmb_make.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("8096")
    cmb_make.SetFocus
    SSTab1.Tab = 0
    validasi2 = False
    Exit Function
Else
    j = 0
    If cmb_make.ListCount < 1 Then Exit Function
       For i = 0 To cmb_make.ListCount - 1
           If Trim(cmb_make.Text) = cmb_make.List(i, 0) Then
               j = 1
           End If
       Next
       If j = 0 Then
            lbl_pesan.Caption = DisplayMsg("8096")
            cmb_make.SetFocus
            SSTab1.Tab = 0
            Lbl_Make.Caption = ""
            Exit Function
        End If
End If

'===========================================================================================================================================
   If Trim(cmbbox_control.Text) <> "" Then
        If rs_control_cls.EOF = False Or rs_control_cls.BOF = False Then
            rs_control_cls.MoveFirst
            rs_control_cls.Find "control_cls='" & Trim(cmbbox_control.Text) & "'"
            If rs_control_cls.EOF = True Then
                lbl_pesan.Caption = DisplayMsg("0020")
                lbl_control = ""
                SSTab1.Tab = 0
                cmbbox_control.SetFocus
                Exit Function
            End If
        End If
    Else
        lbl_pesan.Caption = DisplayMsg("0020")
        lbl_control = ""
        SSTab1.Tab = 0
        cmbbox_control.SetFocus
        Exit Function
    End If
'===========================================================================================================================================

If Trim(cmbbox_unit.Text) = "" Then
    lbl_pesan.Caption = DisplayMsg("1030")
    cmbbox_unit.SetFocus
    SSTab1.Tab = 0
    validasi2 = False
    Exit Function
Else
    j = 0
    If cmbbox_unit.ListCount < 1 Then Exit Function
       For i = 0 To cmbbox_unit.ListCount - 1
           If Trim(cmbbox_unit.Text) = cmbbox_unit.List(i, 0) Then
               j = 1
           End If
       Next
       If j = 0 Then
            lbl_pesan.Caption = DisplayMsg("1010")
            cmbbox_unit.SetFocus
            SSTab1.Tab = 0
            lbl_unit.Caption = ""
            Exit Function
        End If
End If
'===========================================================================================================================================
    If Trim(cmbbox_packing2.Text) <> "" Then
        If rs_packingstyle_cls.EOF = False Or rs_packingstyle_cls.BOF = False Then
                rs_packingstyle_cls.MoveFirst
                rs_packingstyle_cls.Find "packingstyle_cls='" & Trim(cmbbox_packing2.Text) & "'"
                If rs_packingstyle_cls.EOF = True Then
                    lbl_pesan.Caption = DisplayMsg("0025")
                    cmbbox_packing2.SetFocus
                    SSTab1.Tab = 0
                    Exit Function
                End If
        Else
            lbl_pesan.Caption = DisplayMsg("0025")
            cmbbox_packing2.SetFocus
            SSTab1.Tab = 0
            Exit Function
        End If
    End If
'===========================================================================================================================================

'If Trim(cboModel.Text) = "" Then
'    Lbl_pesan.Caption = DisplayMsg("1030")
'    cboModel.SetFocus
'    SSTab1.Tab = 0
'    validasi2 = False
'    Exit Function
'Else
'    j = 0
'    If cboModel.ListCount < 1 Then Exit Function
'       For i = 0 To cboModel.ListCount - 1
'           If Trim(cboModel.Text) = cboModel.List(i, 0) Then
'               j = 1
'           End If
'       Next
'       If j = 0 Then
'            Lbl_pesan.Caption = DisplayMsg("1010")
'            cboModel.SetFocus
'            SSTab1.Tab = 0
'            lblModel.Caption = ""
'            Exit Function
'        End If
'End If
'===========================================================================================================================================

'    If Trim(cmbox_manufacture.Text) <> "" Then
'        If rs_manufacture_line.EOF = False Or rs_manufacture_line.BOF = False Then
'            rs_manufacture_line.MoveFirst
'            rs_manufacture_line.Find "manufacture_code='" & Trim(cmbox_manufacture.Text) & "'"
'            If rs_manufacture_line.EOF = True Then
'no3:
'                txt_factory = ""
'                cmbbox_line.clear
'                lbl_line.Caption = ""
'                cmbbox_line.Text = ""
'                lbl_pesan.Caption = DisplayMsg("0036")
'                cmbox_manufacture.SetFocus
'                SSTab1.Tab = 0
'                Exit Function
'            End If
'        Else
'            GoTo no3
'        End If
'
'        If Trim(cmbbox_line.Text) = "" Then
'            lbl_pesan.Caption = DisplayMsg("0037")
'            cmbbox_line.SetFocus
'            lbl_line.Caption = ""
'            SSTab1.Tab = 0
'            Exit Function
'        Else
'            lbl_pesan.Caption = ""
'        End If
'    End If
'======================================================================================================================================
'   If Trim(cmbbox_line.Text) <> "" Then
'        If Trim(cmbox_manufacture.Text) = "" Then
'             lbl_pesan.Caption = DisplayMsg("1052")
'             txt_factory = ""
'             cmbox_manufacture.SetFocus
'             SSTab1.Tab = 0
'             Exit Function
'         Else
'             lbl_pesan.Caption = ""
'         End If
'
'        If rs_line.EOF = False Or rs_line.BOF = False Then
'            rs_line.MoveFirst
'            rs_line.Find "line_code ='" & Trim(cmbbox_line.Text) & "'"
'            If rs_line.EOF = True Then
'no4:
'                lbl_line.Caption = ""
'                lbl_pesan.Caption = DisplayMsg("8009")
'                cmbbox_line.SetFocus
'                 SSTab1.Tab = 0
'                Exit Function
'            End If
'        Else
'            GoTo no4:
'        End If
'    End If
'======================================================================================================================================
    If Trim(cmbbox_material.Text) <> "" Then
    'On Error Resume Next
        If rs_material_cls.EOF = False Or rs_material_cls.BOF = False Then
            rs_material_cls.MoveFirst
            rs_material_cls.Find "material_cls='" & Trim(cmbbox_material.Text) & "'"
            If rs_material_cls.EOF = True Then
no5:
                lbl_material.Caption = ""
                lbl_pesan.Caption = DisplayMsg("8095")
                cmbbox_material.SetFocus
                SSTab1.Tab = 1
                Exit Function
            End If
        Else
            GoTo no5:
        End If
    End If
'===========================================================================================================================================

If IsNumeric(txt_thickness) = False Then
    txt_thickness.SetFocus
    lbl_pesan = DisplayMsg(8026)
    Exit Function
End If
If CDbl(txt_thickness) > gd_MaxThickness Then
    txt_thickness.SetFocus
    lbl_pesan = DisplayMsg(8027) & " " & gd_MaxThickness
    Exit Function
End If

If IsNumeric(txt_width) = False Then
    txt_width.SetFocus
    lbl_pesan = DisplayMsg(8029)
    Exit Function
End If
If CDbl(txt_width) > gd_MaxWidth Then
    txt_width.SetFocus
    lbl_pesan = DisplayMsg(8029) & " " & gd_MaxWidth
    Exit Function
End If


'----weight
If Trim(txt_weight) = "" Then GoTo pass
txt_weight = Format(txt_weight, gs_formatWeight)

If CDbl(txt_weight) > gd_MaxWeight Then _
lbl_pesan = DisplayMsg(8030) & " " & gd_MaxWeight & " !": Exit Function

If Trim(txt_gross) = "" Then GoTo pass
txt_gross = Format(txt_gross, gs_formatWeight)

If CDbl(txt_gross) > gd_MaxWeight Then _
lbl_pesan = DisplayMsg(8030) & " " & gd_MaxWeight & " !": Exit Function

'------------

pass:

If IsNumeric(txt_length) = False Then
    txt_length.SetFocus
    lbl_pesan = DisplayMsg(8031)
    Exit Function
End If
If CDbl(txt_length) > gd_MaxLength Then
    txt_length.SetFocus
    lbl_pesan = DisplayMsg(8032) & " " & gd_MaxLength
    Exit Function
End If


'===========================================================================================================================================
    If Trim(cmbbox_sheet.Text) <> "" Then
        If rs_sheetcoil_cls.EOF = False Or rs_sheetcoil_cls.BOF = False Then
            rs_sheetcoil_cls.MoveFirst
            rs_sheetcoil_cls.Find "sheetcoil_cls='" & Trim(cmbbox_sheet.Text) & "'"
            If rs_sheetcoil_cls.EOF = True Then
no9:
                lbl_sheet.Caption = ""
                lbl_pesan.Caption = DisplayMsg("0027")
                cmbbox_sheet.SetFocus
                SSTab1.Tab = 1
                Exit Function
            End If
        Else
            GoTo no9:
        End If
    End If
'===========================================================================================================================================

If IsNumeric(txt_pitch) = False Then
    txt_pitch.SetFocus
    lbl_pesan = DisplayMsg(8033)
    Exit Function
End If
If CDbl(txt_pitch) > gd_MaxPitch Then
    txt_pitch.SetFocus
    lbl_pesan = DisplayMsg(8034) & " " & gd_MaxPitch
    Exit Function
End If

If IsNumeric(txt_number_producible) = False Then
    txt_number_producible.SetFocus
    lbl_pesan = DisplayMsg(8035)
    Exit Function
End If
If CDbl(txt_number_producible) > gd_MaxQty Then
    txt_number_producible.SetFocus
    lbl_pesan = DisplayMsg(8036) & " " & gd_MaxQty
    Exit Function
End If

If IsNumeric(txt_scrap_weight) = False Then
    txt_scrap_weight.SetFocus
    lbl_pesan = DisplayMsg(8037)
    Exit Function
End If
If CDbl(txt_scrap_weight) > gd_MaxWeight Then
    txt_scrap_weight.SetFocus
    lbl_pesan = DisplayMsg(8038) & " " & gd_MaxWeight
    Exit Function
End If

'===========================================================================================================================================
    If Trim(cmbbox_drawing.Text) <> "" Then
        If rs_drawingmaterial_cls.EOF = False Or rs_drawingmaterial_cls.BOF = False Then
            rs_drawingmaterial_cls.MoveFirst
            rs_drawingmaterial_cls.Find "drawingmaterial_cls='" & Trim(cmbbox_drawing.Text) & "'"
            If rs_drawingmaterial_cls.EOF = True Then
no6:
                lbl_drawing.Caption = ""
                lbl_pesan.Caption = DisplayMsg("0021")
                cmbbox_drawing.SetFocus
                SSTab1.Tab = 1
                Exit Function
            End If
        Else
            GoTo no6:
        End If
    End If
'===========================================================================================================================================
    If Trim(cmbbox_surface.Text) <> "" Then
        If rs_surfacetreatment_cls.EOF = False Or rs_surfacetreatment_cls.BOF = False Then
            rs_surfacetreatment_cls.MoveFirst
            rs_surfacetreatment_cls.Find "surfacetreatment_cls='" & Trim(cmbbox_surface.Text) & "'"
            If rs_surfacetreatment_cls.EOF = True Then
no7:
                lbl_surface.Caption = ""
                lbl_pesan.Caption = DisplayMsg("0028")
                cmbbox_surface.SetFocus
                SSTab1.Tab = 1
                Exit Function
            End If
        Else
            GoTo no7
        End If
    End If
'===========================================================================================================================================
    If Trim(cmbbox_heat.Text) <> "" Then
        If rs_heattreatment_cls.EOF = False Or rs_heattreatment_cls.BOF = False Then
            rs_heattreatment_cls.MoveFirst
            rs_heattreatment_cls.Find "heattreatment_cls='" & Trim(cmbbox_heat.Text) & "'"
            If rs_heattreatment_cls.EOF = True Then
no8:
                lbl_heat.Caption = ""
                lbl_pesan.Caption = DisplayMsg("0022")
                cmbbox_heat.SetFocus
                SSTab1.Tab = 1
                Exit Function
            End If
        Else
            GoTo no8:
        End If
    End If
'===========================================================================================================================================
If IsNumeric(txt_sample) = False Then
    txt_sample.SetFocus
    lbl_pesan = DisplayMsg(8039)
    Exit Function
End If
If CDbl(txt_sample) > gd_MaxSample Then
    txt_sample.SetFocus
    lbl_pesan = DisplayMsg(8040) & " " & gd_MaxSample
    Exit Function
End If

If IsNumeric(txt_sw) = False Then
    txt_sw.SetFocus
    lbl_pesan = DisplayMsg(8041)
    Exit Function
End If
If CDbl(txt_sw) > gd_MaxSW Then
    txt_sw.SetFocus
    lbl_pesan = DisplayMsg(8042) & " " & gd_MaxSW
    Exit Function
End If

If IsNumeric(txt_ew) = False Then
    txt_ew.SetFocus
    lbl_pesan = DisplayMsg(8043)
    Exit Function
End If
If CDbl(txt_ew) > gd_MaxEW Then
    txt_ew.SetFocus
    lbl_pesan = DisplayMsg(8044) & " " & gd_MaxEW
    Exit Function
End If


If IsNumeric(txt_mc) = False Then
    txt_mc.SetFocus
    lbl_pesan = DisplayMsg(8046)
    Exit Function
End If
If CDbl(txt_mc) > gd_MaxCoefficient Then
    txt_mc.SetFocus
    lbl_pesan = DisplayMsg(8047) & " " & gd_MaxCoefficient
    Exit Function
End If

If IsNumeric(txt_pc) = False Then
    txt_pc.SetFocus
    lbl_pesan = DisplayMsg(8046)
    Exit Function
End If
If CDbl(txt_pc) > gd_MaxCoefficient Then
    txt_pc.SetFocus
    lbl_pesan = DisplayMsg(8047) & " " & gd_MaxCoefficient
    Exit Function
End If


If IsNumeric(txt_yp) = False Then
    txt_yp.SetFocus
    lbl_pesan = DisplayMsg(8048)
    Exit Function
End If
If CDbl(txt_yp) > gd_MaxPercentage Then
    txt_yp.SetFocus
    lbl_pesan = DisplayMsg(8049) & " " & gd_MaxPercentage
    Exit Function
End If

'===========================================================================================================================================
validasi2 = True
End Function

Private Sub set_delivery()
If rs_delivery.State <> adStateClosed Then rs_delivery.Close
rs_delivery.Open "select * from delivery_place where trade_code='" & Trim(cmbox_suplier.Text) & " '", Db, adOpenKeyset, adLockOptimistic
'=====================Setting Combo delivery=========================
cmb_delivery.clear
cmb_delivery.columnCount = 2
cmb_delivery.TextColumn = 1
i = 1

If rs_delivery.EOF = False Or rs_delivery.BOF = False Then
    rs_delivery.MoveFirst
    cmb_delivery.AddItem ""
        cmb_delivery.List(0, 0) = ""
        cmb_delivery.List(0, 1) = ""
    While rs_delivery.EOF = False
        cmb_delivery.AddItem ""
        cmb_delivery.List(i, 0) = rs_delivery!location_code
        cmb_delivery.List(i, 1) = rs_delivery!location_Name
        rs_delivery.MoveNext
        i = i + 1
    Wend
    
    cmb_delivery.ColumnWidths = "50 pt; 90 pt"
    cmb_delivery.ListWidth = 140
    cmb_delivery.ListIndex = 0
    cmb_delivery.Text = cmb_delivery.List(0, 0)
End If
'==============================================================
rs_delivery.Close
End Sub



Private Sub cmbox_manufacture_GotFocus()
SSTab1.Tab = 1
End Sub

Private Sub cmbox_manufacture_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim j As Integer
If KeyCode = vbKeyDelete Then txt_factory = ""
If KeyCode = vbKeyBack Then txt_factory = ""

If KeyCode = 13 Then
    If cmbox_manufacture.ListCount < 0 Then Exit Sub
    j = 0
    For i = 0 To cmbox_manufacture.ListCount - 1
        If UCase(Trim(cmbox_manufacture.Text)) = UCase(Trim(cmbox_manufacture.List(i, 0))) Then j = 1: Exit For
    Next
    If j = 1 Then
        cmbox_manufacture = Trim(cmbox_manufacture.List(i, 0))
        txt_factory = Trim(cmbox_manufacture.List(i, 1))
        Call set_line
        lbl_pesan.Caption = ""
    Else
        txt_factory = ""
        cmbbox_line.clear
        cmbbox_line.Text = ""
        lbl_line.Caption = ""
        lbl_pesan.Caption = DisplayMsg("0036")
        cmbox_manufacture.SetFocus
    End If
End If

End Sub

Private Sub cmbox_manufacture_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmbox_suplier_Click()

txt_supplier = cmbox_suplier.List(cmbox_suplier.ListIndex, 1)
lbl_pesan.Caption = ""
cmb_delivery.clear
cmb_delivery.Text = ""
lbl_delivery_place.Caption = ""
Call set_delivery
End Sub

Private Sub cmbox_suplier_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub cmbox_suplier_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)


If KeyCode = vbKeyDelete Then txt_supplier = ""
If KeyCode = vbKeyBack Then txt_supplier = ""

If KeyCode = 13 Then
    If rs_trade_master.EOF = False Or rs_trade_master.BOF = False Then
        rs_trade_master.MoveFirst
        rs_trade_master.Find "trade_code='" & Trim(cmbox_suplier.Text) & "'"
        If rs_trade_master.EOF = False Then
            cmbox_suplier = Trim(rs_trade_master!Trade_Code)
            txt_supplier = Trim(rs_trade_master!trade_name)
            Call set_delivery
            lbl_pesan.Caption = ""
        Else
            txt_supplier = ""
            lbl_pesan.Caption = DisplayMsg("0032")
            cmb_delivery.clear
            cmb_delivery.Text = ""
            lbl_delivery_place.Caption = ""
            cmbox_suplier.SetFocus
        End If
    End If
End If
End Sub

Public Sub cmd_clear_Click()
Call combo_clear
Call text_clear
Call label_clear
Status = "insert"
txt_item_code.Enabled = True
lbl_record.Caption = "Record 0 of 0"
MaskEdBox1.Text = "99/99/9999"
'txt_item_code.SetFocus
End Sub


Public Sub Cmd_Submit_Click()
If hakUpdate(Me.Name) = 0 Then _
     lbl_pesan = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
     
Dim spos As Integer
spos = SSTab1.Tab
     
If cmb_stock_control2.Text = "02" Then If Not ValidStock(Trim(txt_item_code.Text)) Then lbl_pesan = DisplayMsg(1061): cmb_stock_control2.SetFocus: Exit Sub
If Status = "insert" Then
    If rs_item_master.EOF = False Or rs_item_master.BOF = False Then
        rs_item_master.MoveFirst
            rs_item_master.Find "item_code='" & Trim(txt_item_code.Text) & "'"
        If rs_item_master.EOF = True Then
            If validasi2 = True Then
            Me.MousePointer = vbHourglass
            
            rs_item_master.AddNew
                rs_item_master!Item_Code = Trim(txt_item_code.Text)
                Call isi_field
            rs_item_master.update

            lbl_pesan.Caption = DisplayMsg(1000) '"Input data success !"
            txt_item_code.Enabled = False
            End If
            Me.MousePointer = vbDefault
        Else
            Call txt_item_code_KeyPress(13)
        End If
    Else
          If validasi2 = True Then
            Me.MousePointer = vbHourglass
            rs_item_master.AddNew
                rs_item_master!Item_Code = Trim(txt_item_code.Text)
                Call isi_field
            rs_item_master.update
                      
            
            lbl_pesan.Caption = DisplayMsg(1000) '"Input data success !"
       
            txt_item_code.Enabled = False
            End If
            Me.MousePointer = vbDefault
    End If
Else
    Me.MousePointer = vbHourglass
    rs_item_master.MoveFirst
    rs_item_master.Find "item_code='" & l_txt_item_code & "'"
    If rs_item_master.EOF = False Then
        If validasi2 = True Then
                    Call isi_field
                    rs_item_master.update
                    lbl_pesan.Caption = DisplayMsg(1101) '"Update data success !"
        Else
            Me.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    Call data_tampil
    lbl_pesan.Caption = DisplayMsg(1101) '"Update data success !"
    Me.MousePointer = vbDefault
End If

'SSTab1.Tab = spos
l_txt_item_code = Trim(txt_item_code)
End Sub

Private Sub isi_field()
            rs_item_master!item_name = Trim(txt_item_name.Text)
            rs_item_master!finishgoodpart_cls = Trim(cmb_finish_good.Text)
            rs_item_master!Drawing_Number = Trim(txt_drawingCode)
            rs_item_master!wh_code = Trim(cmbbox_warehouse.Text)
            rs_item_master!Address = Trim(Text1.Text)
            rs_item_master!Supplier_Code = Trim(cmbox_suplier.Text): rs_item_master!delivery_place = Trim(cmb_delivery.Text)
            rs_item_master!HS_Code = Trim(cmb_hs.Text)
        
        If Trim(cmbox_manufacture.Text) <> "" Then
            rs_item_master!Manufacture_Code = Trim(cmbox_manufacture.Text)
        Else
            rs_item_master!Manufacture_Code = Null
        End If

         If Trim(cmbbox_line.Text) <> "" Then
             rs_item_master!line_code = Trim(cmbbox_line.Text)
        Else
             rs_item_master!line_code = Null
        End If
        
        
            rs_item_master!part_cls = Trim(cmb_part2.Text): rs_item_master!MakerItem_Code = Trim(txt_maker_item_code.Text)
            rs_item_master!reserve_cls = Trim(cmb_reserve2.Text): rs_item_master!suply_cls = Trim(cmb_supply2.Text)
            rs_item_master!Provision_Cls = Trim(cmb_provision2.Text)
        
        If Trim(cmbbox_material.Text) <> "" Then
            rs_item_master!Material_Cls = Trim(cmbbox_material.Text)
        Else
            rs_item_master!Material_Cls = Null
        End If
                
        If Trim(txt_thickness.Text) <> "" Then
            rs_item_master!Thickness = CDbl(Trim(txt_thickness.Text))
        Else
            rs_item_master!Thickness = "0"
        End If
        
        If Trim(txt_width.Text) <> "" Then
            rs_item_master!Width = CDbl(Trim(txt_width.Text))
        Else
            rs_item_master!Width = "0"
        End If
        
        If Trim(txt_weight.Text) <> "" Then
           rs_item_master!Weight = CDbl(txt_weight.Text)
        Else
            rs_item_master!Weight = "0"
        End If
        
        If Trim(txt_gross.Text) <> "" Then
           rs_item_master!GrossWeight = CDbl(txt_gross.Text)
        Else
            rs_item_master!GrossWeight = "0"
        End If
        
        If Trim(txt_length.Text) <> "" Then
            rs_item_master!Length = CDbl(Trim(txt_length.Text))
        Else
            rs_item_master!Length = "0"
        End If
            
        If Trim(cmbbox_sheet.Text) <> "" Then
            rs_item_master!sheetcoil_cls = Trim(cmbbox_sheet.Text)
        Else
            rs_item_master!sheetcoil_cls = Null
        End If
        
        If Trim(txt_pitch.Text) <> "" Then
            rs_item_master!pitch = CDbl(Trim(txt_pitch.Text))
        Else
            rs_item_master!pitch = "0"
        End If
        
        If Trim(txt_number_producible.Text) <> "" Then
            rs_item_master!number_producible = CDbl(Trim(txt_number_producible.Text))
        Else
            rs_item_master!number_producible = "0"
        End If
            
        If Trim(txt_scrap_weight.Text) <> "" Then
            rs_item_master!scrap_weight = CDbl(Trim(txt_scrap_weight.Text))
        Else
            rs_item_master!scrap_weight = "0"
        End If
        
        If Trim(cmbbox_drawing.Text) <> "" Then
            rs_item_master!drawingmaterial_cls = Trim(cmbbox_drawing.Text)
        Else
            rs_item_master!drawingmaterial_cls = Null
        End If
        
        If Trim(cmbbox_surface.Text) <> "" Then
            rs_item_master!surfacetreatment_cls = Trim(cmbbox_surface.Text)
        Else
            rs_item_master!surfacetreatment_cls = Null
        End If
        
        If Trim(cmbbox_heat.Text) <> "" Then
            rs_item_master!heattreatment_cls = Trim(cmbbox_heat.Text)
        Else
                rs_item_master!heattreatment_cls = Null
        End If
        
        If Trim(txt_sample.Text) <> "" Then
            rs_item_master!Sample = CDbl(Trim(txt_sample.Text))
        Else
            rs_item_master!Sample = "0"
        End If
        
        If Trim(txt_sw.Text) <> "" Then
            rs_item_master!SW_Qty = CDbl(Trim(txt_sw.Text))
        Else
            rs_item_master!SW_Qty = "0"
        End If
        
        If Trim(txt_ew.Text) <> "" Then
            rs_item_master!EW_Qty = CDbl(Trim(txt_ew.Text))
        Else
            rs_item_master!EW_Qty = "0"
        End If
        
        If Trim(txt_np.Text) <> "" Then
            rs_item_master!Number_Process = CDbl(Trim(txt_np.Text))
        Else
            rs_item_master!Number_Process = "0"
        End If
        
        If Trim(txt_mc.Text) <> "" Then
            rs_item_master!Material_Coefficient = CDbl(Trim(txt_mc.Text))
        Else
            rs_item_master!Material_Coefficient = "0"
        End If
        
        If Trim(txt_pc.Text) <> "" Then
            rs_item_master!Process_Coefficient = CDbl(Trim(txt_pc.Text))
        Else
            rs_item_master!Process_Coefficient = "0"
        End If
        
        If Trim(txt_min_lot.Text) <> "" Then
            rs_item_master!Min_Lot = CDbl(Trim(txt_min_lot.Text))
        Else
            rs_item_master!Min_Lot = "0"
        End If
                
        If Trim(txt_lot.Text) <> "" Then
            rs_item_master!Lot_Qty = CDbl(Trim(txt_lot.Text))
        Else
            rs_item_master!Lot_Qty = "0"
        End If
        
        If Trim(txt_lot_coef.Text) <> "" Then
            rs_item_master!Lot_Coefficience = CDbl(Trim(txt_lot_coef.Text))
        Else
            rs_item_master!Lot_Coefficience = "0"
        End If
        
         If Trim(txt_surface_qty.Text) <> "" Then
            rs_item_master!Surface_OrderPointQty = CDbl(Trim(txt_surface_qty.Text))
        Else
            rs_item_master!Surface_OrderPointQty = "0"
        End If
        
         If Trim(txt_heat_qty.Text) <> "" Then
            rs_item_master!Heat_OrderPointQty = CDbl(Trim(txt_heat_qty.Text))
        Else
            rs_item_master!Heat_OrderPointQty = "0"
        End If
        
        If Trim(txt_prt.Text) <> "" Then
            rs_item_master!Product_ReadTime = CDbl(Trim(txt_prt.Text))
        Else
            rs_item_master!Product_ReadTime = "0"
        End If
        
        If Trim(txt_yp.Text) <> "" Then
            rs_item_master!Yield_Percentage = CDbl(Trim(txt_yp.Text))
        Else
            rs_item_master!Yield_Percentage = "0"
        End If
        
        If Trim(txt_ne.Text) <> "" Then
            rs_item_master!number_entering = CDbl(Trim(txt_ne.Text))
        Else
            rs_item_master!number_entering = "0"
        End If
        
        If Trim(cmbbox_packing.Text) <> "" Then
            rs_item_master!PackingStyle_Cls = Trim(cmbbox_packing.Text)
        Else
            rs_item_master!PackingStyle_Cls = Null
        End If
        
        If Trim(cmbbox_group.Text) <> "" Then
            rs_item_master!group_cls = Trim(cmbbox_group.Text)
        Else
            rs_item_master!group_cls = Null
        End If
        
        If Trim(cmb_prod.Text) <> "" Then
            rs_item_master!production_cls = Trim(cmb_prod.Text)
        Else
            rs_item_master!production_cls = Null
        End If
        
        If Trim(txt_standart_stock.Text) <> "" Then
            rs_item_master!Standard_Stock = CDbl(Trim(txt_standart_stock.Text))
        Else
            rs_item_master!Standard_Stock = "0"
        End If
        
        If Trim(txt_safety_stock.Text) <> "" Then
            rs_item_master!Safety_Stock = CDbl(Trim(txt_safety_stock.Text))
        Else
            rs_item_master!Safety_Stock = "0"
        End If
        
        '20100917 tambahana edi
        
        
        
        If Trim(cboTypeAccs.Text) <> "" Then
             rs_item_master!TypeAccs = Trim(cboTypeAccs.Text)
        Else
             rs_item_master!TypeAccs = Null
        End If

        
        
        If Trim(txtSafetyStock2.Text) <> "" Then
            rs_item_master!Safety_Stock_percentage = CDbl(Trim(txtSafetyStock2.Text))
        Else
            rs_item_master!Safety_Stock_percentage = "0"
        End If
        
        If Trim(txt_max_stock.Text) <> "" Then
            rs_item_master!Max_Stock = CDbl(Trim(txt_max_stock.Text))
        Else
            rs_item_master!Max_Stock = "0"
        End If
        
        If Trim(txt_min_stock.Text) <> "" Then
            rs_item_master!Min_Stock = CDbl(Trim(txt_min_stock.Text))
        Else
            rs_item_master!Min_Stock = "0"
        End If
        
        If Trim(txt_allowance_day.Text) <> "" Then
            rs_item_master!Alowance_Day = CDbl(Trim(txt_allowance_day.Text))
        Else
            rs_item_master!Alowance_Day = "0"
        End If
        
        If Trim(txt_delivery_read_time.Text) <> "" Then
            rs_item_master!Delivery_ReadTime = CDbl(Trim(txt_delivery_read_time.Text))
        Else
            rs_item_master!Delivery_ReadTime = "0"
        End If
        
            rs_item_master!MakeBuy_Cls = Trim(cmb_make.Text)
            
        If Trim(cmbbox_control.Text) <> "" Then
            rs_item_master!control_cls = Trim(cmbbox_control.Text)
        Else
            rs_item_master!control_cls = Null
        End If
        
        If Trim(txt_order_point_qty.Text) <> "" Then
            rs_item_master!OrderPoint_Qty = CDbl(Trim(txt_order_point_qty.Text))
        Else
            rs_item_master!OrderPoint_Qty = "0"
        End If
        
        If Trim(TxtMinOrder.Text) <> "" Then
            rs_item_master!MinOrder = CDbl(Trim(TxtMinOrder.Text))
        Else
            rs_item_master!MinOrder = "0"
        End If
        
          If Trim(cmbbox_packing2.Text) <> "" Then
            rs_item_master!packingstylematerial_cls = Trim(cmbbox_packing2.Text)
        Else
            rs_item_master!packingstylematerial_cls = Null
        End If
        
            rs_item_master!Unit_cls = Trim(cmbbox_unit.Text)
        
        If Trim(txt_number_of_box.Text) <> "" Then
            rs_item_master!Number_Box = CDbl(Trim(txt_number_of_box.Text))
        Else
            rs_item_master!Number_Box = "0"
        End If

            rs_item_master!Accounting_Code = Trim(txt_accounting_code.Text)
            rs_item_master!explosion_cls = Trim(cmb_explosion2.Text)

        If Trim(cmbbox_purchase_person.Text) <> "" Then
            rs_item_master!personincharge_cls = Trim(cmbbox_purchase_person.Text)
        Else
            rs_item_master!personincharge_cls = Null
        End If

        If Trim(cboModel.Text) <> "" Then
            rs_item_master!Model_Cls = Trim(cboModel.Text)
        Else
            rs_item_master!Model_Cls = Null
        End If
        
        If Trim(cbo_POType.Text) <> "" Then
            rs_item_master!POType_Cls = Trim(cbo_POType.Text)
        Else
            rs_item_master!POType_Cls = Null
        End If
        
        If Trim(cboClasificationPart.Text) <> "" Then
            rs_item_master!ClasificationPart_Cls = Trim(cboClasificationPart.Text)
        Else
            rs_item_master!ClasificationPart_Cls = Null
        End If

         If Trim(cboDestination.Text) <> "" Then
            rs_item_master!Destination_Cls = Trim(cboDestination.Text)
        Else
            rs_item_master!Destination_Cls = Null
        End If
        
        If Trim(cboColor.Text) <> "" Then
            rs_item_master!Color_Cls = Trim(cboColor.Text)
        Else
            rs_item_master!Color_Cls = Null
        End If
        
            rs_item_master!stockcontrol_cls = Trim(cmb_stock_control2.Text)
                        
            MaskEdBox1.PromptInclude = False
            
            rs_item_master!Use_EndDay = Right(Trim(MaskEdBox1.Text), 4) & Left(Trim(MaskEdBox1.Text), 2) & Right(Left(Trim(MaskEdBox1.Text), 4), 2)
            rs_item_master!Last_Update = Now
            rs_item_master!last_user = userLogin

End Sub

Private Sub cmd_Browser_Click()
 If txt_item_code.Enabled = True Then
  Me.MousePointer = vbHourglass
  frm_BrowseItem.getItemCode = txt_item_code.Text
  frm_BrowseItem.Show 1
  txt_item_code.Text = frm_BrowseItem.getItemCode
  txt_item_code.SetFocus
  txt_item_code.SelStart = Len(Trim(txt_item_code))
  Me.MousePointer = vbDefault
 End If
End Sub

Private Sub command2_Click()
Dim sqlOrder As String
Dim recAff As Double
Dim tombol As Integer
If Trim(txt_item_code.Text) = "" Then
    lbl_pesan.Caption = "There is no data to delete !": Exit Sub
Else
    If rs_item_master.EOF = False Or rs_item_master.BOF = False Then
        rs_item_master.MoveFirst
            rs_item_master.Find "item_code='" & Trim(txt_item_code.Text) & "'"
        If rs_item_master.EOF = False Then
            GoTo nexto
        Else
            lbl_record.Caption = "Record 0 of 0"
            lbl_pesan.Caption = "There is no data with Product Code " & Trim(txt_item_code.Text) & " !"
            rs_item_master.MoveFirst
            Exit Sub
        End If
    End If
End If

nexto:

If hakUpdate(Me.Name) = 0 Then _
     lbl_pesan = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
Me.MousePointer = vbHourglass
If rs_bom_master.EOF = False Or rs_bom_master.BOF = False Then
    rs_bom_master.MoveFirst
        rs_bom_master.filter = "parent_itemcode ='" & l_txt_item_code & "' or item_code='" & l_txt_item_code & "'"
        If rs_bom_master.EOF = True Then
        
            sqlOrder = "select * from orderEntry_detail where item_code='" & l_txt_item_code & "'"
            Db.Execute sqlOrder, recAff
            
            If recAff = 0 Then
                tombol = MsgBox("Are you sure want to delete this data ?", vbQuestion + vbYesNo, "Warning")
                If tombol = vbYes Then
                    Db.Execute "delete from item_master where  item_code='" & l_txt_item_code & "'"
                    Call cmd_clear_Click
                    lbl_pesan.Caption = DisplayMsg(1201) '"Delete data success !"
            
                End If
            Else
                lbl_pesan.Caption = DisplayMsg("0038") & " Order Entry !"
            End If
           
        Else
            lbl_pesan.Caption = DisplayMsg("0038") & " BOM Master !"
        End If

Set rs_bom_master = Db.Execute("select* from bom_master")

Else
    sqlOrder = "select * from orderEntry_detail where item_code='" & l_txt_item_code & "'"
    Db.Execute sqlOrder, recAff
    
    If recAff = 0 Then
        tombol = MsgBox("Are you sure want to delete this data ?", vbQuestion + vbYesNo, "Warning")
        If tombol = vbYes Then
            Db.Execute "delete from item_master where  item_code='" & l_txt_item_code & "'"
            Call cmd_clear_Click
            lbl_pesan.Caption = DisplayMsg(1201) '"Delete data success !"
        End If
    Else
        lbl_pesan.Caption = DisplayMsg("0038") & " Order Entry !"
    End If

End If
Me.MousePointer = vbDefault
rs_item_master.Requery
End Sub

Private Sub Command1_Click(Index As Integer)

Select Case Index
Case 0:
    frmMainMenu.Show
    Unload Me
    
Case 1:
    If rs_item_master.EOF = False Or rs_item_master.BOF = False Then
    rs_item_master.MoveFirst
    Call data_tampil
    lbl_pesan.Caption = DisplayMsg("4020")
    End If
Case 2:
    If rs_item_master.EOF = False Or rs_item_master.BOF = False Then
    rs_item_master.MovePrevious: lbl_pesan.Caption = ""
    If rs_item_master.BOF Then rs_item_master.MoveFirst: lbl_pesan.Caption = DisplayMsg("4020") ': f_pesan = False
    Call data_tampil
    If rs_item_master.AbsolutePosition = 1 Then lbl_pesan.Caption = DisplayMsg("4020")
    End If
Case 3:
    If k_pertama = True Then
        If rs_item_master.EOF = False Or rs_item_master.BOF = False Then
        rs_item_master.MoveFirst
        Call data_tampil
        lbl_pesan.Caption = DisplayMsg("4020")
        k_pertama = False
        End If
    Else
        If rs_item_master.EOF = False Or rs_item_master.BOF = False Then
        rs_item_master.MoveNext: lbl_pesan.Caption = ""
        If rs_item_master.EOF Then rs_item_master.MoveLast: lbl_pesan.Caption = DisplayMsg("4021") ': f_pesan = False
        Call data_tampil
        If rs_item_master.AbsolutePosition = rs_item_master.RecordCount Then lbl_pesan.Caption = DisplayMsg("4021")
        End If
    End If
Case 4:
    If rs_item_master.EOF = False Or rs_item_master.BOF = False Then
    rs_item_master.MoveLast
    Call data_tampil
    lbl_pesan.Caption = DisplayMsg("4021")
    End If
End Select
End Sub

Private Sub command3_Click()

With frm_item_inquiry

frm_item_inquiry.CmdSubMenu.Caption = "&Back"
If .rs_item.State <> adStateClosed Then .rs_item.Requery
.uf_settingGrid (.is_sql)
.Show
End With
Me.Hide
End Sub

Private Sub DTPicker1_Change()
MaskEdBox1.Text = Format(DTPicker1.Value, "MM") & "/" & Format(DTPicker1.Value, "dd") & "/" & Format(DTPicker1.Value, "yyyy")
End Sub

Private Sub DTPicker1_Click()
MaskEdBox1.Text = Format(DTPicker1.Value, "MM") & "/" & Format(DTPicker1.Value, "dd") & "/" & Format(DTPicker1.Value, "yyyy")
End Sub

Private Sub Form_Load()
 If gb_Simulation = True Then Call up_InitSimulation(Me)
Label63 = ""
Label65 = ""

Dim ctr As CtrlMenu
txtMenu.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & txtMenu.MenuText & ")"
Call koneksi
Call setting
Call cmd_clear_Click
Call set_delivery
lbl_pesan.Caption = ""
Status = "insert"
SSTab1.Tab = 0
DTPicker1.Value = Format(Date, "dd MMM yyyy")
Label4.Caption = Format(Now, "dd MMM yyyy hh:mm:ss")
lbl_record.Caption = "Record 0 of 0"
k_pertama = True
MaskEdBox1.Text = "99/99/9999"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs_delivery.State <> adStateClosed Then rs_delivery.Close
If rs_item_master.State <> adStateClosed Then rs_item_master.Close
If rs_bom_master.State <> adStateClosed Then rs_bom_master.Close
If rs_hs_master.State <> adStateClosed Then rs_hs_master.Close
If rs_trade_master.State <> adStateClosed Then rs_trade_master.Close
If rs_material_cls.State <> adStateClosed Then rs_material_cls.Close
If rs_manufacture_line2.State <> adStateClosed Then rs_manufacture_line2.Close
If rs_manufacture_line.State <> adStateClosed Then rs_manufacture_line.Close
If rs_sheetcoil_cls.State <> adStateClosed Then rs_sheetcoil_cls.Close
If rs_drawingmaterial_cls.State <> adStateClosed Then rs_drawingmaterial_cls.Close
If rs_heattreatment_cls.State <> adStateClosed Then rs_heattreatment_cls.Close
If rs_surfacetreatment_cls.State <> adStateClosed Then rs_surfacetreatment_cls.Close
If rs_group_cls.State <> adStateClosed Then rs_group_cls.Close
If rs_packingstyle_cls.State <> adStateClosed Then rs_packingstyle_cls.Close
If rs_control_cls.State <> adStateClosed Then rs_control_cls.Close
If rs_personincharge_cls.State <> adStateClosed Then rs_personincharge_cls.Close
If rs_line.State <> adStateClosed Then rs_line.Close
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = Len(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub txt_accounting_code_GotFocus()
txt_accounting_code.SelStart = Len(txt_accounting_code)
End Sub

Private Sub txt_accounting_code_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub txt_allowance_day_GotFocus()
txt_allowance_day.SelLength = Len(txt_allowance_day)
End Sub

Private Sub txt_allowance_day_keypress(KeyAscii As Integer)
If Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txt_allowance_day_LostFocus()
If IsNumeric(txt_allowance_day) = False Then txt_allowance_day = 0
End Sub

Private Sub txt_delivery_read_time_GotFocus()
txt_delivery_read_time.SelLength = Len(txt_delivery_read_time)
End Sub

Private Sub txt_delivery_read_time_keypress(KeyAscii As Integer)
If Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txt_delivery_read_time_LostFocus()
If IsNumeric(txt_delivery_read_time) = False Then txt_delivery_read_time = 0
End Sub

Private Sub txt_ew_keypress(KeyAscii As Integer)
KeyAscii = 0: Exit Sub '#Enable False
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txt_ew_LostFocus()
If IsNumeric(txt_ew) = False Then txt_ew = 0
txt_ew = Format(txt_ew, gs_formatEW)
End Sub

Private Sub txt_gross_GotFocus()
txt_gross.SelLength = Len(txt_gross)
End Sub

Private Sub txt_gross_KeyPress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_gross, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_gross_LostFocus()
If IsNumeric(txt_gross) = False Then txt_gross = 0
txt_gross = Format(txt_gross, gs_formatWeight)
End Sub

Private Sub txt_heat_qty_GotFocus()
txt_heat_qty.SelLength = Len(txt_heat_qty)
End Sub

Private Sub txt_heat_qty_KeyPress(KeyAscii As Integer)
KeyAscii = 0: Exit Sub '#Enable False
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_heat_qty, ".") > 0 Then KeyAscii = 0
If Trim(txt_heat_qty) <> "" Then _
If Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 And CDbl(txt_heat_qty) > gs_formatQty Then KeyAscii = 0
End Sub

Private Sub txt_heat_qty_LostFocus()
If IsNumeric(txt_heat_qty) = False Then txt_heat_qty = 0
txt_heat_qty = Format(txt_heat_qty, gs_formatQty)
End Sub

Private Sub txt_item_code_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2

If KeyAscii = 13 Then
    If rs_item_master.EOF = False Or rs_item_master.BOF = False Then
    rs_item_master.MoveFirst
        rs_item_master.Find "item_code='" & Trim(txt_item_code.Text) & "'"
        If rs_item_master.EOF = False Then
            Call data_tampil: Status = "update"
            txt_item_code.Enabled = False
            lbl_record.Caption = "Record " & rs_item_master.AbsolutePosition & " of " & rs_item_master.RecordCount
            txt_item_name.SetFocus
        Else
            rs_item_master.MoveFirst
            l_txt_item_code2 = txt_item_code.Text
            Call cmd_clear_Click
            txt_item_code.Text = l_txt_item_code2
            lbl_pesan.Caption = DisplayMsg("8051")
            lbl_record.Caption = "Record 0 of 0"
            txt_item_code.Enabled = True
        End If
    End If
End If
End Sub

Private Sub detail_info()

'====== Suplier========================================
 If rs_trade_master.EOF = False Or rs_trade_master.BOF = False Then
    rs_trade_master.MoveFirst
    rs_trade_master.Find "trade_code='" & Trim(cmbox_suplier.Text) & "'"
    If rs_trade_master.EOF = False Then
        'lbl_suplier.Caption = Trim(rs_trade_master!trade_name)
        txt_supplier = Trim(rs_trade_master!trade_name)
        lbl_pesan.Caption = ""
    Else
        'lbl_suplier.Caption = ""
        txt_supplier = ""
    End If
End If
'====== HS Code ========================================
 If rs_hs_master.EOF = False Or rs_hs_master.BOF = False Then
    rs_hs_master.MoveFirst
    rs_hs_master.Find "hs_code='" & Trim(cmb_hs.Text) & "'"
    If rs_hs_master.EOF = False Then
        lbl_hs.Caption = Trim(rs_hs_master!tax)
        lbl_pesan.Caption = ""
    Else
        lbl_hs.Caption = ""
    End If
End If
'====== delivery========================================
If cmb_delivery.ListCount > 0 Then
  For i = 0 To cmb_delivery.ListCount - 1
            If Trim(cmb_delivery.Text) = Trim(cmb_delivery.List(i, 0)) Then
                  lbl_delivery_place.Caption = Trim(cmb_delivery.List(i, 1)): Exit For
            Else
                lbl_delivery_place.Caption = ""
            End If
  Next
Else
    lbl_delivery_place.Caption = ""
End If
'====== packing========================================
 If rs_packingstyle_cls.EOF = False Or rs_packingstyle_cls.BOF = False Then
    rs_packingstyle_cls.MoveFirst
    rs_packingstyle_cls.Find "packingstyle_cls='" & Trim(cmbbox_packing.Text) & "'"
    If rs_packingstyle_cls.EOF = False Then
        lbl_packing.Caption = Trim(rs_packingstyle_cls!Description)
        lbl_pesan.Caption = ""
    Else
        lbl_packing.Caption = ""
    End If
End If
'====== packing2========================================
 If rs_packingstyle_cls.EOF = False Or rs_packingstyle_cls.BOF = False Then
    rs_packingstyle_cls.MoveFirst
    rs_packingstyle_cls.Find "packingstyle_cls='" & Trim(cmbbox_packing2.Text) & "'"
    If rs_packingstyle_cls.EOF = False Then
        lbl_packing2.Caption = Trim(rs_packingstyle_cls!Description)
        lbl_pesan.Caption = ""
    Else
        lbl_packing2.Caption = ""
    End If
End If

'====== group========================================
If rs_group_cls.EOF = False Or rs_group_cls.BOF = False Then
    rs_group_cls.MoveFirst
    rs_group_cls.Find "group_cls='" & Trim(cmbbox_group.Text) & "'"
    If rs_group_cls.EOF = False Then
        lbl_group.Caption = Trim(rs_group_cls!Description)
        lbl_pesan.Caption = ""
    Else
        lbl_group.Caption = ""
    End If
End If

'====== production========================================
If Trim(cmb_prod.Text) = "" Then
     lbl_prod.Caption = ""
ElseIf Trim(cmb_prod.Text) = "01" Then
    lbl_prod.Caption = "Yes"
Else
    lbl_prod.Caption = "No"
End If
'====== purchase person========================================
If rs_personincharge_cls.EOF = False Or rs_personincharge_cls.BOF = False Then
    rs_personincharge_cls.MoveFirst
    rs_personincharge_cls.Find "personincharge_cls='" & Trim(cmbbox_purchase_person.Text) & "'"
    If rs_personincharge_cls.EOF = False Then
        lbl_purchase.Caption = Trim(rs_personincharge_cls!Description)
        lbl_pesan.Caption = ""
    Else
        lbl_purchase.Caption = ""
    End If
End If

'====== material========================================
If rs_material_cls.EOF = False Or rs_material_cls.BOF = False Then
    rs_material_cls.MoveFirst
    rs_material_cls.Find "material_cls='" & Trim(cmbbox_material.Text) & "'"
    If rs_material_cls.EOF = False Then
        lbl_material.Caption = Trim(rs_material_cls!Description)
        lbl_pesan.Caption = ""
    Else
        lbl_material.Caption = ""
    End If
End If

'====== Material Packing========================================
 If rs_packingstyle_cls.EOF = False Or rs_packingstyle_cls.BOF = False Then
    rs_packingstyle_cls.MoveFirst
    rs_packingstyle_cls.Find "packingstyle_cls='" & Trim(cmbbox_mPacking.Text) & "'"
    If rs_packingstyle_cls.EOF = False Then
        lbl_mPacking.Caption = Trim(rs_packingstyle_cls!Description)
        lbl_pesan.Caption = ""
    Else
        lbl_mPacking.Caption = ""
    End If
End If

'====== sheetcoil========================================
If rs_sheetcoil_cls.EOF = False Or rs_sheetcoil_cls.BOF = False Then
    rs_sheetcoil_cls.MoveFirst
    rs_sheetcoil_cls.Find "sheetcoil_cls='" & Trim(cmbbox_sheet.Text) & "'"
    If rs_sheetcoil_cls.EOF = False Then
        lbl_sheet.Caption = Trim(rs_sheetcoil_cls!Description)
        lbl_pesan.Caption = ""
    Else
        lbl_sheet.Caption = ""
    End If
End If

'====== drawing material========================================
If rs_drawingmaterial_cls.EOF = False Or rs_drawingmaterial_cls.BOF = False Then
    rs_drawingmaterial_cls.MoveFirst
    rs_drawingmaterial_cls.Find "drawingmaterial_cls='" & Trim(cmbbox_drawing.Text) & "'"
    If rs_drawingmaterial_cls.EOF = False Then
        lbl_drawing.Caption = Trim(rs_drawingmaterial_cls!Description)
        lbl_pesan.Caption = ""
    Else
        lbl_drawing.Caption = ""
    End If
End If

'====== surface treatment========================================
If rs_surfacetreatment_cls.EOF = False Or rs_surfacetreatment_cls.BOF = False Then
    rs_surfacetreatment_cls.MoveFirst
    rs_surfacetreatment_cls.Find "surfacetreatment_cls='" & Trim(cmbbox_surface.Text) & "'"
    If rs_surfacetreatment_cls.EOF = False Then
        lbl_surface.Caption = Trim(rs_surfacetreatment_cls!Description)
        lbl_pesan.Caption = ""
    Else
        lbl_surface.Caption = ""
    End If
End If

'====== heat treatment========================================
If rs_heattreatment_cls.EOF = False Or rs_heattreatment_cls.BOF = False Then
    rs_heattreatment_cls.MoveFirst
    rs_heattreatment_cls.Find "heattreatment_cls='" & Trim(cmbbox_heat.Text) & "'"
    If rs_heattreatment_cls.EOF = False Then
        lbl_heat.Caption = Trim(rs_heattreatment_cls!Description)
        lbl_pesan.Caption = ""
    Else
        lbl_heat.Caption = ""
    End If
End If
End Sub

Private Sub data_tampil()
    
    Status = "update"
    l_txt_item_code = Trim(rs_item_master!Item_Code)
    
    txt_item_code.Text = Trim(rs_item_master!Item_Code): txt_item_name.Text = Trim(rs_item_master!item_name)
    
    cmbbox_warehouse.Text = Trim(IIf(IsNull(rs_item_master!wh_code), "", rs_item_master!wh_code))
           
    If Trim(rs_item_master!Drawing_Number) <> "null" Then
        txt_drawingCode = Trim(rs_item_master!Drawing_Number)
    Else
        txt_drawingCode = ""
    End If
    
    If Trim(rs_item_master!Address) <> "null" Then
        Text1.Text = Trim(rs_item_master!Address)
    Else
        Text1.Text = ""
    End If
    Call cmbbox_warehouse_Change
    
    
    cmbox_suplier.Text = Trim(rs_item_master!Supplier_Code)
    cmb_hs.Text = IIf(IsNull(rs_item_master!HS_Code), "", Trim(rs_item_master!HS_Code))
    Call cmb_hs_Change
    Call set_delivery
    cmb_delivery.Text = Trim(rs_item_master!delivery_place & "")

    
    If Trim(rs_item_master!Manufacture_Code) <> "null" Then
        cmbox_manufacture.Text = Trim(rs_item_master!Manufacture_Code)
    Else
        cmbox_manufacture.Text = "": txt_factory = ""
    End If
        
        
    If Trim(rs_item_master!TypeAccs) <> "null" Then
        cboTypeAccs.Text = Trim(rs_item_master!TypeAccs)
    Else
        cboTypeAccs.Text = ""
    End If
        
        
        
    If Trim(rs_item_master!line_code) <> "null" Then
        cmbbox_line.Text = Trim(rs_item_master!line_code)
        If cmbbox_line.ListCount > 0 Then
            For i = 0 To cmbbox_line.ListCount - 1
                If Trim(cmbbox_line.List(i, 0)) = Trim(rs_item_master!line_code) Then
                    lbl_line.Caption = cmbbox_line.List(i, 1): Exit For
                Else
                    lbl_line.Caption = ""
                End If
            Next
        End If
    Else
        cmbbox_line.Text = "": lbl_line.Caption = ""
    End If

    txt_maker_item_code.Text = Trim(rs_item_master!MakerItem_Code)
    
    If Trim(rs_item_master!finishgoodpart_cls) <> "null" Then
        cmb_finish_good.Text = Trim(rs_item_master!finishgoodpart_cls)
    Else
        cmb_finish_good.Text = ""
        lbl_finish.Caption = ""
    End If
    
    If Trim(rs_item_master!part_cls) <> "null" Then
        cmb_part2.Text = Trim(rs_item_master!part_cls)
    Else
        cmb_part2.Text = ""
    End If
        
    cmb_reserve2.Text = Trim(rs_item_master!reserve_cls):  cmb_supply2.Text = Trim(rs_item_master!suply_cls)
    cmb_provision2.Text = Trim(rs_item_master!Provision_Cls)
    
    If Trim(rs_item_master!Material_Cls) <> "null" Then
        cmbbox_material.Text = Trim(rs_item_master!Material_Cls)
    Else
        cmbbox_material.Text = ""
    End If
    
    If Trim(rs_item_master!drawingmaterial_cls) <> "null" Then
        cmbbox_drawing.Text = Trim(rs_item_master!drawingmaterial_cls)
    Else
        cmbbox_drawing.Text = ""
    End If
    
    If Trim(rs_item_master!surfacetreatment_cls) <> "null" Then
        cmbbox_surface.Text = Trim(rs_item_master!surfacetreatment_cls)
    Else
        cmbbox_surface.Text = ""
    End If
    
    If Trim(rs_item_master!heattreatment_cls) <> "null" Then
        cmbbox_heat.Text = Trim(rs_item_master!heattreatment_cls)
    Else
        cmbbox_heat.Text = ""
    End If
    
    If Trim(rs_item_master!sheetcoil_cls) <> "null" Then
        cmbbox_sheet.Text = Trim(rs_item_master!sheetcoil_cls)
    Else
        cmbbox_sheet.Text = ""
    End If
    
    txt_sample.Text = Format(Trim(rs_item_master!Sample), gs_formatQty)
    txt_sw.Text = Format(Trim(rs_item_master!SW_Qty), gs_formatQty)
    txt_ew.Text = Format(Trim(rs_item_master!EW_Qty), gs_formatQty)
    txt_np.Text = Trim(rs_item_master!Number_Process & "")
    txt_mc.Text = Format(Trim(rs_item_master!Material_Coefficient), gs_formatCoefficient)
    txt_pc.Text = Format(Trim(rs_item_master!Process_Coefficient), gs_formatCoefficient)
    txt_min_lot.Text = Format(Trim(rs_item_master!Min_Lot), gs_formatLot)
    txt_lot.Text = Format(Trim(rs_item_master!Lot_Qty), gs_formatLot)
    txt_lot_coef.Text = Format(Trim(rs_item_master!Lot_Coefficience), gs_formatCoefficient)
    txt_surface_qty.Text = Format(Trim(rs_item_master!Surface_OrderPointQty), gs_formatQty)
    txt_heat_qty.Text = Format(Trim(rs_item_master!Heat_OrderPointQty), gs_formatQty)
    txt_prt.Text = Trim(rs_item_master!Product_ReadTime & "")
    txt_yp.Text = Format(Trim(rs_item_master!Yield_Percentage), gs_formatPercentage)
    txt_thickness.Text = Format(Trim(rs_item_master!Thickness), gs_formatThickness)
    txt_width.Text = Format(Trim(rs_item_master!Width), gs_formatWidth)
    txt_weight.Text = Format(Trim(rs_item_master!Weight), gs_formatWeight)
    txt_gross.Text = Format(Trim(rs_item_master!GrossWeight), gs_formatWeight)
    txt_length.Text = Format(Trim(rs_item_master!Length), gs_formatLength)
    txt_ne.Text = Format(Trim(rs_item_master!number_entering), gs_formatQty)
    txt_pitch.Text = Format(Trim(rs_item_master!pitch), gs_formatPitch)
    txt_number_producible.Text = Format(Trim(rs_item_master!number_producible), gs_formatQty)
    txt_scrap_weight.Text = Format(Trim(rs_item_master!scrap_weight), gs_formatWeight)

    If Trim(rs_item_master!PackingStyle_Cls) <> "null" Then
        cmbbox_packing.Text = Trim(rs_item_master!PackingStyle_Cls)
    Else
        cmbbox_packing.Text = ""
    End If
    
    If Trim(rs_item_master!group_cls) <> "null" Then
        cmbbox_group.Text = Trim(rs_item_master!group_cls)
    Else
        cmbbox_group.Text = ""
    End If
    
    If Trim(rs_item_master!production_cls) <> "null" Then
        cmb_prod.Text = Trim(rs_item_master!production_cls)
    Else
        cmb_prod.Text = ""
    End If
        
    txt_standart_stock.Text = Format(Trim(rs_item_master!Standard_Stock), gs_formatQty)
    txt_safety_stock.Text = Format(Trim(rs_item_master!Safety_Stock), gs_formatQty)
    txtSafetyStock2.Text = Format(Trim(rs_item_master!Safety_Stock_percentage), gs_formatQty)
    txt_max_stock.Text = Format(Trim(rs_item_master!Max_Stock), gs_formatQty)
    txt_min_stock.Text = Format(Trim(rs_item_master!Min_Stock), gs_formatQty)
    txt_allowance_day.Text = Trim(rs_item_master!Alowance_Day & "")
    txt_delivery_read_time.Text = Trim(rs_item_master!Delivery_ReadTime & "")
    cmb_make.Text = Trim(rs_item_master!MakeBuy_Cls): cmbbox_control.Text = Trim(rs_item_master!control_cls)
    txt_order_point_qty.Text = Format(Trim(rs_item_master!OrderPoint_Qty), gs_formatQty)
    ' Add For Kawai - 20090501
    TxtMinOrder.Text = Format(Trim(rs_item_master!MinOrder), gs_formatQty)
    
    cmbbox_unit.Text = IIf(IsNull(Trim(rs_item_master!Unit_cls)), "", Trim(rs_item_master!Unit_cls))
    
    If Trim(rs_item_master!packingstylematerial_cls) <> "null" Then
        cmbbox_packing2.Text = Trim(rs_item_master!packingstylematerial_cls)
    Else
        cmbbox_packing2.Text = ""
    End If
    
    txt_number_of_box.Text = Format(Trim(rs_item_master!Number_Box), gs_formatBox)
    txt_accounting_code.Text = Trim(rs_item_master!Accounting_Code & "")
    cmb_explosion2.Text = IIf(IsNull(Trim(rs_item_master!explosion_cls)), "", Trim(rs_item_master!explosion_cls))

    If Trim(rs_item_master!personincharge_cls) <> "null" Then
        cmbbox_purchase_person.Text = Trim(rs_item_master!personincharge_cls)
    Else
        cmbbox_purchase_person.Text = ""
    End If

    If Trim(rs_item_master!Model_Cls) <> "null" Then
        cboModel.Text = Trim(rs_item_master!Model_Cls)
    Else
        cboModel.Text = ""
        lblModel.Caption = ""
    End If
    
    If Trim(rs_item_master!POType_Cls) <> "null" Then
        cbo_POType.Text = Trim(rs_item_master!POType_Cls)
    Else
        cbo_POType.Text = ""
        lbl_POType.Caption = ""
    End If
    
    If Trim(rs_item_master!Destination_Cls) <> "null" Then
        cboDestination.Text = Trim(rs_item_master!Destination_Cls)
    Else
        cboDestination.Text = ""
        lbl_Destination.Caption = ""
    End If
    
    If Trim(rs_item_master!Color_Cls) <> "null" Then
        cboColor.Text = Trim(rs_item_master!Color_Cls)
    Else
        cboColor.Text = ""
        lbl_Color.Caption = ""
    End If
    
     If Trim(rs_item_master!ClasificationPart_Cls) <> "null" Then
        cboClasificationPart.Text = Trim(rs_item_master!ClasificationPart_Cls)
    Else
        cboClasificationPart.Text = ""
        lblClasificationPart.Caption = ""
    End If
    
    cmb_stock_control2.Text = Trim(rs_item_master!stockcontrol_cls)
    If IsDate(Left(Right(Trim(rs_item_master!Use_EndDay), 4), 2) & " " & Right(Trim(rs_item_master!Use_EndDay), 2) & " " & Left(Trim(rs_item_master!Use_EndDay), 4)) = True Then _
        DTPicker1.Value = Left(Right(Trim(rs_item_master!Use_EndDay), 4), 2) & " " & Right(Trim(rs_item_master!Use_EndDay), 2) & " " & Left(Trim(rs_item_master!Use_EndDay), 4)

    MaskEdBox1.Text = Left(Right(Trim(rs_item_master!Use_EndDay), 4), 2) & "/" & Right(Trim(rs_item_master!Use_EndDay), 2) & "/" & Left(Trim(rs_item_master!Use_EndDay), 4)
    Label4.Caption = Format(Trim(rs_item_master!Last_Update), "dd MMM yyyy hh:mm:ss")
    txt_item_code.Enabled = False
    Call detail_info
    lbl_record.Caption = "Record " & rs_item_master.AbsolutePosition & " of " & rs_item_master.RecordCount

End Sub

Private Sub txt_item_code_LostFocus()
'If rs_item_master.EOF = False Or rs_item_master.BOF = False Then
'    rs_item_master.MoveFirst
'        rs_item_master.Find "item_code='" & Trim(txt_item_code.Text) & "'"
'    If rs_item_master.EOF = False Then
'        Call data_tampil: status = "update"
'        lbl_record.Caption = "Record " & rs_item_master.AbsolutePosition & " of " & rs_item_master.RecordCount
'        txt_item_code.Enabled = False
'    Else
'        lbl_record.Caption = "Record 0 of 0"
'        rs_item_master.MoveFirst
'    End If
'End If
End Sub

Private Sub txt_item_name_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Or KeyAscii = 34 Then KeyAscii = 0
End Sub

Private Sub txt_Length_GotFocus()
txt_length.SelLength = Len(txt_length)
End Sub

Private Sub txt_Length_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_length, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_Length_LostFocus()
If IsNumeric(txt_length) = False Then txt_length = 0
txt_length = Format(txt_length, gs_formatLength)
End Sub

Private Sub txt_lot_coef_GotFocus()
txt_lot_coef.SelLength = Len(txt_lot_coef)
End Sub

Private Sub txt_lot_coef_KeyPress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_lot_coef, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_lot_coef_LostFocus()
If IsNumeric(txt_lot_coef) = False Then txt_lot_coef = 0
txt_lot_coef = Format(txt_lot_coef, gs_formatCoefficient)
End Sub

Private Sub txt_lot_GotFocus()
txt_lot.SelLength = Len(txt_lot)
End Sub

Private Sub txt_lot_keypress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_lot, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_lot_LostFocus()
If IsNumeric(txt_lot) = False Then txt_lot = 0
txt_lot = Format(txt_lot, gs_formatQty)
End Sub

Private Sub txt_maker_item_code_GotFocus()
txt_maker_item_code.SelStart = Len(txt_maker_item_code)
End Sub

Private Sub txt_maker_item_code_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub

Private Sub txt_max_stock_GotFocus()
txt_max_stock.SelLength = Len(txt_max_stock)
End Sub

Private Sub txt_max_stock_keypress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_max_stock, ".") > 0 Then KeyAscii = 0
If Trim(txt_max_stock) <> "" Then _
If Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 And CDbl(txt_max_stock) > gd_MaxQty Then KeyAscii = 0
End Sub

Private Sub txt_max_stock_LostFocus()
If IsNumeric(txt_max_stock) = False Then txt_max_stock = 0
txt_max_stock = Format(txt_max_stock, gs_formatQty)
End Sub

Private Sub txt_mc_GotFocus()
txt_mc.SelLength = Len(txt_mc)
End Sub

Private Sub txt_mc_keypress(KeyAscii As Integer)
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_mc, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_mc_LostFocus()
If IsNumeric(txt_mc) = False Then txt_mc = 0
txt_mc = Format(txt_mc, gs_formatCoefficient)
End Sub

Private Sub txt_min_lot_GotFocus()
txt_min_lot.SelLength = Len(txt_min_lot)
End Sub

Private Sub txt_min_lot_keypress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_min_lot, ".") > 0 Then KeyAscii = 0
If Trim(txt_min_lot) <> "" Then _
If Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 And CDbl(txt_min_lot) > gd_MaxLot Then KeyAscii = 0
End Sub

Private Sub txt_min_lot_LostFocus()
If IsNumeric(txt_min_lot) = False Then txt_min_lot = 0
txt_min_lot = Format(txt_min_lot, gs_formatLot)
End Sub

Private Sub txt_min_stock_GotFocus()
txt_min_stock.SelLength = Len(txt_min_stock)
End Sub

Private Sub txt_min_stock_keypress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_min_stock, ".") > 0 Then KeyAscii = 0
If Trim(txt_min_stock) <> "" Then _
If Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 And CDbl(txt_min_stock) > gd_MaxQty Then KeyAscii = 0
End Sub

Private Sub txt_min_stock_LostFocus()
If IsNumeric(txt_min_stock) = False Then txt_min_stock = 0
txt_min_stock = Format(txt_min_stock, gs_formatQty)
End Sub

Private Sub txt_ne_GotFocus()
txt_ne.SelLength = Len(txt_ne)
End Sub

Private Sub txt_np_GotFocus()
txt_np.SelLength = Len(txt_np)
End Sub

Private Sub txt_np_keypress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_np, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_np_LostFocus()
If IsNumeric(txt_np) = False Then txt_np = 0
End Sub

Private Sub txt_number_of_box_GotFocus()
txt_number_of_box.SelLength = Len(txt_number_of_box)
End Sub

Private Sub txt_number_of_box_keypress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_number_of_box, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_number_of_box_LostFocus()
If IsNumeric(txt_number_of_box) = False Then txt_number_of_box = 0
txt_number_of_box = Format(txt_number_of_box, gs_formatQty)
End Sub

Private Sub txt_number_producible_GotFocus()
txt_number_producible.SelLength = Len(txt_number_producible)
End Sub

Private Sub txt_number_producible_keypress(KeyAscii As Integer)
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txt_number_producible_LostFocus()
If IsNumeric(txt_number_producible) = False Then txt_number_producible = 0
txt_number_producible = Format(txt_number_producible, gs_formatQty)
End Sub

Private Sub txt_order_point_qty_GotFocus()
txt_order_point_qty.SelLength = Len(txt_order_point_qty)
End Sub

Private Sub txt_order_point_qty_keypress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_order_point_qty, ".") > 0 Then KeyAscii = 0
If Trim(txt_order_point_qty) <> "" Then _
If Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 And CDbl(txt_order_point_qty) > gd_MaxQty Then KeyAscii = 0
End Sub

Private Sub txt_order_point_qty_LostFocus()
If IsNumeric(txt_order_point_qty) = False Then txt_order_point_qty = 0
txt_order_point_qty = Format(txt_order_point_qty, gs_formatQty)
End Sub

Private Sub txt_pc_GotFocus()
txt_pc.SelLength = Len(txt_pc)
End Sub

Private Sub txt_pc_keypress(KeyAscii As Integer)
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
  If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_pc, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_pc_LostFocus()
If IsNumeric(txt_pc) = False Then txt_pc = 0
txt_pc = Format(txt_pc, gs_formatCoefficient)
End Sub

Private Sub txt_pitch_GotFocus()
txt_pitch.SelLength = Len(txt_pitch)
End Sub

Private Sub txt_pitch_keypress(KeyAscii As Integer)
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
  If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_pitch, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_pitch_LostFocus()
If IsNumeric(txt_pitch) = False Then txt_pitch = 0
txt_pitch = Format(txt_pitch, gs_formatPitch)
End Sub

Private Sub txt_prt_GotFocus()
txt_prt.SelLength = Len(txt_prt)
End Sub

Private Sub txt_prt_keypress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_prt, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_prt_LostFocus()
If IsNumeric(txt_prt) = False Then txt_prt = 0
End Sub

Private Sub txt_safety_stock_Change()
    If txt_safety_stock <> "" Then
        If CDbl(txt_safety_stock) > 0 Then
            txtSafetyStock2 = "0.00"
        End If
    End If
End Sub

Private Sub txt_safety_stock_GotFocus()
txt_safety_stock.SelLength = Len(txt_safety_stock)
End Sub

Private Sub txt_safety_stock_keypress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_safety_stock, ".") > 0 Then KeyAscii = 0
If Trim(txt_safety_stock) <> "" Then _
If Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 And CDbl(txt_safety_stock) > gd_MaxQty Then KeyAscii = 0
End Sub

Private Sub txt_safety_stock_LostFocus()
If IsNumeric(txt_safety_stock) = False Then txt_safety_stock = 0
txt_safety_stock = Format(txt_safety_stock, gs_formatQty)
End Sub

Private Sub txt_sample_GotFocus()
txt_sample.SelLength = Len(txt_sample)
End Sub

Private Sub txt_sample_keypress(KeyAscii As Integer)
KeyAscii = 0: Exit Sub '#Enable False
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
  If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_sample, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_sample_LostFocus()
If IsNumeric(txt_sample) = False Then txt_sample = 0
txt_sample = Format(txt_sample, gs_formatQty)
End Sub

Private Sub txt_scrap_weight_GotFocus()
txt_scrap_weight.SelLength = Len(txt_scrap_weight)
End Sub

Private Sub txt_scrap_weight_keypress(KeyAscii As Integer)
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
  If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_scrap_weight, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_scrap_weight_LostFocus()
If IsNumeric(txt_scrap_weight) = False Then txt_scrap_weight = 0
txt_scrap_weight = Format(txt_scrap_weight, gs_formatWeight)
End Sub

Private Sub txt_standart_stock_GotFocus()
txt_standart_stock.SelLength = Len(txt_standart_stock)
End Sub

Private Sub txt_standart_stock_keypress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_standart_stock, ".") > 0 Then KeyAscii = 0
If Trim(txt_standart_stock) <> "" Then _
If Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 And CDbl(txt_standart_stock) > gd_MaxQty Then KeyAscii = 0
End Sub

Private Sub txt_standart_stock_LostFocus()
If IsNumeric(txt_standart_stock) = False Then txt_standart_stock = 0
txt_standart_stock = Format(txt_standart_stock, gs_formatQty)
End Sub

Private Sub txt_surface_qty_GotFocus()
txt_surface_qty.SelLength = Len(txt_surface_qty)
End Sub

Private Sub txt_surface_qty_KeyPress(KeyAscii As Integer)
KeyAscii = 0: Exit Sub '#Enable False
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_surface_qty, ".") > 0 Then KeyAscii = 0
If Trim(txt_surface_qty) <> "" Then _
If Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 And CDbl(txt_surface_qty) > gd_MaxQty Then KeyAscii = 0
End Sub

Private Sub txt_surface_qty_LostFocus()
If IsNumeric(txt_surface_qty) = False Then txt_surface_qty = 0
txt_surface_qty = Format(txt_surface_qty, gs_formatQty)
End Sub

Private Sub txt_sw_GotFocus()
txt_sw.SelLength = Len(txt_sw)
End Sub

Private Sub txt_sw_keypress(KeyAscii As Integer)
KeyAscii = 0: Exit Sub '#Enable False
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
  If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_sw, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_sw_LostFocus()
If IsNumeric(txt_sw) = False Then txt_sw = 0
txt_sw = Format(txt_sw, gs_formatSW)
End Sub

Private Sub txt_thickness_GotFocus()
txt_thickness.SelLength = Len(txt_thickness)
End Sub

Private Sub txt_thickness_KeyPress(KeyAscii As Integer)
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_thickness, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_ne_LostFocus()
If IsNumeric(txt_ne) = False Then txt_ne = 0
txt_ne = Format(txt_ne, gs_formatQty)
End Sub

Private Sub txt_ne_keypress(KeyAscii As Integer)
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_ne, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_thickness_LostFocus()
If IsNumeric(txt_thickness) = False Then txt_thickness = 0
txt_thickness = Format(txt_thickness, gs_formatThickness)
End Sub

Private Sub txt_weight_GotFocus()
txt_weight.SelLength = Len(txt_weight)
End Sub

Private Sub txt_weight_KeyPress(KeyAscii As Integer)
If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_weight, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_weight_LostFocus()
If IsNumeric(txt_weight) = False Then txt_weight = 0
txt_weight = Format(txt_weight, gs_formatWeight)
End Sub

Private Sub txt_width_GotFocus()
txt_width.SelLength = Len(txt_width)
End Sub

Private Sub txt_width_KeyPress(KeyAscii As Integer)
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_width, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_width_LostFocus()
If IsNumeric(txt_width) = False Then txt_width = 0
txt_width = Format(txt_width, gs_formatWidth)
End Sub

Private Sub txt_yp_GotFocus()
txt_yp.SelLength = Len(txt_yp)
End Sub

Private Sub txt_yp_KeyPress(KeyAscii As Integer)
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_yp, ".") > 0 Then KeyAscii = 0
End Sub

Private Sub txt_yp_LostFocus()
If IsNumeric(txt_yp) = False Then txt_yp = 0
txt_yp = Format(txt_yp, gs_formatPercentage)
End Sub

Private Sub txtmenu_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    lbl_pesan.Caption = ErrMsg
End If
End Sub

Function ValidStock(Item As String) As Boolean
Dim RsStock As Recordset
sql = "select sum(isnull(lm_inventory,0)) + sum(isnull(tm_Current,0)) + sum(isnull(nm_current,0)) stock from stock_master where item_code ='" & Item & "'"
Set RsStock = Db.Execute(sql)
If RsStock.EOF Then
    ValidStock = True
Else
    If IsNull(RsStock!Stock) Then
        ValidStock = True
    ElseIf RsStock!Stock > 0 Then
        ValidStock = False
    Else
        ValidStock = True
    End If
End If
End Function

Private Sub TxtMinOrder_GotFocus()
TxtMinOrder.SelLength = Len(TxtMinOrder)
End Sub

Private Sub TxtMinOrder_KeyPress(KeyAscii As Integer)
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = 46 And InStr(1, txt_yp, ".") > 0 Then KeyAscii = 0

End Sub

Private Sub TxtMinOrder_LostFocus()
If IsNumeric(TxtMinOrder) = False Then TxtMinOrder = 0
TxtMinOrder = Format(TxtMinOrder, gs_formatQty)

End Sub

Private Sub txtModel_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSafetyStock2_Change()
    If txtSafetyStock2 <> "" Then
        If CDbl(txtSafetyStock2) > 0 Then
            txt_safety_stock = "0.00"
        End If
    End If
End Sub

Private Sub txtSafetyStock2_GotFocus()
    txtSafetyStock2.SelLength = Len(txtSafetyStock2)
End Sub

Private Sub txtSafetyStock2_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
          KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then SendKeys vbTab
    If (txtSafetyStock2.Text & Chr(KeyAscii)) > 100 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtSafetyStock2_LostFocus()
    If IsNumeric(txtSafetyStock2) = False Then txtSafetyStock2 = 0
    txtSafetyStock2 = Format(txtSafetyStock2, gs_formatQty)
End Sub
