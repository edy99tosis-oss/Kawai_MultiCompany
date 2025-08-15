VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmBC41Detail 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BC 41 Detail"
   ClientHeight    =   10680
   ClientLeft      =   2835
   ClientTop       =   450
   ClientWidth     =   15255
   Icon            =   "FrmBC41Detail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleMode       =   0  'User
   ScaleWidth      =   50582.37
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2535
      Left            =   360
      TabIndex        =   74
      Tag             =   "TFTF*/"
      Top             =   720
      Width           =   14565
      Begin VB.TextBox txtTampung 
         Appearance      =   0  'Flat
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
         Left            =   10320
         TabIndex        =   203
         Tag             =   "TTFF*/"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtJabatan 
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
         Height          =   315
         Left            =   10320
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   1605
         Width           =   2415
      End
      Begin VB.TextBox txtPemberitahu 
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
         Height          =   315
         Left            =   10320
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtTempat 
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
         Height          =   315
         Left            =   10320
         TabIndex        =   6
         Tag             =   "TTFF*/"
         Top             =   780
         Width           =   2415
      End
      Begin VB.TextBox txtNoDaftar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "TTFF*/"
         Top             =   800
         Width           =   1815
      End
      Begin VB.CommandButton cmdAction 
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
         Index           =   3
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   75
         Tag             =   "TTFF*/"
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtKantorPabean 
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
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   1600
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtNoPengajuan 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   29
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "######-######-########-######"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpTglDaftar 
         Height          =   315
         Left            =   1920
         TabIndex        =   76
         Tag             =   "TTFF*/"
         Top             =   1200
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483648
         CalendarTitleBackColor=   -2147483648
         CustomFormat    =   "dd MMM yyyy"
         Format          =   288882691
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   315
         Left            =   12840
         TabIndex        =   7
         Tag             =   "TTFF*/"
         Top             =   780
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
         Format          =   288882691
         CurrentDate     =   37798
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
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
         Left            =   8520
         TabIndex        =   86
         Tag             =   "TTFF*/"
         Top             =   1665
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pemberitahu"
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
         Left            =   8520
         TabIndex        =   85
         Tag             =   "TTFF*/"
         Top             =   1260
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat/Tanggal"
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
         Left            =   8520
         TabIndex        =   84
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1395
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3360
         X2              =   6000
         Y1              =   1900
         Y2              =   1900
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tujuan Pengiriman"
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
         Left            =   8520
         TabIndex        =   83
         Tag             =   "TTFF*/"
         Top             =   390
         Width           =   1695
      End
      Begin MSForms.ComboBox cboTujuan 
         Height          =   315
         Left            =   10320
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   3975
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7011;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Daftar"
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
         Left            =   120
         TabIndex        =   82
         Tag             =   "TTFF*/"
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Daftar"
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
         TabIndex        =   81
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Pengajuan"
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
         TabIndex        =   80
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label lblKantorPabean 
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
         Left            =   3360
         TabIndex        =   79
         Tag             =   "TTFF*/"
         Top             =   1630
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Kantor Pabean"
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
         Left            =   120
         TabIndex        =   78
         Tag             =   "TTFF*/"
         Top             =   1630
         Width           =   1455
      End
      Begin MSForms.ComboBox cboJenisTPB 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   2040
         Width           =   3975
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7011;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis TPB"
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
         Left            =   120
         TabIndex        =   77
         Tag             =   "TTFF*/"
         Top             =   2070
         Width           =   1695
      End
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
      Left            =   13740
      Style           =   1  'Graphical
      TabIndex        =   64
      Tag             =   "FFTT*/"
      Top             =   10185
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_SubMenu 
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
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   66
      Tag             =   "TFFT*/"
      Top             =   10185
      Width           =   1125
   End
   Begin VB.CommandButton CmdSyncronize 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Syncronize"
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   65
      Tag             =   "FFTT*/"
      Top             =   10185
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   360
      TabIndex        =   68
      Tag             =   "TFTT*/"
      Top             =   9465
      Width           =   14505
      Begin VB.Label LblErrMsg 
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
         Left            =   90
         TabIndex        =   69
         Tag             =   "TTTF*/"
         Top             =   180
         Width           =   14325
      End
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
      Left            =   11220
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "FFTT*/"
      Top             =   10200
      Visible         =   0   'False
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13140
      TabIndex        =   70
      TabStop         =   0   'False
      Tag             =   "FTTF*/"
      Top             =   120
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2055
      Left            =   360
      TabIndex        =   49
      Tag             =   "TFTF*/"
      Top             =   3360
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   3625
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   5
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
      TabCaption(0)   =   "Pengusaha TPB"
      TabPicture(0)   =   "FrmBC41Detail.frx":0E42
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Pengirim Barang"
      TabPicture(1)   =   "FrmBC41Detail.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   " Pengangkutan"
      TabPicture(2)   =   "FrmBC41Detail.frx":0E7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Harga"
      TabPicture(3)   =   "FrmBC41Detail.frx":0E96
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame9"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   96
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14280
         Begin VB.TextBox txtAlamatPengusaha 
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
            Height          =   315
            Left            =   8400
            TabIndex        =   14
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   5415
         End
         Begin VB.TextBox txtNoIzinPengusaha 
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
            Height          =   315
            Left            =   2160
            TabIndex        =   13
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtNamaPengusaha 
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
            Height          =   315
            Left            =   2160
            TabIndex        =   12
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   4935
         End
         Begin VB.TextBox txtNPWPPengusaha 
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
            Height          =   315
            Left            =   4440
            MaxLength       =   20
            TabIndex        =   11
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
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
            Left            =   7320
            TabIndex        =   100
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "No Izin"
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
            TabIndex        =   99
            Tag             =   "TTFF*/"
            Top             =   1110
            Width           =   975
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
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
            TabIndex        =   98
            Tag             =   "TTFF*/"
            Top             =   750
            Width           =   975
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP"
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
            TabIndex        =   97
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1575
         End
         Begin MSForms.ComboBox cboNPWPPengusaha 
            Height          =   315
            Left            =   2160
            TabIndex        =   10
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   2175
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3836;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   92
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14175
         Begin VB.TextBox txtNamaPengirim 
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
            Height          =   315
            Left            =   2040
            TabIndex        =   17
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   4935
         End
         Begin VB.TextBox txtAlamatPengirim 
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
            Height          =   315
            Left            =   2040
            TabIndex        =   18
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   4935
         End
         Begin VB.TextBox txtNPWPPengirim 
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
            Height          =   315
            Left            =   4320
            TabIndex        =   16
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Identitas"
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
            TabIndex        =   95
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
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
            TabIndex        =   94
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
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
            TabIndex        =   93
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1575
         End
         Begin MSForms.ComboBox cboNPWPPengirim 
            Height          =   315
            Left            =   2040
            TabIndex        =   15
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2175
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3836;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   89
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14295
         Begin VB.TextBox txtNamaPengangkut 
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
            Height          =   315
            Left            =   3000
            TabIndex        =   19
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox txtNoPolisi 
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
            Height          =   315
            Left            =   3000
            TabIndex        =   20
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Saran Pengangkut Darat"
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
            Left            =   120
            TabIndex        =   91
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Nomor Polisi"
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
            Left            =   120
            TabIndex        =   90
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1575
         Left            =   120
         TabIndex        =   87
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14295
         Begin VB.TextBox txtHarga 
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
            Height          =   315
            Left            =   3000
            TabIndex        =   21
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga"
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
            Left            =   120
            TabIndex        =   88
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   3255
         End
      End
   End
   Begin TabDlg.SSTab bb 
      Height          =   3945
      Left            =   360
      TabIndex        =   101
      TabStop         =   0   'False
      Tag             =   "TTFF*/"
      Top             =   5520
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   6959
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   617
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
      TabCaption(0)   =   "Dokumen"
      TabPicture(0)   =   "FrmBC41Detail.frx":0EB2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3(15)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Shape2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3(17)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3(18)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3(19)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3(20)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3(21)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3(22)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDokumen(23)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Shape11"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label3(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "DTPBC40"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "DTPSkep"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DTPFakturPajak"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "DTPKontrak"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "DTPPackingList"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "GridDokumen"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "DTPDokumen"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtDokumen"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtNoDokumen(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtPackingList(2)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtKontrak(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtFakturPajak(4)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtSkep(5)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdSubmitDokumen"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtBC40(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Kemasan"
      TabPicture(1)   =   "FrmBC41Detail.frx":0ECE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape10"
      Tab(1).Control(1)=   "Line1(3)"
      Tab(1).Control(2)=   "Shape5"
      Tab(1).Control(3)=   "Label3(25)"
      Tab(1).Control(4)=   "Label3(24)"
      Tab(1).Control(5)=   "Label3(23)"
      Tab(1).Control(6)=   "lblKemasan(0)"
      Tab(1).Control(7)=   "Label3(27)"
      Tab(1).Control(8)=   "Shape13"
      Tab(1).Control(9)=   "GridKemasan"
      Tab(1).Control(10)=   "txtJumlahKemasan(0)"
      Tab(1).Control(11)=   "txtJenisKemasan"
      Tab(1).Control(12)=   "txtMerkKemasan(2)"
      Tab(1).Control(13)=   "cmSubimtKemasan"
      Tab(1).Control(14)=   "txtSeriKemasan(0)"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Barang"
      TabPicture(2)   =   "FrmBC41Detail.frx":0EEA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblBrgId"
      Tab(2).Control(1)=   "Frame7"
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(3)=   "Frame3"
      Tab(2).Control(4)=   "cmdSubmitBarang"
      Tab(2).Control(5)=   "cmdBrgNew"
      Tab(2).Control(6)=   "Command1(1)"
      Tab(2).Control(7)=   "Command1(2)"
      Tab(2).Control(8)=   "Command1(3)"
      Tab(2).Control(9)=   "Command1(4)"
      Tab(2).Control(10)=   "Command2"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Data Bahan Baku"
      TabPicture(3)   =   "FrmBC41Detail.frx":0F06
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame10"
      Tab(3).Control(1)=   "Frame11"
      Tab(3).Control(2)=   "cmdDeleteBahan"
      Tab(3).Control(3)=   "cmdBahanBakuNew"
      Tab(3).Control(4)=   "cmdSubmitBahanBaku"
      Tab(3).Control(5)=   "Frame12"
      Tab(3).Control(6)=   "txtIdBahan"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Respon"
      TabPicture(4)   =   "FrmBC41Detail.frx":0F22
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "gridRespon"
      Tab(4).Control(1)=   "gridStatus"
      Tab(4).Control(2)=   "Label29"
      Tab(4).Control(3)=   "Label30"
      Tab(4).ControlCount=   4
      Begin VB.TextBox txtIdBahan 
         Appearance      =   0  'Flat
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
         Left            =   -64560
         TabIndex        =   202
         Tag             =   "TTFF*/"
         Top             =   1920
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame Frame12 
         Height          =   1455
         Left            =   -64560
         TabIndex        =   192
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   3960
         Begin VB.TextBox txtHargaPenyerhan 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   2040
            TabIndex        =   198
            Tag             =   "FFTF*/"
            Text            =   "0.00"
            Top             =   675
            Width           =   1815
         End
         Begin VB.TextBox txtJmlSatuan 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   2040
            TabIndex        =   195
            Tag             =   "FFTF*/"
            Text            =   "0.00"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Penyerahan Rp"
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
            Left            =   120
            TabIndex        =   197
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   1875
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Satuan"
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
            Left            =   120
            TabIndex        =   196
            Tag             =   "TTFF*/"
            Top             =   255
            Width           =   1260
         End
         Begin VB.Label lblKPPBC 
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
            Left            =   6960
            TabIndex        =   194
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   60
         End
         Begin VB.Label lblDokAsal 
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
            Left            =   2400
            TabIndex        =   193
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   60
         End
      End
      Begin VB.CommandButton cmdSubmitBahanBaku 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Add"
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
         Left            =   -65880
         Style           =   1  'Graphical
         TabIndex        =   62
         Tag             =   "FFTT*/"
         Top             =   3360
         Width           =   1125
      End
      Begin VB.CommandButton cmdBahanBakuNew 
         BackColor       =   &H0080FFFF&
         Caption         =   "&New"
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
         Left            =   -67080
         Style           =   1  'Graphical
         TabIndex        =   61
         Tag             =   "FFTT*/"
         Top             =   3360
         Width           =   1125
      End
      Begin VB.CommandButton cmdDeleteBahan 
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
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   63
         Tag             =   "FFTT*/"
         Top             =   3360
         Width           =   1125
      End
      Begin VB.Frame Frame11 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   173
         Tag             =   "TFTF*/"
         Top             =   1800
         Width           =   10185
         Begin VB.TextBox txtJenisSatuan 
            Appearance      =   0  'Flat
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
            Left            =   5400
            TabIndex        =   56
            Tag             =   "FFTF*/"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtUkuran 
            Appearance      =   0  'Flat
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
            Left            =   3840
            TabIndex        =   58
            Tag             =   "FFTF*/"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtSPF 
            Appearance      =   0  'Flat
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
            Left            =   8760
            TabIndex        =   60
            Tag             =   "FFTF*/"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtTipe 
            Appearance      =   0  'Flat
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
            Left            =   6120
            TabIndex        =   59
            Tag             =   "FFTF*/"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtMerk 
            Appearance      =   0  'Flat
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
            Left            =   1440
            TabIndex        =   57
            Tag             =   "FFTF*/"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtUraianBrg 
            Appearance      =   0  'Flat
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
            Left            =   1440
            TabIndex        =   55
            Tag             =   "FFTF*/"
            Top             =   675
            Width           =   8655
         End
         Begin VB.TextBox txtKodeBrg 
            Appearance      =   0  'Flat
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
            Left            =   1440
            TabIndex        =   54
            Tag             =   "FFTF*/"
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Satuan"
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
            Index           =   11
            Left            =   4200
            TabIndex        =   201
            Tag             =   "TTFF*/"
            Top             =   285
            Width           =   1080
         End
         Begin VB.Line Line4 
            X1              =   6240
            X2              =   10080
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label lblBahanJenis 
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
            Left            =   6240
            TabIndex        =   200
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   3180
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ukuran"
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
            Left            =   3000
            TabIndex        =   179
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SPF Lain"
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
            Left            =   7680
            TabIndex        =   178
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipe"
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
            Left            =   5520
            TabIndex        =   177
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Merk"
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
            Left            =   120
            TabIndex        =   176
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uraian Barang"
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
            Left            =   120
            TabIndex        =   175
            Tag             =   "TTFF*/"
            Top             =   690
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang"
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
            Left            =   120
            TabIndex        =   174
            Tag             =   "TTFF*/"
            Top             =   285
            Width           =   1110
         End
      End
      Begin VB.TextBox txtBC40 
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
         Height          =   315
         Index           =   0
         Left            =   10200
         TabIndex        =   161
         Tag             =   "FFTF*/"
         Top             =   2318
         Width           =   1935
      End
      Begin VB.TextBox txtSeriKemasan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   -63360
         TabIndex        =   133
         Tag             =   "FFTF*/"
         Top             =   728
         Width           =   1095
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
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   44
         Tag             =   "FFTT*/"
         Top             =   3488
         Width           =   1125
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
         Left            =   -68550
         Style           =   1  'Graphical
         TabIndex        =   48
         Tag             =   "FFTT*/"
         Top             =   3488
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
         Left            =   -69780
         Style           =   1  'Graphical
         TabIndex        =   47
         Tag             =   "FFTT*/"
         Top             =   3488
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
         Left            =   -71010
         Style           =   1  'Graphical
         TabIndex        =   46
         Tag             =   "FFTT*/"
         Top             =   3488
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
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   45
         Tag             =   "FFTT*/"
         Top             =   3488
         Width           =   1140
      End
      Begin VB.CommandButton cmdBrgNew 
         BackColor       =   &H0080FFFF&
         Caption         =   "&New"
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
         Left            =   -66720
         Style           =   1  'Graphical
         TabIndex        =   42
         Tag             =   "FFTT*/"
         Top             =   3488
         Width           =   1125
      End
      Begin VB.CommandButton cmdSubmitBarang 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Add"
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
         Left            =   -65520
         Style           =   1  'Graphical
         TabIndex        =   43
         Tag             =   "FFTT*/"
         Top             =   3488
         Width           =   1125
      End
      Begin VB.CommandButton cmSubimtKemasan 
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
         Left            =   -63000
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "FFTT*/"
         Top             =   3383
         Width           =   1125
      End
      Begin VB.CommandButton cmdSubmitDokumen 
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
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "FFTT*/"
         Top             =   3383
         Width           =   1125
      End
      Begin VB.TextBox txtMerkKemasan 
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
         Height          =   315
         Index           =   2
         Left            =   -67920
         TabIndex        =   28
         Tag             =   "FFTF*/"
         Top             =   3368
         Width           =   2775
      End
      Begin VB.TextBox txtJenisKemasan 
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
         Height          =   315
         Left            =   -74640
         TabIndex        =   26
         Tag             =   "FFTF*/"
         Top             =   3368
         Width           =   1575
      End
      Begin VB.TextBox txtJumlahKemasan 
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
         Height          =   315
         Index           =   0
         Left            =   -70200
         TabIndex        =   27
         Tag             =   "FFTF*/"
         Top             =   3368
         Width           =   2175
      End
      Begin VB.TextBox txtSkep 
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
         Height          =   315
         Index           =   5
         Left            =   10200
         TabIndex        =   132
         Tag             =   "FFTF*/"
         Top             =   1943
         Width           =   1935
      End
      Begin VB.TextBox txtFakturPajak 
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
         Height          =   315
         Index           =   4
         Left            =   10200
         TabIndex        =   131
         Tag             =   "FFTF*/"
         Top             =   1568
         Width           =   1935
      End
      Begin VB.TextBox txtKontrak 
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
         Height          =   315
         Index           =   3
         Left            =   10200
         TabIndex        =   130
         Tag             =   "FFTF*/"
         Top             =   1208
         Width           =   1935
      End
      Begin VB.TextBox txtPackingList 
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
         Height          =   315
         Index           =   2
         Left            =   10200
         TabIndex        =   129
         Tag             =   "FFTF*/"
         Top             =   848
         Width           =   1935
      End
      Begin VB.TextBox txtNoDokumen 
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
         Height          =   315
         Index           =   1
         Left            =   5640
         TabIndex        =   23
         Tag             =   "FFTF*/"
         Top             =   3368
         Width           =   2175
      End
      Begin VB.TextBox txtDokumen 
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
         Height          =   315
         Left            =   480
         TabIndex        =   22
         Tag             =   "FFTF*/"
         Top             =   3368
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Height          =   1695
         Left            =   -74760
         TabIndex        =   118
         Tag             =   "TFTF*/"
         Top             =   503
         Width           =   10440
         Begin VB.TextBox txtId 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1440
            TabIndex        =   120
            Tag             =   "FFTF*/"
            Top             =   315
            Width           =   735
         End
         Begin VB.TextBox txtIdEnd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   2760
            TabIndex        =   119
            Tag             =   "FFTF*/"
            Top             =   315
            Width           =   735
         End
         Begin VB.TextBox txtBrgKode 
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
            Height          =   315
            Left            =   1440
            TabIndex        =   30
            Tag             =   "FFTF*/"
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtBrgUraian 
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
            Height          =   315
            Left            =   1440
            TabIndex        =   31
            Tag             =   "FFTF*/"
            Top             =   1155
            Width           =   5895
         End
         Begin VB.TextBox txtBrgMerk 
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
            Height          =   315
            Left            =   4440
            TabIndex        =   32
            Tag             =   "FFTF*/"
            Top             =   315
            Width           =   2895
         End
         Begin VB.TextBox txtBrgType 
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
            Height          =   315
            Left            =   4440
            TabIndex        =   33
            Tag             =   "FFTF*/"
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txtBrgSpf 
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
            Height          =   315
            Left            =   8520
            TabIndex        =   35
            Tag             =   "FFTF*/"
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtBrgUkuran 
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
            Height          =   315
            Left            =   8520
            TabIndex        =   34
            Tag             =   "FFTF*/"
            Top             =   315
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Barang"
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
            Index           =   30
            Left            =   120
            TabIndex        =   128
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "dari"
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
            Index           =   31
            Left            =   2280
            TabIndex        =   127
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang"
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
            Index           =   32
            Left            =   120
            TabIndex        =   126
            Tag             =   "TTFF*/"
            Top             =   765
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uraian Barang"
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
            Index           =   33
            Left            =   120
            TabIndex        =   125
            Tag             =   "TTFF*/"
            Top             =   1170
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Merk"
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
            Index           =   39
            Left            =   3840
            TabIndex        =   124
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipe"
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
            Index           =   40
            Left            =   3840
            TabIndex        =   123
            Tag             =   "TTFF*/"
            Top             =   765
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SPF Lain"
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
            Index           =   41
            Left            =   7680
            TabIndex        =   122
            Tag             =   "TTFF*/"
            Top             =   765
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ukuran"
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
            Index           =   26
            Left            =   7680
            TabIndex        =   121
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   111
         Tag             =   "TFTF*/"
         Top             =   2183
         Width           =   10440
         Begin VB.TextBox txtBahanBaku 
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
            Height          =   315
            Left            =   9480
            TabIndex        =   41
            Tag             =   "FFTF*/"
            Text            =   "0.00"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtBrgJenisSatuan 
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
            Height          =   315
            Left            =   1440
            TabIndex        =   37
            Tag             =   "FFTF*/"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtBrgJumlahSatuan 
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
            Height          =   315
            Left            =   1440
            TabIndex        =   36
            Tag             =   "FFTF*/"
            Text            =   "0.00"
            Top             =   285
            Width           =   1935
         End
         Begin VB.TextBox txtBrgNetto 
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
            Height          =   315
            Left            =   4800
            TabIndex        =   38
            Tag             =   "FFTF*/"
            Text            =   "0.00"
            Top             =   285
            Width           =   1935
         End
         Begin VB.TextBox txtBrgVolume 
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
            Height          =   315
            Left            =   4800
            TabIndex        =   39
            Tag             =   "FFTF*/"
            Text            =   "0.00"
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtBrgHarga 
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
            Height          =   315
            Left            =   8880
            TabIndex        =   40
            Tag             =   "FFTF*/"
            Text            =   "0.00"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Bahan Baku"
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
            Left            =   6840
            TabIndex        =   199
            Tag             =   "TTFF*/"
            Top             =   735
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Satuan"
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
            Index           =   34
            Left            =   120
            TabIndex        =   117
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Satuan"
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
            Index           =   35
            Left            =   120
            TabIndex        =   116
            Tag             =   "TTFF*/"
            Top             =   735
            Width           =   1080
         End
         Begin VB.Line Line3 
            X1              =   2280
            X2              =   3360
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Netto (Kgm)"
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
            Index           =   36
            Left            =   3600
            TabIndex        =   115
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Volume(m3)"
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
            Index           =   37
            Left            =   3600
            TabIndex        =   114
            Tag             =   "TTFF*/"
            Top             =   735
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Penyerahan Rp"
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
            Index           =   38
            Left            =   6840
            TabIndex        =   113
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1875
         End
         Begin VB.Label lblBrgJenis 
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
            Index           =   27
            Left            =   2280
            TabIndex        =   112
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   1140
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1695
         Left            =   -64200
         TabIndex        =   102
         Tag             =   "TFTF*/"
         Top             =   503
         Width           =   3480
         Begin VB.TextBox txtJumahBrg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   2505
            TabIndex        =   106
            Tag             =   "FFTF*/"
            Text            =   "0.00"
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtNetto 
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
            Height          =   315
            Left            =   1785
            TabIndex        =   105
            Tag             =   "FFTF*/"
            Text            =   "0.00"
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtBruto 
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
            Height          =   315
            Left            =   1785
            TabIndex        =   104
            Tag             =   "FFTF*/"
            Text            =   "0.00"
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtVolume 
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
            Height          =   315
            Left            =   1800
            TabIndex        =   103
            Tag             =   "FFTF*/"
            Text            =   "0.00"
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Barang"
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
            Index           =   52
            Left            =   240
            TabIndex        =   110
            Tag             =   "TTFF*/"
            Top             =   1320
            Width           =   1275
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Berat Bersih (Kg)"
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
            Index           =   51
            Left            =   240
            TabIndex        =   109
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1500
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Berat Kotor (Kg)"
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
            Index           =   50
            Left            =   240
            TabIndex        =   108
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Volume (m3)"
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
            Index           =   49
            Left            =   240
            TabIndex        =   107
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1125
         End
      End
      Begin MSComCtl2.DTPicker DTPDokumen 
         Height          =   345
         Left            =   8040
         TabIndex        =   24
         Tag             =   "TTFF*/"
         Top             =   3368
         Width           =   1530
         _ExtentX        =   2699
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   293666819
         CurrentDate     =   37798
      End
      Begin VSFlex8Ctl.VSFlexGrid grid 
         Height          =   2535
         Left            =   -74910
         TabIndex        =   134
         Tag             =   "TTFF*/"
         Top             =   390
         Width           =   14925
         _cx             =   26326
         _cy             =   4471
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
         FixedRows       =   1
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
         Begin MSComCtl2.DTPicker DelDate 
            Height          =   315
            Left            =   3720
            TabIndex        =   135
            Tag             =   "TTFF*/"
            Top             =   480
            Visible         =   0   'False
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
            Format          =   293666819
            CurrentDate     =   37798
         End
         Begin MSForms.ComboBox cbocurr 
            Height          =   285
            Left            =   5640
            TabIndex        =   137
            TabStop         =   0   'False
            Tag             =   "TTFF*/"
            Top             =   480
            Visible         =   0   'False
            Width           =   855
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "1508;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboprice 
            Height          =   285
            Left            =   4920
            TabIndex        =   136
            TabStop         =   0   'False
            Tag             =   "TTFF*/"
            Top             =   1140
            Visible         =   0   'False
            Width           =   2055
            VariousPropertyBits=   746604571
            MaxLength       =   19
            DisplayStyle    =   3
            Size            =   "3625;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid GridDokumen 
         Height          =   1935
         Left            =   360
         TabIndex        =   138
         TabStop         =   0   'False
         Tag             =   "TTTT*/"
         Top             =   743
         Width           =   7605
         _cx             =   13414
         _cy             =   3413
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
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
         Begin VB.TextBox Text14 
            Height          =   315
            Index           =   0
            Left            =   10080
            TabIndex        =   139
            Tag             =   "FFTF*/"
            Top             =   0
            Width           =   2175
         End
      End
      Begin MSComCtl2.DTPicker DTPPackingList 
         Height          =   345
         Left            =   12240
         TabIndex        =   140
         Tag             =   "TTFF*/"
         Top             =   848
         Width           =   1530
         _ExtentX        =   2699
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   293666819
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DTPKontrak 
         Height          =   345
         Left            =   12240
         TabIndex        =   141
         Tag             =   "TTFF*/"
         Top             =   1208
         Width           =   1530
         _ExtentX        =   2699
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   293666819
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DTPFakturPajak 
         Height          =   345
         Left            =   12240
         TabIndex        =   142
         Tag             =   "TTFF*/"
         Top             =   1568
         Width           =   1530
         _ExtentX        =   2699
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   293666819
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DTPSkep 
         Height          =   345
         Left            =   12240
         TabIndex        =   143
         Tag             =   "TTFF*/"
         Top             =   1928
         Width           =   1530
         _ExtentX        =   2699
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   293666819
         CurrentDate     =   37798
      End
      Begin VSFlex8Ctl.VSFlexGrid GridKemasan 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   144
         TabStop         =   0   'False
         Tag             =   "TTTT*/"
         Top             =   623
         Width           =   8445
         _cx             =   14896
         _cy             =   3836
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
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
      Begin MSComCtl2.DTPicker DTPBC40 
         Height          =   345
         Left            =   12240
         TabIndex        =   162
         Tag             =   "TTFF*/"
         Top             =   2303
         Width           =   1530
         _ExtentX        =   2699
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   293666819
         CurrentDate     =   37798
      End
      Begin VSFlex8Ctl.VSFlexGrid gridRespon 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   164
         TabStop         =   0   'False
         Tag             =   "TTFT*/"
         Top             =   840
         Width           =   6525
         _cx             =   11509
         _cy             =   5106
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
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
      Begin VSFlex8Ctl.VSFlexGrid gridStatus 
         Height          =   2895
         Left            =   -67440
         TabIndex        =   165
         TabStop         =   0   'False
         Tag             =   "TTTT*/"
         Top             =   840
         Width           =   6525
         _cx             =   11509
         _cy             =   5106
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
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
      Begin VB.Frame Frame10 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   168
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   10200
         Begin VB.TextBox txtBahanId 
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
            Height          =   315
            Left            =   5400
            TabIndex        =   67
            Tag             =   "FFTF*/"
            Top             =   195
            Width           =   735
         End
         Begin VB.TextBox txtBahanEnd 
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
            Height          =   315
            Left            =   6720
            TabIndex        =   189
            Tag             =   "FFTF*/"
            Top             =   195
            Width           =   735
         End
         Begin VB.TextBox txtNoDok 
            Appearance      =   0  'Flat
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
            Left            =   1440
            TabIndex        =   50
            Tag             =   "FFTF*/"
            Top             =   1035
            Width           =   855
         End
         Begin VB.TextBox txtDokAsal 
            Appearance      =   0  'Flat
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
            Left            =   1440
            TabIndex        =   204
            Tag             =   "FFTF*/"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtKPPBC 
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
            Height          =   315
            Left            =   5400
            TabIndex        =   51
            Tag             =   "FFTF*/"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtNoAju 
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
            Height          =   315
            Left            =   5400
            TabIndex        =   52
            Tag             =   "FFTF*/"
            Top             =   1035
            Width           =   3135
         End
         Begin VB.TextBox txtUrut 
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
            Height          =   315
            Left            =   9360
            TabIndex        =   53
            Tag             =   "FFTF*/"
            Top             =   1035
            Width           =   735
         End
         Begin VB.TextBox txtBrgEnd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   2760
            TabIndex        =   170
            Tag             =   "FFTF*/"
            Top             =   195
            Width           =   735
         End
         Begin VB.TextBox txtBrgId 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1440
            TabIndex        =   169
            Tag             =   "FFTF*/"
            Top             =   195
            Width           =   735
         End
         Begin MSComCtl2.DTPicker DTPDok 
            Height          =   315
            Left            =   2640
            TabIndex        =   180
            Tag             =   "TTFF*/"
            Top             =   1035
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   293732355
            CurrentDate     =   37798
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bahan Baku Ke"
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
            Index           =   61
            Left            =   3960
            TabIndex        =   191
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "dari"
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
            Index           =   62
            Left            =   6240
            TabIndex        =   190
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   330
         End
         Begin VB.Label lblDokAsal 
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
            Index           =   65
            Left            =   2280
            TabIndex        =   188
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   60
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dok Asal"
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
            Index           =   63
            Left            =   120
            TabIndex        =   187
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No / Tgl Dok"
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
            Index           =   64
            Left            =   120
            TabIndex        =   186
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Line Line9 
            X1              =   2280
            X2              =   4080
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Index           =   66
            Left            =   2400
            TabIndex        =   185
            Tag             =   "TTFF*/"
            Top             =   1125
            Width           =   75
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "KPPBC Dok"
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
            Index           =   67
            Left            =   4320
            TabIndex        =   184
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   990
         End
         Begin VB.Line Line10 
            X1              =   6600
            X2              =   9960
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Aju"
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
            Index           =   69
            Left            =   4320
            TabIndex        =   182
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Urut Ke"
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
            Index           =   70
            Left            =   8640
            TabIndex        =   181
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "dari"
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
            Left            =   2280
            TabIndex        =   172
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Barang"
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
            Left            =   120
            TabIndex        =   171
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblKPPBC 
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
            Index           =   68
            Left            =   6600
            TabIndex        =   183
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   60
         End
      End
      Begin VB.Label Label29 
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
         Height          =   255
         Left            =   -67440
         TabIndex        =   167
         Tag             =   "TTFF*/"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Respon"
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
         Left            =   -74760
         TabIndex        =   166
         Tag             =   "TTFF*/"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BC 4.0 Asal"
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
         Left            =   8400
         TabIndex        =   163
         Tag             =   "TTFF*/"
         Top             =   2303
         Width           =   1005
      End
      Begin VB.Label lblBrgId 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -67140
         TabIndex        =   160
         Tag             =   "TTFF*/"
         Top             =   2768
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H00808080&
         Height          =   495
         Left            =   -66120
         Tag             =   "TTTF*/"
         Top             =   623
         Width           =   4005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kemasan"
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
         Index           =   27
         Left            =   -66000
         TabIndex        =   159
         Tag             =   "TTFF*/"
         Top             =   728
         Width           =   795
      End
      Begin VB.Label lblIdKemasan 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -74850
         TabIndex        =   158
         Tag             =   "TTFF*/"
         Top             =   360
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblIdDokumen 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -68700
         TabIndex        =   157
         Tag             =   "TTFF*/"
         Top             =   2400
         Width           =   75
      End
      Begin VB.Label lblKemasan 
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
         Left            =   -72960
         TabIndex        =   156
         Tag             =   "TTFF*/"
         Top             =   3383
         Width           =   2460
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00808080&
         Height          =   2055
         Left            =   8160
         Tag             =   "TTTF*/"
         Top             =   743
         Width           =   5835
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Merk"
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
         Index           =   23
         Left            =   -67920
         TabIndex        =   155
         Tag             =   "TTFF*/"
         Top             =   3023
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
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
         Index           =   24
         Left            =   -70200
         TabIndex        =   154
         Tag             =   "TTFF*/"
         Top             =   3023
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
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
         Left            =   -74640
         TabIndex        =   153
         Tag             =   "TTFF*/"
         Top             =   3023
         Width           =   420
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00808080&
         Height          =   495
         Left            =   -74760
         Tag             =   "TTTF*/"
         Top             =   3263
         Width           =   10725
      End
      Begin VB.Line Line2 
         X1              =   -72960
         X2              =   -70440
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00A6D2FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   375
         Index           =   1
         Left            =   -74760
         Tag             =   "TTTF*/"
         Top             =   3720
         Width           =   10695
      End
      Begin VB.Label lblDokumen 
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
         Index           =   23
         Left            =   2760
         TabIndex        =   152
         Tag             =   "TTFF*/"
         Top             =   3368
         Width           =   60
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SKEP"
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
         Index           =   22
         Left            =   8400
         TabIndex        =   151
         Tag             =   "TTFF*/"
         Top             =   1928
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faktur Pajak"
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
         Index           =   21
         Left            =   8400
         TabIndex        =   150
         Tag             =   "TTFF*/"
         Top             =   1583
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kontrak"
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
         Index           =   20
         Left            =   8400
         TabIndex        =   149
         Tag             =   "TTFF*/"
         Top             =   1208
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing List"
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
         Index           =   19
         Left            =   8400
         TabIndex        =   148
         Tag             =   "TTFF*/"
         Top             =   848
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Index           =   18
         Left            =   8040
         TabIndex        =   147
         Tag             =   "TTFF*/"
         Top             =   3008
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor"
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
         Index           =   17
         Left            =   5640
         TabIndex        =   146
         Tag             =   "TTFF*/"
         Top             =   3008
         Width           =   570
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   -72360
         X2              =   -69840
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   495
         Left            =   360
         Tag             =   "TTTF*/"
         Top             =   3263
         Width           =   10815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dokumen"
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
         Index           =   15
         Left            =   480
         TabIndex        =   145
         Tag             =   "TTFF*/"
         Top             =   3008
         Width           =   825
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2760
         X2              =   5280
         Y1              =   3623
         Y2              =   3623
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   -72960
         X2              =   -70440
         Y1              =   3623
         Y2              =   3623
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00A6D2FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   360
         Tag             =   "TTTF*/"
         Top             =   2903
         Width           =   10815
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00A6D2FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   -74760
         Tag             =   "TTTF*/"
         Top             =   2888
         Width           =   10725
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BC 41 Detail"
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
      TabIndex        =   73
      Tag             =   "TTTF*/"
      Top             =   120
      Width           =   14610
   End
   Begin VB.Label lblNoId 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   72
      Tag             =   "TTFF*/"
      Top             =   0
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblNPWPPengirim 
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
      Index           =   1
      Left            =   6660
      TabIndex        =   71
      Tag             =   "TTFF*/"
      Top             =   3060
      Width           =   60
   End
End
Attribute VB_Name = "FrmBC41Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tampung As String
Public sql As String
Public rs_barang As New ADODB.Recordset
Public rs_detail As New ADODB.Recordset
Public RSiD As New ADODB.Recordset
Public rsDokumen As ADODB.Recordset
Public rsKemasan As ADODB.Recordset
Public rsBarang As ADODB.Recordset
Public rsBahan As ADODB.Recordset
Dim RS As Recordset
Dim X As Integer, xx As Boolean, hapus As Boolean
Dim k_pertama As Boolean, Syncronize As Boolean

Dim bteColSelect As Byte, bteColNo As Byte, bteColjumlah As Byte, bteColKode As Byte, bteColUraian As Byte
Dim bteColMerkKemasan As Byte, bteColId As Byte

'---------------------------------------
Const colSelect As Integer = 0
Const colId As Integer = 1
Const colSeri As Integer = 2
Const colKodeDokumen As Integer = 3
Const colJenisDokumen As Integer = 4
Const colNomor As Integer = 5
Const colTanggal As Integer = 6
Const colcount As Integer = 7
'---------------------------------------

'---------------------------------------
Const colKodeStatus As Integer = 0
Const colUraianStatus As Integer = 1
Const colWaktuStatus As Integer = 2
Const colCountStatus As Integer = 3
'---------------------------------------

'---------------------------------------
Const colKodeRespon As Integer = 0
Const colUraianRespon As Integer = 1
Const colWaktuRespon As Integer = 2
Const colCountRespon As Integer = 3
'---------------------------------------

Private Sub clear()
    
    dtpTanggal = Format(Now, "dd MMM yyyy")
    DTPDokumen = Format(Now, "dd MMM yyyy")
    DTPFakturPajak = Format(Now, "dd MMM yyyy")
    DTPPackingList = Format(Now, "dd MMM yyyy")
    DTPSkep = Format(Now, "dd MMM yyyy")
    DTPKontrak = Format(Now, "dd MMM yyyy")
    dtpTglDaftar = Format(Now, "dd MMM yyyy")
     
    LblErrMsg.Caption = ""
    
    up_IsiGridDokumen
    
    up_IsiGridKemasan
    
    koneksi

End Sub

Private Sub koneksi()
Dim sql As String
Dim RS As New Recordset

    If txtNoPengajuan = "______-______-________-______" Then
      sql = "Select * from Bea_Cukai_TPB_Header WHERE NO_PENGAJUAN='" & Replace(FrmBC41List.txtTampung, "-", "") & "'"
            If RS.State <> adStateClosed Then RS.Close
            RS.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        
            If RS.EOF = False Then
                tampung = Trim(RS("Id"))
            End If
    Else
      sql = "Select * from Bea_Cukai_TPB_Header WHERE NO_PENGAJUAN='" & Replace(txtNoPengajuan.Text, "-", "") & "'"
            If RS.State <> adStateClosed Then RS.Close
            RS.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        
            If RS.EOF = False Then
                tampung = Trim(RS("Id"))
            End If
    End If
    
    sql = " SELECT Id, KODE_BARANG, URAIAN, MERK, TIPE, UKURAN, SPESIFIKASI_LAIN, JUMLAH_SATUAN, KODE_SATUAN," & vbCrLf & _
          " NETTO, VOLUME, HARGA_PENYERAHAN, SERI_BARANG, ID_HEADER FROM Bea_Cukai_TPB_Barang Where ID_HEADER='" & tampung & "'"
    If rs_barang.State <> adStateClosed Then rs_barang.Close
    rs_barang.Open sql, Db, adOpenKeyset, adLockOptimistic

End Sub

Private Sub data_tampil()
    If rs_barang.EOF = False Then
        lblBrgId.Caption = Trim(rs_barang("Id"))
        txtID.Text = rs_barang.AbsolutePosition
        txtIdEnd.Text = rs_barang.RecordCount
        'cboBrgStatus.Text = Trim(rs_barang("KODE_STATUS"))
        txtBrgKode.Text = Trim(rs_barang("KODE_BARANG"))
        txtBrgUraian.Text = Trim(rs_barang("Uraian"))
        txtBrgMerk.Text = Trim(rs_barang("Merk"))
        txtBrgType.Text = Trim(rs_barang("Tipe"))
        txtBrgUkuran.Text = Trim(rs_barang("Ukuran"))
        txtBrgSpf.Text = Trim(rs_barang("Spesifikasi_Lain"))
        txtBrgJumlahSatuan.Text = Format(rs_barang("Jumlah_Satuan"), gs_formatQty)
        txtBrgJenisSatuan.Text = Format(rs_barang("Kode_Satuan"), gs_formatQty)
        txtBrgNetto.Text = Format(rs_barang("Netto"), gs_formatQty)
        txtBrgVolume.Text = Format(rs_barang("Volume"), gs_formatQty)
        txtBrgHarga.Text = Format(rs_barang("Harga_Penyerahan"), gs_formatAmountIDR)
        txtJumahBrg.Text = rs_barang.RecordCount
        txtBrgId.Text = txtID
        txtBrgEnd.Text = txtIdEnd
    End If
End Sub

Private Sub up_FillComboTujuan()
Dim sql As String
Dim RS As New Recordset

    sql = "Select * From Bea_Cukai_Referensi_Tujuan_Pengiriman where Kode_Dokumen='" & 41 & "' "
    Set RS = Db.Execute(sql)

    With cboTujuan
        .clear
        .columnCount = 1
        .ColumnWidths = "50pt;300pt"
        .ListWidth = 350
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS(2)) & " - " & IIf(IsNull(RS(3)), "", Trim(RS(3)))
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = -1
    End With
End Sub

Private Sub up_FillComboJenisTPB()
Dim sql As String
Dim RS As New Recordset

    sql = "Select * From Bea_Cukai_Jenis_TPB"
    Set RS = Db.Execute(sql)

    With cboJenisTPB
        .clear
        .columnCount = 1
        .ColumnWidths = "50pt;300pt"
        .ListWidth = 350
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS(1)) & " - " & IIf(IsNull(RS(2)), "", Trim(RS(2)))
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = -1
    End With
End Sub

Private Sub up_FillComboNPWP()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41NPWP_Sel"
    
    Set RS = cmd.Execute

    With cboNPWPPengusaha
        .clear
        .columnCount = 1
        .ColumnWidths = "30pt;80pt"
        .ListWidth = 110
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS(0)) & " - " & IIf(IsNull(RS(1)), "", Trim(RS(1)))
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = 0
    End With
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41NPWP_Sel"
    
    Set RS = cmd.Execute

    With cboNPWPPengirim
        .clear
        .columnCount = 1
        .ColumnWidths = "30pt;80pt"
        .ListWidth = 110
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS(0)) & " - " & IIf(IsNull(RS(1)), "", Trim(RS(1)))
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = 0
    End With
End Sub

Private Sub up_LoadKantorPabean(pKode As String)
Dim sql As String
Dim RS As New Recordset

        sql = "Select * From Bea_Cukai_Kantor_Pabean Where Kode_Kantor = '" & pKode & "'"
        Set RS = Db.Execute(sql)
            
        If Not RS.EOF Then
            lblKantorPabean.Caption = RS.Fields("Nama_Kantor")
        Else
            lblKantorPabean.Caption = ""
        End If

End Sub

Private Sub up_LoadKPPBC(pKode As String)
Dim sql As String
Dim RS As New Recordset

        sql = "Select * From Bea_Cukai_Kantor_Pabean Where Kode_Kantor = '" & pKode & "'"
        Set RS = Db.Execute(sql)
            
        If Not RS.EOF Then
            lblKPPBC(68).Caption = RS.Fields("Nama_Kantor")
        Else
            lblKPPBC(68).Caption = ""
        End If

End Sub

Private Sub Cmd_SubMenu_Click()
    Unload Me
    FrmBC41List.Show
End Sub

Private Sub cmdBahan_Click()

End Sub

Private Sub cmdBahanBakuNew_Click()
Dim sql As String
Dim RS As New ADODB.Recordset
Dim Aa As String

Aa = MsgBox("Do you want to get data from the last document?", vbYesNo + vbQuestion + vbDefaultButton2, "Question")
    If Aa = vbYes Then
    
        Me.MousePointer = vbHourglass
            
        Clear_Bahan
        
        FrmBC41LoadBarang.Show
        
        Me.MousePointer = vbDefault
    Else
        
    End If

End Sub

Private Sub Clear_Bahan()
    txtDokAsal.Enabled = True
    txtNoDok.Enabled = True
    'DTPDok.Enabled = True
    txtKPPBC.Enabled = True
    txtUrut.Enabled = True
    txtKodeBrg.Enabled = True
    txtUraianBrg.Enabled = True
    txtJmlSatuan.Enabled = True
    txtJenisSatuan.Enabled = True
    txtTipe.Enabled = True
    txtMerk.Enabled = True
    txtUkuran.Enabled = True
    txtSPF.Enabled = True
    txtHargaPenyerhan.Enabled = True
    txtNoAju.Enabled = True
    txtIdBahan.Text = ""
    txtDokAsal.Text = ""
    txtNoDok.Text = ""
    'DTPDok.Value = ""
    txtKPPBC.Text = ""
    txtUrut.Text = ""
    txtKodeBrg.Text = ""
    txtUraianBrg.Text = ""
    txtJmlSatuan.Text = Format(0, gs_formatQty)
    txtJenisSatuan.Text = ""
    txtTipe.Text = ""
    txtMerk.Text = ""
    txtUkuran.Text = ""
    txtSPF.Text = ""
    txtHargaPenyerhan.Text = Format(0, gs_formatAmount)
    txtNoAju.Text = ""
    lblKPPBC(68).Caption = ""
End Sub

Private Sub BeforeDeleteBahan()
    txtDokAsal.Text = ""
    txtNoDok.Text = ""
    'DTPDok.Value = ""
    txtKPPBC.Text = ""
    txtUrut.Text = ""
    txtKodeBrg.Text = ""
    txtUraianBrg.Text = ""
    txtJmlSatuan.Text = Format(0, gs_formatQty)
    txtJenisSatuan.Text = ""
    txtTipe.Text = ""
    txtMerk.Text = ""
    txtUkuran.Text = ""
    txtSPF.Text = ""
    txtHargaPenyerhan.Text = Format(0, gs_formatAmount)
    txtNoAju.Text = ""
    lblKPPBC(68).Caption = ""
End Sub


Private Sub cmdBrgNew_Click()
    'up_fillCboStatus
    lblBrgId.Caption = ""
    txtBrgKode.Text = ""
    txtBrgUraian.Text = ""
    txtBrgMerk.Text = ""
    txtBrgType.Text = ""
    txtBrgSpf.Text = ""
    txtBrgUkuran.Text = ""
    txtBrgJumlahSatuan.Text = Format(0, gs_formatQty)
    txtBrgJenisSatuan.Text = ""
    txtBrgNetto.Text = Format(0, gs_formatQty)
    txtBrgVolume = Format(0, gs_formatQty)
    txtBrgHarga = Format(0, gs_formatAmountIDR)
    txtBrgId = txtID
    txtBrgEnd = txtIdEnd
    txtBrgKode.SetFocus
    
    LblErrMsg.Caption = ""
End Sub

Private Sub cmdDeleteBahan_Click()
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim tanya

    If IsEmpty(tanya) Then tanya = MsgBox("Do you really want to delete this data ?", vbQuestion & vbYesNo, "Confirmation")
        If tanya = vbYes Then
            
            Set cmd = New ADODB.Command
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 0
            cmd.ActiveConnection = Db
            cmd.CommandText = "sp_BC41Bahan_Del"
            
            cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 100, txtIdBahan)
            
            Set RS = cmd.Execute
            
            BeforeDeleteBahan
            
            up_DetailBahan
            
            tampil_detail
                        
            LblErrMsg = DisplayMsg(1201)
            
            Exit Sub
            
        Else
            Exit Sub
        End If
End Sub

Private Sub CmdSubmit_Click()
    Dim strSQL As String
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim prm1 As ADODB.Parameter
    Dim prm2 As ADODB.Parameter
    Dim prm3 As ADODB.Parameter
    Dim prm4 As ADODB.Parameter
    
    Set RS = New ADODB.Recordset
    
    LblErrMsg.Caption = ""
    
    Me.MousePointer = vbHourglass
    Insert_Update
    up_Form
    Me.MousePointer = vbDefault
    LblErrMsg = DisplayMsg(1000)
End Sub

Private Sub Insert_Update()
Dim strSQL As String
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim prm3 As ADODB.Parameter
Dim prm4 As ADODB.Parameter
Dim prm5 As ADODB.Parameter


    Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC41Header_InsertUpdate"
        
        cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 100, lblNoId.Caption)
        cmd.Parameters.append cmd.CreateParameter("Nomor_Aju", adVarChar, adParamInput, 100, Replace(txtNoPengajuan.Text, "-", ""))
        cmd.Parameters.append cmd.CreateParameter("Nomor_Daftar", adVarChar, adParamInput, 50, txtNoDaftar)
        cmd.Parameters.append cmd.CreateParameter("Tanggal_Daftar", adDBTime, adParamInput, , dtpTanggal.Value)
        cmd.Parameters.append cmd.CreateParameter("Tanggal_Aju", adDBTime, adParamInput, , dtpTanggal.Value)
        cmd.Parameters.append cmd.CreateParameter("Kode_Kantor", adVarChar, adParamInput, 50, txtKantorPabean.Text)
        'cmd.Parameters.append cmd.CreateParameter("Kode_Status", adVarChar, adParamInput, 50, cboBrgStatus.Text)
        cmd.Parameters.append cmd.CreateParameter("Kode_Dokumen_Pabean", adVarChar, adParamInput, 50, Left(txtNoPengajuan.Text, 6))
        cmd.Parameters.append cmd.CreateParameter("Id_Modul", adVarChar, adParamInput, 50, Mid$(txtNoPengajuan.Text, 9, 5))
        cmd.Parameters.append cmd.CreateParameter("Kode_Jenis_TPB", adVarChar, adParamInput, 50, Left(cboJenisTPB.Text, 1))
        cmd.Parameters.append cmd.CreateParameter("Kode_Tujuan_Pengiriman", adVarChar, adParamInput, 50, Left(cboTujuan.Text, 1))
        cmd.Parameters.append cmd.CreateParameter("Kode_Id_Pengusaha", adVarChar, adParamInput, 50, Left(cboNPWPPengusaha.Text, 1))
        cmd.Parameters.append cmd.CreateParameter("Id_Pengusaha", adVarChar, adParamInput, 50, txtNPWPPengusaha.Text)
        cmd.Parameters.append cmd.CreateParameter("Nama_Pengusaha", adVarChar, adParamInput, 100, txtNamaPengusaha.Text)
        cmd.Parameters.append cmd.CreateParameter("Alamat_Pengusaha", adVarChar, adParamInput, 200, txtAlamatPengusaha.Text)
        cmd.Parameters.append cmd.CreateParameter("Nomor_Ijin_TPB", adVarChar, adParamInput, 50, txtNoIzinPengusaha.Text)
        cmd.Parameters.append cmd.CreateParameter("Kode_Id_Penerima_Barang", adVarChar, adParamInput, 50, Left(cboNPWPPengirim.Text, 1))
        cmd.Parameters.append cmd.CreateParameter("Id_Penerima_Barang", adVarChar, adParamInput, 50, txtNPWPPengirim.Text)
        cmd.Parameters.append cmd.CreateParameter("Nama_Penerima_Barang", adVarChar, adParamInput, 50, txtNamaPengirim.Text)
        cmd.Parameters.append cmd.CreateParameter("Alamat_Penerimaa_Barang", adVarChar, adParamInput, 200, txtAlamatPengirim.Text)
        cmd.Parameters.append cmd.CreateParameter("Nama_Pengangkut", adVarChar, adParamInput, 100, txtNamaPengangkut.Text)
        cmd.Parameters.append cmd.CreateParameter("Nomor_Polisi", adVarChar, adParamInput, 50, txtNoPolisi.Text) 'Harus diganti dengan nopol
        Set prm = cmd.CreateParameter("Harga_Penyerahan", adDecimal, adParamInput)
        prm.Precision = 38
        prm.NumericScale = 2
        prm.Value = CDec(IIf(txtHarga = "", 0, txtHarga))
        cmd.Parameters.append prm
        Set prm1 = cmd.CreateParameter("VOLUME", adDecimal, adParamInput)
        prm1.Precision = 38
        prm1.NumericScale = 2
        prm1.Value = CDec(IIf(txtVolume = "", 0, txtVolume))
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("BRUTO", adDecimal, adParamInput)
        prm2.Precision = 38
        prm2.NumericScale = 2
        prm2.Value = CDec(IIf(txtBruto = "", 0, txtBruto))
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("Netto", adDecimal, adParamInput)
        prm3.Precision = 38
        prm3.NumericScale = 2
        prm3.Value = CDec(IIf(txtNetto = "", 0, txtNetto))
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("JUMLAH_BARANG", adDecimal, adParamInput)
        prm4.Precision = 38
        prm4.NumericScale = 2
        prm4.Value = CDec(IIf(txtJumahBrg = "", 0, txtJumahBrg))
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("JUMLAH_Kemasan", adDecimal, adParamInput)
        prm5.Precision = 38
        prm5.NumericScale = 2
        prm5.Value = CDec(IIf(txtSeriKemasan(0) = "", 0, txtSeriKemasan(0)))
        cmd.Parameters.append prm5
        cmd.Parameters.append cmd.CreateParameter("Kota_TTD", adVarChar, adParamInput, 100, txtTempat.Text)
        cmd.Parameters.append cmd.CreateParameter("Tanggal_TTD", adDBTime, adParamInput, , dtpTanggal.Value)
        cmd.Parameters.append cmd.CreateParameter("Nama_TTD", adVarChar, adParamInput, 50, txtPemberitahu.Text)
        cmd.Parameters.append cmd.CreateParameter("Jabatan_TTD", adVarChar, adParamInput, 50, txtJabatan.Text)
        cmd.Parameters.append cmd.CreateParameter("NO_PENGAJUAN", adVarChar, adParamInput, 50, Replace(txtNoPengajuan.Text, "-", ""))
        
        Set RS = cmd.Execute
End Sub

Private Sub cmdSubmitBahanBaku_Click()
Dim strSQL As String
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim prm1 As ADODB.Parameter
Dim prm3 As ADODB.Parameter
Dim prm4 As ADODB.Parameter
Dim prm5 As ADODB.Parameter

Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41Bahan_InsertUpdate"
    
    cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 50, txtIdBahan)
    Set prm = cmd.CreateParameter("Harga_Penyerahan", adDecimal, adParamInput)
    prm.Precision = 38
    prm.NumericScale = 2
    prm.Value = CDec(IIf(txtHargaPenyerhan = "", 0, txtHargaPenyerhan))
    cmd.Parameters.append prm
    cmd.Parameters.append cmd.CreateParameter("Jenis_Satuan", adVarChar, adParamInput, 50, txtJenisSatuan.Text)
    Set prm1 = cmd.CreateParameter("Jumlah_Satuan", adDecimal, adParamInput)
    prm1.Precision = 38
    prm1.NumericScale = 2
    prm1.Value = CDec(IIf(txtJmlSatuan = "", 0, txtJmlSatuan))
    cmd.Parameters.append prm1
    cmd.Parameters.append cmd.CreateParameter("Kode_Barang", adVarChar, adParamInput, 50, txtKodeBrg.Text)
    cmd.Parameters.append cmd.CreateParameter("Kode_Jenis_Dok_Asal", adVarChar, adParamInput, 50, txtDokAsal.Text)
    cmd.Parameters.append cmd.CreateParameter("Kode_Kantor", adVarChar, adParamInput, 50, txtKPPBC.Text)
    cmd.Parameters.append cmd.CreateParameter("Nomor_Aju_Dok_Asal", adVarChar, adParamInput, 50, Replace(txtNoAju.Text, "-", ""))
    cmd.Parameters.append cmd.CreateParameter("Nomor_Daftar_Dok_Asal", adVarChar, adParamInput, 50, txtNoDok.Text)
    'cmd.Parameters.append cmd.CreateParameter("Seri_Bahan_Baku", adVarChar, adParamInput, 50, txtBrgSpf.Text)
    cmd.Parameters.append cmd.CreateParameter("Seri_Barang", adVarChar, adParamInput, 50, txtBrgId.Text)
    cmd.Parameters.append cmd.CreateParameter("Tanggal_Daftar_Dok_Asal", adDBTime, adParamInput, , DTPDok.Value)
    cmd.Parameters.append cmd.CreateParameter("Tipe", adVarChar, adParamInput, 50, txtTipe.Text)
    cmd.Parameters.append cmd.CreateParameter("Uraian", adVarChar, adParamInput, 100, txtUraianBrg.Text)
    cmd.Parameters.append cmd.CreateParameter("Merk", adVarChar, adParamInput, 50, txtMerk.Text)
    cmd.Parameters.append cmd.CreateParameter("Ukuran", adVarChar, adParamInput, 50, txtUkuran.Text)
    cmd.Parameters.append cmd.CreateParameter("Spesifikasi_Lain", adVarChar, adParamInput, 50, txtSPF.Text)
    'cmd.Parameters.append cmd.CreateParameter("Id_Barang", adVarChar, adParamInput, 50, txtBrgSpf.Text)
    cmd.Parameters.append cmd.CreateParameter("Id_Header", adVarChar, adParamInput, 50, lblNoId.Caption)
    
    Set RS = cmd.Execute
      
    up_DetailBahan
              
    tampil_detail
                         
    LblErrMsg = DisplayMsg(1000)
End Sub

Private Sub tampil_detail()
    If rs_detail.EOF = False Then
        txtIdBahan = Trim(rs_detail("Id"))
        txtBahanId.Text = rs_detail.AbsolutePosition
        txtBahanEnd.Text = rs_detail.RecordCount
        'cboBrgStatus.Text = Trim(rs_detail_barang("KODE_STATUS"))
        txtDokAsal.Text = Trim(rs_detail("Kode_Jenis_Dok_Asal"))
        txtNoDok.Text = Trim(rs_detail("Nomor_Daftar_Dok_Asal"))
        DTPDok.Value = Trim(rs_detail("Tanggal_Daftar_Dok_Asal"))
        txtKPPBC.Text = Trim(rs_detail("Kode_Kantor"))
        txtNoAju.Text = Format(rs_detail("Nomor_Aju_Dok_Asal"), gs_formatNoAju)
        txtKodeBrg.Text = Trim(rs_detail("Kode_Barang"))
        txtUraianBrg.Text = Trim(rs_detail("Uraian"))
        txtJmlSatuan.Text = Trim(rs_detail("Jumlah_Satuan"))
        txtJenisSatuan.Text = Trim(rs_detail("Jenis_Satuan"))
        txtTipe.Text = IIf(IsNull(Trim(rs_detail("Tipe"))) = True, "", Trim(rs_detail("Tipe")))
        txtMerk.Text = Trim(rs_detail("Merk"))
        txtUkuran.Text = Trim(rs_detail("Ukuran"))
        txtSPF.Text = Trim(rs_detail("Spesifikasi_Lain"))
        txtHargaPenyerhan.Text = Format(rs_detail("Harga_Penyerahan"), gs_formatAmountIDR)
    End If
End Sub

Private Sub cmdSubmitBarang_Click()
Dim strSQL As String
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim prm3 As ADODB.Parameter
Dim prm4 As ADODB.Parameter
Dim prm5 As ADODB.Parameter

Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41Barang_InsertUpdate"
    'lblBrgId.Caption = ""
    
   cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 50, lblBrgId.Caption)
    cmd.Parameters.append cmd.CreateParameter("Kode_Barang", adVarChar, adParamInput, 50, txtBrgKode.Text)
    cmd.Parameters.append cmd.CreateParameter("URAIAN", adVarChar, adParamInput, 50, txtBrgUraian.Text)
    cmd.Parameters.append cmd.CreateParameter("Merk", adVarChar, adParamInput, 50, txtBrgMerk.Text)
    cmd.Parameters.append cmd.CreateParameter("Tipe", adVarChar, adParamInput, 50, txtBrgType.Text)
    cmd.Parameters.append cmd.CreateParameter("Ukuran", adVarChar, adParamInput, 50, txtBrgUkuran.Text)
    cmd.Parameters.append cmd.CreateParameter("Spesifikasi_Lain", adVarChar, adParamInput, 50, txtBrgSpf.Text)
    Set prm1 = cmd.CreateParameter("Jumlah_Satuan", adDecimal, adParamInput)
    prm1.Precision = 38
    prm1.NumericScale = 2
    prm1.Value = CDec(IIf(txtBrgJumlahSatuan = "", 0, txtBrgJumlahSatuan))
    cmd.Parameters.append prm1
    cmd.Parameters.append cmd.CreateParameter("Kode_Satuan", adVarChar, adParamInput, 50, txtBrgJenisSatuan.Text)
    Set prm2 = cmd.CreateParameter("Netto", adDecimal, adParamInput)
    prm2.Precision = 38
    prm2.NumericScale = 2
    prm2.Value = CDec(IIf(txtBrgNetto = "", 0, txtBrgNetto))
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("Volume", adDecimal, adParamInput)
    prm3.Precision = 38
    prm3.NumericScale = 2
    prm3.Value = CDec(IIf(txtBrgVolume = "", 0, txtBrgVolume))
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("Harga_Penyerahan", adDecimal, adParamInput)
    prm4.Precision = 38
    prm4.NumericScale = 2
    prm4.Value = CDec(IIf(txtBrgHarga = "", 0, txtBrgHarga))
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("Harga_Penyerahan", adDecimal, adParamInput)
    prm5.Precision = 38
    prm5.NumericScale = 2
    prm5.Value = CDec(IIf(txtBahanBaku = "", 0, txtBahanBaku))
    cmd.Parameters.append prm5

    cmd.Parameters.append cmd.CreateParameter("Id_Header", adVarChar, adParamInput, 50, lblNoId.Caption)
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan.Text, "-", ""))
  
        
    Set RS = cmd.Execute
    
    koneksi
    
    data_tampil
    
    up_HargaPenyerahan
    
    up_Netto
    
    up_Volume
                       
    LblErrMsg = DisplayMsg(1000)
End Sub

Private Sub cmdSubmitDokumen_Click()
    Dim strSQL As String
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim id As String
Dim tanya

LblErrMsg.Caption = ""

    hapus = False
    With gridDokumen
        For i = 1 To .Rows - 1
            If .TextMatrix(i, colSelect) = "D" Then
                If IsEmpty(tanya) Then tanya = MsgBox("Do you really want to delete this data ?", vbQuestion & vbYesNo, "Confirmation")
                If tanya = vbYes Then
                
                Set cmd = New ADODB.Command
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 0
                cmd.ActiveConnection = Db
                cmd.CommandText = "sp_BC41Dokumen_Del"
                
                cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 100, gridDokumen.TextMatrix(i, colId))
                
                Set RS = cmd.Execute
                    
                    Clear_dokumen
                    
                    LblErrMsg = DisplayMsg(1201)
                    
                    Exit Sub
                    
                Else
                    Exit Sub
                End If
            End If
        Next i
                
    End With
    'Exit Sub
    
  
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41Dokumen_InsertUpdate"
    
     With gridDokumen
      For i = 1 To .Rows - 1
      If .TextMatrix(i, colSelect) = "S" Then
        lblIdDokumen.Caption = gridDokumen.TextMatrix(i, colId)
      End If
    Next i
    End With
    
    If txtDokumen.Text = "" Then
        txtDokumen.SetFocus
        LblErrMsg = DisplayMsg(8123)  '"Please Input Kode Dokumen !"
        Exit Sub
    End If
    If txtNoDokumen(1).Text = "" Then
        txtNoDokumen(1).SetFocus
        LblErrMsg = DisplayMsg(8124)  '"Please Input No Dokumen !"
        Exit Sub
    End If

    With gridDokumen
      For i = 1 To .Rows - 1
      If .TextMatrix(i, colSelect) = "S" Then
        lblIdDokumen.Caption = gridDokumen.TextMatrix(i, colId)
      End If
    Next i
    End With
    
    cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 100, lblIdDokumen.Caption)
    cmd.Parameters.append cmd.CreateParameter("Kode_Jenis_Dokumen", adVarChar, adParamInput, 100, txtDokumen.Text)
    cmd.Parameters.append cmd.CreateParameter("Nomor_Daftar", adVarChar, adParamInput, 50, txtNoDokumen(1).Text)
    cmd.Parameters.append cmd.CreateParameter("Tanggal_Dokumen", adDBTime, adParamInput, , DTPDokumen.Value)
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    
    Set RS = cmd.Execute
    
    Clear_dokumen
    
    LblErrMsg = DisplayMsg(1000)
End Sub

Private Sub up_IsiGridDokumen()
Dim sql As String
Dim cmd As ADODB.Command
Dim li_Row As Integer

up_HeaderDokumen

If txtNoPengajuan.Text = "______-______-________-______" Then

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41Dokumen_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 100, Replace(FrmBC41List.txtTampung, "-", ""))
    
    Set rsDokumen = cmd.Execute
    
        
        i = 1
        With gridDokumen
            While Not rsDokumen.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(i, colSelect) = ""
                .TextMatrix(i, colId) = Trim(rsDokumen("ID"))
                .TextMatrix(i, colSeri) = Trim(rsDokumen("Seri_Dokumen"))
                .TextMatrix(i, colKodeDokumen) = Trim(rsDokumen("Kode_Jenis_Dokumen"))
                    If (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 388 Then
                        txtFakturPajak(4).Text = Trim(rsDokumen("Nomor_Dokumen"))
                        DTPFakturPajak.Value = Trim(rsDokumen("Tanggal_Dokumen"))
                    ElseIf (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 217 Then
                        txtPackingList(2).Text = Trim(rsDokumen("Nomor_Dokumen"))
                        DTPPackingList.Value = Trim(rsDokumen("Tanggal_Dokumen"))
                    ElseIf (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 315 Then
                        txtKontrak(3).Text = Trim(rsDokumen("Nomor_Dokumen"))
                        DTPKontrak.Value = Trim(RS("Tanggal_Dokumen"))
                    ElseIf (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 912 Then
                        txtSkep(5).Text = Trim(rsDokumen("Nomor_Dokumen"))
                        DTPSkep.Value = Trim(rsDokumen("Tanggal_Dokumen"))
                    End If
                .TextMatrix(i, colJenisDokumen) = Trim(rsDokumen("Jenis_Dokumen"))
                .TextMatrix(i, colNomor) = Trim(rsDokumen("Nomor_Dokumen"))
                .TextMatrix(i, colTanggal) = Format(Trim(rsDokumen("Tanggal_Dokumen")), "dd MMM yyyy")
                i = i + 1
            rsDokumen.MoveNext
            Wend
        End With
Else
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41Dokumen_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 100, Replace(txtNoPengajuan, "-", ""))
    
    Set rsDokumen = cmd.Execute
    
        
        i = 1
        With gridDokumen
            While Not rsDokumen.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(i, colSelect) = ""
                .TextMatrix(i, colId) = Trim(rsDokumen("ID"))
                .TextMatrix(i, colSeri) = Trim(rsDokumen("Seri_Dokumen"))
                .TextMatrix(i, colKodeDokumen) = Trim(rsDokumen("Kode_Jenis_Dokumen"))
                    If (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 388 Then
                        txtFakturPajak(4).Text = Trim(rsDokumen("Nomor_Dokumen"))
                        DTPFakturPajak.Value = Trim(rsDokumen("Tanggal_Dokumen"))
                    ElseIf (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 217 Then
                        txtPackingList(2).Text = Trim(rsDokumen("Nomor_Dokumen"))
                        DTPPackingList.Value = Trim(rsDokumen("Tanggal_Dokumen"))
                    ElseIf (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 315 Then
                        txtKontrak(3).Text = Trim(rsDokumen("Nomor_Dokumen"))
                        DTPKontrak.Value = Trim(RS("Tanggal_Dokumen"))
                    ElseIf (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 912 Then
                        txtSkep(5).Text = Trim(rsDokumen("Nomor_Dokumen"))
                        DTPSkep.Value = Trim(rsDokumen("Tanggal_Dokumen"))
                    ElseIf (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 911 Then
                        txtSkep(5).Text = Trim(rsDokumen("Nomor_Dokumen"))
                        DTPSkep.Value = Trim(rsDokumen("Tanggal_Dokumen"))
                    ElseIf (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 40 Then
                        txtBC40(0).Text = Trim(rsDokumen("Nomor_Dokumen"))
                        DTPBC40.Value = Trim(rsDokumen("Tanggal_Dokumen"))
                    End If
                .TextMatrix(i, colJenisDokumen) = Trim(rsDokumen("Jenis_Dokumen"))
                .TextMatrix(i, colNomor) = Trim(rsDokumen("Nomor_Dokumen"))
                .TextMatrix(i, colTanggal) = Format(Trim(rsDokumen("Tanggal_Dokumen")), "dd MMM yyyy")
                i = i + 1
            rsDokumen.MoveNext
            Wend
        End With
    End If
End Sub

Private Sub cmdSyncronize_Click()
Dim strSQL As String
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim prm3 As ADODB.Parameter
Dim prm4 As ADODB.Parameter
Dim tanya

Set RS = New ADODB.Recordset

   
    If IsEmpty(tanya) Then tanya = MsgBox("Do you want to syncronize to Mysql ?", vbQuestion & vbYesNo, "Confirmation")
        If tanya = vbYes Then
        
        Me.MousePointer = vbHourglass
    
        KoneksiMysql
        
        strSQL = "Select * from tpbdb.tpb_header where Nomor_Aju='" & Replace(txtNoPengajuan.Text, "-", "") & "'"
                
                'rsId.Open strSQL, ConnStr
                If RSiD.State <> adStateClosed Then RSiD.Close
                RSiD.Open strSQL, ConnStr, adOpenForwardOnly, adLockReadOnly, adCmdText
                
            
            If RSiD.EOF = False Then
               tampung = Trim(RSiD("Id"))
               
                 
               strSQL = " Update tpbdb.tpb_header " & vbCrLf & _
                         " set Nomor_Aju='" & Replace(txtNoPengajuan.Text, "-", "") & "', Tanggal_Daftar='" & Format(Trim(dtpTanggal.Value), "yyyy-mm-dd") & "', TANGGAL_AJU='" & Format(Trim(dtpTanggal.Value), "yyyy-mm-dd") & "', " & vbCrLf & _
                         " Kode_Kantor='" & txtKantorPabean.Text & "', Id_Modul=  '" & Mid$(txtNoPengajuan.Text, 8, 5) & "',  Kode_Jenis_TPB='" & Left(cboJenisTPB.Text, 1) & "', Kode_Tujuan_Pengiriman='" & Left(cboTujuan.Text, 1) & "', " & vbCrLf & _
                         " Kode_Id_Pengusaha='" & Left(cboNPWPPengusaha.Text, 1) & "', ID_PENGUSAHA='" & txtNPWPPengusaha.Text & "', Nama_Pengusaha='" & txtNamaPengusaha.Text & "', " & vbCrLf & _
                         " Alamat_Pengusaha='" & txtAlamatPengusaha.Text & "', Nomor_Ijin_TPB='" & txtNoIzinPengusaha.Text & "', Kode_Id_Penerima_Barang='" & Left(cboNPWPPengirim.Text, 1) & "', " & vbCrLf & _
                         " Id_Penerima_Barang='" & txtNPWPPengirim.Text & "', Nama_Penerima_Barang ='" & txtNamaPengirim.Text & "', Alamat_Penerima_Barang='" & txtAlamatPengirim.Text & "', Nama_Pengangkut='" & txtNamaPengangkut.Text & "', " & vbCrLf & _
                         " Nomor_Polisi='" & txtNoPolisi.Text & "', Harga_Penyerahan ='" & CDec(txtHarga.Text) & "', VOLUME='" & txtVolume.Text & "', BRUTO='" & txtBruto.Text & "', " & vbCrLf & _
                         " NETTO='" & txtNetto.Text & "', JUMLAH_BARANG='" & txtJumahBrg.Text & "', Kota_TTD='" & txtTempat.Text & "', Tanggal_TTD ='" & Format(Trim(dtpTanggal.Value), "yyyy-mm-dd") & "', " & vbCrLf & _
                         " Nama_TTD='" & txtPemberitahu.Text & "', Jabatan_TTD='" & txtJabatan.Text & "' where Id='" & tampung & "'"
                         
                         
                'rs.Close
                 RS.Open strSQL, ConnStr
                            
    '            'Insert Dokumen
    
                   strSQL = " Delete from tpbdb.tpb_dokumen Where Id_Header='" & Trim(RSiD("Id")) & "' "
                   RS.Open strSQL, ConnStr
    
                Set cmd = New ADODB.Command
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 0
                cmd.ActiveConnection = Db
                cmd.CommandText = "sp_BC41Dokumen_Sel"
                
                cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 100, Replace(txtNoPengajuan, "-", ""))
                
                Set rsDokumen = cmd.Execute
                
                'With GridDokumen
                'rs.Close
                Do While Not rsDokumen.EOF
                    strSQL = " Insert into tpbdb.tpb_dokumen " & vbCrLf & _
                             " (Kode_Jenis_Dokumen, Nomor_Dokumen, Seri_Dokumen, Tanggal_Dokumen, Tipe_Dokumen, Id_Header) " & vbCrLf & _
                             " values ('" & Trim(rsDokumen("Kode_Jenis_Dokumen")) & "', '" & Trim(rsDokumen("Nomor_Dokumen")) & "', '" & Trim(rsDokumen("Seri_Dokumen")) & "', " & vbCrLf & _
                             " '" & Format(rsDokumen("Tanggal_Dokumen"), "yyyy-mm-dd") & "', '" & Trim(rsDokumen("Tipe_Dokumen")) & "','" & tampung & "')"
    '
                    RS.Open strSQL, ConnStr
                    rsDokumen.MoveNext
                Loop
                
                'Insert Kemasan
                strSQL = " Delete from tpbdb.tpb_kemasan Where Id_Header='" & tampung & "' "
                RS.Open strSQL, ConnStr
                
                   
                Set cmd = New ADODB.Command
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 0
                cmd.ActiveConnection = Db
                cmd.CommandText = "sp_BC41Kemasan_Sel"
                 
                cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 100, Replace(txtNoPengajuan, "-", ""))
        
                Set rsKemasan = cmd.Execute
                
                Do While Not rsKemasan.EOF
                    strSQL = " Insert into tpbdb.tpb_kemasan " & vbCrLf & _
                             " (Jumlah_Kemasan, Kode_Jenis_Kemasan, Merk_Kemasan, Seri_Kemasan, Id_Header) " & vbCrLf & _
                             " Values('" & Trim(rsKemasan("Jumlah_Kemasan")) & "','" & Trim(rsKemasan("Kode_Jenis_Kemasan")) & "','" & Trim(rsKemasan("Merk_Kemasan")) & "', " & vbCrLf & _
                             " '" & Trim(rsKemasan("Seri_Kemasan")) & "', '" & tampung & "' )"
                    RS.Open strSQL, ConnStr
                    rsKemasan.MoveNext
                Loop
                
                'Insert Barang
                strSQL = " Delete from tpbdb.tpb_bahan_baku Where Id_Header='" & tampung & "' "
                RS.Open strSQL, ConnStr
                
                strSQL = " Delete from tpbdb.tpb_barang Where Id_Header='" & tampung & "' "
                RS.Open strSQL, ConnStr
                
                Set cmd = New ADODB.Command
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 0
                cmd.ActiveConnection = Db
                cmd.CommandText = "sp_BC41Barang_Sel"
                 
                cmd.Parameters.append cmd.CreateParameter("Id_Header", adVarChar, adParamInput, 50, lblNoId.Caption)
        
                Set rsBarang = cmd.Execute
                
                Do While Not rsBarang.EOF
                    strSQL = " Insert into tpbdb.tpb_barang " & vbCrLf & _
                             " (Harga_Penyerahan, Jumlah_Satuan, Kode_Barang, Kode_Satuan, Netto, Seri_Barang, Uraian, Volume, Id_Header) " & vbCrLf & _
                             " Values('" & Trim(rsBarang("Harga_Penyerahan")) & "','" & Trim(rsBarang("Jumlah_Satuan")) & "', '" & Trim(rsBarang("Kode_Barang")) & "' , " & vbCrLf & _
                             " '" & Trim(rsBarang("Kode_Satuan")) & "', '" & Trim(rsBarang("Netto")) & "', '" & Trim(rsBarang("Seri_Barang")) & "', '" & Trim(rsBarang("Uraian")) & "', " & vbCrLf & _
                             " '" & Trim(rsBarang("Volume")) & "', '" & tampung & "' ) "
                             
                    RS.Open strSQL, ConnStr
                    rsBarang.MoveNext
                Loop
                
                'Insert Bahan Baku
'                strSQL = " Delete from tpbdb.tpb_bahan_baku Where Id_Header='" & tampung & "' "
'                rs.Open strSQL, ConnStr
                
                Set cmd = New ADODB.Command
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 0
                cmd.ActiveConnection = Db
                cmd.CommandText = "sp_BC41Bahan_Sel"
                 
                cmd.Parameters.append cmd.CreateParameter("Id_Header", adVarChar, adParamInput, 50, lblNoId.Caption)
        
                Set rsBahan = cmd.Execute
                
                'Get Id Barang
                strSQL = " Select Id from tpbdb.tpb_barang Where Id_Header='" & tampung & "' "
                If rsBarang.State <> adStateClosed Then rsBarang.Close
                rsBarang.Open strSQL, ConnStr, adOpenForwardOnly, adLockReadOnly, adCmdText
                
               If rsBarang.EOF = False Then
                    Do While Not rsBahan.EOF
                        strSQL = " Insert into tpbdb.tpb_bahan_baku " & vbCrLf & _
                                 " (Harga_Penyerahan, Jenis_Satuan, Jumlah_Satuan, Kode_Asal_Bahan_Baku, KODE_BARANG, " & vbCrLf & _
                                 "  Kode_Jenis_Dok_Asal, Kode_Kantor, Nomor_Aju_Dok_Asal, Nomor_Daftar_Dok_Asal, Seri_Bahan_Baku, Seri_Barang, " & vbCrLf & _
                                 "  Tanggal_Daftar_Dok_Asal, Uraian, TIPE, MERK, UKURAN, SPESIFIKASI_LAIN, Id_Barang, Id_Header)" & vbCrLf & _
                                 " Values('" & Trim(rsBahan("Harga_Penyerahan")) & "','" & Trim(rsBahan("Jenis_Satuan")) & "','" & Trim(rsBahan("Jumlah_Satuan")) & "', " & vbCrLf & _
                                 " '" & Trim(rsBahan("Kode_Asal_Bahan_Baku")) & "', '" & Trim(rsBahan("KODE_BARANG")) & "', '" & Trim(rsBahan("Kode_Jenis_Dok_Asal")) & "', " & vbCrLf & _
                                 " '" & Trim(rsBahan("Kode_Kantor")) & "', '" & Trim(rsBahan("Nomor_Aju_Dok_Asal")) & "', '" & Trim(rsBahan("Nomor_Daftar_Dok_Asal")) & "', " & vbCrLf & _
                                 " '" & Trim(rsBahan("Seri_Bahan_Baku")) & "', '" & Trim(rsBahan("Seri_Barang")) & "', '" & Trim(rsBahan("Tanggal_Daftar_Dok_Asal")) & "', " & vbCrLf & _
                                 " '" & Trim(rsBahan("Uraian")) & "', '" & Trim(rsBahan("TIPE")) & "', '" & Trim(rsBahan("MERK")) & "', '" & Trim(rsBahan("MERK")) & "', " & vbCrLf & _
                                 " '" & Trim(rsBahan("SPESIFIKASI_LAIN")) & "', '" & Trim(rsBarang("Id")) & "', '" & tampung & "' ) "
                                 
                        RS.Open strSQL, ConnStr
                        rsBahan.MoveNext
                        rsBarang.MoveNext
                    Loop
                End If
                    
                
                Insert_Update
                       
                Me.MousePointer = vbDefault
                
                LblErrMsg = DisplayMsg(9007)
            Else
                'rs.Close
                'PErcontohan
                strSQL = " Insert into tpbdb.tpb_header " & vbCrLf & _
                         " (Nomor_Aju, Tanggal_Daftar, TANGGAL_AJU, Kode_Kantor, KODE_DOKUMEN_PABEAN, ID_MODUL, Kode_Jenis_TPB, Kode_Tujuan_Pengiriman, " & vbCrLf & _
                         " Kode_Id_Pengusaha,  ID_PENGUSAHA, Nama_Pengusaha, Alamat_Pengusaha, Nomor_Ijin_TPB, Kode_Id_Penerima_Barang, Id_Penerima_Barang, Nama_Penerima_Barang, " & vbCrLf & _
                         " Alamat_Penerima_Barang, Nama_Pengangkut,Nomor_Polisi, Harga_Penyerahan, VOLUME, BRUTO, NETTO, JUMLAH_BARANG, Kota_TTD, Tanggal_TTD, Nama_TTD, Jabatan_TTD) " & vbCrLf & _
                         " values('" & Replace(txtNoPengajuan.Text, "-", "") & "','" & Format(Trim(dtpTanggal.Value), "yyyy-mm-dd") & "' , '" & Format(Trim(dtpTanggal.Value), "yyyy-mm-dd") & "', '" & txtKantorPabean.Text & "', '" & Mid$(txtNoPengajuan.Text, 5, 2) & "', '" & Mid$(txtNoPengajuan.Text, 8, 5) & "' , '" & Left(cboJenisTPB.Text, 1) & "', '" & Left(cboTujuan.Text, 1) & "', " & vbCrLf & _
                         "  '" & Left(cboNPWPPengusaha.Text, 1) & "', '" & txtNPWPPengusaha.Text & "', '" & txtNamaPengusaha.Text & "', '" & txtAlamatPengusaha.Text & "', '" & txtNoIzinPengusaha.Text & "', " & vbCrLf & _
                         " '" & Left(cboNPWPPengirim.Text, 1) & "', '" & txtNPWPPengirim.Text & "', '" & txtNamaPengirim.Text & "', '" & txtAlamatPengirim.Text & "', '" & txtNamaPengangkut.Text & "',  " & vbCrLf & _
                         " '" & txtNoPolisi.Text & "', '" & txtHarga.Text & "', '" & txtVolume.Text & "', '" & txtBruto.Text & "', '" & txtNetto.Text & "', '" & txtJumahBrg.Text & "', '" & txtTempat.Text & "', " & vbCrLf & _
                         " '" & Format(Trim(dtpTanggal.Value), "yyyy-mm-dd") & "', '" & txtPemberitahu.Text & "', '" & txtJabatan.Text & "')"
                         
                RS.Open strSQL, ConnStr
                
                
                strSQL = "Select * from tpbdb.tpb_header where Nomor_Aju='" & Replace(txtNoPengajuan.Text, "-", "") & "'"
                
                'Insert Dokumen
                If RSiD.State <> adStateClosed Then RSiD.Close
                RSiD.Open strSQL, ConnStr, adOpenForwardOnly, adLockReadOnly, adCmdText
                tampung = Trim(RSiD("Id"))
                
                Set cmd = New ADODB.Command
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 0
                cmd.ActiveConnection = Db
                cmd.CommandText = "sp_BC41Dokumen_Sel"
                
                cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 100, Replace(txtNoPengajuan, "-", ""))
                
                Set rsDokumen = cmd.Execute
                
                Do While Not rsDokumen.EOF
                    strSQL = " Insert into tpbdb.tpb_dokumen " & vbCrLf & _
                             " (Kode_Jenis_Dokumen, Nomor_Dokumen, Seri_Dokumen, Tanggal_Dokumen, Tipe_Dokumen, Id_Header) " & vbCrLf & _
                             " values ('" & Trim(rsDokumen("Kode_Jenis_Dokumen")) & "', '" & Trim(rsDokumen("Nomor_Dokumen")) & "', '" & Trim(rsDokumen("Seri_Dokumen")) & "', " & vbCrLf & _
                             " '" & Format(rsDokumen("Tanggal_Dokumen"), "yyyy-mm-dd") & "', '" & Trim(rsDokumen("Tipe_Dokumen")) & "','" & tampung & "')"
                          
                    RS.Open strSQL, ConnStr
                    rsDokumen.MoveNext
                Loop
                
                
                'Insert Kemasan
                Set cmd = New ADODB.Command
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 0
                cmd.ActiveConnection = Db
                cmd.CommandText = "sp_BC41Kemasan_Sel"
                 
                cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 100, Replace(txtNoPengajuan, "-", ""))

        
                Set rsKemasan = cmd.Execute
                
                Do While Not rsKemasan.EOF
                    strSQL = " Insert into tpbdb.tpb_kemasan " & vbCrLf & _
                             " (Jumlah_Kemasan, Kode_Jenis_Kemasan, Merk_Kemasan, Seri_Kemasan, Id_Header) " & vbCrLf & _
                             " Values('" & Trim(rsKemasan("Jumlah_Kemasan")) & "','" & Trim(rsKemasan("Kode_Jenis_Kemasan")) & "','" & Trim(rsKemasan("Merk_Kemasan")) & "', " & vbCrLf & _
                             " '" & Trim(rsKemasan("Seri_Kemasan")) & "', '" & tampung & "' )"
                    RS.Open strSQL, ConnStr
                    rsKemasan.MoveNext
                Loop
                
                'Insert Barang
                Set cmd = New ADODB.Command
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 0
                cmd.ActiveConnection = Db
                cmd.CommandText = "sp_BC41Barang_Sel"
                 
                cmd.Parameters.append cmd.CreateParameter("Id_Header", adVarChar, adParamInput, 50, lblNoId.Caption)
        
                Set rsBarang = cmd.Execute
                
                Do While Not rsBarang.EOF
                    strSQL = " Insert into tpbdb.tpb_barang " & vbCrLf & _
                             " (Harga_Penyerahan, Jumlah_Satuan, Kode_Barang, Kode_Satuan, Netto, Seri_Barang, Uraian, Volume, Id_Header) " & vbCrLf & _
                             " Values('" & Trim(rsBarang("Harga_Penyerahan")) & "','" & Trim(rsBarang("Jumlah_Satuan")) & "', '" & Trim(rsBarang("Kode_Barang")) & "' , " & vbCrLf & _
                             " '" & Trim(rsBarang("Kode_Satuan")) & "', '" & Trim(rsBarang("Netto")) & "', '" & Trim(rsBarang("Seri_Barang")) & "', '" & Trim(rsBarang("Uraian")) & "',  " & vbCrLf & _
                             " '" & Trim(rsBarang("Volume")) & "', '" & tampung & "' ) "
                             
                    RS.Open strSQL, ConnStr
                    rsBarang.MoveNext
                Loop
                
                Set cmd = New ADODB.Command
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 0
                cmd.ActiveConnection = Db
                cmd.CommandText = "sp_BC41Bahan_Sel"
                 
                cmd.Parameters.append cmd.CreateParameter("Id_Header", adVarChar, adParamInput, 50, lblNoId.Caption)
        
                Set rsBahan = cmd.Execute
                
                'Get Id Barang
                strSQL = " Select Id from tpbdb.tpb_barang Where Id_Header='" & tampung & "' "
                If rsBarang.State <> adStateClosed Then rsBarang.Close
                rsBarang.Open strSQL, ConnStr, adOpenForwardOnly, adLockReadOnly, adCmdText
                
                If rsBarang.EOF = False Then
                
                    Do While Not rsBahan.EOF
                        strSQL = " Insert into tpbdb.tpb_bahan_baku " & vbCrLf & _
                                 " (Harga_Penyerahan, Jenis_Satuan, Jumlah_Satuan, Kode_Asal_Bahan_Baku, KODE_BARANG, " & vbCrLf & _
                                 "  Kode_Jenis_Dok_Asal, Kode_Kantor, Nomor_Aju_Dok_Asal, Nomor_Daftar_Dok_Asal, Seri_Bahan_Baku, Seri_Barang, " & vbCrLf & _
                                 "  Tanggal_Daftar_Dok_Asal, Uraian, TIPE, MERK, UKURAN, SPESIFIKASI_LAIN, Id_Barang, Id_Header)" & vbCrLf & _
                                 " Values('" & Trim(rsBahan("Harga_Penyerahan")) & "','" & Trim(rsBahan("Jenis_Satuan")) & "','" & Trim(rsBahan("Jumlah_Satuan")) & "', " & vbCrLf & _
                                 " '" & Trim(rsBahan("Kode_Asal_Bahan_Baku")) & "', '" & Trim(rsBahan("KODE_BARANG")) & "', '" & Trim(rsBahan("Kode_Jenis_Dok_Asal")) & "', " & vbCrLf & _
                                 " '" & Trim(rsBahan("Kode_Kantor")) & "', '" & Trim(rsBahan("Nomor_Aju_Dok_Asal")) & "', '" & Trim(rsBahan("Nomor_Daftar_Dok_Asal")) & "', " & vbCrLf & _
                                 " '" & Trim(rsBahan("Seri_Bahan_Baku")) & "', '" & Trim(rsBahan("Seri_Barang")) & "', '" & Trim(rsBahan("Tanggal_Daftar_Dok_Asal")) & "', " & vbCrLf & _
                                 " '" & Trim(rsBahan("Uraian")) & "', '" & Trim(rsBahan("TIPE")) & "', '" & Trim(rsBahan("MERK")) & "', '" & Trim(rsBahan("MERK")) & "', " & vbCrLf & _
                                 " '" & Trim(rsBahan("SPESIFIKASI_LAIN")) & "', '" & Trim(rsBarang("Id")) & "', '" & tampung & "' ) "
                                 
                        RS.Open strSQL, ConnStr
                        rsBahan.MoveNext
                        rsBarang.MoveNext
                    Loop
                End If
                
            End If
            
            Insert_Update
            
            Me.MousePointer = vbDefault
            
            LblErrMsg = DisplayMsg(9006)
            
    End If
    
End Sub

Private Sub cmSubimtKemasan_Click()
    Dim strSQL As String
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim tanya

LblErrMsg.Caption = ""


    hapus = False
    
    With gridKemasan
        For i = 1 To .Rows - 1
            If .TextMatrix(i, colSelect) = "D" Then
                If IsEmpty(tanya) Then tanya = MsgBox("Do you really want to delete this data ?", vbQuestion & vbYesNo, "Confirmation")
                If tanya = vbYes Then
                
                    Set cmd = New ADODB.Command
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandTimeout = 0
                    cmd.ActiveConnection = Db
                    cmd.CommandText = "sp_BC41Kemasan_Del"
                    
                    cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 100, gridKemasan.TextMatrix(i, bteColId))
                    
                    Set RS = cmd.Execute
                    
                    Clear_Kemasan
                    
                    LblErrMsg = DisplayMsg(1201)
                    
                    Exit Sub
                    
                Else
                   Exit Sub
                End If
            End If
        Next i
                
    End With

Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41Kemasan_InsertUpdate"
    
    With gridKemasan
      For i = 1 To .Rows - 1
      If .TextMatrix(i, colSelect) = "S" Then
        lblIdKemasan.Caption = gridKemasan.TextMatrix(i, bteColId)
      End If
    Next i
    End With
    
    If txtJenisKemasan.Text = "" Then
        txtJenisKemasan.SetFocus
        LblErrMsg = DisplayMsg(8125)  '"Please Input Kode Dokumen !"
        Exit Sub
    End If
    If txtJumlahKemasan(0).Text = "" Then
        txtJumlahKemasan(0).SetFocus
        LblErrMsg = DisplayMsg(8126)  '"Please Input No Dokumen !"
        Exit Sub
    End If
    
    cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 50, lblIdKemasan.Caption)
    cmd.Parameters.append cmd.CreateParameter("Jumlah_Kemasan", adVarChar, adParamInput, 50, txtJumlahKemasan(0).Text)
    cmd.Parameters.append cmd.CreateParameter("Kode_Jenis_Kemasan", adVarChar, adParamInput, 50, txtJenisKemasan.Text)
    cmd.Parameters.append cmd.CreateParameter("MERK_KEMASAN", adVarChar, adParamInput, 50, txtMerkKemasan(2).Text)
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    
    Set RS = cmd.Execute
    
    Clear_Kemasan
    
    LblErrMsg = DisplayMsg(1000)
    
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
    
    Case 1:
        If rs_barang.EOF = False Or rs_barang.BOF = False Then
        rs_barang.MoveFirst
        If rs_detail.EOF = False Or rs_detail.BOF = False Then
        rs_detail.MoveFirst
        End If
        Call tampil_detail
        Call data_tampil
        LblErrMsg.Caption = DisplayMsg("4020")
    End If
    Case 2:
        If rs_barang.EOF = False Or rs_barang.BOF = False Then
        rs_barang.MovePrevious: LblErrMsg.Caption = ""
        If rs_detail.EOF = False Or rs_detail.BOF = False Then
        rs_detail.MovePrevious: LblErrMsg.Caption = ""
        End If
        If rs_barang.BOF Then rs_barang.MoveFirst: LblErrMsg.Caption = DisplayMsg("4020") ': f_pesan = False
        Call tampil_detail
        Call data_tampil
        If rs_barang.AbsolutePosition = 1 Then LblErrMsg.Caption = DisplayMsg("4020")
        End If
    Case 3:
        If k_pertama = True Then
            If rs_barang.EOF = False Or rs_barang.BOF = False Then
            rs_barang.MoveFirst
            If rs_detail.EOF = False Or rs_detail.BOF = False Then
            rs_detail.MoveFirst
            End If
            Call tampil_detail
            Call data_tampil
            LblErrMsg.Caption = DisplayMsg("4020")
            k_pertama = False
            End If
        Else
            If rs_barang.EOF = False Or rs_barang.BOF = False Then
            rs_barang.MoveNext: LblErrMsg.Caption = ""
            
            If rs_detail.EOF = False Or rs_detail.BOF = False Then
            rs_detail.MoveNext: LblErrMsg.Caption = ""
            End If
            
            If rs_barang.EOF Then rs_barang.MoveLast: LblErrMsg.Caption = DisplayMsg("4121") ': f_pesan = False
            Call tampil_detail
            Call data_tampil
            If rs_barang.AbsolutePosition = rs_barang.RecordCount Then LblErrMsg.Caption = DisplayMsg("4121")
            End If
        End If
    Case 4:
        If rs_barang.EOF = False Or rs_barang.BOF = False Then
        rs_barang.MoveLast
        If rs_detail.EOF = False Or rs_detail.BOF = False Then
        rs_detail.MoveLast
        End If
        Call tampil_detail
        Call data_tampil
        LblErrMsg.Caption = DisplayMsg("4021")
        End If
    End Select
End Sub

Private Sub command2_Click()
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim tanya

    If IsEmpty(tanya) Then tanya = MsgBox("Do you really want to delete this data ?", vbQuestion & vbYesNo, "Confirmation")
        If tanya = vbYes Then
            
            Set cmd = New ADODB.Command
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 0
            cmd.ActiveConnection = Db
            cmd.CommandText = "sp_BC41Barang_Del"
            
            cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 100, lblBrgId.Caption)
            
            Set RS = cmd.Execute
            
            koneksi
            
            data_tampil
                        
            LblErrMsg = DisplayMsg(1201)
            
            Exit Sub
            
        Else
            Exit Sub
        End If
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub command3_Click()

End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub Form_Load()
  clear
  up_FillComboTujuan
  up_FillComboJenisTPB
  up_FillComboNPWP
  up_GridHeaderRespon
  up_GridHeaderStatus
  'up_HeaderDokumen
  'up_HeaderKemasan
  up_Form
     
  data_tampil
     
  CtrlMenu1.FormName = Me.Name
  Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
End Sub

Private Sub up_GridHeaderRespon()
    
    With gridRespon
        .ColS = colCountRespon
        .Rows = 1

        .TextMatrix(0, colKodeRespon) = "Kode"
        .TextMatrix(0, colUraianRespon) = "Uraian Respon"
        .TextMatrix(0, colWaktuRespon) = "Waktu"

        .ColWidth(colKodeRespon) = 800
        .ColWidth(colUraianRespon) = 3000
        .ColWidth(colWaktuRespon) = 1500
        
        .ColFormat(colWaktuRespon) = "dd MMM yyyy"
        .ColAlignment(colKodeRespon) = flexAlignLeftCenter
    End With
End Sub

Private Sub up_GridHeaderStatus()
    
    With gridStatus
        .ColS = colCountStatus
        .Rows = 1

        .TextMatrix(0, colKodeStatus) = "Kode"
        .TextMatrix(0, colUraianStatus) = "Uraian Status"
        .TextMatrix(0, colWaktuStatus) = "Waktu"

        .ColWidth(colKodeStatus) = 800
        .ColWidth(colUraianStatus) = 3000
        .ColWidth(colWaktuStatus) = 1500
        
        .ColFormat(colWaktuStatus) = "dd MMM yyyy"
        .ColAlignment(colKodeStatus) = flexAlignLeftCenter
    End With
End Sub

'**Tab Dokumen
Sub up_HeaderDokumen()
   
    With gridDokumen
        .ColS = colcount
        .Rows = 1
        
        .TextMatrix(0, colSelect) = ""
        .TextMatrix(0, colId) = "Id"
        .TextMatrix(0, colSeri) = "Seri"
        .TextMatrix(0, colKodeDokumen) = "Kode Dokumen"
        .TextMatrix(0, colJenisDokumen) = "Jenis Dokumen"
        .TextMatrix(0, colNomor) = "Nomor"
        .TextMatrix(0, colTanggal) = "Tanggal"
        
        .ColWidth(colSelect) = 250
        .ColWidth(colSeri) = 500
        .ColWidth(colKodeDokumen) = 1410
        .ColWidth(colJenisDokumen) = 1800
        .ColWidth(colNomor) = 2200
        .ColWidth(colTanggal) = 1410
        
        .ColHidden(colId) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, 5) = flexAlignCenterCenter
        
    End With
End Sub

Private Sub GridDokumen_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If gridDokumen.Col > colId Then Cancel = True
End Sub

Private Sub GridDokumen_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim TextGrid As String
    Dim k As Boolean
    Dim j As Integer
    
   k = False
    With gridDokumen
        TextGrid = gridDokumen.Text
        If TextGrid = "S" Then
            txtDokumen.Text = .TextMatrix(Row, colKodeDokumen)
            txtNoDokumen(1).Text = .TextMatrix(Row, colNomor)
            DTPDokumen.Value = Format(.TextMatrix(Row, colTanggal), "mm/dd/yyyy")
         Call ClearColGrid
        ElseIf TextGrid = "D" Then
            'Call ClearColGrid("S")
        End If
        
        .TextMatrix(Row, Col) = TextGrid
        For j = 1 To .Rows - 1
            If .TextMatrix(j, bteColSelect) <> "" Then k = True
        Next j
        If k = False Then Clear_dokumen
    End With
End Sub

Private Sub Clear_dokumen()
    up_HeaderDokumen
    txtDokumen.Text = ""
    lblDokumen(23).Caption = ""
    txtNoDokumen(1).Text = ""
    txtPackingList(2).Text = ""
    txtKontrak(3).Text = ""
    txtFakturPajak(4).Text = ""
    txtSkep(5).Text = ""
    DTPDokumen = Format(Now, "dd MMM yyyy")
    up_IsiGridDokumen
End Sub

Private Sub ClearColGrid(Optional Kolom As String)
 Dim i As Integer
    With gridDokumen
        .Col = bteColSelect
        If Kolom <> "" Then
            For i = 1 To .Rows - 1
                If .Text = Kolom Then .Text = ""
                If .TextMatrix(i, bteColSelect) <> "D" Then .TextMatrix(i, bteColSelect) = ""
            Next i
            clear
        Else
            For i = 1 To .Rows - 1
                If .TextMatrix(i, bteColSelect) <> "" Then .TextMatrix(i, bteColSelect) = ""
            Next i
        End If
    End With
End Sub

'**Kemasan
Sub up_HeaderKemasan()

bteColSelect = 0
bteColId = 1
bteColNo = 2
bteColjumlah = 3
bteColKode = 4
bteColUraian = 5
bteColMerkKemasan = 6
   
    With gridKemasan
         .clear
         .Rows = 1
         .ColS = 7
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColId) = "Id"
        .TextMatrix(0, bteColNo) = "No"
        .TextMatrix(0, bteColjumlah) = "Jumlah"
        .TextMatrix(0, bteColKode) = "Kode"
        .TextMatrix(0, bteColUraian) = "Uraian"
        .TextMatrix(0, bteColMerkKemasan) = "Merk Kemasan"
        
        .ColWidth(bteColSelect) = 250
        .ColWidth(bteColNo) = 500
        .ColWidth(bteColjumlah) = 1000
        .ColWidth(bteColKode) = 2000
        .ColWidth(bteColUraian) = 2000
        .ColWidth(bteColMerkKemasan) = 2500
        
        .ColHidden(bteColId) = True
        '.ColWidth(colTanggal) = 1800
        
        .Cell(flexcpAlignment, 0, 0, 0, 6) = flexAlignCenterCenter
        
    End With
End Sub

Private Sub GridKemasan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If gridKemasan.Col > bteColId Then Cancel = True
End Sub

Private Sub GridKemasan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim TextGrid As String
    Dim k As Boolean
    Dim j As Integer
    
   k = False
    With gridKemasan
        TextGrid = gridKemasan.Text
        If TextGrid = "S" Then
            txtJenisKemasan.Text = .TextMatrix(Row, bteColKode)
            txtJumlahKemasan(0).Text = .TextMatrix(Row, bteColjumlah)
            txtMerkKemasan(2).Text = .TextMatrix(Row, bteColMerkKemasan)
         Call ClearColGridKemasan
        ElseIf TextGrid = "D" Then
            'Call ClearColGrid("S")
        End If
        
        .TextMatrix(Row, Col) = TextGrid
        For j = 1 To .Rows - 1
            If .TextMatrix(j, bteColSelect) <> "" Then k = True
        Next j
        If k = False Then Clear_Kemasan
    End With
End Sub

Private Sub ClearColGridKemasan(Optional Kolom As String)
 Dim i As Integer
    With gridKemasan
        .Col = bteColSelect
        If Kolom <> "" Then
            For i = 1 To .Rows - 1
                If .Text = Kolom Then .Text = ""
                If .TextMatrix(i, bteColSelect) <> "D" Then .TextMatrix(i, bteColSelect) = ""
            Next i
            clear
        Else
            For i = 1 To .Rows - 1
                If .TextMatrix(i, bteColSelect) <> "" Then .TextMatrix(i, bteColSelect) = ""
            Next i
        End If
    End With
End Sub

Private Sub Clear_Kemasan()
    txtJenisKemasan.Text = ""
    txtJumlahKemasan(0).Text = ""
    txtMerkKemasan(2).Text = ""
    
    up_IsiGridKemasan
    
End Sub

Private Sub up_IsiGridKemasan()
Dim sql As String
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim li_Row As Integer

up_HeaderKemasan

    If txtNoPengajuan.Text = "______-______-________-______" Then
    
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC41Kemasan_Sel"
        
        cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, Replace(FrmBC41List.txtTampung, "-", ""))
        
        Set RS = cmd.Execute
        
        If RS.EOF = False Then
        i = 1
        With gridKemasan
            While Not RS.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(i, colSelect) = ""
                .TextMatrix(i, bteColId) = Trim(RS("ID"))
                .TextMatrix(i, bteColNo) = Trim(RS("Seri_Kemasan"))
                .TextMatrix(i, bteColjumlah) = Trim(RS("JUMLAH_KEMASAN"))
                .TextMatrix(i, bteColKode) = Trim(RS("Kode_Jenis_Kemasan"))
                .TextMatrix(i, bteColUraian) = Trim(RS("Uraian_Kemasan"))
                .TextMatrix(i, bteColMerkKemasan) = Trim(RS("Merk_Kemasan"))
                i = i + 1
            RS.MoveNext
            Wend
        End With
        End If
        
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC41SeriKemasan_Max"
        
        cmd.Parameters.append cmd.CreateParameter("Id_Header", adVarChar, adParamInput, 50, Replace(txtNoPengajuan.Text, "-", ""))
        
        Set RS = cmd.Execute
    
        If RS.EOF = False Then
            txtSeriKemasan(0).Text = IIf(IsNull(Trim(RS("SERI_KEMASAN"))) = True, "", Trim(RS("Seri_Kemasan")))
            
        Else
            txtSeriKemasan(0).Text = ""
            Exit Sub
        End If
    Else
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC41Kemasan_Sel"
        
        cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
        
        Set RS = cmd.Execute
        
        If RS.EOF = False Then
        i = 1
        With gridKemasan
            While Not RS.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(i, colSelect) = ""
                .TextMatrix(i, bteColId) = Trim(RS("ID"))
                .TextMatrix(i, bteColNo) = Trim(RS("Seri_Kemasan"))
                .TextMatrix(i, bteColjumlah) = Trim(RS("JUMLAH_KEMASAN"))
                .TextMatrix(i, bteColKode) = Trim(RS("Kode_Jenis_Kemasan"))
                .TextMatrix(i, bteColUraian) = Trim(RS("Uraian_Kemasan"))
                .TextMatrix(i, bteColMerkKemasan) = Trim(RS("Merk_Kemasan"))
                i = i + 1
            RS.MoveNext
            Wend
        End With
        End If
        
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC41SeriKemasan_Max"
        
        cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan.Text, "-", ""))
        
        Set RS = cmd.Execute
    
        If RS.EOF = False Then
            txtSeriKemasan(0).Text = IIf(IsNull(Trim(RS("SERI_KEMASAN"))) = True, "", Trim(RS("Seri_Kemasan")))
            
        Else
            txtSeriKemasan(0).Text = ""
            Exit Sub
        End If
    End If
End Sub

'** Tab Barang
Private Sub up_TabBarang()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41Barang_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("Id_Header", adVarChar, adParamInput, 50, lblNoId.Caption)
        
    Set RS = cmd.Execute
           
   If RS.EOF = False Then
        lblBrgId.Caption = Trim(RS("Id"))
        txtID.Text = RS.AbsolutePosition
        txtIdEnd.Text = RS.RecordCount
        txtBrgKode.Text = Trim(RS("KODE_BARANG"))
        txtBrgUraian.Text = Trim(RS("Uraian"))
        txtBrgMerk.Text = Trim(RS("Merk"))
        txtBrgType.Text = Trim(RS("Tipe"))
        txtBrgUkuran.Text = Trim(RS("Ukuran"))
        txtBrgSpf.Text = Trim(RS("Spesifikasi_Lain"))
        txtBrgJumlahSatuan.Text = Format(RS("Jumlah_Satuan"), gs_formatQty)
        txtBrgJenisSatuan.Text = Format(RS("KODE_SATUAN"), gs_formatQty)
        txtBrgNetto.Text = Format(RS("Netto"), gs_formatQty)
        txtBrgVolume.Text = Format(RS("Volume"), gs_formatQty)
        txtBrgHarga.Text = Format(RS("Harga_Penyerahan"), gs_formatAmountIDR)
        txtBahanBaku.Text = IIf(IsNull(RS("Jumlah_Bahan_Baku")), "", RS("Jumlah_Bahan_Baku"))
        txtJumahBrg = RS.RecordCount
    End If
    
End Sub

Private Sub up_Form()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command
    
    If lblNoId.Caption <> "" Then
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41Header_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("Nomor_Aju", adVarChar, adParamInput, 100, Replace(txtNoPengajuan.Text, "-", ""))
            
    Set RS = cmd.Execute
'
        If RS.EOF = False Then
            clear
            
            lblNoId.Caption = Trim(RS("Id"))
            txtNoPengajuan.Text = Format(RS("Nomor_Aju"), gs_formatNoAju)
            txtNoDaftar.Text = Trim(RS("Nomor_Daftar"))
            dtpTanggal.Value = Trim(RS("Tanggal_Daftar"))
            txtKantorPabean.Text = Trim(RS("Kode_Kantor"))
            'cboBrgStatus.Text = Trim(rs("Kode_Status"))
            cboJenisTPB.Text = Trim(RS("Kode_Jenis_TPB"))
            cboTujuan.Text = Trim(RS("Kode_Tujuan_Pengiriman"))
            cboNPWPPengusaha.Text = Trim(RS("Kode_Id_Pengusaha"))
            txtNPWPPengusaha.Text = Trim(RS("Id_Pengusaha"))
            txtNamaPengusaha.Text = Trim(RS("Nama_Pengusaha"))
            txtAlamatPengusaha.Text = Trim(RS("Alamat_Pengusaha"))
            txtNoIzinPengusaha.Text = Trim(RS("Nomor_Ijin_TPB"))
            cboNPWPPengirim.Text = IIf(IsNull(Trim(RS("Kode_Id_Penerima_Barang"))) = True, "", Trim(RS("Kode_Id_Penerima_Barang")))
            txtNPWPPengirim.Text = IIf(IsNull(Trim(RS("Kode_Id_Penerima_Barang"))) = True, "", Trim(RS("Kode_Id_Penerima_Barang")))
            txtNamaPengirim.Text = IIf(IsNull(Trim(RS("Nama_Penerima_Barang"))) = True, "", Trim(RS("Nama_Penerima_Barang")))
            txtAlamatPengirim.Text = IIf(IsNull(Trim(RS("Alamat_Penerima_Barang"))) = True, "", Trim(RS("Alamat_Penerima_Barang")))
            txtNamaPengangkut.Text = Trim(RS("Nama_Pengangkut"))
            txtNoPolisi.Text = Trim(RS("Nomor_Polisi"))
            txtHarga.Text = Format(RS("Harga_Penyerahan"), gs_formatAmountIDR)
            txtVolume.Text = Format(RS("Volume"), gs_formatQty)
            txtBruto.Text = Format(RS("Bruto"), gs_formatQty)
            txtNetto.Text = Format(RS("Netto"), gs_formatQty)
            txtJumahBrg.Text = Format(RS("Jumlah_Barang"), gs_formatQty)
            txtTempat.Text = Trim(RS("Kota_TTD"))
            dtpTanggal.Value = Trim(RS("Tanggal_TTD"))
            txtPemberitahu.Text = Trim(RS("Nama_TTD"))
            txtJabatan.Text = Trim(RS("Jabatan_TTD"))
            
        Else
            clear
        End If
        
        up_IsiGridDokumen
        
        up_IsiGridKemasan
        
        up_TabBarang
        
        up_DetailBahan
        
        tampil_detail
        
        
    ElseIf txtNoPengajuan.Text = "______-______-________-______" Then
    
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC41Header_Sel"
        
        cmd.Parameters.append cmd.CreateParameter("Nomor_Aju", adVarChar, adParamInput, 100, Replace(FrmBC41List.txtTampung, "-", ""))
                
        Set RS = cmd.Execute

        If RS.EOF = False Then
            clear
            
            lblNoId.Caption = Trim(RS("Id"))
            txtNoPengajuan.Text = Format(RS("Nomor_Aju"), gs_formatNoAju)
            txtNoDaftar.Text = Trim(RS("Nomor_Daftar"))
            dtpTanggal.Value = IIf(IsNull(Trim(RS("Tanggal_Daftar"))) = True, dtpTanggal.Value, Trim(RS("Tanggal_Daftar")))
            Trim (RS("Tanggal_Daftar"))
            txtKantorPabean.Text = Trim(RS("Kode_Kantor"))
            cboJenisTPB.Text = Trim(RS("Kode_Jenis_TPB"))
            cboTujuan.Text = IIf(IsNull(Trim(RS("Kode_Tujuan_Pengiriman"))) = True, "", Trim(RS("Kode_Tujuan_Pengiriman")))
            cboNPWPPengusaha.Text = Trim(RS("Kode_Id_Pengusaha"))
            txtNPWPPengusaha.Text = Trim(RS("Id_Pengusaha"))
            txtNamaPengusaha.Text = Trim(RS("Nama_Pengusaha"))
            txtAlamatPengusaha.Text = Trim(RS("Alamat_Pengusaha"))
            txtNoIzinPengusaha.Text = Trim(RS("Nomor_Ijin_TPB"))
            cboNPWPPengirim.Text = IIf(IsNull(Trim(RS("Kode_Id_Penerima_Barang"))) = True, "", Trim(RS("Kode_Id_Penerima_Barang")))
            txtNPWPPengirim.Text = IIf(IsNull(Trim(RS("Id_Penerima_Barang"))) = True, "", Trim(RS("Id_Penerima_Barang")))
            txtNamaPengirim.Text = IIf(IsNull(Trim(RS("Nama_Penerima_Barang"))) = True, "", Trim(RS("Nama_Penerima_Barang")))
            txtAlamatPengirim.Text = IIf(IsNull(Trim(RS("Alamat_Penerima_Barang"))) = True, "", Trim(RS("Alamat_Penerima_Barang")))
            txtNamaPengangkut.Text = Trim(RS("Nama_Pengangkut"))
            txtNoPolisi.Text = Trim(RS("Nomor_Polisi"))
            txtHarga.Text = Format(RS("Harga_Penyerahan"), gs_formatAmountIDR)
            txtVolume.Text = Format(RS("Volume"), gs_formatQty)
            txtBruto.Text = Format(RS("Bruto"), gs_formatQty)
            txtNetto.Text = Format(RS("Netto"), gs_formatQty)
            txtJumahBrg.Text = Format(RS("Jumlah_Barang"), gs_formatQty)
            txtTempat.Text = Trim(RS("Kota_TTD"))
            dtpTanggal.Value = IIf(IsNull(Trim(RS("Tanggal_TTD"))) = True, dtpTanggal.Value, Trim(RS("Tanggal_TTD")))
            txtPemberitahu.Text = Trim(RS("Nama_TTD"))
            txtJabatan.Text = Trim(RS("Jabatan_TTD"))
        Else
            clear
        End If
        
        up_IsiGridDokumen
        
        up_IsiGridKemasan
            
        up_TabBarang
        
        up_DetailBahan
        
        tampil_detail
    Else
      
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC41Header_Sel"
        
        cmd.Parameters.append cmd.CreateParameter("Nomor_Aju", adVarChar, adParamInput, 100, Replace(txtNoPengajuan.Text, "-", ""))
                
        Set RS = cmd.Execute
        
        If RS.EOF = False Then
            clear
            
            lblNoId.Caption = Trim(RS("Id"))
            txtNoPengajuan.Text = Format(RS("Nomor_Aju"), gs_formatNoAju)
            txtNoDaftar.Text = Trim(RS("Nomor_Daftar"))
            dtpTanggal.Value = IIf(IsNull(Trim(RS("Tanggal_Daftar"))) = True, dtpTanggal.Value, Trim(RS("Tanggal_Daftar")))
            Trim (RS("Tanggal_Daftar"))
            txtKantorPabean.Text = Trim(RS("Kode_Kantor"))
            'cboBrgStatus.Text = Trim(rs("Kode_Status"))
            cboJenisTPB.Text = Trim(RS("Kode_Jenis_TPB"))
            cboTujuan.Text = Trim(RS("Kode_Tujuan_Pengiriman"))
            cboNPWPPengusaha.Text = Trim(RS("Kode_Id_Pengusaha"))
            txtNPWPPengusaha.Text = Trim(RS("Id_Pengusaha"))
            txtNamaPengusaha.Text = Trim(RS("Nama_Pengusaha"))
            txtAlamatPengusaha.Text = Trim(RS("Alamat_Pengusaha"))
            txtNoIzinPengusaha.Text = Trim(RS("Nomor_Ijin_TPB"))
            cboNPWPPengirim.Text = Trim(RS("Kode_Id_Penerima_Barang"))
            txtNPWPPengirim.Text = Trim(RS("Id_Penerima_Barang"))
            txtNamaPengirim.Text = Trim(RS("Nama_Penerima_Barang"))
            txtAlamatPengirim.Text = Trim(RS("Alamat_Penerima_Barang"))
            txtNamaPengangkut.Text = Trim(RS("Nama_Pengangkut"))
            txtNoPolisi.Text = Trim(RS("Nomor_Polisi"))
            txtHarga.Text = Format(RS("Harga_Penyerahan"), gs_formatAmountIDR)
            txtVolume.Text = Format(RS("Volume"), gs_formatQty)
            txtBruto.Text = Format(RS("Bruto"), gs_formatQty)
            txtNetto.Text = Format(RS("Netto"), gs_formatQty)
            txtJumahBrg.Text = Format(RS("Jumlah_Barang"), gs_formatQty)
            txtTempat.Text = Trim(RS("Kota_TTD"))
            dtpTanggal.Value = IIf(IsNull(Trim(RS("Tanggal_TTD"))) = True, dtpTanggal.Value, Trim(RS("Tanggal_TTD")))
            txtPemberitahu.Text = Trim(RS("Nama_TTD"))
            txtJabatan.Text = Trim(RS("Jabatan_TTD"))
            
        Else
            clear
            
            sql = "SELECT MAX(id)+1 Id FROM Bea_Cukai_TPB_Header"
            Set RS = Db.Execute(sql)
'
            If RS.EOF = False Then
                lblNoId.Caption = IIf(IsNull(Trim(RS("Id"))) = True, 1, Trim(RS("Id")))
            End If
    
        End If
        
        up_IsiGridDokumen
        
        up_IsiGridKemasan
        
        up_TabBarang
        
        up_DetailBahan
        
        tampil_detail
        
    End If
    
    up_HargaPenyerahan
    
    up_Netto
    
    up_Volume
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BCCompanyProfile_Sel"
     
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
        txtTempat.Text = Trim(RS("City"))
        txtPemberitahu.Text = Trim(RS("SJ_Person"))
        txtJabatan.Text = Trim(RS("SJ_Position"))
        txtNPWPPengusaha.Text = Trim(RS("NPWP_No"))
        txtNamaPengusaha.Text = Trim(RS("Company_Name"))
        txtAlamatPengusaha.Text = Trim(RS("Address1")) & ", " & Trim(RS("Address2")) & ", " & Trim(RS("City")) & ", " & Trim(RS("Province"))
    End If
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41KantorPabean_Sel"
     
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
        txtKantorPabean.Text = Trim(RS("Kode_Kantor"))
    End If
End Sub

Private Sub up_HargaPenyerahan()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41Harga"
    
    cmd.Parameters.append cmd.CreateParameter("Id_Header", adVarChar, adParamInput, 50, lblNoId.Caption)
     
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
        txtHarga.Text = Format(RS("Harga_Penyerahan"), gs_formatAmountIDR)
    Else
        txtHarga.Text = IIf(IsNull(Trim(RS("Harga_Penyerahan"))) = True, 0, Trim(RS("Harga_Penyerahan")))
    End If
    
End Sub

Private Sub up_Netto()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41Netto"
    
    cmd.Parameters.append cmd.CreateParameter("Id_Header", adVarChar, adParamInput, 50, lblNoId.Caption)
     
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
        txtNetto.Text = Format(RS("Netto"), gs_formatQty)
        txtBruto.Text = Format(RS("Netto"), gs_formatQty)
    Else
        txtNetto.Text = IIf(IsNull(Trim(RS("Netto"))) = True, 0, Trim(RS("Netto")))
        txtBruto.Text = IIf(IsNull(Trim(RS("Netto"))) = True, 0, Trim(RS("Netto")))
    End If
    
End Sub

Private Sub up_Volume()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41Volume"
    
    cmd.Parameters.append cmd.CreateParameter("Id_Header", adVarChar, adParamInput, 50, lblNoId.Caption)
     
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
        txtVolume.Text = Format(RS("Volume"), gs_formatQty)
    Else
        txtVolume.Text = IIf(IsNull(Trim(RS("Volume"))) = True, 0, Trim(RS("Volume")))
    End If
    
End Sub

Private Sub up_DetailBahan()
Dim sql As String
Dim RS As New Recordset

If txtNoPengajuan = "______-______-________-______" Then
      sql = "Select * from Bea_Cukai_TPB_Header WHERE NO_PENGAJUAN='" & Replace(FrmBC41List.txtTampung, "-", "") & "'"
            If RS.State <> adStateClosed Then RS.Close
            RS.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        
            If RS.EOF = False Then
                tampung = Trim(RS("Id"))
            End If
    Else
      sql = "Select * from Bea_Cukai_TPB_Header WHERE NO_PENGAJUAN='" & Replace(txtNoPengajuan.Text, "-", "") & "'"
            If RS.State <> adStateClosed Then RS.Close
            RS.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        
            If RS.EOF = False Then
                tampung = Trim(RS("Id"))
            End If
    End If
    
  sql = " SELECT Id, Harga_Penyerahan, Jenis_Satuan, Jumlah_Satuan, Kode_Barang, Kode_Jenis_Dok_Asal," & vbCrLf & _
        " Kode_Kantor, Nomor_Aju_Dok_Asal, Nomor_Daftar_Dok_Asal, Seri_Bahan_Baku, Seri_Barang, " & vbCrLf & _
        " Tanggal_Daftar_Dok_Asal, Tipe, Merk, Ukuran, Spesifikasi_Lain, Uraian, Id_Barang, Id_Header FROM Bea_Cukai_TPB_Bahan_Baku Where ID_HEADER='" & tampung & "'"
  If rs_detail.State <> adStateClosed Then rs_detail.Close
  rs_detail.Open sql, Db, adOpenKeyset, adLockOptimistic

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    LblErrMsg.Caption = ""
End Sub

Private Sub cboNPWPPengirim_Change()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command
    
    sql = "SELECT * from Bea_Cukai_Kode_Id Where Kode_Id='" & cboNPWPPengirim.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
         cboNPWPPengirim.Text = Trim(RS(1)) & " - " & IIf(IsNull(RS(2)), "", Trim(RS(2)))
    End If
End Sub

Private Sub cboNPWPPengusaha_Change()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command
    
    sql = "SELECT * from Bea_Cukai_Kode_Id Where Kode_Id='" & cboNPWPPengusaha.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
         cboNPWPPengusaha.Text = Trim(RS(1)) & " - " & IIf(IsNull(RS(2)), "", Trim(RS(2)))
    End If
End Sub

Private Sub cboJenisTPB_Change()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command
    
    sql = "SELECT * FROM Bea_Cukai_Referensi_Jenis_TPB Where Kode_Jenis_TPB='" & cboJenisTPB.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
         cboJenisTPB.Text = Trim(RS(1)) & " - " & IIf(IsNull(RS(2)), "", Trim(RS(2)))
    End If
End Sub

Private Sub cboTujuan_Change()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command
    
    sql = "SELECT * FROM Bea_Cukai_Referensi_Tujuan_Pengiriman Where Kode_Tujuan_Pengiriman='" & Left(cboTujuan.Text, 1) & "' and Kode_Dokumen='" & 41 & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
         cboTujuan.Text = Trim(RS(2)) & " - " & IIf(IsNull(RS(3)), "", Trim(RS(3)))
    End If
End Sub

Private Sub txtBrgJenisSatuan_Change()
Dim sql As String
Dim RS As New Recordset
    
    sql = "Select Uraian_Satuan From Bea_Cukai_Referensi_Satuan Where Kode_Satuan =  '" & txtBrgJenisSatuan.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lblBrgJenis(27).Caption = Trim(RS("Uraian_Satuan"))
    Else
        lblBrgJenis(27).Caption = ""
        Exit Sub
    End If
End Sub

Private Sub txtDokAsal_Change()
    LblErrMsg.Caption = ""
    
    sql = "Select Uraian_Dokumen From Bea_Cukai_Dokumen Where Kode_Dokumen =  '" & txtDokAsal.Text & "'"
        Set RS = Db.Execute(sql)
        
        If RS.EOF = False Then
            lblDokAsal(65).Caption = Trim(RS("Uraian_Dokumen"))
        Else
            lblDokAsal(65).Caption = ""
            Exit Sub
        End If
End Sub

Private Sub txtJenisKemasan_Change()
Dim sql As String
Dim RS As New Recordset

sql = "Select Uraian_Kemasan From Bea_Cukai_Kemasan Where Kode_Kemasan =  '" & txtJenisKemasan.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lblKemasan(0).Caption = Trim(RS("Uraian_Kemasan"))
    Else
        lblKemasan(0).Caption = ""
        'Exit Sub
    End If
End Sub

Private Sub txtPabean_Change()
Dim sql As String
Dim RS As New Recordset
    
'    sql = "SELECT Nama_Kantor FROM Bea_Cukai_Kantor_pabean Where Kode_Kantor like  '%" & txtKantorPabean.Text & "%'"
'    Set rs = Db.Execute(sql)
'
'    If rs.EOF = False Then
'        lblPabean(30).Caption = Trim(rs("Nama_Kantor"))
'    Else
'        lblPabean(30).Caption = ""
'        Exit Sub
'    End If
End Sub

Private Sub txtDokumen_Change()
Dim sql As String
Dim RS As New Recordset

LblErrMsg.Caption = ""
    
    sql = "Select Uraian_Dokumen From Bea_Cukai_Dokumen Where Kode_Dokumen =  '" & txtDokumen.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lblDokumen(23).Caption = Trim(RS("Uraian_Dokumen"))
    Else
        lblDokumen(23).Caption = ""
        Exit Sub
    End If
End Sub

Private Sub txtJenisSatuan_Change()
Dim sql As String
Dim RS As New Recordset
    
    sql = "Select Uraian_Satuan From Bea_Cukai_Referensi_Satuan Where Kode_Satuan =  '" & txtJenisSatuan.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lblBahanJenis(0).Caption = Trim(RS("Uraian_Satuan"))
    Else
        lblBahanJenis(0).Caption = ""
        Exit Sub
    End If
End Sub

Private Sub txtKantorPabean_Change()
     up_LoadKantorPabean txtKantorPabean
End Sub

Private Sub txtKPPBC_Change()
     up_LoadKPPBC txtKPPBC
End Sub

Private Sub txtSeriKemasan_Change(Index As Integer)
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC41SeriKemasan_Max"
    
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan.Text, "-", ""))
    
    Set RS = cmd.Execute
    If RS.EOF = False Then
        txtSeriKemasan(0).Text = IIf(IsNull(Trim(RS("SERI_KEMASAN"))) = True, "", Trim(RS("SERI_KEMASAN")))
    Else
        txtSeriKemasan(0).Text = ""
        Exit Sub
    End If
End Sub

Private Sub txtKantorPabean_LostFocus()
    up_LoadKantorPabean txtKantorPabean
End Sub

Private Sub txtBrgHarga_LostFocus()
    txtBrgHarga.Text = Format(txtBrgHarga.Text, gs_formatAmountIDR)
End Sub

Private Sub txtBrgJumlahSatuan_LostFocus()
    txtBrgJumlahSatuan.Text = Format(txtBrgJumlahSatuan.Text, gs_formatQty)
End Sub

Private Sub txtBrgNetto_LostFocus()
    txtBrgNetto.Text = Format(txtBrgNetto.Text, gs_formatQty)
End Sub

Private Sub txtBrgVolume_LostFocus()
    txtBrgVolume.Text = Format(txtBrgVolume.Text, gs_formatQty)
End Sub

Private Sub txtBruto_LostFocus()
    txtBruto.Text = Format(txtBruto.Text, gs_formatQty)
End Sub

Private Sub txtNetto_LostFocus()
    txtNetto.Text = Format(txtNetto.Text, gs_formatQty)
End Sub

Private Sub txtVolume_LostFocus()
    txtVolume.Text = Format(txtVolume.Text, gs_formatQty)
End Sub

Private Sub txtBrgHarga_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBrgJenisSatuan_KeyPress(KeyAscii As Integer)
Dim Index As Integer

If KeyAscii = Asc("'") Then KeyAscii = 0
If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtBrgJumlahSatuan_KeyPress(KeyAscii As Integer)
      If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBrgNetto_KeyPress(KeyAscii As Integer)
      If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBrgVolume_KeyPress(KeyAscii As Integer)
      If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBruto_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        KeyAscii = 0
        
    End If
End Sub

Private Sub txtHarga_KeyPress(KeyAscii As Integer)
      If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtJenisKemasan_KeyPress(KeyAscii As Integer)
Dim Index As Integer

If KeyAscii = Asc("'") Then KeyAscii = 0
If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then SendKeys vbTab
End Sub


Private Sub txtJumahBrg_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNetto_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNoPengajuan_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNPWPPengirim_KeyPress(KeyAscii As Integer)
       If Left(cboNPWPPengirim, 1) = 1 Then
        txtNPWPPengirim.MaxLength = 20
    Else
        txtNPWPPengirim.MaxLength = 24
    End If
End Sub

Private Sub txtNPWPPengusaha_KeyPress(KeyAscii As Integer)
    If Left(cboNPWPPengusaha, 1) = 1 Then
        If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
            KeyAscii = 0
        End If
        txtNPWPPengusaha.MaxLength = 20
    Else
        If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
            KeyAscii = 0
        End If
        txtNPWPPengusaha.MaxLength = 24
    End If
End Sub

Private Sub txtJenisSatuan_KeyPress(KeyAscii As Integer)
    Dim Index As Integer
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub GridDokumen_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If gridDokumen.Col = bteColSelect Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then
          KeyAscii = 0
       End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
    End If
End Sub

Private Sub GridKemasan_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If gridKemasan.Col = bteColSelect Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then
          KeyAscii = 0
       End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
    End If
End Sub

Private Sub txtPabean_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtVolume_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub

Private Sub KoneksiMysql()
Dim ConnString As String
Dim db_name As String
Dim db_server As String
Dim db_port As String
Dim db_user As String
Dim db_pass As String
Dim Conn As New ADODB.Connection
'//error traping
On Error GoTo buat_koneksi_Error
'/isi variable
db_name = "tpbdb"
db_server = "localhost"
db_port = "3306"
db_user = "beacukai"
db_pass = "beacukai"
'/buat connection string
ConnStr = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_user & ";PWD=" & db_pass & ";PORT=" & db_port & ""
'/buka koneksi
With Conn
    .ConnectionString = ConnStr
    .Open
   'MsgBox "Koneksi Berhasil"
End With
'___________________________________________________________
On Error GoTo 0
Exit Sub

buat_koneksi_Error:
    MsgBox "Ada kesalahan dengan server, periksa apakah server sudah berjalan !", vbInformation, "Cek Server"
End Sub


