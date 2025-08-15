VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmBC40Detail 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BC 40 Detail"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15045
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBC40Detail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11415
   ScaleMode       =   0  'User
   ScaleWidth      =   15045
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdSubmit 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "FFTT*/"
      Top             =   10440
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_SubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "TFFT*/"
      Top             =   10440
      Width           =   1125
   End
   Begin VB.CommandButton CmdSyncronize 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Send to CEISA"
      Height          =   375
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "FFTT*/"
      Top             =   10440
      Width           =   1605
   End
   Begin VB.Frame Frame2 
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
      Height          =   600
      Left            =   240
      TabIndex        =   37
      Tag             =   "TFTT*/"
      Top             =   9720
      Width           =   14490
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
         TabIndex        =   38
         Tag             =   "TFTF*/"
         Top             =   180
         Width           =   14325
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Tag             =   "TFTF*/"
      Top             =   720
      Width           =   14565
      Begin VB.TextBox txtBCNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   12960
         TabIndex        =   61
         Tag             =   "TTFF*/"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtBCType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   12960
         TabIndex        =   59
         Tag             =   "TTFF*/"
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txtKantorPabean 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Tag             =   "TTFF*/"
         Top             =   1600
         Width           =   1335
      End
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
         Height          =   375
         Index           =   3
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "TTFF*/"
         Top             =   840
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtNoDaftar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   24
         Tag             =   "TTFF*/"
         Top             =   800
         Width           =   1815
      End
      Begin VB.TextBox txtTempat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7560
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   780
         Width           =   2415
      End
      Begin VB.TextBox txtPemberitahu 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7560
         TabIndex        =   7
         Tag             =   "TTFF*/"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtJabatan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7560
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   1605
         Width           =   2415
      End
      Begin MSMask.MaskEdBox txtNoPengajuan 
         Height          =   315
         Left            =   1560
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
         Left            =   1560
         TabIndex        =   26
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
         Format          =   138149891
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   315
         Left            =   10080
         TabIndex        =   6
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
         Format          =   138149891
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DTPBCDate 
         Height          =   315
         Left            =   12960
         TabIndex        =   57
         Tag             =   "TTFF*/"
         Top             =   360
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
         Format          =   138149891
         CurrentDate     =   37798
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BC No"
         Height          =   195
         Index           =   7
         Left            =   11880
         TabIndex        =   60
         Tag             =   "TTFF*/"
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BC Type"
         Height          =   195
         Index           =   6
         Left            =   11880
         TabIndex        =   58
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BC Date"
         Height          =   195
         Index           =   5
         Left            =   11880
         TabIndex        =   56
         Tag             =   "TTFF*/"
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis TPB"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Tag             =   "TTFF*/"
         Top             =   2070
         Width           =   1695
      End
      Begin MSForms.ComboBox cboJenisTPB 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Kantor Pabean"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Tag             =   "TTFF*/"
         Top             =   1630
         Width           =   1455
      End
      Begin VB.Label lblKantorPabean 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3000
         TabIndex        =   35
         Tag             =   "TTFF*/"
         Top             =   1635
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Pengajuan"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Daftar"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Daftar"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Tag             =   "TTFF*/"
         Top             =   1260
         Width           =   1275
      End
      Begin MSForms.ComboBox cboTujuan 
         Height          =   315
         Left            =   7560
         TabIndex        =   4
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
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tujuan Pengiriman"
         Height          =   255
         Left            =   5880
         TabIndex        =   30
         Tag             =   "TTFF*/"
         Top             =   390
         Width           =   1695
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3000
         X2              =   5640
         Y1              =   1905
         Y2              =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat/Tanggal"
         Height          =   195
         Index           =   2
         Left            =   5880
         TabIndex        =   29
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pemberitahu"
         Height          =   195
         Index           =   3
         Left            =   5880
         TabIndex        =   28
         Tag             =   "TTFF*/"
         Top             =   1260
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         Height          =   195
         Index           =   4
         Left            =   5880
         TabIndex        =   27
         Tag             =   "TTFF*/"
         Top             =   1665
         Width           =   660
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2055
      Left            =   240
      TabIndex        =   39
      Tag             =   "TFTF*/"
      Top             =   3360
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   3625
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   6
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
      TabPicture(0)   =   "FrmBC40Detail.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Pengirim Barang"
      TabPicture(1)   =   "FrmBC40Detail.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   " Pengangkutan"
      TabPicture(2)   =   "FrmBC40Detail.frx":0E7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Harga"
      TabPicture(3)   =   "FrmBC40Detail.frx":0E96
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Pemilik"
      TabPicture(4)   =   "FrmBC40Detail.frx":0EB2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame10"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         TabIndex        =   154
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14295
         Begin VB.TextBox txtNamaPemilik 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   157
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   4935
         End
         Begin VB.TextBox txtAlamatPemilik 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   156
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   4935
         End
         Begin VB.TextBox txtNPWPPemilik 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4320
            TabIndex        =   155
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Identitas"
            Height          =   255
            Left            =   240
            TabIndex        =   161
            Tag             =   "TTFF*/"
            Top             =   390
            Width           =   1575
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            Height          =   255
            Left            =   240
            TabIndex        =   160
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            Height          =   255
            Left            =   240
            TabIndex        =   159
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   1575
         End
         Begin MSForms.ComboBox cboNPWPPemilik 
            Height          =   315
            Left            =   2040
            TabIndex        =   158
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
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         TabIndex        =   54
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14295
         Begin VB.TextBox txtHarga 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3000
            TabIndex        =   20
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         TabIndex        =   49
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14295
         Begin VB.TextBox txtNoPolisi 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3000
            TabIndex        =   19
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtNamaPengangkut 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3000
            TabIndex        =   18
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Nomor Polisi"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Saran Pengangkut Darat"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         TabIndex        =   45
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14175
         Begin VB.TextBox txtNPWPPengirim 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4320
            TabIndex        =   15
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtAlamatPengirim 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   17
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   4935
         End
         Begin VB.TextBox txtNamaPengirim 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   16
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   4935
         End
         Begin MSForms.ComboBox cboNPWPPengirim 
            Height          =   315
            Left            =   2040
            TabIndex        =   14
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
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Identitas"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   40
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14280
         Begin VB.TextBox txtNPWPPengusaha 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4440
            MaxLength       =   20
            TabIndex        =   10
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtNamaPengusaha 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   11
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   4935
         End
         Begin VB.TextBox txtNoIzinPengusaha 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   12
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   4935
         End
         Begin VB.TextBox txtAlamatPengusaha 
            Appearance      =   0  'Flat
            Height          =   1035
            Left            =   8400
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   13
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   5415
         End
         Begin MSForms.ComboBox cboNPWPPengusaha 
            Height          =   315
            Left            =   2160
            TabIndex        =   9
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
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Tag             =   "TTFF*/"
            Top             =   750
            Width           =   975
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "No Izin"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Tag             =   "TTFF*/"
            Top             =   1110
            Width           =   975
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            Height          =   255
            Left            =   7320
            TabIndex        =   41
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   975
         End
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12960
      TabIndex        =   62
      TabStop         =   0   'False
      Tag             =   "FTTF*/"
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin EZRunnerv3.Anchor Anchor1 
      Left            =   2880
      Top             =   10320
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3825
      Left            =   240
      TabIndex        =   63
      TabStop         =   0   'False
      Tag             =   "TTTF*/"
      Top             =   5520
      Width           =   14565
      _ExtentX        =   25691
      _ExtentY        =   6747
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
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
      TabPicture(0)   =   "FrmBC40Detail.frx":0ECE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3(19)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3(20)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3(21)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3(22)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDokumen(23)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Shape1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3(15)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label3(17)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label3(18)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "DTPSkep"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "DTPFakturPajak"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DTPKontrak"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "DTPPackingList"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "GridDokumen"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "DTPDokumen"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtDokumen"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtNoDokumen(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtPackingList(2)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtKontrak(3)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtFakturPajak(4)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtSkep(5)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdSubmitDokumen"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Kemasan"
      TabPicture(1)   =   "FrmBC40Detail.frx":0EEA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape10"
      Tab(1).Control(1)=   "Label3(1)"
      Tab(1).Control(2)=   "Shape13"
      Tab(1).Control(3)=   "Label3(24)"
      Tab(1).Control(4)=   "Line1(3)"
      Tab(1).Control(5)=   "Shape5"
      Tab(1).Control(6)=   "Label3(25)"
      Tab(1).Control(7)=   "Label3(23)"
      Tab(1).Control(8)=   "Label3(27)"
      Tab(1).Control(9)=   "lblKemasan(0)"
      Tab(1).Control(10)=   "GridKemasan"
      Tab(1).Control(11)=   "txtJumlahKemasan(0)"
      Tab(1).Control(12)=   "txtJenisKemasan"
      Tab(1).Control(13)=   "txtMerkKemasan(2)"
      Tab(1).Control(14)=   "cmSubimtKemasan"
      Tab(1).Control(15)=   "txtSeriKemasan(0)"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Barang"
      TabPicture(2)   =   "FrmBC40Detail.frx":0F06
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command1(4)"
      Tab(2).Control(1)=   "Command1(3)"
      Tab(2).Control(2)=   "Command1(2)"
      Tab(2).Control(3)=   "Command1(1)"
      Tab(2).Control(4)=   "Frame3"
      Tab(2).Control(5)=   "Frame4"
      Tab(2).Control(6)=   "Frame7"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Respon"
      TabPicture(3)   =   "FrmBC40Detail.frx":0F22
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label29"
      Tab(3).Control(1)=   "Label30"
      Tab(3).Control(2)=   "gridStatus"
      Tab(3).Control(3)=   "gridRespon"
      Tab(3).ControlCount=   4
      Begin VB.TextBox txtSeriKemasan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   -64575
         TabIndex        =   119
         Tag             =   "FFFF*/"
         Top             =   735
         Width           =   3645
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Last Page"
         Height          =   375
         Index           =   4
         Left            =   -61680
         Style           =   1  'Graphical
         TabIndex        =   118
         Tag             =   "FFFF*/"
         Top             =   3000
         Width           =   1020
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Next Page"
         Height          =   375
         Index           =   3
         Left            =   -62760
         Style           =   1  'Graphical
         TabIndex        =   117
         Tag             =   "FFFF*/"
         Top             =   3000
         Width           =   1020
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Prev Page"
         Height          =   375
         Index           =   2
         Left            =   -63840
         Style           =   1  'Graphical
         TabIndex        =   116
         Tag             =   "FFFF*/"
         Top             =   3000
         Width           =   1020
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "First Page"
         Height          =   375
         Index           =   1
         Left            =   -64920
         Style           =   1  'Graphical
         TabIndex        =   115
         Tag             =   "FFFF*/"
         Top             =   3000
         Width           =   1020
      End
      Begin VB.CommandButton cmSubimtKemasan 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Submit"
         Height          =   375
         Left            =   -65820
         Style           =   1  'Graphical
         TabIndex        =   114
         Tag             =   "FFFF*/"
         Top             =   3315
         Width           =   1125
      End
      Begin VB.CommandButton cmdSubmitDokumen 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Submit"
         Height          =   375
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   113
         Tag             =   "TTFF*/"
         Top             =   3120
         Width           =   1125
      End
      Begin VB.TextBox txtMerkKemasan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   -68970
         TabIndex        =   112
         Tag             =   "FFFF*/"
         Top             =   3225
         Width           =   2745
      End
      Begin VB.TextBox txtJenisKemasan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73320
         TabIndex        =   111
         Tag             =   "FFFF*/"
         Top             =   3240
         Width           =   1170
      End
      Begin VB.TextBox txtSkep 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   10635
         TabIndex        =   110
         Tag             =   "TFFF*/"
         Top             =   1860
         Width           =   1935
      End
      Begin VB.TextBox txtFakturPajak 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   10635
         TabIndex        =   109
         Tag             =   "TFFF*/"
         Top             =   1485
         Width           =   1935
      End
      Begin VB.TextBox txtKontrak 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   10635
         TabIndex        =   108
         Tag             =   "TFFF*/"
         Top             =   1110
         Width           =   1935
      End
      Begin VB.TextBox txtPackingList 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   10635
         TabIndex        =   107
         Tag             =   "TFFF*/"
         Top             =   735
         Width           =   1935
      End
      Begin VB.TextBox txtNoDokumen 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   4290
         TabIndex        =   106
         Tag             =   "TFFF*/"
         Top             =   3120
         Width           =   2935
      End
      Begin VB.TextBox txtDokumen 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         TabIndex        =   105
         Tag             =   "TFFF*/"
         Top             =   3120
         Width           =   750
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74880
         TabIndex        =   86
         Tag             =   "FFFF*/"
         Top             =   360
         Width           =   9960
         Begin VB.TextBox txtId 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1530
            TabIndex        =   95
            Tag             =   "TFFF*/"
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox txtIdEnd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2850
            TabIndex        =   94
            Tag             =   "TFFF*/"
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox txtBrgKode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1530
            TabIndex        =   93
            Tag             =   "TFFF*/"
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtBrgUraian 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1530
            TabIndex        =   92
            Tag             =   "TFFF*/"
            Top             =   1155
            Width           =   5445
         End
         Begin VB.TextBox txtBrgMerk 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4440
            TabIndex        =   91
            Tag             =   "TFFF*/"
            Top             =   285
            Width           =   2535
         End
         Begin VB.TextBox txtBrgType 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4440
            TabIndex        =   90
            Tag             =   "TFFF*/"
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtBrgSpf 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8220
            TabIndex        =   89
            Tag             =   "TFFF*/"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtBrgUkuran 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8220
            TabIndex        =   88
            Tag             =   "TFFF*/"
            Top             =   285
            Width           =   1575
         End
         Begin VB.TextBox txtHSCode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8220
            TabIndex        =   87
            Tag             =   "TFFF*/"
            Top             =   1155
            Width           =   1575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Barang"
            Height          =   195
            Index           =   30
            Left            =   150
            TabIndex        =   104
            Tag             =   "TTFF*/"
            Top             =   345
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "dari"
            Height          =   195
            Index           =   31
            Left            =   2370
            TabIndex        =   103
            Tag             =   "TTFF*/"
            Top             =   345
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang"
            Height          =   195
            Index           =   32
            Left            =   150
            TabIndex        =   102
            Tag             =   "TTFF*/"
            Top             =   765
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uraian Barang"
            Height          =   195
            Index           =   33
            Left            =   150
            TabIndex        =   101
            Tag             =   "TTFF*/"
            Top             =   1170
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Merk"
            Height          =   195
            Index           =   39
            Left            =   3840
            TabIndex        =   100
            Tag             =   "TTFF*/"
            Top             =   345
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipe"
            Height          =   195
            Index           =   40
            Left            =   3840
            TabIndex        =   99
            Tag             =   "TTFF*/"
            Top             =   765
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SPF Lain"
            Height          =   195
            Index           =   41
            Left            =   7320
            TabIndex        =   98
            Tag             =   "TTFF*/"
            Top             =   780
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ukuran"
            Height          =   195
            Index           =   26
            Left            =   7320
            TabIndex        =   97
            Tag             =   "TTFF*/"
            Top             =   345
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HS Code"
            Height          =   195
            Index           =   2
            Left            =   7320
            TabIndex        =   96
            Tag             =   "TTFF*/"
            Top             =   1200
            Width           =   750
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74880
         TabIndex        =   74
         Tag             =   "FFFF*/"
         Top             =   2160
         Width           =   9840
         Begin VB.TextBox txtBrgJenisSatuan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1530
            TabIndex        =   79
            Tag             =   "TFFF*/"
            Top             =   720
            Width           =   705
         End
         Begin VB.TextBox txtBrgJumlahSatuan 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1530
            TabIndex        =   78
            Tag             =   "TFFF*/"
            Text            =   "0.00"
            Top             =   285
            Width           =   1485
         End
         Begin VB.TextBox txtBrgNetto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4560
            TabIndex        =   77
            Tag             =   "TFFF*/"
            Text            =   "0.00"
            Top             =   285
            Width           =   1365
         End
         Begin VB.TextBox txtBrgVolume 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4560
            TabIndex        =   76
            Tag             =   "TFFF*/"
            Text            =   "0.00"
            Top             =   720
            Width           =   1365
         End
         Begin VB.TextBox txtBrgHarga 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8160
            TabIndex        =   75
            Tag             =   "TFFF*/"
            Text            =   "0.00"
            Top             =   285
            Width           =   1605
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Satuan"
            Height          =   195
            Index           =   34
            Left            =   150
            TabIndex        =   85
            Tag             =   "TTFF*/"
            Top             =   345
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Satuan"
            Height          =   195
            Index           =   35
            Left            =   150
            TabIndex        =   84
            Tag             =   "TTFF*/"
            Top             =   780
            Width           =   1080
         End
         Begin VB.Line Line3 
            X1              =   2280
            X2              =   3060
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Netto (Kgm)"
            Height          =   195
            Index           =   36
            Left            =   3285
            TabIndex        =   83
            Tag             =   "TTFF*/"
            Top             =   345
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Volume(m3)"
            Height          =   195
            Index           =   37
            Left            =   3285
            TabIndex        =   82
            Tag             =   "TTFF*/"
            Top             =   780
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Penyerahan Rp"
            Height          =   195
            Index           =   38
            Left            =   6120
            TabIndex        =   81
            Tag             =   "TTFF*/"
            Top             =   345
            Width           =   1875
         End
         Begin VB.Label lblBrgJenis 
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   27
            Left            =   2355
            TabIndex        =   80
            Tag             =   "TTFF*/"
            Top             =   780
            Width           =   660
         End
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -64800
         TabIndex        =   65
         Tag             =   "FFFF*/"
         Top             =   360
         Width           =   4005
         Begin VB.TextBox txtBruto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1785
            TabIndex        =   69
            Tag             =   "TFFF*/"
            Text            =   "0.00"
            Top             =   555
            Width           =   1575
         End
         Begin VB.TextBox txtVolume 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1785
            TabIndex        =   68
            Tag             =   "TFFF*/"
            Text            =   "0.00"
            Top             =   195
            Width           =   1575
         End
         Begin VB.TextBox txtNetto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1785
            TabIndex        =   67
            Tag             =   "TFFF*/"
            Text            =   "0.00"
            Top             =   915
            Width           =   1575
         End
         Begin VB.TextBox txtJumahBrg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1785
            TabIndex        =   66
            Tag             =   "TFFF*/"
            Text            =   "0.00"
            Top             =   1275
            Width           =   1575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Berat Bersih (Kg)"
            Height          =   195
            Index           =   51
            Left            =   165
            TabIndex        =   73
            Tag             =   "TTFF*/"
            Top             =   975
            Width           =   1500
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Berat Kotor (Kg)"
            Height          =   195
            Index           =   50
            Left            =   165
            TabIndex        =   72
            Tag             =   "TTFF*/"
            Top             =   615
            Width           =   1425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Volume (m3)"
            Height          =   195
            Index           =   49
            Left            =   165
            TabIndex        =   71
            Tag             =   "TTFF*/"
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Barang"
            Height          =   195
            Index           =   52
            Left            =   165
            TabIndex        =   70
            Tag             =   "TTFF*/"
            Top             =   1335
            Width           =   1575
         End
      End
      Begin VB.TextBox txtJumlahKemasan 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   -74640
         TabIndex        =   64
         Tag             =   "FFFF*/"
         Top             =   3225
         Width           =   1170
      End
      Begin MSComCtl2.DTPicker DTPDokumen 
         Height          =   345
         Left            =   7320
         TabIndex        =   120
         Tag             =   "TFFF*/"
         Top             =   3120
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
         Format          =   151846915
         CurrentDate     =   37798
      End
      Begin VSFlex8Ctl.VSFlexGrid grid 
         Height          =   2535
         Left            =   -74910
         TabIndex        =   121
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
            TabIndex        =   122
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
            Format          =   151846915
            CurrentDate     =   37798
         End
         Begin MSForms.ComboBox cbocurr 
            Height          =   285
            Left            =   5640
            TabIndex        =   124
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
            TabIndex        =   123
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
         Height          =   2055
         Left            =   240
         TabIndex        =   125
         TabStop         =   0   'False
         Tag             =   "TFFF*/"
         Top             =   480
         Width           =   8730
         _cx             =   15399
         _cy             =   3625
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
            TabIndex        =   126
            Tag             =   "FFTF*/"
            Top             =   0
            Width           =   2175
         End
      End
      Begin MSComCtl2.DTPicker DTPPackingList 
         Height          =   345
         Left            =   12630
         TabIndex        =   127
         Tag             =   "TTFF*/"
         Top             =   720
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
         Format          =   151846915
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DTPKontrak 
         Height          =   345
         Left            =   12630
         TabIndex        =   128
         Tag             =   "TTFF*/"
         Top             =   1095
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
         Format          =   151846915
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DTPFakturPajak 
         Height          =   345
         Left            =   12630
         TabIndex        =   129
         Tag             =   "TTFF*/"
         Top             =   1470
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
         Format          =   151846915
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DTPSkep 
         Height          =   345
         Left            =   12630
         TabIndex        =   130
         Tag             =   "TTFF*/"
         Top             =   1845
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
         Format          =   151846915
         CurrentDate     =   37798
      End
      Begin VSFlex8Ctl.VSFlexGrid gridRespon 
         Height          =   2880
         Left            =   -74760
         TabIndex        =   131
         TabStop         =   0   'False
         Tag             =   "TTFT*/"
         Top             =   795
         Width           =   6945
         _cx             =   12250
         _cy             =   5080
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
         Height          =   2880
         Left            =   -67620
         TabIndex        =   132
         TabStop         =   0   'False
         Tag             =   "TTTT*/"
         Top             =   795
         Width           =   6945
         _cx             =   12250
         _cy             =   5080
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
      Begin VSFlex8Ctl.VSFlexGrid GridKemasan 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   133
         TabStop         =   0   'False
         Tag             =   "FFFF*/"
         Top             =   480
         Width           =   8730
         _cx             =   15399
         _cy             =   3625
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
      Begin VB.Label lblKemasan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   -72000
         TabIndex        =   149
         Tag             =   "FFFF*/"
         Top             =   3240
         Width           =   2925
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   195
         Index           =   18
         Left            =   7335
         TabIndex        =   141
         Tag             =   "TTFF*/"
         Top             =   2700
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Dokumen"
         Height          =   195
         Index           =   17
         Left            =   4290
         TabIndex        =   140
         Tag             =   "TTFF*/"
         Top             =   2700
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   15
         Left            =   375
         TabIndex        =   139
         Tag             =   "FFFF*/"
         Top             =   2700
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Dokumen"
         Height          =   195
         Index           =   0
         Left            =   1230
         TabIndex        =   135
         Tag             =   "FFFF*/"
         Top             =   2700
         Width           =   1305
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00A6D2FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   240
         Tag             =   "TFFF*/"
         Top             =   2640
         Width           =   8730
      End
      Begin VB.Label lblBrgId 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -67140
         TabIndex        =   153
         Tag             =   "TTFF*/"
         Top             =   2745
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kemasan"
         Height          =   195
         Index           =   27
         Left            =   -65625
         TabIndex        =   152
         Tag             =   "FFFF*/"
         Top             =   795
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
         TabIndex        =   151
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
         TabIndex        =   150
         Tag             =   "TTFF*/"
         Top             =   2400
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Merk"
         Height          =   195
         Index           =   23
         Left            =   -69000
         TabIndex        =   148
         Tag             =   "FFFF*/"
         Top             =   2820
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   25
         Left            =   -73320
         TabIndex        =   147
         Tag             =   "FFFF*/"
         Top             =   2820
         Width           =   435
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00808080&
         Height          =   540
         Left            =   -74760
         Tag             =   "FFFF*/"
         Top             =   3120
         Width           =   8730
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
         Height          =   195
         Index           =   23
         Left            =   1200
         TabIndex        =   146
         Tag             =   "TTFF*/"
         Top             =   3200
         Width           =   2940
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SKEP"
         Height          =   195
         Index           =   22
         Left            =   9360
         TabIndex        =   145
         Tag             =   "TTFF*/"
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faktur Pajak"
         Height          =   195
         Index           =   21
         Left            =   9360
         TabIndex        =   144
         Tag             =   "TTFF*/"
         Top             =   1545
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kontrak"
         Height          =   195
         Index           =   20
         Left            =   9360
         TabIndex        =   143
         Tag             =   "TTFF*/"
         Top             =   1170
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing List"
         Height          =   195
         Index           =   19
         Left            =   9375
         TabIndex        =   142
         Tag             =   "TTFF*/"
         Top             =   795
         Width           =   1005
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   -72360
         X2              =   -69840
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1200
         X2              =   4120
         Y1              =   3400
         Y2              =   3400
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   -72000
         X2              =   -69090
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Respon"
         Height          =   195
         Left            =   -74760
         TabIndex        =   138
         Tag             =   "TTFF*/"
         Top             =   510
         Width           =   630
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Left            =   -67620
         TabIndex        =   137
         Tag             =   "TTFF*/"
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         Height          =   195
         Index           =   24
         Left            =   -74640
         TabIndex        =   136
         Tag             =   "FFFF*/"
         Top             =   2820
         Width           =   600
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H00808080&
         Height          =   2145
         Left            =   -65820
         Tag             =   "FFFF*/"
         Top             =   570
         Width           =   5070
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00808080&
         Height          =   2085
         Left            =   9120
         Tag             =   "TFFF*/"
         Top             =   480
         Width           =   5130
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uraian"
         Height          =   195
         Index           =   1
         Left            =   -72000
         TabIndex        =   134
         Tag             =   "FFFF*/"
         Top             =   2820
         Width           =   555
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   540
         Left            =   240
         Tag             =   "TFFF*/"
         Top             =   3000
         Width           =   8730
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00A6D2FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   -74760
         Tag             =   "FFFF*/"
         Top             =   2760
         Width           =   8730
      End
   End
   Begin VB.Label lblNoId 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Tag             =   "TTFF*/"
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BC 40 Detail"
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
      Height          =   495
      Left            =   240
      TabIndex        =   34
      Tag             =   "TTTF*/"
      Top             =   120
      Width           =   14535
   End
End
Attribute VB_Name = "FrmBC40Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MySQLCon As New ADODB.Connection

Dim rs_barang As New ADODB.Recordset
Dim k_pertama As Boolean

Dim SuratJalan As String, NoPengajuan As String, SupplierCode As String
Dim bteColSelect As Byte, bteColNo As Byte, bteColjumlah As Byte, bteColKode As Byte, bteColUraian As Byte
Dim bteColMerkKemasan As Byte, bteColId As Byte

Const colSelect As Integer = 0
Const colId As Integer = 1
Const colSeri As Integer = 2
Const colKodeDokumen As Integer = 3
Const colJenisDokumen As Integer = 4
Const colNomor As Integer = 5
Const colTanggal As Integer = 6
Const colcount As Integer = 7

Const colKodeStatus As Integer = 0
Const colUraianStatus As Integer = 1
Const colWaktuStatus As Integer = 2
Const colCountStatus As Integer = 3

Const colKodeRespon As Integer = 0
Const colUraianRespon As Integer = 1
Const colWaktuRespon As Integer = 2
Const colCountRespon As Integer = 3

Private Sub clear()
    LblErrMsg.Caption = ""
    
    dtpTanggal = Format(Now, "dd MMM yyyy")
    DTPDokumen = Format(Now, "dd MMM yyyy")
    DTPFakturPajak = Format(Now, "dd MMM yyyy")
    DTPPackingList = Format(Now, "dd MMM yyyy")
    DTPSkep = Format(Now, "dd MMM yyyy")
    DTPKontrak = Format(Now, "dd MMM yyyy")
    
    txtNoDaftar.Text = ""
    txtNPWPPengusaha.Text = ""
    txtNamaPengusaha.Text = ""
    txtAlamatPengusaha.Text = ""
    txtNoIzinPengusaha.Text = ""
    txtNPWPPengirim.Text = ""
    txtNamaPengirim.Text = ""
    txtAlamatPengirim.Text = ""
    txtNPWPPemilik.Text = ""
    txtNamaPemilik.Text = ""
    txtAlamatPemilik.Text = ""
    txtNamaPengangkut.Text = ""
    
    'Tab Barang
    txtId.Text = ""
    txtIdEnd.Text = ""
    txtBrgKode.Text = ""
    txtBrgUraian.Text = ""
    txtBrgMerk.Text = ""
    txtBrgType.Text = ""
    txtBrgUkuran.Text = ""
    txtBrgSpf.Text = ""
    txtBrgJumlahSatuan.Text = ""
    txtBrgJenisSatuan.Text = ""
    lblBrgJenis(27).Caption = ""
    txtBrgNetto.Text = ""
    txtBrgVolume.Text = ""
    txtBrgHarga.Text = ""
End Sub

Private Sub data_tampil()
    If rs_barang.EOF = False Then
        lblBrgId.Caption = Trim(rs_barang("Id"))
        txtId.Text = rs_barang.AbsolutePosition
        txtIdEnd.Text = rs_barang.RecordCount
        txtBrgKode.Text = Trim(rs_barang("KODE_BARANG"))
        txtBrgUraian.Text = Trim(rs_barang("Uraian"))
        txtBrgMerk.Text = Trim(rs_barang("Merk"))
        txtBrgType.Text = Trim(rs_barang("Tipe"))
        txtBrgUkuran.Text = Trim(rs_barang("Ukuran"))
        txtBrgSpf.Text = Trim(rs_barang("Spesifikasi_Lain"))
        txtBrgJumlahSatuan.Text = Format(rs_barang("Jumlah_Satuan"), gs_formatQty)
        txtBrgJenisSatuan.Text = IIf(IsNull(Trim(rs_barang("Kode_Satuan"))), "", Trim(rs_barang("Kode_Satuan")))
        txtBrgNetto.Text = IIf(IsNull(Trim(rs_barang("Netto"))) = False, "0.00", Trim(rs_barang("Netto")))
        txtBrgVolume.Text = IIf(IsNull(Trim(rs_barang("Volume"))) = False, "0.00", Trim(rs_barang("Volume")))
        txtBrgHarga.Text = Format(rs_barang("Harga_Penyerahan"), gs_formatAmountIDR)
        txtJumahBrg.Text = rs_barang.RecordCount
        txtHSCode.Text = Trim(rs_barang("HS_Code"))
    End If
End Sub

Private Sub up_FillComboTujuan()
Dim sql As String
Dim RS As New Recordset

    sql = "Select * From Bea_Cukai_Tujuan"
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
            .List(i, 0) = Trim(RS(0)) & " - " & IIf(IsNull(RS(1)), "", Trim(RS(1)))
            
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
    cmd.CommandText = "sp_BC40NPWP_Sel"
    
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
    cmd.CommandText = "sp_BC40NPWP_Sel"
    
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
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC40NPWP_Sel"
    
    Set RS = cmd.Execute
    
    With cboNPWPPemilik
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

Private Sub cboNPWPPemilik_Change()
Dim sql As String
    Dim RS As New Recordset
    Dim cmd As ADODB.Command
    
    sql = "SELECT * from Bea_Cukai_Kode_Id Where Kode_Id='" & cboNPWPPemilik.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
         cboNPWPPemilik.Text = Trim(RS(1)) & " - " & IIf(IsNull(RS(2)), "", Trim(RS(2)))
    End If
End Sub

Private Sub Cmd_SubMenu_Click()
    Unload Me
    FrmBC40List.Show
End Sub

Private Sub CmdSubmit_Click()
    LblErrMsg.Caption = ""
    Me.MousePointer = vbHourglass
    
    Insert_Update
    up_Form
    
    Me.MousePointer = vbDefault
    LblErrMsg = DisplayMsg(1000)
End Sub

Private Sub Insert_Update()
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
    cmd.CommandText = "sp_BC40Header_InsertUpdate"
    
    cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 100, lblNoId.Caption)
    cmd.Parameters.append cmd.CreateParameter("Nomor_Aju", adVarChar, adParamInput, 100, NoPengajuan)
    cmd.Parameters.append cmd.CreateParameter("Nomor_Daftar", adVarChar, adParamInput, 50, txtNoDaftar)
    cmd.Parameters.append cmd.CreateParameter("Tanggal_Daftar", adDBTime, adParamInput, , dtpTanggal.Value)
    cmd.Parameters.append cmd.CreateParameter("Tanggal_Aju", adDBTime, adParamInput, , dtpTanggal.Value)
    cmd.Parameters.append cmd.CreateParameter("Kode_Kantor", adVarChar, adParamInput, 50, txtKantorPabean.Text)
    cmd.Parameters.append cmd.CreateParameter("Kode_Dokumen_Pabean", adVarChar, adParamInput, 50, Left(txtNoPengajuan.Text, 6))
    cmd.Parameters.append cmd.CreateParameter("Id_Modul", adVarChar, adParamInput, 50, Mid$(txtNoPengajuan.Text, 9, 5))
    cmd.Parameters.append cmd.CreateParameter("Kode_Jenis_TPB", adVarChar, adParamInput, 50, Left(cboJenisTPB.Text, 1))
    cmd.Parameters.append cmd.CreateParameter("Kode_Tujuan_Pengiriman", adVarChar, adParamInput, 50, Left(cboTujuan.Text, 1))
    cmd.Parameters.append cmd.CreateParameter("Kode_Id_Pengusaha", adVarChar, adParamInput, 50, Left(cboNPWPPengusaha.Text, 1))
    cmd.Parameters.append cmd.CreateParameter("Id_Pengusaha", adVarChar, adParamInput, 50, txtNPWPPengusaha.Text)
    cmd.Parameters.append cmd.CreateParameter("Nama_Pengusaha", adVarChar, adParamInput, 100, txtNamaPengusaha.Text)
    cmd.Parameters.append cmd.CreateParameter("Alamat_Pengusaha", adVarChar, adParamInput, 200, txtAlamatPengusaha.Text)
    cmd.Parameters.append cmd.CreateParameter("Nomor_Ijin_TPB", adVarChar, adParamInput, 150, txtNoIzinPengusaha.Text)
    cmd.Parameters.append cmd.CreateParameter("Kode_Id_Pengirim", adVarChar, adParamInput, 50, Left(cboNPWPPengirim.Text, 1))
    cmd.Parameters.append cmd.CreateParameter("Id_Pengirim", adVarChar, adParamInput, 50, txtNPWPPengirim.Text)
    cmd.Parameters.append cmd.CreateParameter("Nama_Pengirim", adVarChar, adParamInput, 50, txtNamaPengirim.Text)
    cmd.Parameters.append cmd.CreateParameter("Alamat_Pengirim", adVarChar, adParamInput, 200, txtAlamatPengirim.Text)
    cmd.Parameters.append cmd.CreateParameter("Nama_Pengangkut", adVarChar, adParamInput, 100, txtNamaPengangkut.Text)
    cmd.Parameters.append cmd.CreateParameter("Nomor_Polisi", adVarChar, adParamInput, 50, txtNoPolisi.Text)
    cmd.Parameters.append cmd.CreateParameter("Kode_Id_Pemilik", adVarChar, adParamInput, 50, Left(cboNPWPPemilik.Text, 1))
    cmd.Parameters.append cmd.CreateParameter("Id_Pemilik", adVarChar, adParamInput, 50, txtNPWPPemilik.Text)
    cmd.Parameters.append cmd.CreateParameter("Nama_Pemilik", adVarChar, adParamInput, 50, txtNamaPemilik.Text)
    cmd.Parameters.append cmd.CreateParameter("Alamat_Pemilik", adVarChar, adParamInput, 200, txtAlamatPemilik.Text)
    
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
    cmd.Parameters.append cmd.CreateParameter("NO_PENGAJUAN", adVarChar, adParamInput, 50, NoPengajuan)
    cmd.Parameters.append cmd.CreateParameter("SURATJALAN_NO", adVarChar, adParamInput, 25, SuratJalan)
    
    Set RS = cmd.Execute
End Sub

Private Sub cmdSubmitDokumen_Click()
Dim strSQL As String
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim id As String
Dim tanya

LblErrMsg.Caption = ""

    With GridDokumen
        For i = 1 To .Rows - 1
            If .TextMatrix(i, colSelect) = "D" Then
                If IsEmpty(tanya) Then tanya = MsgBox("Do you really want to delete this data ?", vbQuestion & vbYesNo, "Confirmation")
                If tanya = vbYes Then
                
                Set cmd = New ADODB.Command
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 0
                cmd.ActiveConnection = Db
                cmd.CommandText = "sp_BC40Dokumen_Del"
                
                cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 100, GridDokumen.TextMatrix(i, colId))
                
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
    cmd.CommandText = "sp_BC40Dokumen_InsertUpdate"
    
     With GridDokumen
      For i = 1 To .Rows - 1
      If .TextMatrix(i, colSelect) = "S" Then
        lblIdDokumen.Caption = GridDokumen.TextMatrix(i, colId)
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

    With GridDokumen
      For i = 1 To .Rows - 1
      If .TextMatrix(i, colSelect) = "S" Then
        lblIdDokumen.Caption = GridDokumen.TextMatrix(i, colId)
      End If
    Next i
    End With
    
    cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 100, lblIdDokumen.Caption)
    cmd.Parameters.append cmd.CreateParameter("Kode_Jenis_Dokumen", adVarChar, adParamInput, 100, txtDokumen.Text)
    cmd.Parameters.append cmd.CreateParameter("Nomor_Daftar", adVarChar, adParamInput, 50, txtNoDokumen(1).Text)
    cmd.Parameters.append cmd.CreateParameter("Tanggal_Dokumen", adDBTime, adParamInput, , DTPDokumen.Value)
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, NoPengajuan)
    
    Set RS = cmd.Execute
    
    Clear_dokumen
    
    LblErrMsg = DisplayMsg(1000)
End Sub

Private Sub up_IsiGridDokumen()
    Dim sql As String
    Dim cmd As ADODB.Command
    Dim rsDokumen As ADODB.Recordset
    Dim li_Row As Integer
 
    up_HeaderDokumen
        
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC40Dokumen_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 100, NoPengajuan)
    
    Set rsDokumen = cmd.Execute
            
    i = 1
    With GridDokumen
        While Not rsDokumen.EOF
            If (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 388 Then
                txtFakturPajak(4).Text = Trim(rsDokumen("Nomor_Dokumen"))
                DTPFakturPajak.Value = Format(rsDokumen("Tanggal_Dokumen"), "DD-MMM-YYYY")
            ElseIf (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 217 Then
                txtPackingList(2).Text = Trim(rsDokumen("Nomor_Dokumen"))
                DTPPackingList.Value = Format(rsDokumen("Tanggal_Dokumen"), "DD-MMM-YYYY")
            ElseIf (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 315 Then
                txtKontrak(3).Text = Trim(rsDokumen("Nomor_Dokumen"))
                DTPKontrak.Value = Format(rsDokumen("Tanggal_Dokumen"), "DD-MMM-YYYY")
            ElseIf (Trim(rsDokumen("Kode_Jenis_Dokumen"))) = 912 Then
                txtSkep(5).Text = Trim(rsDokumen("Nomor_Dokumen"))
                DTPSkep.Value = Format(rsDokumen("Tanggal_Dokumen"), "DD-MMM-YYYY")
            Else
                .Rows = .Rows + 1
                
                .TextMatrix(i, colSelect) = ""
                .TextMatrix(i, colId) = Trim(rsDokumen("ID"))
                .TextMatrix(i, colSeri) = Trim(rsDokumen("Seri_Dokumen"))
                .TextMatrix(i, colKodeDokumen) = Trim(rsDokumen("Kode_Jenis_Dokumen"))
                .TextMatrix(i, colJenisDokumen) = Trim(rsDokumen("Jenis_Dokumen"))
                .TextMatrix(i, colNomor) = Trim(rsDokumen("Nomor_Dokumen"))
                .TextMatrix(i, colTanggal) = Format(rsDokumen("Tanggal_Dokumen"), "DD-MMM-YYYY")
            
                i = i + 1
            End If
                        
            rsDokumen.MoveNext
        Wend
    End With
End Sub

Private Sub cmdSyncronize_Click()
    Dim strSQL As String
    
    Dim RS As ADODB.Recordset
    Dim rsDokumen As ADODB.Recordset
    Dim rsKemasan As ADODB.Recordset
    Dim rsBarang As ADODB.Recordset
    Dim RSiD As New ADODB.Recordset
    
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim prm1 As ADODB.Parameter
    Dim prm2 As ADODB.Parameter
    Dim prm3 As ADODB.Parameter
    Dim prm4 As ADODB.Parameter
    
    Dim tanya
    Dim versi_modul As String
    Dim asal_data As String
    
    Set RS = New ADODB.Recordset

    versi_modul = "3.1.8"
    asal_data = "I"
    
    If IsEmpty(tanya) Then tanya = MsgBox("Do you want to send this data to CEISA ?", vbQuestion & vbYesNo, "Confirmation")
    If tanya = vbYes Then
        Me.MousePointer = vbHourglass
                            
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_UpdateStatus_Interface"
                       
        cmd.Parameters.append cmd.CreateParameter("SuratJalan", adVarChar, adParamInput, 50, SuratJalan)
        
        Set RS = cmd.Execute
        
        KoneksiMysql
        
        strSQL = "Select * from tpbdb.tpb_header where Nomor_Aju='" & NoPengajuan & "'"
            
        If RSiD.State <> adStateClosed Then RSiD.Close
        RSiD.Open strSQL, MySQLCon, adOpenForwardOnly, adLockReadOnly, adCmdText
        If RSiD.EOF = False Then
        
            strSQL = " Update tpbdb.tpb_header " & vbCrLf & _
                      " set Nomor_Aju='" & NoPengajuan & "', Tanggal_Daftar='" & Format(Trim(dtpTanggal.Value), "yyyy-mm-dd") & "', TANGGAL_AJU='" & Format(Trim(dtpTanggal.Value), "yyyy-mm-dd") & "', " & vbCrLf & _
                      " Kode_Kantor='" & txtKantorPabean.Text & "', Id_Modul=  '" & Mid$(txtNoPengajuan.Text, 8, 5) & "',  Kode_Jenis_TPB='" & Left(cboJenisTPB.Text, 1) & "', Kode_Tujuan_Pengiriman='" & Left(cboTujuan.Text, 1) & "', " & vbCrLf & _
                      " Kode_Id_Pengusaha='" & Left(cboNPWPPengusaha.Text, 1) & "', ID_PENGUSAHA='" & txtNPWPPengusaha.Text & "', Nama_Pengusaha='" & txtNamaPengusaha.Text & "', Alamat_Pengusaha='" & txtAlamatPengusaha.Text & "', " & vbCrLf & _
                      " Nomor_Ijin_TPB='" & txtNoIzinPengusaha.Text & "', Kode_Id_Pengirim='" & Left(cboNPWPPengirim.Text, 1) & "', Id_Pengirim='" & txtNPWPPengirim.Text & "', Nama_Pengirim ='" & txtNamaPengirim.Text & "'," & vbCrLf & _
                      " Alamat_Pengirim='" & txtAlamatPengirim.Text & "', Kode_Id_Pemilik ='" & Left(cboNPWPPemilik.Text, 1) & "', Harga_Penyerahan ='" & CDec(txtHarga.Text) & "', VOLUME='" & CDbl(txtVolume.Text) & "', " & vbCrLf & _
                      " BRUTO='" & CDbl(txtBruto.Text) & "', Id_Pemilik='" & txtNPWPPemilik.Text & "', Nama_Pemilik ='" & txtNamaPemilik.Text & "', Alamat_Pemilik='" & txtAlamatPemilik.Text & "', Nama_Pengangkut='" & txtNamaPengangkut.Text & "'," & vbCrLf & _
                      " Nomor_Polisi='" & txtNoPolisi.Text & "', NETTO='" & CDbl(txtNetto.Text) & "', JUMLAH_BARANG='" & txtJumahBrg.Text & "', Kota_TTD='" & txtTempat.Text & "', Tanggal_TTD ='" & Format(Trim(dtpTanggal.Value), "yyyy-mm-dd") & "'," & vbCrLf & _
                      " Nama_TTD='" & txtPemberitahu.Text & "', Jabatan_TTD='" & txtJabatan.Text & "', asal_data='" & asal_data & "', versi_modul='" & versi_modul & "' where Id='" & Trim(RSiD("Id")) & "'"
                         
            RS.Open strSQL, MySQLCon
                                                
            strSQL = " Delete from tpbdb.tpb_dokumen Where Id_Header='" & Trim(RSiD("Id")) & "' "
            RS.Open strSQL, MySQLCon
    
            Set cmd = New ADODB.Command
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 0
            cmd.ActiveConnection = Db
            cmd.CommandText = "sp_BC40Dokumen_Sel"
                
            cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 100, NoPengajuan)
                
            Set rsDokumen = cmd.Execute
                
            Do While Not rsDokumen.EOF
                strSQL = " Insert into tpbdb.tpb_dokumen " & vbCrLf & _
                         " (Kode_Jenis_Dokumen, Nomor_Dokumen, Seri_Dokumen, Tanggal_Dokumen, Tipe_Dokumen, Id_Header) " & vbCrLf & _
                         " values ('" & Trim(rsDokumen("Kode_Jenis_Dokumen")) & "', '" & Trim(rsDokumen("Nomor_Dokumen")) & "', '" & Trim(rsDokumen("Seri_Dokumen")) & "', " & vbCrLf & _
                         " '" & Format(rsDokumen("Tanggal_Dokumen"), "yyyy-mm-dd") & "', '" & Trim(rsDokumen("Tipe_Dokumen")) & "','" & Trim(RSiD("Id")) & "')"
                RS.Open strSQL, MySQLCon
                                    
                rsDokumen.MoveNext
            Loop
                
            'Insert Kemasan
            strSQL = " Delete from tpbdb.tpb_kemasan Where Id_Header='" & Trim(RSiD("Id")) & "' "
            RS.Open strSQL, MySQLCon
               
            Set cmd = New ADODB.Command
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 0
            cmd.ActiveConnection = Db
            cmd.CommandText = "sp_BC40Kemasan_Sel"
             
            cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 100, NoPengajuan)
    
            Set rsKemasan = cmd.Execute
                
            Do While Not rsKemasan.EOF
                strSQL = " Insert into tpbdb.tpb_kemasan " & vbCrLf & _
                         " (Jumlah_Kemasan, Kode_Jenis_Kemasan, Merk_Kemasan, Seri_Kemasan, Id_Header) " & vbCrLf & _
                         " Values('" & Trim(rsKemasan("Jumlah_Kemasan")) & "','" & Trim(rsKemasan("Kode_Jenis_Kemasan")) & "','" & Trim(rsKemasan("Merk_Kemasan")) & "', " & vbCrLf & _
                         " '" & Trim(rsKemasan("Seri_Kemasan")) & "', '" & Trim(RSiD("Id")) & "' )"
                RS.Open strSQL, MySQLCon
                rsKemasan.MoveNext
            Loop
                
            'Insert Barang
            strSQL = " Delete from tpbdb.tpb_barang Where Id_Header='" & Trim(RSiD("Id")) & "' "
            RS.Open strSQL, MySQLCon
            
            Set cmd = New ADODB.Command
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 0
            cmd.ActiveConnection = Db
            cmd.CommandText = "sp_BC40Barang_Sel"
             
            cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, NoPengajuan)
    
            Set rsBarang = cmd.Execute
            
            Do While Not rsBarang.EOF
                strSQL = " Insert into tpbdb.tpb_barang " & vbCrLf & _
                         " (Harga_Penyerahan, Jumlah_Satuan, Kode_Barang, Kode_Satuan, Netto, POS_Tarif, Seri_Barang, Uraian, Volume, Id_Header) " & vbCrLf & _
                         " Values('" & Trim(rsBarang("Harga_Penyerahan")) & "','" & Trim(rsBarang("Jumlah_Satuan")) & "', '" & Trim(rsBarang("Kode_Barang")) & "' , " & vbCrLf & _
                         " '" & Trim(rsBarang("Kode_Satuan")) & "', '" & Trim(rsBarang("Netto")) & "', '" & Trim(rsBarang("HS_Code")) & "', '" & Trim(rsBarang("Seri_Barang")) & "', '" & Trim(rsBarang("Uraian")) & "', " & vbCrLf & _
                         " '" & Trim(rsBarang("Volume")) & "', '" & Trim(RSiD("Id")) & "' ) "
                         
                RS.Open strSQL, MySQLCon
                rsBarang.MoveNext
            Loop
            
        Else

            strSQL = " INSERT INTO tpbdb.tpb_header  " & vbCrLf & _
                    " (Nomor_Aju, TANGGAL_AJU, Kode_Kantor,  " & vbCrLf & _
                    " KODE_DOKUMEN_PABEAN, ID_MODUL, Kode_Jenis_TPB,  " & vbCrLf & _
                    " Kode_Tujuan_Pengiriman, Kode_Id_Pengusaha, ID_PENGUSAHA,  " & vbCrLf & _
                    " Nama_Pengusaha, Alamat_Pengusaha, Nomor_Ijin_TPB,  " & vbCrLf & _
                    " Kode_Id_Pengirim, Id_Pengirim, Nama_Pengirim, Alamat_Pengirim, Kode_Id_Pemilik, " & vbCrLf & _
                    " Id_Pemilik, Nama_Pemilik, Alamat_Pemilik, Nama_Pengangkut, Nomor_Polisi,  " & vbCrLf & _
                    " Harga_Penyerahan, VOLUME, BRUTO,  " & vbCrLf & _
                    " NETTO, JUMLAH_BARANG, Kota_TTD,  " & vbCrLf & _
                    " Tanggal_TTD, Nama_TTD, Jabatan_TTD,  " & vbCrLf & _
                    " Versi_Modul, Asal_data)  "

            strSQL = strSQL + " VALUES " & vbCrLf & _
                    " ('" & NoPengajuan & "', '" & Format(Trim(dtpTanggal.Value), "yyyy-mm-dd") & "', '" & txtKantorPabean.Text & "',  " & vbCrLf & _
                    " '" & Mid$(NoPengajuan, 5, 2) & "', '" & Mid$(NoPengajuan, 8, 5) & "' , '" & Left(cboJenisTPB.Text, 1) & "',  " & vbCrLf & _
                    " '" & Left(cboTujuan.Text, 1) & "', '" & Left(cboNPWPPengusaha.Text, 1) & "', '" & txtNPWPPengusaha.Text & "',  " & vbCrLf & _
                    " '" & txtNamaPengusaha.Text & "', '" & txtAlamatPengusaha.Text & "', '" & txtNoIzinPengusaha.Text & "',  " & vbCrLf & _
                    " '" & Left(cboNPWPPengirim.Text, 1) & "', '" & txtNPWPPengirim.Text & "', '" & txtNamaPengirim.Text & "',  " & vbCrLf & _
                    " '" & txtAlamatPengirim.Text & "', '" & Left(cboNPWPPemilik.Text, 1) & "', '" & txtNPWPPemilik.Text & "',  " & vbCrLf & _
                    " '" & txtAlamatPemilik.Text & "','" & txtNamaPengirim.Text & "',   '" & txtNamaPengangkut.Text & "', '" & txtNoPolisi.Text & "', " & vbCrLf & _
                    " '" & CDec(IIf(txtHarga.Text = "", 0, txtHarga.Text)) & "', '" & CDbl(txtVolume.Text) & "', '" & CDbl(txtBruto.Text) & "',  " & vbCrLf & _
                    " '" & CDbl(txtNetto.Text) & "', '" & CDbl(txtJumahBrg.Text) & "', '" & txtTempat.Text & "',  " & vbCrLf & _
                    "  '" & Format(Trim(dtpTanggal.Value), "yyyy-mm-dd") & "', '" & txtPemberitahu.Text & "', '" & txtJabatan.Text & "',  " & vbCrLf & _
                    "  '" & versi_modul & "', '" & asal_data & "') "
                  
            RS.Open strSQL, MySQLCon
            
            strSQL = "Select * from tpbdb.tpb_header where Nomor_Aju='" & NoPengajuan & "'"
                
            'Insert Dokumen
            If RSiD.State <> adStateClosed Then RSiD.Close
            RSiD.Open strSQL, MySQLCon, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            Set cmd = New ADODB.Command
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 0
            cmd.ActiveConnection = Db
            cmd.CommandText = "sp_BC40Dokumen_Sel"
            
            cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 100, NoPengajuan)
            
            Set rsDokumen = cmd.Execute
            
            Do While Not rsDokumen.EOF
                strSQL = " Insert into tpbdb.tpb_dokumen " & vbCrLf & _
                         " (Kode_Jenis_Dokumen, Nomor_Dokumen, Seri_Dokumen, Tanggal_Dokumen, Tipe_Dokumen, Id_Header) " & vbCrLf & _
                         " values ('" & Trim(rsDokumen("Kode_Jenis_Dokumen")) & "', '" & Trim(rsDokumen("Nomor_Dokumen")) & "', '" & Trim(rsDokumen("Seri_Dokumen")) & "', " & vbCrLf & _
                         " '" & Format(rsDokumen("Tanggal_Dokumen"), "yyyy-mm-dd") & "', '" & Trim(rsDokumen("Tipe_Dokumen")) & "','" & Trim(RSiD("Id")) & "')"
                      
                RS.Open strSQL, MySQLCon
                rsDokumen.MoveNext
            Loop
                
            'Insert Kemasan
            Set cmd = New ADODB.Command
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 0
            cmd.ActiveConnection = Db
            cmd.CommandText = "sp_BC40Kemasan_Sel"
             
            cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 100, NoPengajuan)

            Set rsKemasan = cmd.Execute
            
            Do While Not rsKemasan.EOF
                strSQL = " Insert into tpbdb.tpb_kemasan " & vbCrLf & _
                         " (Jumlah_Kemasan, Kode_Jenis_Kemasan, Merk_Kemasan, Seri_Kemasan, Id_Header) " & vbCrLf & _
                         " Values('" & Trim(rsKemasan("Jumlah_Kemasan")) & "','" & Trim(rsKemasan("Kode_Jenis_Kemasan")) & "','" & Trim(rsKemasan("Merk_Kemasan")) & "', " & vbCrLf & _
                         " '" & Trim(rsKemasan("Seri_Kemasan")) & "', '" & Trim(RSiD("Id")) & "' )"
                RS.Open strSQL, MySQLCon
                rsKemasan.MoveNext
            Loop
                
            'Insert Barang
            Set cmd = New ADODB.Command
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 0
            cmd.ActiveConnection = Db
            cmd.CommandText = "sp_BC40Barang_Sel"
             
            cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, NoPengajuan)
    
            Set rsBarang = cmd.Execute
            
            Do While Not rsBarang.EOF
                strSQL = " Insert into tpbdb.tpb_barang " & vbCrLf & _
                         " (Harga_Penyerahan, Jumlah_Satuan, Kode_Barang, Kode_Satuan, Netto, Seri_Barang, Uraian, Volume, Id_Header) " & vbCrLf & _
                         " Values('" & Trim(rsBarang("Harga_Penyerahan")) & "','" & Trim(rsBarang("Jumlah_Satuan")) & "', '" & Trim(rsBarang("Kode_Barang")) & "' , " & vbCrLf & _
                         " '" & Trim(rsBarang("Kode_Satuan")) & "', '" & Trim(rsBarang("Netto")) & "', '" & Trim(rsBarang("Seri_Barang")) & "', '" & Trim(rsBarang("Uraian")) & "',  " & vbCrLf & _
                         " '" & Trim(rsBarang("Volume")) & "', '" & Trim(RSiD("Id")) & "' ) "
                         
                RS.Open strSQL, MySQLCon
                rsBarang.MoveNext
            Loop
                
        End If
        
        Insert_Update
        
        MySQLCon.Close
        Set MySQLCon = Nothing
        
        Me.MousePointer = vbDefault
        LblErrMsg = DisplayMsg(1000)
    End If
End Sub

Private Sub cmSubimtKemasan_Click()
    Dim strSQL As String
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim tanya

    LblErrMsg.Caption = ""
    
    With GridKemasan
        For i = 1 To .Rows - 1
            If .TextMatrix(i, colSelect) = "D" Then
                If IsEmpty(tanya) Then tanya = MsgBox("Do you really want to delete this data ?", vbQuestion & vbYesNo, "Confirmation")
                If tanya = vbYes Then
                
                    Set cmd = New ADODB.Command
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandTimeout = 0
                    cmd.ActiveConnection = Db
                    cmd.CommandText = "sp_BC40Kemasan_Del"
                    
                    cmd.Parameters.append cmd.CreateParameter("Id", adVarChar, adParamInput, 100, GridKemasan.TextMatrix(i, bteColId))
                    
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
    cmd.CommandText = "sp_BC40Kemasan_InsertUpdate"
    
    With GridKemasan
      For i = 1 To .Rows - 1
      If .TextMatrix(i, colSelect) = "S" Then
        lblIdKemasan.Caption = GridKemasan.TextMatrix(i, bteColId)
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
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, NoPengajuan)
    
    Set RS = cmd.Execute
    
    Clear_Kemasan
    
    LblErrMsg = DisplayMsg(1000)
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
    Case 1:
        If rs_barang.EOF = False Or rs_barang.BOF = False Then
            rs_barang.MoveFirst
            Call data_tampil
            LblErrMsg.Caption = DisplayMsg("4020")
        End If
    Case 2:
        If rs_barang.EOF = False Or rs_barang.BOF = False Then
            rs_barang.MovePrevious: LblErrMsg.Caption = ""
            If rs_barang.BOF Then rs_barang.MoveFirst: LblErrMsg.Caption = DisplayMsg("4020")
            Call data_tampil
            If rs_barang.AbsolutePosition = 1 Then LblErrMsg.Caption = DisplayMsg("4020")
        End If
    Case 3:
        If k_pertama = True Then
            If rs_barang.EOF = False Or rs_barang.BOF = False Then
                rs_barang.MoveFirst
                Call data_tampil
                LblErrMsg.Caption = DisplayMsg("4020")
                k_pertama = False
            End If
        Else
            If rs_barang.EOF = False Or rs_barang.BOF = False Then
                rs_barang.MoveNext: LblErrMsg.Caption = ""
                If rs_barang.EOF Then rs_barang.MoveLast: LblErrMsg.Caption = DisplayMsg("4021")
                Call data_tampil
                If rs_barang.AbsolutePosition = rs_barang.RecordCount Then LblErrMsg.Caption = DisplayMsg("4021")
            End If
        End If
    Case 4:
        If rs_barang.EOF = False Or rs_barang.BOF = False Then
            rs_barang.MoveLast
            Call data_tampil
            LblErrMsg.Caption = DisplayMsg("4021")
        End If
    End Select
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
  up_Form
     
  data_tampil
     
  CtrlMenu1.FormName = Me.Name
  Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
  
  With Anchor1
    .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
    .DoInit
  End With
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

Sub up_HeaderDokumen()
    With GridDokumen
        .ColS = colcount
        .Rows = 1
        
        .TextMatrix(0, colSelect) = ""
        .TextMatrix(0, colId) = "ID"
        .TextMatrix(0, colSeri) = "Seri"
        .TextMatrix(0, colKodeDokumen) = "Kode"
        .TextMatrix(0, colJenisDokumen) = "Jenis Dokumen"
        .TextMatrix(0, colNomor) = "Nomor"
        .TextMatrix(0, colTanggal) = "Tanggal"
        
        .ColWidth(colSelect) = 300
        .ColWidth(colSeri) = 600
        .ColWidth(colKodeDokumen) = 800
        .ColWidth(colJenisDokumen) = 2700
        .ColWidth(colNomor) = 2700
        .ColWidth(colTanggal) = 1200
        
        .ColHidden(colId) = True
        .ColAlignment(colSelect) = flexAlignCenterCenter
        .ColAlignment(colSeri) = flexAlignCenterCenter
        .ColAlignment(colKodeDokumen) = flexAlignCenterCenter
        .ColAlignment(colJenisDokumen) = flexAlignLeftCenter
        .ColAlignment(colNomor) = flexAlignLeftCenter
        .ColAlignment(colTanggal) = flexAlignCenterCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
    End With
End Sub


Private Sub GridDokumen_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If GridDokumen.Col > colId Then Cancel = True
End Sub

Private Sub GridDokumen_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim TextGrid As String
    Dim k As Boolean
    Dim j As Integer
    
   k = False
    With GridDokumen
        TextGrid = GridDokumen.Text
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
    With GridDokumen
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

Sub up_HeaderKemasan()
    bteColSelect = 0
    bteColId = 1
    bteColNo = 2
    bteColjumlah = 3
    bteColKode = 4
    bteColUraian = 5
    bteColMerkKemasan = 6
   
    With GridKemasan
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
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColNo) = 500
        .ColWidth(bteColjumlah) = 1000
        .ColWidth(bteColKode) = 2000
        .ColWidth(bteColUraian) = 2000
        .ColWidth(bteColMerkKemasan) = 2500
        
        .ColHidden(bteColId) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, 6) = flexAlignCenterCenter
    End With
End Sub

Private Sub GridKemasan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If GridKemasan.Col > bteColId Then Cancel = True
End Sub

Private Sub GridKemasan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim TextGrid As String
    Dim k As Boolean
    Dim j As Integer
    
   k = False
    With GridKemasan
        TextGrid = GridKemasan.Text
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
    With GridKemasan
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
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC40Kemasan_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, NoPengajuan)
    
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
    
    i = 1
    With GridKemasan
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
    cmd.CommandText = "sp_BC40SeriKemasan_Max"
    
    cmd.Parameters.append cmd.CreateParameter("Id_Header", adVarChar, adParamInput, 50, NoPengajuan)
    
    Set RS = cmd.Execute

    If RS.EOF = False Then
        txtSeriKemasan(0).Text = IIf(IsNull(Trim(RS("SERI_KEMASAN"))) = True, "", Trim(RS("Seri_Kemasan")))
        
    Else
        txtSeriKemasan(0).Text = ""
    End If
End Sub

Private Sub up_TabBarang()
    Dim strSQL As String

    strSQL = "EXEC sp_BC40Barang_Sel '" & NoPengajuan & "'"
    
    If rs_barang.State <> adStateClosed Then rs_barang.Close
    rs_barang.CursorLocation = adUseClient
    rs_barang.Open strSQL, Db, adOpenDynamic, adLockOptimistic
           
    data_tampil
End Sub

Private Sub up_Form()
    Dim sql As String
    Dim RS As New Recordset
    Dim cmd As ADODB.Command
    
    SuratJalan = FrmBC40List.SuratJalanNo
    SupplierCode = FrmBC40List.SupplierCode
    NoPengajuan = FrmBC40List.NoPengajuan
    
    txtNoPengajuan.Text = Format(NoPengajuan, gs_formatNoAju)
    
    up_PartReceipt
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC40Header_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("Nomor_Aju", adVarChar, adParamInput, 100, NoPengajuan)
            
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
        clear
        
        lblNoId.Caption = Trim(RS("Id"))
        txtNoDaftar.Text = Trim(RS("Nomor_Daftar"))
        dtpTglDaftar.Value = Trim(RS("Tanggal_Daftar"))
        txtKantorPabean.Text = Trim(RS("Kode_Kantor"))
        cboJenisTPB.Text = Trim(RS("Kode_Jenis_TPB"))
        cboTujuan.Text = Trim(RS("Kode_Tujuan_Pengiriman"))
        cboNPWPPengusaha.Text = Trim(RS("Kode_Id_Pengusaha"))
        txtNPWPPengusaha.Text = Trim(RS("Id_Pengusaha"))
        txtNamaPengusaha.Text = Trim(RS("Nama_Pengusaha"))
        txtAlamatPengusaha.Text = Trim(RS("Alamat_Pengusaha"))
        txtNoIzinPengusaha.Text = Trim(RS("Nomor_Ijin_TPB") & "")
        cboNPWPPengirim.Text = Trim(RS("Kode_Id_Pengirim"))
        txtNPWPPengirim.Text = Trim(RS("Id_Pengirim"))
        txtNamaPengirim.Text = Trim(RS("Nama_Pengirim"))
        txtAlamatPengirim.Text = Trim(RS("Alamat_Pengirim"))
        txtNamaPengangkut.Text = Trim(RS("Nama_Pengangkut"))
        txtNoPolisi.Text = Trim(RS("Nomor_Polisi"))
        txtHarga.Text = Format(RS("Harga_Penyerahan"), gs_formatAmountIDR)
        cboNPWPPemilik.Text = Trim(RS("KODE_ID_PEMILIK"))
        txtNPWPPemilik.Text = Trim(RS("Id_Pemilik"))
        txtNamaPemilik.Text = Trim(RS("Nama_Pemilik"))
        txtAlamatPemilik.Text = Trim(RS("Alamat_Pemilik"))
        txtVolume.Text = Format(RS("Volume"), gs_formatQty)
        txtBruto.Text = Format(RS("Bruto"), gs_formatQty)
        txtNetto.Text = Format(RS("Netto"), gs_formatQty)
        txtJumahBrg.Text = Format(RS("Jumlah_Barang"), gs_formatQty)
        txtTempat.Text = Trim(RS("Kota_TTD"))
        dtpTanggal.Value = Trim(RS("Tanggal_TTD"))
        txtPemberitahu.Text = Trim(RS("Nama_TTD"))
        txtJabatan.Text = Trim(RS("Jabatan_TTD"))
        txtBCNo.Text = IIf(IsNull(Trim(RS("BC40_NO"))) = True, txtBCNo.Text, Trim(RS("BC40_NO")))
        txtBCType = IIf(IsNull(Trim(RS("BC_Type"))) = True, txtBCNo.Text, Trim(RS("BC_Type")))
        DTPBCDate.Value = IIf(IsNull(Trim(RS("BC40_Date"))) = True, dtpTanggal.Value, Trim(RS("BC40_Date")))
    End If
    
    up_IsiGridDokumen
    
    up_IsiGridKemasan
    
    up_TabBarang
    
    up_HargaPenyerahan
    
End Sub

Private Sub up_HargaPenyerahan()
    Dim sql As String
    Dim RS As New Recordset
    Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC40Harga"
    
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, NoPengajuan)
     
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
    cmd.CommandText = "sp_BC40Netto"
    
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, NoPengajuan)
     
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
        txtNetto.Text = Format(RS("Netto"), gs_formatQty)
        txtBruto.Text = Format(RS("Netto"), gs_formatQty)
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
    cmd.CommandText = "sp_BC40Volume"
    
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, NoPengajuan)
     
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
        txtVolume.Text = Format(RS("Volume"), gs_formatQty)
    Else
        txtVolume.Text = IIf(IsNull(Trim(RS("Volume"))) = True, 0, Trim(RS("Volume")))
    End If
End Sub

Private Sub up_PartReceipt()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC40Get_PartReceipt"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, NoPengajuan)
    cmd.Parameters.append cmd.CreateParameter("SuratJalanNo", adVarChar, adParamInput, 50, SuratJalan)
    cmd.Parameters.append cmd.CreateParameter("SupplierCode", adVarChar, adParamInput, 50, SupplierCode)
    
    Set RS = cmd.Execute
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
    
    sql = "SELECT * FROM Bea_Cukai_Tujuan Where Kode_Tujuan='" & Left(cboTujuan.Text, 1) & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
         cboTujuan.Text = Trim(RS(0)) & " - " & IIf(IsNull(RS(1)), "", Trim(RS(1)))
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

Private Sub txtJenisKemasan_Change()
    Dim sql As String
    Dim RS As New Recordset

    sql = "Select Uraian_Kemasan From Bea_Cukai_Kemasan Where Kode_Kemasan =  '" & txtJenisKemasan.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lblKemasan(0).Caption = Trim(RS("Uraian_Kemasan"))
    Else
        lblKemasan(0).Caption = ""
    End If
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

Private Sub txtKantorPabean_Change()
     up_LoadKantorPabean txtKantorPabean
End Sub

Private Sub txtNoPolisi_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSeriKemasan_Change(Index As Integer)
    Dim sql As String
    Dim RS As New Recordset
    Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC40SeriKemasan_Max"
    
    cmd.Parameters.append cmd.CreateParameter("No_Pengajuan", adVarChar, adParamInput, 50, NoPengajuan)
    
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

Private Sub txtNPWPPemilik_KeyPress(KeyAscii As Integer)
       If Left(cboNPWPPemilik, 1) = 1 Then
        If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
            KeyAscii = 0
        End If
        txtNPWPPemilik.MaxLength = 20
    Else
        If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
            KeyAscii = 0
        End If
        txtNPWPPemilik.MaxLength = 24
    End If
End Sub

Private Sub GridDokumen_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If GridDokumen.Col = bteColSelect Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then
          KeyAscii = 0
       End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
    End If
End Sub

Private Sub GridKemasan_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If GridKemasan.Col = bteColSelect Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then
          KeyAscii = 0
       End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
    End If
End Sub

Private Sub txtVolume_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub

Private Sub KoneksiMysql()
    Dim db_name As String
    Dim db_server As String
    Dim db_port As String
    Dim db_user As String
    Dim db_pass As String
    
    Dim sql As String
    Dim RS As New Recordset
    
    '//error traping
    On Error GoTo buat_koneksi_Error
    
    sql = "SELECT * FROM Connection_Mysql"
    Set RS = Db.Execute(sql)
        
    '/variable localhost
    db_name = Trim(RS("DatabaseName"))
    db_server = Trim(RS("ServerName"))
    db_port = Trim(RS("Port"))
    db_user = Trim(RS("UserId"))
    db_pass = fc_Decrypt(Trim(RS("Password")))
    
    '/buka koneksi my sql
    With MySQLCon
        .ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_user & ";PWD=" & db_pass & ";PORT=" & db_port & ""
        .Open
    End With
            
    On Error GoTo 0
    Exit Sub
    
buat_koneksi_Error:
    MsgBox "[" & err.number & "] " & err.Description, vbCritical, "Error"
End Sub


