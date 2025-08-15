VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmBC23Detail 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BC 23 Detail"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmBC23Detail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2535
      Left            =   240
      TabIndex        =   57
      Tag             =   "TFTF*/"
      Top             =   840
      Width           =   14565
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
         Tag             =   "TTFF*/"
         Top             =   780
         Width           =   2415
      End
      Begin VB.TextBox txtKPBBCPengawas 
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
         TabIndex        =   61
         Tag             =   "TTFF*/"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtKPBBCBongkar 
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
         TabIndex        =   60
         Tag             =   "TTFF*/"
         Top             =   1600
         Width           =   1335
      End
      Begin VB.TextBox txtNoDaftar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Locked          =   -1  'True
         TabIndex        =   59
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
         TabIndex        =   58
         Tag             =   "TTFF*/"
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSMask.MaskEdBox txtNoPengajuan 
         Height          =   315
         Left            =   1920
         TabIndex        =   65
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
         TabIndex        =   66
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
         Format          =   293994499
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   315
         Left            =   12840
         TabIndex        =   67
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
         Format          =   293994499
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
         Index           =   4
         Left            =   8520
         TabIndex        =   79
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
         Index           =   3
         Left            =   8520
         TabIndex        =   78
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
         Index           =   2
         Left            =   8520
         TabIndex        =   77
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1395
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3360
         X2              =   6000
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3360
         X2              =   6000
         Y1              =   1900
         Y2              =   1900
      End
      Begin VB.Label lblKPPBCBongkar 
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
         TabIndex        =   76
         Tag             =   "TTFF*/"
         Top             =   1630
         Width           =   2535
      End
      Begin VB.Label lblKPBBCPengawas 
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
         TabIndex        =   75
         Tag             =   "TTFF*/"
         Top             =   2070
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tujuan"
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
         TabIndex        =   74
         Tag             =   "TTFF*/"
         Top             =   390
         Width           =   1575
      End
      Begin MSForms.ComboBox cboTujuan 
         Height          =   315
         Left            =   10320
         TabIndex        =   73
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "KPPBC Pengawas"
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
         TabIndex        =   72
         Tag             =   "TTFF*/"
         Top             =   2070
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "KPPBC Bongkar"
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
         TabIndex        =   71
         Tag             =   "TTFF*/"
         Top             =   1630
         Width           =   1455
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
         Index           =   1
         Left            =   120
         TabIndex        =   70
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
         Index           =   0
         Left            =   120
         TabIndex        =   69
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
         TabIndex        =   68
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Back"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   56
      Tag             =   "TFFT*/"
      Top             =   10320
      Width           =   1140
   End
   Begin VB.CommandButton cmdAction 
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
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   55
      Tag             =   "FFTT*/"
      Top             =   10320
      Width           =   1140
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "Syn&cronize"
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   54
      Tag             =   "FFTT*/"
      Top             =   10320
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   240
      TabIndex        =   52
      Tag             =   "TFTT*/"
      Top             =   9600
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
         TabIndex        =   53
         Tag             =   "TFTF*/"
         Top             =   180
         Width           =   14325
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   495
      Left            =   12960
      TabIndex        =   0
      Tag             =   "FTTF*/"
      Top             =   120
      Width           =   1935
      _extentx        =   3413
      _extenty        =   873
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Tag             =   "TFTF*/"
      Top             =   3480
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   3625
      _Version        =   393216
      Style           =   1
      Tabs            =   5
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
      TabCaption(0)   =   "PEMASOK"
      TabPicture(0)   =   "FrmBC23Detail.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "IMPORTIR"
      TabPicture(1)   =   "FrmBC23Detail.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PEMILIK"
      TabPicture(2)   =   "FrmBC23Detail.frx":0E7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "PPJK"
      TabPicture(3)   =   "FrmBC23Detail.frx":0E96
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "PENGANGKUTAN"
      TabPicture(4)   =   "FrmBC23Detail.frx":0EB2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame8"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame8 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   34
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14295
         Begin VB.TextBox txtPelabuhanBongkar 
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
            Left            =   8280
            TabIndex        =   40
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtPelabuhanTransit 
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
            Left            =   8280
            TabIndex        =   39
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtPelabuhanMuat 
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
            Left            =   8280
            TabIndex        =   38
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtNegaraPengangkut 
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
            Left            =   3600
            TabIndex        =   37
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtVoyFlight 
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
            Left            =   2520
            TabIndex        =   36
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   975
         End
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
            Left            =   2520
            TabIndex        =   35
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   3495
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   9480
            X2              =   14040
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   9480
            X2              =   14040
            Y1              =   900
            Y2              =   900
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   9480
            X2              =   14040
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   4320
            X2              =   6120
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Label lblPelabuhanBongkar 
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
            Left            =   9480
            TabIndex        =   51
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   4575
         End
         Begin VB.Label lblPelabuhanTransit 
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
            Left            =   9480
            TabIndex        =   50
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label lblPelabuhanMuat 
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
            Left            =   9480
            TabIndex        =   49
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Pelabuhan Bongkar"
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
            Left            =   6480
            TabIndex        =   48
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Pelabuhan Transit"
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
            Left            =   6480
            TabIndex        =   47
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Pelabuhan Muat"
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
            Left            =   6480
            TabIndex        =   46
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblNegaraPengangkut 
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
            Left            =   4320
            TabIndex        =   45
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   1695
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Voy Flight && Negara"
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
            TabIndex        =   44
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Sarana Pengangkut"
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
            TabIndex        =   43
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Cara Pengangkutan"
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
            TabIndex        =   42
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1695
         End
         Begin MSForms.ComboBox cboCaraAngkut 
            Height          =   315
            Left            =   2520
            TabIndex        =   41
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2415
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4260;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   33
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14295
      End
      Begin VB.Frame Frame6 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   22
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14175
         Begin VB.TextBox txtIdentitasPemilik 
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
            TabIndex        =   26
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtAPIPemilik 
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
            Left            =   11640
            TabIndex        =   25
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtAlamatPemilik 
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
            TabIndex        =   24
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   4935
         End
         Begin VB.TextBox txtNamaPemilik 
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
            TabIndex        =   23
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   4935
         End
         Begin MSForms.ComboBox cboIDPemilik 
            Height          =   315
            Left            =   2040
            TabIndex        =   32
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
         Begin MSForms.ComboBox cboPemilikAPI 
            Height          =   315
            Left            =   10080
            TabIndex        =   31
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1455
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2566;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "API"
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
            Left            =   8280
            TabIndex        =   30
            Tag             =   "TTFF*/"
            Top             =   240
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
            TabIndex        =   29
            Tag             =   "TTFF*/"
            Top             =   960
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
            TabIndex        =   28
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1575
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
            TabIndex        =   27
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   10
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14280
         Begin VB.TextBox txtIdentitasImportir 
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
            MaxLength       =   20
            TabIndex        =   15
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txtNamaImportir 
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
            TabIndex        =   14
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   3615
         End
         Begin VB.TextBox txtNoIzin 
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
         Begin VB.TextBox txtAlamatImportir 
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
            Left            =   8280
            TabIndex        =   12
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   5535
         End
         Begin VB.TextBox txtAPIImportir 
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
            Left            =   9840
            TabIndex        =   11
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Identitas - NPWP"
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
            TabIndex        =   21
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1575
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
            TabIndex        =   20
            Tag             =   "TTFF*/"
            Top             =   750
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
            TabIndex        =   19
            Tag             =   "TTFF*/"
            Top             =   1110
            Width           =   975
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
            Left            =   6480
            TabIndex        =   18
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "API"
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
            Left            =   6480
            TabIndex        =   17
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   1575
         End
         Begin MSForms.ComboBox cboImportirAPI 
            Height          =   315
            Left            =   8280
            TabIndex        =   16
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   1455
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2566;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   14280
         Begin VB.TextBox txtNegara 
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
            MaxLength       =   2
            TabIndex        =   5
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtAlamatPemasok 
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
            TabIndex        =   4
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   7815
         End
         Begin VB.TextBox txtNamaPemasok 
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
            TabIndex        =   3
            Tag             =   "TTFF*/"
            Top             =   330
            Width           =   7815
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   2520
            X2              =   5160
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label lblNegara 
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
            Left            =   2520
            TabIndex        =   9
            Tag             =   "TTFF*/"
            Top             =   1140
            Width           =   2655
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Negara"
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
            TabIndex        =   8
            Tag             =   "TTFF*/"
            Top             =   1110
            Width           =   975
         End
         Begin VB.Label Label11 
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
            TabIndex        =   7
            Tag             =   "TTFF*/"
            Top             =   750
            Width           =   975
         End
         Begin VB.Label Label10 
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
            TabIndex        =   6
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   975
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   240
      TabIndex        =   80
      Tag             =   "TTTT*/"
      Top             =   5640
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   6
      TabsPerRow      =   9
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Barang"
      TabPicture(0)   =   "FrmBC23Detail.frx":0ECE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdDetailBarang"
      Tab(0).Control(1)=   "cmdAddBarang"
      Tab(0).Control(2)=   "Frame11"
      Tab(0).Control(3)=   "gridBarang"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Harga"
      TabPicture(1)   =   "FrmBC23Detail.frx":0EEA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame13"
      Tab(1).Control(1)=   "Frame14"
      Tab(1).Control(2)=   "Frame15"
      Tab(1).Control(3)=   "cmdSaveHarga"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Penimbunan"
      TabPicture(2)   =   "FrmBC23Detail.frx":0F06
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame12"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Dokumen"
      TabPicture(3)   =   "FrmBC23Detail.frx":0F22
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "gridDokumen"
      Tab(3).Control(1)=   "Frame9"
      Tab(3).Control(2)=   "btnBrowseDokumen"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Kemasan"
      TabPicture(4)   =   "FrmBC23Detail.frx":0F3E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "gridKemasan"
      Tab(4).Control(1)=   "Frame3"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Kontainer"
      TabPicture(5)   =   "FrmBC23Detail.frx":0F5A
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "gridKontainer"
      Tab(5).Control(1)=   "Frame10"
      Tab(5).Control(2)=   "txtJDataKontainer"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Respon"
      TabPicture(6)   =   "FrmBC23Detail.frx":0F76
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Label29"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label30"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "gridStatus"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "gridRespon"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Pungutan"
      TabPicture(7)   =   "FrmBC23Detail.frx":0F92
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "gridPungutan"
      Tab(7).ControlCount=   1
      Begin VB.CommandButton cmdSaveHarga 
         BackColor       =   &H0080FFFF&
         Caption         =   "Save"
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
         Left            =   -61680
         Style           =   1  'Graphical
         TabIndex        =   186
         Tag             =   "FFFF*/"
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame15 
         Caption         =   "Nilai ini yang akan tercantum di PIB"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -69240
         TabIndex        =   173
         Tag             =   "TFTF*/"
         Top             =   1920
         Width           =   8535
         Begin VB.TextBox txtAsuransi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            Locked          =   -1  'True
            TabIndex        =   178
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtFreightPIB 
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
            TabIndex        =   177
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtFOBPIB 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   176
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtCIF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   175
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtCIFRp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   174
            Tag             =   "TTFF*/"
            Text            =   "0.0000"
            Top             =   960
            Width           =   1935
         End
         Begin MSForms.ComboBox cboAsuransi 
            Height          =   315
            Left            =   1800
            TabIndex        =   185
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1935
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3413;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Asuransi bayar di"
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
            TabIndex        =   184
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "Nilai Asuransi"
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
            TabIndex        =   183
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "Freight"
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
            TabIndex        =   182
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "FOB"
            Height          =   255
            Left            =   4800
            TabIndex        =   181
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "CIF"
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
            Left            =   4800
            TabIndex        =   180
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "CIF Rp."
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
            Left            =   4800
            TabIndex        =   179
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Isi Sesuai Invoice / Dok. Lain"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -69240
         TabIndex        =   159
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   8535
         Begin VB.TextBox txtDiskon 
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
            Left            =   6360
            TabIndex        =   165
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtBiayaTambahan 
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
            Left            =   6360
            TabIndex        =   164
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtHargaCNF 
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
            Left            =   6360
            TabIndex        =   163
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtNDPBMInvoice 
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
            TabIndex        =   162
            Tag             =   "TTFF*/"
            Text            =   "0.0000"
            Top             =   960
            Width           =   1935
         End
         Begin VB.CommandButton cmdGetKurs 
            BackColor       =   &H0080FFFF&
            Caption         =   "Get Kurs"
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
            Style           =   1  'Graphical
            TabIndex        =   161
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtValutaInvoice 
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
            TabIndex        =   160
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "Diskon"
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
            Left            =   4800
            TabIndex        =   172
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "Biaya Tambahan"
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
            Left            =   4800
            TabIndex        =   171
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga CNF"
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
            Left            =   4800
            TabIndex        =   170
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "NDPBM"
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
            TabIndex        =   169
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "Valuta"
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
            TabIndex        =   168
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label50 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Harga"
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
            TabIndex        =   167
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1575
         End
         Begin MSForms.ComboBox cboKodeHarga 
            Height          =   315
            Left            =   1800
            TabIndex        =   166
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1935
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3413;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.TextBox txtJDataKontainer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   -61560
         Locked          =   -1  'True
         TabIndex        =   158
         Tag             =   "FFTF*/"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdDetailBarang 
         BackColor       =   &H0080FFFF&
         Caption         =   "Detail"
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
         Left            =   -62880
         Style           =   1  'Graphical
         TabIndex        =   157
         Tag             =   "FFTT*/"
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdAddBarang 
         BackColor       =   &H0080FFFF&
         Caption         =   "Add"
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
         Left            =   -61800
         Style           =   1  'Graphical
         TabIndex        =   156
         Tag             =   "FFTT*/"
         Top             =   3360
         Width           =   975
      End
      Begin VB.Frame Frame13 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   140
         Tag             =   "TTFT*/"
         Top             =   360
         Width           =   5175
         Begin VB.TextBox txtNilaiCIFRupiah 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            MaxLength       =   30
            TabIndex        =   147
            Tag             =   "TTFF*/"
            Text            =   "0"
            Top             =   2520
            Width           =   2895
         End
         Begin VB.TextBox txtNilaiCIF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            MaxLength       =   30
            TabIndex        =   146
            Tag             =   "TTFF*/"
            Text            =   "0"
            Top             =   2160
            Width           =   2895
         End
         Begin VB.TextBox txtAsuransiLNDN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            MaxLength       =   30
            TabIndex        =   145
            Tag             =   "TTFF*/"
            Text            =   "0"
            Top             =   1800
            Width           =   2895
         End
         Begin VB.TextBox txtFreight 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            MaxLength       =   30
            TabIndex        =   144
            Tag             =   "TTFF*/"
            Text            =   "0"
            Top             =   1440
            Width           =   2895
         End
         Begin VB.TextBox txtValuta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            MaxLength       =   4
            TabIndex        =   143
            Tag             =   "TTFF*/"
            Top             =   330
            Width           =   1095
         End
         Begin VB.TextBox txtNDPBM 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            MaxLength       =   30
            TabIndex        =   142
            Tag             =   "TTFF*/"
            Text            =   "0"
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txtFOB 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            MaxLength       =   30
            TabIndex        =   141
            Tag             =   "TTFF*/"
            Text            =   "0"
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label lblValuta 
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
            Left            =   3000
            TabIndex        =   155
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   3000
            X2              =   4680
            Y1              =   630
            Y2              =   630
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "Nilai CIF Rupiah"
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
            TabIndex        =   154
            Tag             =   "TTFF*/"
            Top             =   2550
            Width           =   1575
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "Nilai CIF"
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
            TabIndex        =   153
            Tag             =   "TTFF*/"
            Top             =   2190
            Width           =   1575
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "Asuransi LN/DN"
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
            TabIndex        =   152
            Tag             =   "TTFF*/"
            Top             =   1830
            Width           =   1575
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Freight"
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
            TabIndex        =   151
            Tag             =   "TTFF*/"
            Top             =   1470
            Width           =   1575
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Valuta"
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
            TabIndex        =   150
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "NDPBM"
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
            TabIndex        =   149
            Tag             =   "TTFF*/"
            Top             =   750
            Width           =   1815
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "FOB"
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
            TabIndex        =   148
            Tag             =   "TTFF*/"
            Top             =   1110
            Width           =   1575
         End
      End
      Begin VB.Frame Frame12 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   135
         Tag             =   "TTFT*/"
         Top             =   360
         Width           =   7215
         Begin VB.CommandButton cmdSavePenimbunan 
            BackColor       =   &H0080FFFF&
            Caption         =   "Save"
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
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   137
            Tag             =   "FFTT*/"
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtKodePenimbunan 
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
            Left            =   2400
            TabIndex        =   136
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   3720
            X2              =   6720
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Label lblPenimbunan 
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
            Left            =   3720
            TabIndex        =   139
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   3015
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Tempat Penimbunan"
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
            TabIndex        =   138
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   2535
         End
      End
      Begin VB.Frame Frame11 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   128
         Tag             =   "TTFT*/"
         Top             =   360
         Width           =   5175
         Begin VB.TextBox txtJumlahBarang 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
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
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   131
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtNettoBarang 
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
            MaxLength       =   20
            TabIndex        =   130
            Tag             =   "TTFF*/"
            Text            =   "0.0000"
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox txtBrutoBarang 
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
            MaxLength       =   20
            TabIndex        =   129
            Tag             =   "TTFF*/"
            Text            =   "0.0000"
            Top             =   330
            Width           =   2175
         End
         Begin VB.Label Label37 
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
            Height          =   255
            Left            =   240
            TabIndex        =   134
            Tag             =   "TTFF*/"
            Top             =   1110
            Width           =   1575
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Netto (Kg)"
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
            TabIndex        =   133
            Tag             =   "TTFF*/"
            Top             =   750
            Width           =   1815
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Bruto (Kg)"
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
            TabIndex        =   132
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame10 
         Height          =   3015
         Left            =   -68040
         TabIndex        =   114
         Tag             =   "TTTT*/"
         Top             =   840
         Width           =   7335
         Begin VB.TextBox txtIDKontainer 
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
            MaxLength       =   4
            TabIndex        =   121
            Tag             =   "TTFF*/"
            Top             =   1800
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancelKontainer 
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
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   120
            Tag             =   "FFTT*/"
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdDeleteKontainer 
            BackColor       =   &H0080FFFF&
            Caption         =   "Delete"
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
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   119
            Tag             =   "FFTT*/"
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdSaveKontainer 
            BackColor       =   &H0080FFFF&
            Caption         =   "Save"
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
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   118
            Tag             =   "FFTT*/"
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox txtKeteranganKontainer 
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
            MaxLength       =   4
            TabIndex        =   117
            Tag             =   "TTFF*/"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtNomorKontainer2 
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
            Left            =   3240
            MaxLength       =   7
            TabIndex        =   116
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtNomorKontainer1 
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
            MaxLength       =   4
            TabIndex        =   115
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
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
            TabIndex        =   127
            Tag             =   "TTFF*/"
            Top             =   1320
            Width           =   1815
         End
         Begin MSForms.ComboBox cboTipeKontainer 
            Height          =   315
            Left            =   2040
            TabIndex        =   126
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1335
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2355;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label33 
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
            Height          =   255
            Left            =   240
            TabIndex        =   125
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1695
         End
         Begin MSForms.ComboBox cboUkuranKontainer 
            Height          =   315
            Left            =   2040
            TabIndex        =   124
            Tag             =   "TTFF*/"
            Top             =   600
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
         Begin VB.Label Label32 
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
            Height          =   255
            Left            =   240
            TabIndex        =   123
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Kontainer"
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
            TabIndex        =   122
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton btnBrowseDokumen 
         BackColor       =   &H0080FFFF&
         Caption         =   "Browse"
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
         Left            =   -61680
         Style           =   1  'Graphical
         TabIndex        =   113
         Tag             =   "FFTT*/"
         Top             =   3360
         Width           =   975
      End
      Begin VB.Frame Frame9 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   92
         Tag             =   "FTFT*/"
         Top             =   360
         Width           =   5895
         Begin VB.TextBox txtFasilitasImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            TabIndex        =   102
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtPos3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4920
            TabIndex        =   101
            Tag             =   "TTFF*/"
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtPos2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4200
            TabIndex        =   100
            Tag             =   "TTFF*/"
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtPos1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3360
            TabIndex        =   99
            Tag             =   "TTFF*/"
            Top             =   2160
            Width           =   615
         End
         Begin VB.CommandButton btnBC 
            BackColor       =   &H0080FFFF&
            Caption         =   "BC 1.1"
            Height          =   315
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   98
            Tag             =   "TTFF*/"
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox txtNomorBC11 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            TabIndex        =   97
            Tag             =   "TTFF*/"
            Top             =   1800
            Width           =   2175
         End
         Begin VB.TextBox txtLCAWB 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            TabIndex        =   96
            Tag             =   "TTFF*/"
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox txtLCDokumen 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            TabIndex        =   95
            Tag             =   "TTFF*/"
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            TabIndex        =   94
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtInvoiceDokumen 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            TabIndex        =   93
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker dtpTglInvoice 
            Height          =   315
            Left            =   4080
            TabIndex        =   103
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
            Format          =   152305667
            CurrentDate     =   37798
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   3240
            TabIndex        =   104
            Tag             =   "TTFF*/"
            Top             =   720
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
            Format          =   152305667
            CurrentDate     =   37798
         End
         Begin MSComCtl2.DTPicker dtpTglLC 
            Height          =   315
            Left            =   4080
            TabIndex        =   105
            Tag             =   "TTFF*/"
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
            Format          =   152305667
            CurrentDate     =   37798
         End
         Begin MSComCtl2.DTPicker dtpTglLCAWB 
            Height          =   315
            Left            =   4080
            TabIndex        =   106
            Tag             =   "TTFF*/"
            Top             =   1440
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
            Format          =   152305667
            CurrentDate     =   37798
         End
         Begin MSComCtl2.DTPicker dtpTglBC11 
            Height          =   315
            Left            =   4080
            TabIndex        =   107
            Tag             =   "TTFF*/"
            Top             =   1800
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
            Format          =   152305667
            CurrentDate     =   37798
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Pos"
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
            Left            =   2640
            TabIndex        =   112
            Tag             =   "TTFF*/"
            Top             =   2190
            Width           =   735
         End
         Begin MSForms.ComboBox cboDokumenBLAWB 
            Height          =   315
            Left            =   120
            TabIndex        =   111
            Tag             =   "TTFF*/"
            Top             =   1440
            Width           =   1455
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2566;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "LC"
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
            TabIndex        =   110
            Tag             =   "TTFF*/"
            Top             =   1110
            Width           =   1455
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Fasilitas Impor"
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
            TabIndex        =   109
            Tag             =   "TTFF*/"
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice"
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
            TabIndex        =   108
            Tag             =   "TTFF*/"
            Top             =   390
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3495
         Left            =   -66360
         TabIndex        =   81
         Tag             =   "TTTT*/"
         Top             =   360
         Width           =   5655
         Begin VB.CommandButton cmdCancelKemasan 
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
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   87
            Tag             =   "FFTT*/"
            Top             =   2880
            Width           =   975
         End
         Begin VB.CommandButton cmdDeleteKemasan 
            BackColor       =   &H0080FFFF&
            Caption         =   "Delete"
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
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   86
            Tag             =   "FFTT*/"
            Top             =   2880
            Width           =   975
         End
         Begin VB.CommandButton cmdSaveKemasan 
            BackColor       =   &H0080FFFF&
            Caption         =   "Save"
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
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   85
            Tag             =   "FFTT*/"
            Top             =   2880
            Width           =   975
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
            Left            =   1680
            MaxLength       =   255
            TabIndex        =   84
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   3135
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
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   83
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtJumlahKemasan 
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
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   82
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label46 
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
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblJenisKemasan 
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
            Left            =   2880
            TabIndex        =   90
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   2415
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   2880
            X2              =   5400
            Y1              =   900
            Y2              =   900
         End
         Begin VB.Label Label48 
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
            Height          =   255
            Left            =   240
            TabIndex        =   89
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   1335
         End
         Begin VB.Label Label47 
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
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   1335
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid gridKemasan 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   187
         TabStop         =   0   'False
         Tag             =   "TTFT*/"
         Top             =   480
         Width           =   8205
         _cx             =   14473
         _cy             =   5741
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
      Begin VSFlex8Ctl.VSFlexGrid gridDokumen 
         Height          =   2775
         Left            =   -68520
         TabIndex        =   188
         TabStop         =   0   'False
         Tag             =   "TTTT*/"
         Top             =   600
         Width           =   7725
         _cx             =   13626
         _cy             =   4895
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
      Begin VSFlex8Ctl.VSFlexGrid gridKontainer 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   189
         TabStop         =   0   'False
         Tag             =   "TTFT*/"
         Top             =   480
         Width           =   6525
         _cx             =   11509
         _cy             =   5741
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
      Begin VSFlex8Ctl.VSFlexGrid gridRespon 
         Height          =   2895
         Left            =   360
         TabIndex        =   190
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
         Left            =   7440
         TabIndex        =   191
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
      Begin VSFlex8Ctl.VSFlexGrid gridBarang 
         Height          =   2775
         Left            =   -69600
         TabIndex        =   192
         TabStop         =   0   'False
         Tag             =   "TTTT*/"
         Top             =   480
         Width           =   8805
         _cx             =   15531
         _cy             =   4895
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
      Begin VSFlex8Ctl.VSFlexGrid gridPungutan 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   193
         TabStop         =   0   'False
         Tag             =   "TTTT*/"
         Top             =   360
         Width           =   9525
         _cx             =   16801
         _cy             =   5953
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
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data"
         Height          =   255
         Left            =   -62760
         TabIndex        =   196
         Tag             =   "FFTF*/"
         Top             =   540
         Width           =   1215
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
         Left            =   360
         TabIndex        =   195
         Tag             =   "TTFF*/"
         Top             =   480
         Width           =   975
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
         Left            =   7440
         TabIndex        =   194
         Tag             =   "TTFF*/"
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BC 23 Detail"
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
      TabIndex        =   197
      Tag             =   "TTTF*/"
      Top             =   240
      Width           =   14535
   End
End
Attribute VB_Name = "FrmBC23Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DbMy As New ADODB.Connection
Dim checkOpenDB As Boolean
Dim checkOKToMysql As Boolean
Public checkAlreadyData As Boolean


'-------------------------------------------
Const colNo As Integer = 0
Const colJumlah As Integer = 1
Const colKodeKemasan As Integer = 2
Const colNamaKemasan As Integer = 3
Const colMerkKemasan As Integer = 4
Const colCountKemasan As Integer = 5

'-------------------------------------------
Const colJenisDokumen As Integer = 0
Const colNomorDokumen As Integer = 1
Const colTanggal As Integer = 2
Const colCountDokumen As Integer = 3


'-------------------------------------------
Const colNoUrutKontainer As Integer = 0
Const colNomorKontainer As Integer = 1
Const colUkuran As Integer = 2
Const colTipe As Integer = 3
Const colIDKontainer As Integer = 4
Const colHideUkuran As Integer = 5
Const colHideTipe As Integer = 6
Const colHideKeterangan As Integer = 7
Const colCountKontainer As Integer = 8

'-------------------------------------------
Const colKodeBarang As Integer = 0
Const colNamaBarang As Integer = 1
Const colSatuan As Integer = 2
Const colHarga As Integer = 3
Const ColQty As Integer = 4
Const colTotal As Integer = 5
Const colHideNoSeri As Integer = 6
Const colCountBarang As Integer = 7

'-------------------------------------------
Const colKodeRespon As Integer = 0
Const colUraianRespon As Integer = 1
Const colWaktuRespon As Integer = 2
Const colCountRespon As Integer = 3

'-------------------------------------------
Const colKodeStatus As Integer = 0
Const colUraianStatus As Integer = 1
Const colWaktuStatus As Integer = 2
Const colCountStatus As Integer = 3

'-------------------------------------------
Const colNoTarif As Integer = 0
Const colJenisPungutan As Integer = 1
Const colDitangguhkan As Integer = 2
Const colDibebaskan As Integer = 3
Const colTidakDipungut As Integer = 4
Const colCountPungutan As Integer = 5


'====================================================================================================================================================================
' 1. Functions (Start)
'====================================================================================================================================================================

Private Function uf_ValidateInput() As Boolean
    If txtKPBBCBongkar = "" Then
        txtKPBBCBongkar.SetFocus
        LblErrMsg = "Please Input Kode KPPBC Bongkar!"
        uf_ValidateInput = False
        Exit Function
    ElseIf cboTujuan = "" Then
        cboTujuan.SetFocus
        LblErrMsg = "Please Input Tujuan!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtNamaPemasok.Text = "" Then
        txtNamaPemasok.SetFocus
        SSTab2.Tab = 0
        LblErrMsg = "Please Input Nama Pemasok!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtAlamatPemasok.Text = "" Then
        txtAlamatPemasok.SetFocus
        SSTab2.Tab = 0
        LblErrMsg = "Please Input Alamat Pemasok!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtNegara.Text = "" Or lblNegara.Caption = "" Then
        txtNegara.SetFocus
        SSTab2.Tab = 0
        LblErrMsg = "Please Input Negara Pemasok!"
        uf_ValidateInput = False
        Exit Function
    ElseIf cboImportirAPI.Text = "" Then
        cboImportirAPI.SetFocus
        SSTab2.Tab = 1
        LblErrMsg = "Please Input Kode API Importir!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtAPIImportir.Text = "" Then
        txtAPIImportir.SetFocus
        SSTab2.Tab = 1
        LblErrMsg = "Please Input Nomor Importir!"
        uf_ValidateInput = False
        Exit Function
    ElseIf cboIDPemilik.Text = "" Then
        cboIDPemilik.SetFocus
        SSTab2.Tab = 2
        LblErrMsg = "Please Input Identitas!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtIdentitasPemilik.Text = "" Then
        txtIdentitasPemilik.SetFocus
        SSTab2.Tab = 2
        LblErrMsg = "Please Input Nomor Identitas!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtNamaPemilik.Text = "" Then
        txtNamaPemilik.SetFocus
        SSTab2.Tab = 2
        LblErrMsg = "Please Input Nama Pemilik!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtAlamatPemilik.Text = "" Then
        txtAlamatPemilik.SetFocus
        SSTab2.Tab = 2
        LblErrMsg = "Please Input Alamat Pemilik!"
        uf_ValidateInput = False
        Exit Function
    ElseIf cboPemilikAPI.Text = "" Then
        cboPemilikAPI.SetFocus
        SSTab2.Tab = 2
        LblErrMsg = "Please Input Kode Jenis API Pemilik!"
        uf_ValidateInput = False
        Exit Function
    ElseIf cboPemilikAPI.Text = "" Then
        cboPemilikAPI.SetFocus
        SSTab2.Tab = 2
        LblErrMsg = "Please Input API Pemilik!"
        uf_ValidateInput = False
        Exit Function
    ElseIf cboKodeHarga.ListIndex = -1 Then
        cboKodeHarga.SetFocus
        SSTab1.Tab = 1
        LblErrMsg = "Please Input Kode Harga!"
        Exit Function
    ElseIf txtValutaInvoice = "" Then
        txtValutaInvoice.SetFocus
        SSTab1.Tab = 1
        LblErrMsg = "Please Input Valuta!"
        Exit Function
    End If
    
    uf_ValidateInput = True
End Function

'====================================================================================================================================================================
' 1. Functions (End)
'====================================================================================================================================================================


Public Sub up_LoadDataBC23(pNoPengajuan As String)
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23LoadData_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(pNoPengajuan, "-", ""))
    Set RS = cmd.Execute
    
    If Not RS.EOF Then
        checkAlreadyData = True
        
        txtNamaPemasok = IIf(IsNull(RS.Fields("NAMA_PEMASOK")), "", RS.Fields("NAMA_PEMASOK"))
        txtAlamatPemasok = IIf(IsNull(RS.Fields("ALAMAT_PEMASOK")), "", RS.Fields("ALAMAT_PEMASOK"))
        txtNegara = IIf(IsNull(RS.Fields("KODE_NEGARA_PEMASOK")), "", RS.Fields("KODE_NEGARA_PEMASOK"))
        lblNegara.Caption = IIf(IsNull(RS.Fields("NEGARA_PEMASOK")), "", RS.Fields("NEGARA_PEMASOK"))
        
        txtIdentitasImportir = IIf(IsNull(RS.Fields("ID_PENGUSAHA")), "", RS.Fields("ID_PENGUSAHA"))
        txtNamaImportir = IIf(IsNull(RS.Fields("NAMA_PENGUSAHA")), "", RS.Fields("NAMA_PENGUSAHA"))
        txtNoIzin = IIf(IsNull(RS.Fields("NOMOR_IJIN_TPB")), "", RS.Fields("NOMOR_IJIN_TPB"))
        txtAlamatImportir = IIf(IsNull(RS.Fields("ALAMAT_PENGUSAHA")), "", RS.Fields("ALAMAT_PENGUSAHA"))
        cboImportirAPI = IIf(IsNull(RS.Fields("JENISIDPENGUSAHA")), "", RS.Fields("JENISIDPENGUSAHA"))
        txtAPIImportir = IIf(IsNull(RS.Fields("API_PENGUSAHA")), "", RS.Fields("API_PENGUSAHA"))
        
        cboIDPemilik = IIf(IsNull(RS.Fields("KODEPEMILIK")), "", RS.Fields("KODEPEMILIK"))
        txtIdentitasPemilik = IIf(IsNull(RS.Fields("ID_PEMILIK")), "", RS.Fields("ID_PEMILIK"))
        txtNamaPemilik = IIf(IsNull(RS.Fields("NAMA_PEMILIK")), "", RS.Fields("NAMA_PEMILIK"))
        txtAlamatPemilik = IIf(IsNull(RS.Fields("ALAMAT_PEMILIK")), "", RS.Fields("ALAMAT_PEMILIK"))
        cboPemilikAPI = IIf(IsNull(RS.Fields("KODEPEMILIK")), "", RS.Fields("KODEPEMILIK"))
        txtAPIPemilik = IIf(IsNull(RS.Fields("API_PEMILIK")), "", RS.Fields("API_PEMILIK"))
        
        cboCaraAngkut = IIf(IsNull(RS.Fields("CARA_ANGKUT")), "", RS.Fields("CARA_ANGKUT"))
        txtNamaPengangkut = IIf(IsNull(RS.Fields("NAMA_PENGANGKUT")), "", RS.Fields("NAMA_PENGANGKUT"))
        txtVoyFlight = IIf(IsNull(RS.Fields("NOMOR_VOY_FLIGHT")), "", RS.Fields("NOMOR_VOY_FLIGHT"))
        txtNegaraPengangkut = IIf(IsNull(RS.Fields("KODE_BENDERA")), "", RS.Fields("KODE_BENDERA"))
        lblNegaraPengangkut.Caption = IIf(IsNull(RS.Fields("NAMA_BENDERA")), "", RS.Fields("NAMA_BENDERA"))
        
        txtPelabuhanMuat = IIf(IsNull(RS.Fields("KODE_PEL_MUAT")), "", RS.Fields("KODE_PEL_MUAT"))
        lblPelabuhanMuat.Caption = IIf(IsNull(RS.Fields("PELABUHAN_MUAT")), "", RS.Fields("PELABUHAN_MUAT"))
        txtPelabuhanTransit = IIf(IsNull(RS.Fields("KODE_PEL_TRANSIT")), "", RS.Fields("KODE_PEL_TRANSIT"))
        lblPelabuhanTransit.Caption = IIf(IsNull(RS.Fields("PELABUHAN_TRANSIT")), "", RS.Fields("PELABUHAN_TRANSIT"))
        txtPelabuhanBongkar = IIf(IsNull(RS.Fields("KODE_PEL_BONGKAR")), "", RS.Fields("KODE_PEL_BONGKAR"))
        lblPelabuhanBongkar.Caption = IIf(IsNull(RS.Fields("PELABUHAN_BONGKAR")), "", RS.Fields("PELABUHAN_BONGKAR"))
        
        If txtIdentitasImportir = "" Then
            txtIdentitasImportir.Text = RS.Fields("NPWP_No")
        End If
        If txtNamaImportir = "" Then
            txtNamaImportir.Text = RS.Fields("Company_Name")
        End If
        If txtAlamatImportir = "" Then
            txtAlamatImportir.Text = RS.Fields("Company_Address")
        End If
        If txtNoIzin = "" Then
            txtNoIzin.Text = RS.Fields("No_Izin")
        End If
        
        txtKPBBCBongkar.Text = IIf(IsNull(RS.Fields("KODE_KANTOR_BONGKAR")), "", RS.Fields("KODE_KANTOR_BONGKAR"))
        lblKPPBCBongkar.Caption = IIf(IsNull(RS.Fields("KANTORBONGKAR")), "", RS.Fields("KANTORBONGKAR"))
        cboTujuan = IIf(IsNull(RS.Fields("TUJUANTPB")), "", RS.Fields("TUJUANTPB"))
                        
        txtKPBBCPengawas.Text = RS.Fields("KPPBC_Pengawas")
        lblKPBBCPengawas.Caption = RS.Fields("Kantor_KPPBC_Pengawas")
        txtTempat.Text = Trim(RS.Fields("City"))
        txtPemberitahu.Text = Trim(RS.Fields("SJ_Person"))
        txtJabatan.Text = Trim(RS.Fields("SJ_Position"))
                
        txtInvoiceDokumen = Trim(RS.Fields("NomorDokumenInvoice"))
        dtpTglInvoice.Value = RS.Fields("TglDokumenInvoice")
        txtLCDokumen = Trim(RS.Fields("NomorDokumenLC"))
        dtpTglLC.Value = RS.Fields("TglDokumenLC")
        txtLCAWB = Trim(RS.Fields("NomorDokumenBLAWB"))
        dtpTglLCAWB.Value = RS.Fields("TglDokumenBLAWB")
        cboDokumenBLAWB = IIf(IsNull(RS.Fields("JenisDokumenBLAWB")), "", RS.Fields("JenisDokumenBLAWB"))
        
        txtNomorBC11 = IIf(IsNull(RS.Fields("NOMOR_BC11")), "", RS.Fields("NOMOR_BC11"))
        dtpTglBC11.Value = RS.Fields("TANGGAL_BC11")
        
        txtValuta = IIf(IsNull(RS.Fields("Kode_Valuta")), "", RS.Fields("Kode_Valuta"))
        txtNDPBM = Format(IIf(IsNull(RS.Fields("NDPBM")), 0, RS.Fields("NDPBM")), "#,0.0000")
        
        txtFreight = Format(IIf(IsNull(RS.Fields("FREIGHT")), 0, RS.Fields("FREIGHT")), "#,0.00")
        txtAsuransiLNDN = Format(IIf(IsNull(RS.Fields("ASURANSI")), 0, RS.Fields("ASURANSI")), "#,0.00")
        txtNilaiCIF = Format(IIf(IsNull(RS.Fields("CIF")), 0, RS.Fields("CIF")), "#,0.00")
        txtNilaiCIFRupiah = Format(IIf(IsNull(RS.Fields("CIF_RUPIAH")), 0, RS.Fields("CIF_RUPIAH")), "#,0.0000")
        
        txtValutaInvoice = IIf(IsNull(RS.Fields("Kode_Valuta")), "", RS.Fields("Kode_Valuta"))
        txtNDPBMInvoice = Format(IIf(IsNull(RS.Fields("NDPBM")), 0, RS.Fields("NDPBM")), "#,0.00")
        
        txtHargaCNF = Format(IIf(IsNull(RS.Fields("FOB")), 0, RS.Fields("FOB")), "#,0.00")
        
        txtBiayaTambahan = Format(IIf(IsNull(RS.Fields("BIAYA_TAMBAHAN")), 0, RS.Fields("BIAYA_TAMBAHAN")), "#,0.00")
        txtDiskon = Format(IIf(IsNull(RS.Fields("DISKON")), 0, RS.Fields("DISKON")), "#,0.00")
        txtFreightPIB = Format(IIf(IsNull(RS.Fields("FREIGHT")), 0, RS.Fields("FREIGHT")), "#,0.00")
        
        txtCIF = Format(IIf(IsNull(RS.Fields("CIF")), 0, RS.Fields("CIF")), "#,0.00")
        txtCIFRp = Format(IIf(IsNull(RS.Fields("CIF_RUPIAH")), 0, RS.Fields("CIF_RUPIAH")), "#,0.00")
        
        txtAsuransi = Format(IIf(IsNull(RS.Fields("ASURANSI")), 0, RS.Fields("ASURANSI")), "#,0.00")
        
        txtBrutoBarang = Format(IIf(IsNull(RS.Fields("BRUTO")), 0, RS.Fields("BRUTO")), "#,0.00")
        txtNettoBarang = Format(IIf(IsNull(RS.Fields("NETTO")), 0, RS.Fields("NETTO")), "#,0.00")
        
        If RS.Fields("KODE_HARGA") = "" Then
            cboKodeHarga = ""
        Else
            cboKodeHarga = IIf(IsNull(RS.Fields("DescHarga")), "", RS.Fields("DescHarga"))
        End If
        
        If RS.Fields("KODE_ASURANSI") = "" Then
            cboAsuransi = ""
        Else
            cboAsuransi = IIf(IsNull(RS.Fields("DescAsuransi")), "", RS.Fields("DescAsuransi"))
        End If
        
        If cboKodeHarga <> "" Then
            If Trim(Split(cboKodeHarga, "-")(0)) = "CIF" Then
                txtFOBPIB = "0.00"
                txtFOB = "0.00"
            Else
                txtFOBPIB = Format(IIf(IsNull(RS.Fields("FOB")), 0, RS.Fields("FOB")), "#,0.00")
                txtFOB = Format(IIf(IsNull(RS.Fields("FOB")), 0, RS.Fields("FOB")), "#,0.00")
            End If
        End If
        
        txtKodePenimbunan = IIf(IsNull(RS.Fields("Kode_TPS")), "", RS.Fields("Kode_TPS"))
        lblPenimbunan = IIf(IsNull(RS.Fields("URAIAN_TPS")), "", RS.Fields("URAIAN_TPS"))
        
        txtPos1 = IIf(IsNull(RS.Fields("POS_BC11")), "", RS.Fields("POS_BC11"))
        txtPos2 = IIf(IsNull(RS.Fields("SUBPOS_BC11")), "", RS.Fields("SUBPOS_BC11"))
        txtPos3 = IIf(IsNull(RS.Fields("SUBSUBPOS_BC11")), "", RS.Fields("SUBSUBPOS_BC11"))
        
    End If
    
End Sub

Private Sub up_LoadKemasan(pKode As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select Kode_Kemasan, Nama_Kemasan = Uraian_Kemasan From Bea_Cukai_Kemasan Where Kode_Kemasan = '" & pKode & "'"
Set RS = Db.Execute(sql)

If Not RS.EOF Then
    lblJenisKemasan.Caption = RS.Fields("Nama_Kemasan")
Else
    lblJenisKemasan.Caption = ""
End If
End Sub

Private Sub up_LoadMasterGlobal(pTable As String, pLabel As Label, pField As String, pFilter As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select * From " & pTable & " " & pFilter & ""
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    pLabel.Caption = RS.Fields(pField)
Else
    pLabel.Caption = ""
End If
End Sub

Private Sub up_LoadNegara(pKode As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select Kode_Negara, Nama_Negara From Bea_Cukai_Negara Where Kode_Negara = '" & pKode & "'"
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    lblNegara.Caption = RS.Fields("Nama_Negara")
Else
    lblNegara.Caption = ""
End If
End Sub

Private Sub gb_LoadDataMaster(pTable As String, pField As String, pLabelName As Label, pFilter As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select " & pField & " From " & pTable & " " & pFilter & ""
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    pLabelName.Caption = RS.Fields(0)
Else
    pLabelName.Caption = ""
End If
End Sub

Private Sub up_LoadKantorKPPBCBongkar(pKode As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select * From Bea_Cukai_Kantor_Pabean Where Kode_Kantor = '" & pKode & "'"
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    lblKPPBCBongkar.Caption = RS.Fields("Nama_Kantor")
Else
    lblKPPBCBongkar.Caption = ""
End If
End Sub

Private Sub up_LoadKantorPabean(pKode As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select * From Bea_Cukai_Kantor_Pabean Where Kode_Kantor = '" & pKode & "'"
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    lblKPBBCPengawas.Caption = RS.Fields("Nama_Kantor")
Else
    lblKPBBCPengawas.Caption = ""
End If
End Sub

Private Sub up_GetBC11()
Dim sql As String
Dim RS As New Recordset

sql = "Select * From Bea_Cukai_TPB_Header"
Set RS = Db.Execute(sql)
    
'If Not rs.EOF Then
'    lblKPBBCPengawas.Caption = rs.fields("Nama_Kantor")
'Else
'    lblKPBBCPengawas.Caption = ""
'End If
End Sub

Private Sub up_GetKurs()
Dim sql As String
Dim RS As New Recordset

sql = "Select * From Bea_Cukai_TPB_Header"
Set RS = Db.Execute(sql)
    
'If Not rs.EOF Then
'    lblKPBBCPengawas.Caption = rs.fields("Nama_Kantor")
'Else
'    lblKPBBCPengawas.Caption = ""
'End If
End Sub

Private Sub up_FillComboKodeID(pcbo As MSForms.ComboBox)
Dim sql As String
Dim RS As New Recordset

    sql = "Select * From Bea_Cukai_Kode_ID"
    Set RS = Db.Execute(sql)

    With pcbo
        .clear
        .columnCount = 1
        .ColumnWidths = "30pt;80pt"
        .ListWidth = 110
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS(1)) & " - " & IIf(IsNull(RS(2)), "", Trim(RS(2)))
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = 0
    End With
End Sub

Private Sub up_FillComboBLAWB(pcbo As MSForms.ComboBox)
Dim sql As String
Dim RS As New Recordset

    sql = "Select * From Bea_Cukai_Dokumen Where Kode_Dokumen In ('705','740')"
    Set RS = Db.Execute(sql)

    With pcbo
        .clear
        .columnCount = 2
        .ColumnWidths = "70pt;0pt"
        .ListWidth = 70
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS(3))
            .List(i, 1) = Trim(RS(1))
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = -1
    End With
End Sub

Private Sub up_FillComboGeneral(pcbo As MSForms.ComboBox, pTable As String, pField1 As String, pField2 As String, pColWidth2 As Integer, pListWidth As Integer)
Dim sql As String
Dim RS As New Recordset

    sql = "Select " & pField1 & ", " & pField2 & " From " & pTable & ""
    Set RS = Db.Execute(sql)

    With pcbo
        .clear
        .columnCount = 1
        .ColumnWidths = "20pt; " & pColWidth2 & "pt"
        .ListWidth = pListWidth
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS(0)) & " - " & IIf(IsNull(RS(1)), "", Trim(RS(1)))
            
            RS.MoveNext
            i = i + 1
        Loop

    End With
End Sub

Private Sub up_FillComboAPI(pcbo As MSForms.ComboBox)
Dim sql As String
Dim RS As New Recordset

    sql = "Select * From Bea_Cukai_Jenis_API"
    Set RS = Db.Execute(sql)

    With pcbo
        .clear
        .columnCount = 1
        .ColumnWidths = "20pt;50pt"
        .ListWidth = 70
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS(1)) & " - " & IIf(IsNull(RS(2)), "", Trim(RS(2)))
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = 0
    End With
End Sub

Private Sub up_FillComboCaraAngkut()
Dim sql As String
Dim RS As New Recordset

    sql = "Select * From Bea_Cukai_Cara_Angkut"
    Set RS = Db.Execute(sql)

    With cboCaraAngkut
        .clear
        .columnCount = 1
        .ColumnWidths = "50pt;100pt"
        .ListWidth = 150
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS(1)) & " - " & IIf(IsNull(RS(2)), "", Trim(RS(2)))
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = 0
    End With
End Sub

Private Sub up_FillComboTujuan()
Dim sql As String
Dim RS As New Recordset

    sql = "Select * From Bea_Cukai_Jenis_TPB"
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
            .List(i, 0) = Trim(RS(1)) & " - " & IIf(IsNull(RS(2)), "", Trim(RS(2)))
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = -1
    End With
End Sub

Private Sub up_FillComboKemasan()
'Dim sql As String
'Dim rs As New Recordset
'
'    sql = "Select Kode_Kemasan, Nama_Kemasan From Bea_Cukai_Kemasan"
'    Set rs = Db.Execute(sql)
'
'    With cbotrade
'        .clear
'        .ColumnCount = 2
'        .ColumnWidths = "50pt;300pt"
'        .ListWidth = 350
'        .ListRows = 15
'
'        i = 1
'
'        Do While Not rs.EOF
'            .AddItem
'            .List(i, 0) = Trim(rs("Kode_Kemasan"))
'            .List(i, 1) = IIf(IsNull(rs("Nama_Kemasan")), "", Trim(rs("Nama_Kemasan")))
'
'            rs.MoveNext
'            i = i + 1
'        Loop
'
'        .ListIndex = 0
'
'    End With
End Sub

Private Sub up_GridHeaderKontainer()
    
    With gridKontainer
        .ColS = colCountKontainer
        .Rows = 1

        .TextMatrix(0, colNoUrutKontainer) = "No"
        .TextMatrix(0, colNomorKontainer) = "Nomor Kontainer"
        .TextMatrix(0, colUkuran) = "Ukuran"
        .TextMatrix(0, colTipe) = "Tipe"

        .ColWidth(colNoUrutKontainer) = 500
        .ColWidth(colNomorKontainer) = 1800
        .ColWidth(colUkuran) = 1000
        .ColWidth(colTipe) = 1500
        
        .ColWidth(colHideUkuran) = 0
        .ColWidth(colHideTipe) = 0
        .ColWidth(colHideKeterangan) = 0
        .ColWidth(colIDKontainer) = 0
    End With
End Sub

Private Sub up_GridHeaderPungutan()
      With gridPungutan
        .ColS = colCountPungutan
        .Rows = 1

        .TextMatrix(0, colNo) = "No"
        .TextMatrix(0, colJenisPungutan) = "Jenis Pungutan"
        .TextMatrix(0, colDitangguhkan) = "Ditangguhkan"
        .TextMatrix(0, colDibebaskan) = "Dibebaskan"
        .TextMatrix(0, colTidakDipungut) = "Tidak Dipungut"
        
        .ColWidth(colNo) = 500
        .ColWidth(colJenisPungutan) = 2500
        .ColWidth(colDitangguhkan) = 1700
        .ColWidth(colDibebaskan) = 1700
        .ColWidth(colTidakDipungut) = 1700
       
        .ColFormat(colJenisPungutan) = "#,0.00"
        .ColFormat(colDitangguhkan) = "#,0.00"
        .ColFormat(colDibebaskan) = "#,0.00"
        .ColFormat(colTidakDipungut) = "#,0.00"
        
        
        .MergeCells = flexMergeRestrictRows
        .WordWrap = True
        
        .AllowUserResizing = flexResizeColumns
        
    End With
End Sub

Private Sub up_GridHeaderBarang()
    
    With gridBarang
        .ColS = colCountBarang
        .Rows = 1

        .TextMatrix(0, colKodeBarang) = "Kode"
        .TextMatrix(0, colNamaBarang) = "Uraian"
        .TextMatrix(0, colSatuan) = "Satuan"
        .TextMatrix(0, colHarga) = "Harga"
        .TextMatrix(0, ColQty) = "Qty"
        .TextMatrix(0, colTotal) = "Total"

        .ColWidth(colKodeBarang) = 1200
        .ColWidth(colNamaBarang) = 2000
        .ColWidth(colSatuan) = 1000
        .ColWidth(colHarga) = 1200
        .ColWidth(ColQty) = 1000
        .ColWidth(colTotal) = 1500
        .ColWidth(colHideNoSeri) = 0
'        .ColFormat(colTanggal) = "dd MMM yyyy"
        .ColAlignment(colKodeBarang) = flexAlignLeftCenter
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


Private Sub up_GridHeaderDokumen()
    
    With gridDokumen
        .ColS = colCountDokumen
        .Rows = 1

        .TextMatrix(0, colJenisDokumen) = "Jenis Dokumen"
        .TextMatrix(0, colNomorDokumen) = "Nomor"
        .TextMatrix(0, colTanggal) = "Tanggal"

        .ColWidth(colJenisDokumen) = 1600
        .ColWidth(colNomorDokumen) = 3000
        .ColWidth(colTanggal) = 1500
        
        .ColFormat(colTanggal) = "dd MMM yyyy"
        .ColAlignment(colNomorDokumen) = flexAlignLeftCenter
    End With
End Sub

Private Sub up_GridHeaderKemasan()
    LblErrMsg.Caption = ""
    
    With gridKemasan
        .ColS = colCountKemasan
        .Rows = 1
        
        .TextMatrix(0, colNo) = "No"
        .TextMatrix(0, colJumlah) = "Jumlah"
        .TextMatrix(0, colKodeKemasan) = "Kode"
        .TextMatrix(0, colNamaKemasan) = "Kemasan"
        .TextMatrix(0, colMerkKemasan) = "Merk Kemasan"
        
        .ColWidth(colNo) = 500
        .ColWidth(colJumlah) = 1000
        .ColWidth(colKodeKemasan) = 1200
        .ColWidth(colNamaKemasan) = 2500
        .ColWidth(colMerkKemasan) = 2500
        
    End With
    
End Sub

Private Sub up_GridLoadKemasan()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer
    Dim i As Integer
    
    up_GridHeaderKemasan
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23LoadDataKemasan_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    Set RS = cmd.Execute
    
    With gridKemasan
    
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
            
            i = i + 1
            .TextMatrix(li_Row, colNo) = i
            .TextMatrix(li_Row, colKodeKemasan) = RS.Fields("Kode_Jenis_Kemasan")
            .TextMatrix(li_Row, colNamaKemasan) = IIf(IsNull(RS.Fields("Uraian_Kemasan")), "", RS.Fields("Uraian_Kemasan"))
            .TextMatrix(li_Row, colJumlah) = RS.Fields("JUMLAH_KEMASAN")
            .TextMatrix(li_Row, colMerkKemasan) = RS.Fields("MERK_KEMASAN")
            
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End With
End Sub

Private Sub up_GridLoadKontainer()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer
    Dim i As Integer
    
    up_GridHeaderKontainer
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23LoadDataKontainer_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    Set RS = cmd.Execute
    
    With gridKontainer
    
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
            
            i = i + 1
            .TextMatrix(li_Row, colNoUrutKontainer) = i
            .TextMatrix(li_Row, colNomorKontainer) = RS.Fields("NOMOR_KONTAINER")
            .TextMatrix(li_Row, colUkuran) = RS.Fields("URAIAN_UKURAN_KONTAINER")
            .TextMatrix(li_Row, colTipe) = RS.Fields("URAIAN_TIPE_KONTAINER")
            .TextMatrix(li_Row, colHideUkuran) = RS.Fields("KODE_UKURAN_KONTAINER") & "-" & RS.Fields("URAIAN_UKURAN_KONTAINER")
            .TextMatrix(li_Row, colHideTipe) = RS.Fields("KODE_TIPE_KONTAINER") & "-" & RS.Fields("URAIAN_TIPE_KONTAINER")
            .TextMatrix(li_Row, colIDKontainer) = RS.Fields("ID_KONTAINER")
            .TextMatrix(li_Row, colHideKeterangan) = RS.Fields("KETERANGAN")
            
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        
        txtJDataKontainer = i
    End With
End Sub

Private Sub up_GridLoadPungutan()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer
    Dim i As Integer
    
    Dim NilaiDitangguhkan As Double
    Dim NilaiDibebaskan As Double
    Dim NilaiTidakDipungut As Double
    
    up_GridHeaderPungutan
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23TPBTarifFasilitas_Sel"

    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))

    Set RS = cmd.Execute

    With gridPungutan
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
            
            i = i + 1
            
            .TextMatrix(li_Row, colNo) = i
            .TextMatrix(li_Row, colJenisPungutan) = IIf(IsNull(RS.Fields("Kode_Pungutan")), "", RS.Fields("Kode_Pungutan"))
            .TextMatrix(li_Row, colDitangguhkan) = IIf(IsNull(RS.Fields("NILAIDITANGGUHKAN")), 0, RS.Fields("NILAIDITANGGUHKAN"))
            .TextMatrix(li_Row, colDibebaskan) = IIf(IsNull(RS.Fields("NILAIDIBEBASKAN")), 0, RS.Fields("NILAIDIBEBASKAN"))
            .TextMatrix(li_Row, colTidakDipungut) = IIf(IsNull(RS.Fields("NILAITIDAKDIPUNGUT")), 0, RS.Fields("NILAITIDAKDIPUNGUT"))
            
            NilaiDitangguhkan = NilaiDitangguhkan + CDbl(.TextMatrix(li_Row, colDitangguhkan))
            NilaiDibebaskan = NilaiDibebaskan + CDbl(.TextMatrix(li_Row, colDibebaskan))
            NilaiTidakDipungut = NilaiTidakDipungut + CDbl(.TextMatrix(li_Row, colTidakDipungut))
            
            RS.MoveNext
        Wend
        
        .Rows = .Rows + 1
        li_Row = .Rows - 1

        
        .TextMatrix(li_Row, colDitangguhkan) = NilaiDitangguhkan
        .TextMatrix(li_Row, colDibebaskan) = NilaiDibebaskan
        .TextMatrix(li_Row, colTidakDipungut) = NilaiTidakDipungut
        
        .Cell(flexcpText, li_Row, colNo, li_Row, colJenisPungutan) = "TOTAL"
        .Cell(flexcpFontBold, li_Row, colNo, li_Row, colTidakDipungut) = True
        
        Const ClrTotal1 = &HFFC0C0
        .Cell(flexcpBackColor, li_Row, colNo, .Rows - 1, colTidakDipungut) = ClrTotal1  '&HFFC0C0
        
        .MergeRow(li_Row) = True
        
        RS.Close
        Set RS = Nothing
    End With
            
End Sub

Private Sub up_GridLoadBarang()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer

    up_GridHeaderBarang
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailBarang_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    
    Set RS = cmd.Execute
    
    With gridBarang
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
                    
            .TextMatrix(li_Row, colKodeBarang) = Trim(RS!Kode_Barang)
            .TextMatrix(li_Row, colNamaBarang) = Trim(RS!URAIAN)
            .TextMatrix(li_Row, colSatuan) = Trim(RS!URAIAN_SATUAN)
            .TextMatrix(li_Row, colHarga) = Format(RS!HARGA_SATUAN, "#,0.00")
            .TextMatrix(li_Row, ColQty) = Format(RS!JUMLAH_SATUAN, "#,0.00")
            .TextMatrix(li_Row, colTotal) = Format(RS!Total, "#,0.00")
            .TextMatrix(li_Row, colHideNoSeri) = Trim(RS!SERI_BARANG)
            
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        
        txtJumlahBarang = .Rows - 1
    End With
End Sub



Private Sub up_GridLoadDokumen()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer

    up_GridHeaderDokumen
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23TPBDokumenWithoutInvoice_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    
    Set RS = cmd.Execute
    
    With gridDokumen
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
                    
            .TextMatrix(li_Row, colJenisDokumen) = Trim(RS!Uraian_Dokumen)
            .TextMatrix(li_Row, colNomorDokumen) = Trim(RS!Nomor_Dokumen)
            .TextMatrix(li_Row, colTanggal) = Trim(RS!Tanggal_Dokumen)
            
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End With
End Sub


Private Sub up_SaveKontainer()
    Dim cmd As ADODB.Command
    Dim prm1 As ADODB.Parameter
    Dim prm2 As ADODB.Parameter
    Dim prm3 As ADODB.Parameter
    Dim prm4 As ADODB.Parameter
    Dim prm5 As ADODB.Parameter
    Dim prm6 As ADODB.Parameter
    
    Dim Y As Integer
    
    If txtNomorKontainer1 = "" Then
        txtNomorKontainer1.SetFocus
        LblErrMsg = "Please Input Nomor Kontainer!"
        Exit Sub
    ElseIf txtNomorKontainer2 = "" Then
        txtNomorKontainer2.SetFocus
        LblErrMsg = "Please Input Nomor Kontainer!"
        Exit Sub
    ElseIf cboUkuranKontainer = "" Then
        cboUkuranKontainer.SetFocus
        LblErrMsg = "Please Input Ukuran Kontainer!"
        Exit Sub
    ElseIf cboTipeKontainer = "" Then
        cboTipeKontainer.SetFocus
        LblErrMsg = "Please Input Tipe Kontainer!"
        Exit Sub
    End If
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailKontainer_Upd"
    
    If txtIDKontainer = "" Then
        txtIDKontainer = 0
    End If
    
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NomorKontainer", adVarChar, adParamInput, 15, txtNomorKontainer1 & txtNomorKontainer2)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("Ukuran", adVarChar, adParamInput, 5, Split(cboUkuranKontainer, "-")(0))
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("Tipe", adVarChar, adParamInput, 5, Split(cboTipeKontainer, "-")(0))
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("Keterangan", adVarChar, adParamInput, 255, txtKeteranganKontainer)
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("IDKontainer", adInteger, adParamInput, , txtIDKontainer)
    cmd.Parameters.append prm6
    
    cmd.Execute Y
    
    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC23DetailKontainer_Ins"
            
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NomorKontainer", adVarChar, adParamInput, 15, txtNomorKontainer1 & txtNomorKontainer2)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("Ukuran", adVarChar, adParamInput, 5, Split(cboUkuranKontainer, "-")(0))
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("Tipe", adVarChar, adParamInput, 5, Split(cboTipeKontainer, "-")(0))
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("Keterangan", adVarChar, adParamInput, 255, txtKeteranganKontainer)
        cmd.Parameters.append prm5

        cmd.Execute
    End If

    up_GridLoadKontainer
        
    txtNomorKontainer1 = ""
    txtNomorKontainer2 = ""
    cboUkuranKontainer = ""
    cboTipeKontainer = ""
    txtKeteranganKontainer = ""
    txtNomorKontainer1.SetFocus
    
    If Y = 0 Then
        LblErrMsg = DisplayMsg(1000)
    Else
        LblErrMsg = DisplayMsg(1101)
    End If
End Sub

Private Sub up_SaveDetailBC23()
    Dim cmd As ADODB.Command
    
    Dim prm1 As ADODB.Parameter
    Dim prm2 As ADODB.Parameter
    Dim prm3 As ADODB.Parameter
    Dim prm4 As ADODB.Parameter
    Dim prm5 As ADODB.Parameter
    Dim prm6 As ADODB.Parameter
    Dim prm7 As ADODB.Parameter
    Dim prm8 As ADODB.Parameter
    Dim prm9 As ADODB.Parameter
    Dim prm10 As ADODB.Parameter
    Dim prm11 As ADODB.Parameter
    Dim prm12 As ADODB.Parameter
    Dim prm13 As ADODB.Parameter
    Dim prm14 As ADODB.Parameter
    Dim prm15 As ADODB.Parameter
    Dim prm16 As ADODB.Parameter
    Dim prm17 As ADODB.Parameter
    Dim prm18 As ADODB.Parameter
    Dim prm19 As ADODB.Parameter
    Dim prm20 As ADODB.Parameter
    Dim prm21 As ADODB.Parameter
    Dim prm22 As ADODB.Parameter
    Dim prm23 As ADODB.Parameter
    Dim prm24 As ADODB.Parameter
    Dim prm25 As ADODB.Parameter
    Dim prm26 As ADODB.Parameter
    Dim prm27 As ADODB.Parameter
    Dim prm28 As ADODB.Parameter
    Dim prm29 As ADODB.Parameter
    Dim prm30 As ADODB.Parameter
    Dim prm31 As ADODB.Parameter
    Dim prm32 As ADODB.Parameter
    Dim prm33 As ADODB.Parameter
    Dim prm34 As ADODB.Parameter
    Dim prm35 As ADODB.Parameter
    Dim prm36 As ADODB.Parameter
    Dim prm37 As ADODB.Parameter
    Dim prm38 As ADODB.Parameter
    Dim prm39 As ADODB.Parameter
    Dim prm40 As ADODB.Parameter
    Dim prm41 As ADODB.Parameter
    Dim prm42 As ADODB.Parameter
    Dim prm43 As ADODB.Parameter
    Dim prm44 As ADODB.Parameter
    Dim prm45 As ADODB.Parameter
    Dim prm46 As ADODB.Parameter
    Dim prm47 As ADODB.Parameter
    Dim prm48 As ADODB.Parameter
    Dim prm49 As ADODB.Parameter
    Dim prm50 As ADODB.Parameter
    
    Dim Y As Integer
    
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23Detail_Upd"
    
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("KodeTujuanTPB", adVarChar, adParamInput, 10, Left(cboTujuan, 1))
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NamaTTD", adVarChar, adParamInput, 200, txtPemberitahu)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("JabatanTTD", adVarChar, adParamInput, 200, txtJabatan)
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KotaTTD", adVarChar, adParamInput, 200, txtTempat)
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("TanggalTTD", adDate, adParamInput, , Format(dtpTanggal, "yyyy-MM-dd"))
    cmd.Parameters.append prm6
    Set prm7 = cmd.CreateParameter("KodeKantorBongkar", adVarChar, adParamInput, 200, txtKPBBCBongkar)
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("KodeKantorPengawas", adVarChar, adParamInput, 200, txtKPBBCPengawas)
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("NamaPemasok", adVarChar, adParamInput, 200, txtNamaPemasok)
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("AlamatPemasok", adVarChar, adParamInput, 200, txtAlamatPemasok)
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("KodeNegaraPemasok", adVarChar, adParamInput, 200, txtNegara)
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("IDPengusaha", adVarChar, adParamInput, 200, Replace(Replace(txtIdentitasImportir, ".", ""), "-", ""))
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("NamaPengusaha", adVarChar, adParamInput, 200, txtNamaImportir)
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("NomorIjinTPB", adVarChar, adParamInput, 200, txtNoIzin)
    cmd.Parameters.append prm14
    Set prm15 = cmd.CreateParameter("AlamatPengusaha", adVarChar, adParamInput, 200, txtAlamatImportir)
    cmd.Parameters.append prm15
    Set prm16 = cmd.CreateParameter("KodeJenisAPIPengusaha", adVarChar, adParamInput, 200, Left(cboImportirAPI, 1))
    cmd.Parameters.append prm16
    Set prm17 = cmd.CreateParameter("APIPengusaha", adVarChar, adParamInput, 200, txtAPIImportir)
    cmd.Parameters.append prm17
    Set prm18 = cmd.CreateParameter("KodeIDPemilik", adVarChar, adParamInput, 200, Left(cboIDPemilik, 1))
    cmd.Parameters.append prm18
    Set prm19 = cmd.CreateParameter("IDPemilik", adVarChar, adParamInput, 200, txtIdentitasPemilik)
    cmd.Parameters.append prm19
    Set prm20 = cmd.CreateParameter("NamaPemilik", adVarChar, adParamInput, 200, txtNamaPemilik)
    cmd.Parameters.append prm20
    Set prm21 = cmd.CreateParameter("AlamatPemilik", adVarChar, adParamInput, 200, txtAlamatPemilik)
    cmd.Parameters.append prm21
    Set prm22 = cmd.CreateParameter("KodeJenisAPIPemilik", adVarChar, adParamInput, 200, Left(cboPemilikAPI, 1))
    cmd.Parameters.append prm22
    Set prm23 = cmd.CreateParameter("APIPemilik", adVarChar, adParamInput, 200, txtAPIPemilik)
    cmd.Parameters.append prm23
    Set prm24 = cmd.CreateParameter("KodeCaraAngkut", adVarChar, adParamInput, 200, Left(cboCaraAngkut, 1))
    cmd.Parameters.append prm24
    Set prm25 = cmd.CreateParameter("NamaPengangkut", adVarChar, adParamInput, 200, txtNamaPengangkut)
    cmd.Parameters.append prm25
    Set prm26 = cmd.CreateParameter("NomorVoyFlight", adVarChar, adParamInput, 200, txtVoyFlight)
    cmd.Parameters.append prm26
    Set prm27 = cmd.CreateParameter("KodeBendera", adVarChar, adParamInput, 200, txtNegaraPengangkut)
    cmd.Parameters.append prm27
    Set prm28 = cmd.CreateParameter("KodePelabuhanMuat", adVarChar, adParamInput, 200, txtPelabuhanMuat)
    cmd.Parameters.append prm28
    Set prm29 = cmd.CreateParameter("KodePelabuhanTransit", adVarChar, adParamInput, 200, txtPelabuhanTransit)
    cmd.Parameters.append prm29
    Set prm30 = cmd.CreateParameter("KodePelabuhanBongkar", adVarChar, adParamInput, 200, txtPelabuhanBongkar)
    cmd.Parameters.append prm30
    Set prm31 = cmd.CreateParameter("KodeTPS", adVarChar, adParamInput, 200, txtKodePenimbunan)
    cmd.Parameters.append prm31
    Set prm32 = cmd.CreateParameter("NomorBC11", adVarChar, adParamInput, 200, txtNomorBC11)
    cmd.Parameters.append prm32
    Set prm33 = cmd.CreateParameter("Bruto", adDecimal, adParamInput, , CDbl(txtBrutoBarang))
    prm33.Precision = 38
    prm33.NumericScale = 4
    cmd.Parameters.append prm33
    Set prm34 = cmd.CreateParameter("Netto", adDecimal, adParamInput, , CDbl(txtNettoBarang))
    prm34.Precision = 38
    prm34.NumericScale = 4
    cmd.Parameters.append prm34
    Set prm35 = cmd.CreateParameter("Pos1", adVarChar, adParamInput, 200, txtPos1)
    cmd.Parameters.append prm35
    Set prm36 = cmd.CreateParameter("Pos2", adVarChar, adParamInput, 200, txtPos2)
    cmd.Parameters.append prm36
    Set prm37 = cmd.CreateParameter("Pos3", adVarChar, adParamInput, 200, txtPos3)
    cmd.Parameters.append prm37
    Set prm38 = cmd.CreateParameter("KodeValuta", adVarChar, adParamInput, 200, txtValuta)
    cmd.Parameters.append prm38
    Set prm39 = cmd.CreateParameter("FOB", adDecimal, adParamInput, , CDbl(txtHargaCNF))
    prm39.Precision = 38
    prm39.NumericScale = 4
    cmd.Parameters.append prm39
    Set prm40 = cmd.CreateParameter("Freight", adDecimal, adParamInput, , CDbl(txtFreight))
    prm40.Precision = 38
    prm40.NumericScale = 4
    cmd.Parameters.append prm40
    Set prm41 = cmd.CreateParameter("Asuransi", adDecimal, adParamInput, , CDbl(txtAsuransi))
    prm41.Precision = 38
    prm41.NumericScale = 4
    cmd.Parameters.append prm41
    Set prm42 = cmd.CreateParameter("HargaInvoice", adDecimal, adParamInput, , CDbl(txtFreight))
    prm42.Precision = 38
    prm42.NumericScale = 4
    cmd.Parameters.append prm42
    Set prm43 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , CDbl(txtCIF))
    prm43.Precision = 38
    prm43.NumericScale = 4
    cmd.Parameters.append prm43
    Set prm44 = cmd.CreateParameter("CIFRupiah", adDecimal, adParamInput, , CDbl(txtCIFRp))
    prm44.Precision = 38
    prm44.NumericScale = 4
    cmd.Parameters.append prm44
    Set prm45 = cmd.CreateParameter("NDPBM", adDecimal, adParamInput, , CDbl(txtNDPBMInvoice))
    prm45.Precision = 38
    prm45.NumericScale = 4
    cmd.Parameters.append prm45
    
    cmd.Execute Y
        
    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC23Detail_Ins"
    
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("KodeTujuanTPB", adVarChar, adParamInput, 10, Left(cboTujuan, 1))
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("NamaTTD", adVarChar, adParamInput, 200, txtPemberitahu)
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("JabatanTTD", adVarChar, adParamInput, 200, txtJabatan)
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KotaTTD", adVarChar, adParamInput, 200, txtTempat)
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("TanggalTTD", adDate, adParamInput, , Format(dtpTanggal, "yyyy-MM-dd"))
        cmd.Parameters.append prm6
        Set prm7 = cmd.CreateParameter("KodeKantorBongkar", adVarChar, adParamInput, 200, txtKPBBCBongkar)
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("KodeKantorPengawas", adVarChar, adParamInput, 200, txtKPBBCPengawas)
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("NamaPemasok", adVarChar, adParamInput, 200, txtNamaPemasok)
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("AlamatPemasok", adVarChar, adParamInput, 200, txtAlamatPemasok)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("KodeNegaraPemasok", adVarChar, adParamInput, 200, txtNegara)
        cmd.Parameters.append prm11
        Set prm12 = cmd.CreateParameter("IDPengusaha", adVarChar, adParamInput, 200, Replace(Replace(txtIdentitasImportir, ".", ""), "-", ""))
        cmd.Parameters.append prm12
        Set prm13 = cmd.CreateParameter("NamaPengusaha", adVarChar, adParamInput, 200, txtNamaImportir)
        cmd.Parameters.append prm13
        Set prm14 = cmd.CreateParameter("NomorIjinTPB", adVarChar, adParamInput, 200, txtNoIzin)
        cmd.Parameters.append prm14
        Set prm15 = cmd.CreateParameter("AlamatPengusaha", adVarChar, adParamInput, 200, txtAlamatImportir)
        cmd.Parameters.append prm15
        Set prm16 = cmd.CreateParameter("KodeJenisAPIPengusaha", adVarChar, adParamInput, 200, Left(cboImportirAPI, 1))
        cmd.Parameters.append prm16
        Set prm17 = cmd.CreateParameter("APIPengusaha", adVarChar, adParamInput, 200, txtAPIImportir)
        cmd.Parameters.append prm17
        Set prm18 = cmd.CreateParameter("KodeIDPemilik", adVarChar, adParamInput, 200, Left(cboIDPemilik, 1))
        cmd.Parameters.append prm18
        Set prm19 = cmd.CreateParameter("IDPemilik", adVarChar, adParamInput, 200, txtIdentitasPemilik)
        cmd.Parameters.append prm19
        Set prm20 = cmd.CreateParameter("NamaPemilik", adVarChar, adParamInput, 200, txtNamaPemilik)
        cmd.Parameters.append prm20
        Set prm21 = cmd.CreateParameter("AlamatPemilik", adVarChar, adParamInput, 200, txtAlamatPemilik)
        cmd.Parameters.append prm21
        Set prm22 = cmd.CreateParameter("KodeJenisAPIPemilik", adVarChar, adParamInput, 200, Left(cboPemilikAPI, 1))
        cmd.Parameters.append prm22
        Set prm23 = cmd.CreateParameter("APIPemilik", adVarChar, adParamInput, 200, txtAPIPemilik)
        cmd.Parameters.append prm23
        Set prm24 = cmd.CreateParameter("KodeCaraAngkut", adVarChar, adParamInput, 200, Left(cboCaraAngkut, 1))
        cmd.Parameters.append prm24
        Set prm25 = cmd.CreateParameter("NamaPengangkut", adVarChar, adParamInput, 200, txtNamaPengangkut)
        cmd.Parameters.append prm25
        Set prm26 = cmd.CreateParameter("NomorVoyFlight", adVarChar, adParamInput, 200, txtVoyFlight)
        cmd.Parameters.append prm26
        Set prm27 = cmd.CreateParameter("KodeBendera", adVarChar, adParamInput, 200, txtNegaraPengangkut)
        cmd.Parameters.append prm27
        Set prm28 = cmd.CreateParameter("KodePelabuhanMuat", adVarChar, adParamInput, 200, txtPelabuhanMuat)
        cmd.Parameters.append prm28
        Set prm29 = cmd.CreateParameter("KodePelabuhanTransit", adVarChar, adParamInput, 200, txtPelabuhanTransit)
        cmd.Parameters.append prm29
        Set prm30 = cmd.CreateParameter("KodePelabuhanBongkar", adVarChar, adParamInput, 200, txtPelabuhanBongkar)
        cmd.Parameters.append prm30
        Set prm31 = cmd.CreateParameter("KodeTPS", adVarChar, adParamInput, 200, txtKodePenimbunan)
        cmd.Parameters.append prm31
        Set prm32 = cmd.CreateParameter("NomorBC11", adVarChar, adParamInput, 200, txtNomorBC11)
        cmd.Parameters.append prm32
        Set prm33 = cmd.CreateParameter("Bruto", adDecimal, adParamInput, , CDbl(txtBrutoBarang))
        prm33.Precision = 38
        prm33.NumericScale = 4
        cmd.Parameters.append prm33
        Set prm34 = cmd.CreateParameter("Netto", adDecimal, adParamInput, , CDbl(txtNettoBarang))
        prm34.Precision = 38
        prm34.NumericScale = 4
        cmd.Parameters.append prm34
        Set prm35 = cmd.CreateParameter("Pos1", adVarChar, adParamInput, 200, txtPos1)
        cmd.Parameters.append prm35
        Set prm36 = cmd.CreateParameter("Pos2", adVarChar, adParamInput, 200, txtPos2)
        cmd.Parameters.append prm36
        Set prm37 = cmd.CreateParameter("Pos3", adVarChar, adParamInput, 200, txtPos3)
        cmd.Parameters.append prm37
        Set prm38 = cmd.CreateParameter("KodeValuta", adVarChar, adParamInput, 200, txtValuta)
        cmd.Parameters.append prm38
        Set prm39 = cmd.CreateParameter("FOB", adDecimal, adParamInput, , CDbl(txtHargaCNF))
        prm39.Precision = 38
        prm39.NumericScale = 4
        cmd.Parameters.append prm39
        Set prm40 = cmd.CreateParameter("Freight", adDecimal, adParamInput, , CDbl(txtFreight))
        prm40.Precision = 38
        prm40.NumericScale = 4
        cmd.Parameters.append prm40
        Set prm41 = cmd.CreateParameter("Asuransi", adDecimal, adParamInput, , CDbl(txtAsuransi))
        prm41.Precision = 38
        prm41.NumericScale = 4
        cmd.Parameters.append prm41
        Set prm42 = cmd.CreateParameter("HargaInvoice", adDecimal, adParamInput, , CDbl(txtFreight))
        prm42.Precision = 38
        prm42.NumericScale = 4
        cmd.Parameters.append prm42
        Set prm43 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , CDbl(txtCIF))
        prm43.Precision = 38
        prm43.NumericScale = 4
        cmd.Parameters.append prm43
        Set prm44 = cmd.CreateParameter("CIFRupiah", adDecimal, adParamInput, , CDbl(txtCIFRp))
        prm44.Precision = 38
        prm44.NumericScale = 4
        cmd.Parameters.append prm44
        Set prm45 = cmd.CreateParameter("NDPBM", adDecimal, adParamInput, , CDbl(txtNDPBMInvoice))
        prm45.Precision = 38
        prm45.NumericScale = 4
        cmd.Parameters.append prm45
    
        cmd.Execute Y

    End If
    
    LblErrMsg = DisplayMsg(1101)

End Sub

Private Sub up_SaveHarga()
    Dim cmd As ADODB.Command
    Dim prm1 As ADODB.Parameter
    Dim prm2 As ADODB.Parameter
    Dim prm3 As ADODB.Parameter
    Dim prm4 As ADODB.Parameter
    Dim prm5 As ADODB.Parameter
    Dim prm6 As ADODB.Parameter
    Dim prm7 As ADODB.Parameter
    Dim prm8 As ADODB.Parameter
    Dim prm9 As ADODB.Parameter
    Dim prm10 As ADODB.Parameter
    Dim prm11 As ADODB.Parameter
    Dim prm12 As ADODB.Parameter
    
    Dim Y As Integer
    

    
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailHarga_Upd"
    
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("KodeHarga", adVarChar, adParamInput, 10, Trim(Split(cboKodeHarga, "-")(0)))
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("KodeValuta", adVarChar, adParamInput, 10, txtValutaInvoice)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("NDPBM", adDecimal, adParamInput, , txtNDPBMInvoice)
    prm4.Precision = 38
    prm4.NumericScale = 4
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("FOB", adDecimal, adParamInput, , txtHargaCNF)
    prm5.Precision = 38
    prm5.NumericScale = 4
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("BiayaTambahan", adDecimal, adParamInput, , txtBiayaTambahan)
    prm6.Precision = 38
    prm6.NumericScale = 4
    cmd.Parameters.append prm6
    Set prm7 = cmd.CreateParameter("Diskon", adDecimal, adParamInput, , txtDiskon)
    prm7.Precision = 38
    prm7.NumericScale = 4
    cmd.Parameters.append prm7
    If cboAsuransi = "" Then
        Set prm8 = cmd.CreateParameter("KodeAsuransi", adVarChar, adParamInput, 10, "")
        cmd.Parameters.append prm8
    Else
        Set prm8 = cmd.CreateParameter("KodeAsuransi", adVarChar, adParamInput, 10, Trim(Split(cboAsuransi, "-")(0)))
        cmd.Parameters.append prm8
    End If

    Set prm9 = cmd.CreateParameter("Asuransi", adDecimal, adParamInput, , txtAsuransi)
    prm9.Precision = 38
    prm9.NumericScale = 4
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("Freight", adDecimal, adParamInput, , txtFreightPIB)
    prm10.Precision = 38
    prm10.NumericScale = 4
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , txtCIF)
    prm11.Precision = 38
    prm11.NumericScale = 4
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("CIFRupiah", adDecimal, adParamInput, , txtCIFRp)
    prm12.Precision = 38
    prm12.NumericScale = 4
    cmd.Parameters.append prm12
    
    cmd.Execute Y
    
    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC23DetailHarga_Ins"
        
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("KodeHarga", adVarChar, adParamInput, 10, Trim(Split(cboKodeHarga, "-")(0)))
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("KodeValuta", adVarChar, adParamInput, 10, txtValutaInvoice)
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("NDPBM", adNumeric, adParamInput, , txtNDPBMInvoice)
        prm4.Precision = 38
        prm4.NumericScale = 4
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("FOB", adDecimal, adParamInput, , txtHargaCNF)
        prm5.Precision = 38
        prm5.NumericScale = 4
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("BiayaTambahan", adDecimal, adParamInput, , txtBiayaTambahan)
        prm6.Precision = 38
        prm6.NumericScale = 4
        cmd.Parameters.append prm6
        Set prm7 = cmd.CreateParameter("Diskon", adDecimal, adParamInput, , txtDiskon)
        prm7.Precision = 38
        prm7.NumericScale = 4
        cmd.Parameters.append prm7
        If cboAsuransi = "" Then
            Set prm8 = cmd.CreateParameter("KodeAsuransi", adVarChar, adParamInput, 10, Null)
            cmd.Parameters.append prm8
        Else
            Set prm8 = cmd.CreateParameter("KodeAsuransi", adVarChar, adParamInput, 10, Trim(Split(cboAsuransi, "-")(0)))
            cmd.Parameters.append prm8
        End If
        Set prm9 = cmd.CreateParameter("Asuransi", adDecimal, adParamInput, , txtAsuransi)
        prm9.Precision = 38
        prm9.NumericScale = 4
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("Freight", adDecimal, adParamInput, , txtFreightPIB)
        prm10.Precision = 38
        prm10.NumericScale = 4
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , txtCIF)
        prm11.Precision = 38
        prm11.NumericScale = 4
        cmd.Parameters.append prm11
        Set prm12 = cmd.CreateParameter("CIFRupiah", adDecimal, adParamInput, , txtCIFRp)
        prm12.Precision = 38
        prm12.NumericScale = 4
        cmd.Parameters.append prm12
    
        cmd.Execute
    End If
    
'    up_LoadDataBC23 Replace(txtNoPengajuan, "-", "")
    
End Sub

Private Sub up_SaveKemasan()
    Dim cmd As ADODB.Command
    Dim prm1 As ADODB.Parameter
    Dim prm2 As ADODB.Parameter
    Dim prm3 As ADODB.Parameter
    Dim prm4 As ADODB.Parameter
    Dim Y As Integer
    
    If txtJumlahKemasan = "" Then
        txtJumlahKemasan.SetFocus
        LblErrMsg = "Please Input Jumlah Kemasan!"
        Exit Sub
    ElseIf txtJenisKemasan = "" Or lblJenisKemasan = "" Then
        txtJenisKemasan.SetFocus
        LblErrMsg = "Please Input Jenis Kemasan!"
        Exit Sub
    ElseIf txtMerkKemasan = "" Then
        txtMerkKemasan.SetFocus
        LblErrMsg = "Please Input Merk Kemasan!"
        Exit Sub
    End If
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailKemasan_Upd"
    
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("KodeKemasan", adVarChar, adParamInput, 2, txtJenisKemasan)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("Jumlah", adNumeric, adParamInput, , txtJumlahKemasan)
    prm3.Precision = 18
    prm3.NumericScale = 2
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("Merk", adVarChar, adParamInput, 255, txtMerkKemasan)
    cmd.Parameters.append prm4
'
'    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
'    cmd.Parameters.append cmd.CreateParameter("KodeKemasan", adVarChar, adParamInput, 2, txtJenisKemasan)
'    cmd.Parameters.append cmd.CreateParameter("Jumlah", adNumeric, adParamInput, 18, txtJumlahKemasan)
'    cmd.Parameters.append cmd.CreateParameter("Merk", adVarChar, adParamInput, 255, txtMerkKemasan)
    
    cmd.Execute Y
    
    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC23DetailKemasan_Ins"
            
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("KodeKemasan", adVarChar, adParamInput, 2, txtJenisKemasan)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("Jumlah", adNumeric, adParamInput, , txtJumlahKemasan)
        prm3.Precision = 18
        prm3.NumericScale = 2
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("Merk", adVarChar, adParamInput, 255, txtMerkKemasan)
        cmd.Parameters.append prm4
                
            
'        cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
'        cmd.Parameters.append cmd.CreateParameter("KodeKemasan", adVarChar, adParamInput, 2, txtJenisKemasan)
'        cmd.Parameters.append cmd.CreateParameter("Jumlah", adNumeric, adParamInput, 18, txtJumlahKemasan)
'        cmd.Parameters.append cmd.CreateParameter("Merk", adVarChar, adParamInput, 255, txtMerkKemasan)


        cmd.Execute
    End If

    up_GridLoadKemasan
    
    
    txtJumlahKemasan = ""
    txtJenisKemasan = ""
    txtMerkKemasan = ""
    lblJenisKemasan.Caption = ""
    txtJenisKemasan.Enabled = True
    txtJumlahKemasan.SetFocus
    If Y = 0 Then
        LblErrMsg = DisplayMsg(1000)
    Else
        LblErrMsg = DisplayMsg(1101)
    End If
End Sub

Private Sub up_DeleteKemasan()
    Dim cmd As ADODB.Command
    Dim prm1 As ADODB.Parameter
    Dim prm2 As ADODB.Parameter
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailKemasan_Del"
    
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("KodeKemasan", adVarChar, adParamInput, 2, txtJenisKemasan)
    cmd.Parameters.append prm2
    
    cmd.Execute
    
    up_GridLoadKemasan
    
    
    txtJumlahKemasan = ""
    txtJenisKemasan = ""
    txtMerkKemasan = ""
    lblJenisKemasan.Caption = ""
    txtJenisKemasan.Enabled = True
    txtJumlahKemasan.SetFocus
    
    LblErrMsg = DisplayMsg(1201)

    
End Sub

Private Sub up_DeleteKontainer()
    Dim cmd As ADODB.Command
    Dim prm1 As ADODB.Parameter
    Dim prm2 As ADODB.Parameter
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailKontainer_Del"
    
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NomorKontainer", adVarChar, adParamInput, 11, txtNomorKontainer1 & txtNomorKontainer2)
    cmd.Parameters.append prm2
    
    cmd.Execute
    
    up_GridLoadKontainer
    
    txtNomorKontainer1 = ""
    txtNomorKontainer2 = ""
    cboUkuranKontainer = ""
    cboTipeKontainer = ""
    txtKeteranganKontainer = ""
  
    LblErrMsg = DisplayMsg(1201)
End Sub

Private Sub btnBC_Click()
    up_GetBC11
End Sub

Private Sub btnBrowseDokumen_Click()
frmBC23BrowseDokumen.txtNoAju = Replace(txtNoPengajuan, "-", "")
frmBC23BrowseDokumen.up_GridLoad
frmBC23BrowseDokumen.Show 1
up_GridLoadDokumen
up_LoadDataBC23 txtNoPengajuan
End Sub

Private Sub cboAsuransi_Change()
If cboAsuransi <> "" Then
    If Trim(Split(cboAsuransi, "-")(0)) = "1" Then
        txtAsuransi.Text = "0.00"
        txtAsuransi.BackColor = vbWhite
        txtAsuransi.locked = False
    ElseIf Trim(Split(cboAsuransi, "-")(0)) = "2" Then
        txtAsuransi.Text = "0.00"
        txtAsuransi.BackColor = &H80000000
        txtAsuransi.locked = True
    Else
        txtAsuransi.Text = "0.00"
        txtAsuransi.BackColor = &H80000000
        txtAsuransi.locked = True
    End If
'    txtCIF = Format((CDbl(txtAsuransi) + CDbl(txtFreightPIB)) + CDbl(txtFOBPIB), "#,0.00")
'    txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
    
    txtAsuransi = Format(txtAsuransi, "#,0.00")
    
End If

End Sub

Private Sub cboKodeHarga_Change()
    If Trim(Split(cboKodeHarga, "-")(0)) = "CIF" Then
        cboAsuransi.locked = True
        cboAsuransi = ""
        txtFreightPIB.locked = True
        txtFreightPIB.BackColor = &H80000000
        cboAsuransi.BackColor = &H80000000
        Label53.Caption = "CIF"
        Label58.Caption = "FOB"
        txtFOBPIB = "0.00"
        
        
        txtCIF = Format((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan)) - CDbl(txtDiskon), "#,0.00")
        txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
        
    ElseIf Trim(Split(cboKodeHarga, "-")(0)) = "CFR" Then
        cboAsuransi.locked = False
        cboAsuransi.ListIndex = 1
        cboAsuransi.BackColor = vbWhite
        txtFreightPIB.locked = True
        txtFreightPIB.BackColor = &H80000000
        Label53.Caption = "CNF"
        Label58.Caption = "CNF"
        
        txtFOBPIB = Format((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan)) - CDbl(txtDiskon), "#,0.00")
        
        txtCIF = Format((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan)) - CDbl(txtDiskon), "#,0.00")
        txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
        
    Else
        cboAsuransi.BackColor = vbWhite
        cboAsuransi.locked = False
        txtFreightPIB.locked = False
        cboAsuransi.ListIndex = 1
        txtFreightPIB.BackColor = vbWhite
        Label53.Caption = "FOB"
        Label58.Caption = "FOB"
        
        txtFOBPIB = Format((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan)) - CDbl(txtDiskon), "#,0.00")
        
        txtCIF = Format((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan)) - CDbl(txtDiskon), "#,0.00")
        txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
    End If
    
'        txtCIF = Format((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan)) - CDbl(txtDiskon), "#,0.00")
'        txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
        
End Sub

Private Sub cmdAction_Click(Index As Integer)
If Index = 0 Then
    FrmBC23List.Show
    Unload Me
ElseIf Index = 1 Then
    If uf_ValidateInput = False Then Exit Sub
    
    
    up_SaveDetailBC23
    up_SaveHarga
    
    up_LoadDataBC23 Replace(txtNoPengajuan, "-", "")
ElseIf Index = 2 Then
    If MsgBox("Are you sure want to synchronize the data?", vbYesNo + vbExclamation, "Delete") = vbYes Then
        up_Syncronize
    End If

End If
End Sub
    
Private Sub up_OpenDBMysql()
Dim ConnString As String
Dim db_name As String
Dim db_server As String
Dim db_port As String
Dim db_user As String
Dim db_pass As String
'Dim Conn As New ADODB.Connection
'//error traping
On Error GoTo buat_koneksi_Error
'/isi variable
db_name = "tpbdb"
db_server = "localhost"
db_port = "3306"
db_user = "beacukai"
db_pass = "beacukai"
'/buat connection string
ConnString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_user & ";PWD=" & db_pass & ";PORT=" & db_port & ""
'/buka koneksi
With DbMy
    .ConnectionString = ConnString
    .Open
   'MsgBox "Koneksi Berhasil"
   checkOpenDB = True
End With
'_____________________
On Error GoTo 0
Exit Sub

buat_koneksi_Error:
    checkOpenDB = False
    MsgBox "Ada kesalahan dengan server, periksa apakah server sudah berjalan !", vbInformation, "Cek Server"
End Sub

Private Sub up_CloseDBMysql()
On Error GoTo buat_koneksi_Error
    DbMy.Close
    Exit Sub
buat_koneksi_Error:
    MsgBox "Ada kesalahan dengan server, periksa apakah server sudah berjalan !", vbInformation, "Cek Server"
End Sub

Private Sub up_Syncronize()

    up_OpenDBMysql
    If checkOpenDB = True Then
        
        up_SaveTPBHeaderMy
        up_SaveTPBBarangMy
        up_SaveTPBDokumenMy
        up_SaveTPBKemasanMy
        up_SaveTPBBarangDokumenMy
        up_SaveTPBKontainerMy
        up_SaveTPBBarangTarifMy
        
        If checkOKToMysql = True Then LblErrMsg = DisplayMsg(1101)
    End If
    
    up_CloseDBMysql
End Sub

Private Sub up_SaveTPBBarangMy()
Dim lirow As Integer
Dim sql As String
Dim rs1 As New Recordset
Dim rs2 As New Recordset



Dim ls_KodeBarang As String
Dim ls_Uraian As String
Dim ls_Merk As String
Dim ls_KodeKategori As String
Dim ls_Tipe As String
Dim ls_Ukuran As String
Dim ls_SpesifikasiLain As String
Dim ls_NomorHS As String
Dim ls_KodeSatuan As String
Dim ls_KodeKemasan As String
Dim ls_KodeNegara As String
Dim ls_KodeFasilitas As String
Dim ls_KodeSkemaTarif As String
Dim ls_JumlahSatuan As Double
Dim ls_JumlahKemasan As Double
Dim ls_JumlahHarga As Double
Dim ls_FOB As Double
Dim ls_HargaInvoice As Double
Dim ls_Diskon As Double
Dim ls_Asuransi As Double
Dim ls_CIF As Double
Dim ls_CIFRupiah As Double

Dim rsHeader As New Recordset
Dim ls_IDHeader As String

sql = "Select * From TPB_Header WHERE NOMOR_AJU = '" & Replace(txtNoPengajuan, "-", "") & "'"
rsHeader.Open sql, DbMy, adOpenDynamic, adLockOptimistic

If Not rsHeader.EOF Then
    ls_IDHeader = rsHeader.Fields("ID")
End If
rsHeader.Close


With gridBarang
    For lirow = 1 To .Rows - 1
        
        sql = "Select * From Bea_Cukai_TPB_Barang WHERE SERI_BARANG = " & .TextMatrix(lirow, colHideNoSeri) & " AND NO_PENGAJUAN = '" & Replace(txtNoPengajuan, "-", "") & "'"
        rs1.Open sql, Db, adOpenDynamic, adLockOptimistic
        
        If Not rs1.EOF Then
            
            ls_KodeBarang = IIf(IsNull(rs1.Fields("KODE_BARANG")), "", rs1.Fields("KODE_BARANG"))
            ls_Uraian = IIf(IsNull(rs1.Fields("URAIAN")), "", rs1.Fields("URAIAN"))
            ls_Merk = IIf(IsNull(rs1.Fields("MERK")), "", rs1.Fields("MERK"))
            ls_Tipe = IIf(IsNull(rs1.Fields("TIPE")), "", rs1.Fields("TIPE"))
            ls_KodeKategori = IIf(IsNull(rs1.Fields("KATEGORI_BARANG")), "", rs1.Fields("KATEGORI_BARANG"))
            ls_Ukuran = IIf(IsNull(rs1.Fields("UKURAN")), "", rs1.Fields("UKURAN"))
            ls_SpesifikasiLain = IIf(IsNull(rs1.Fields("SPESIFIKASI_LAIN")), "", rs1.Fields("SPESIFIKASI_LAIN"))
            ls_NomorHS = IIf(IsNull(rs1.Fields("POS_TARIF")), "", rs1.Fields("POS_TARIF"))
            ls_KodeSatuan = IIf(IsNull(rs1.Fields("KODE_SATUAN")), "", rs1.Fields("KODE_SATUAN"))
            ls_KodeKemasan = IIf(IsNull(rs1.Fields("KODE_KEMASAN")), "", rs1.Fields("KODE_KEMASAN"))
            ls_KodeNegara = IIf(IsNull(rs1.Fields("KODE_NEGARA_ASAL")), "", rs1.Fields("KODE_NEGARA_ASAL"))
            ls_KodeFasilitas = IIf(IsNull(rs1.Fields("KODE_FASILITAS_DOKUMEN")), "", rs1.Fields("KODE_FASILITAS_DOKUMEN"))
            ls_KodeSkemaTarif = IIf(IsNull(rs1.Fields("KODE_SKEMA_TARIF")), "", rs1.Fields("KODE_SKEMA_TARIF"))
            
            ls_JumlahSatuan = IIf(IsNull(rs1.Fields("JUMLAH_SATUAN")), 0, rs1.Fields("JUMLAH_SATUAN"))
            ls_JumlahKemasan = IIf(IsNull(rs1.Fields("JUMLAH_KEMASAN")), 0, rs1.Fields("JUMLAH_KEMASAN"))
            ls_JumlahHarga = IIf(IsNull(rs1.Fields("HARGA_SATUAN")), 0, rs1.Fields("HARGA_SATUAN"))
            ls_FOB = IIf(IsNull(rs1.Fields("FOB")), 0, rs1.Fields("FOB"))
            ls_HargaInvoice = IIf(IsNull(rs1.Fields("HARGA_INVOICE")), 0, rs1.Fields("HARGA_INVOICE"))
            ls_Diskon = IIf(IsNull(rs1.Fields("DISKON")), 0, rs1.Fields("DISKON"))
            ls_Asuransi = IIf(IsNull(rs1.Fields("ASURANSI")), 0, rs1.Fields("ASURANSI"))
            ls_CIF = IIf(IsNull(rs1.Fields("CIF")), 0, rs1.Fields("CIF"))
            ls_CIFRupiah = IIf(IsNull(rs1.Fields("CIF_RUPIAH")), 0, rs1.Fields("CIF_RUPIAH"))
        
            sql = "Select * From TPB_Barang WHERE SERI_BARANG = " & .TextMatrix(lirow, colHideNoSeri) & " AND ID_HEADER = " & ls_IDHeader & ""
            rs2.Open sql, DbMy, adOpenDynamic, adLockOptimistic
            
            If rs2.EOF Then
                sql = "     INSERT INTO TPB_Barang " & vbCrLf & _
                            "   (ID_HEADER, SERI_BARANG, KODE_BARANG, URAIAN, MERK, TIPE, SPESIFIKASI_LAIN, UKURAN, " & vbCrLf & _
                            "   POS_TARIF, KATEGORI_BARANG, KODE_SATUAN, KODE_KEMASAN,  " & vbCrLf & _
                            "   KODE_NEGARA_ASAL, KODE_FASILITAS_DOKUMEN, KODE_SKEMA_TARIF, " & vbCrLf & _
                            "   JUMLAH_SATUAN, JUMLAH_KEMASAN, HARGA_SATUAN, FOB, " & vbCrLf & _
                            "   HARGA_INVOICE, DISKON, ASURANSI, CIF, CIF_RUPIAH " & vbCrLf & _
                            "   ) " & vbCrLf & _
                            "   VALUES  " & vbCrLf & _
                            "   ('" & ls_IDHeader & "', '" & .TextMatrix(lirow, colHideNoSeri) & "', '" & ls_KodeBarang & "', '" & ls_Uraian & "', '" & ls_Merk & "', '" & ls_Tipe & "', '" & ls_SpesifikasiLain & "', '" & ls_Ukuran & "', " & vbCrLf & _
                            "   '" & ls_NomorHS & "', '" & ls_KodeKategori & "', '" & ls_KodeSatuan & "', '" & ls_KodeKemasan & "', " & vbCrLf & _
                            "   '" & ls_KodeNegara & "', '" & ls_KodeFasilitas & "', '" & ls_KodeSkemaTarif & "', "
                
                sql = sql + "   " & ls_JumlahSatuan & ", " & ls_JumlahKemasan & ", " & ls_JumlahHarga & ", " & ls_FOB & ", " & vbCrLf & _
                            "   " & ls_HargaInvoice & ", " & ls_Diskon & ", " & ls_CIF & ", " & ls_CIFRupiah & ", " & ls_CIFRupiah & " " & vbCrLf & _
                            "   ) " & vbCrLf & _
                            "  "
    
            Else
                sql = "     UPDATE TPB_Barang " & vbCrLf & _
                            "   SET KODE_BARANG = '" & ls_KodeBarang & "', " & vbCrLf & _
                            "       URAIAN = '" & ls_Uraian & "', " & vbCrLf & _
                            "       MERK = '" & ls_Merk & "', " & vbCrLf & _
                            "       TIPE = '" & ls_Tipe & "', " & vbCrLf & _
                            "       SPESIFIKASI_LAIN = '" & ls_SpesifikasiLain & "', " & vbCrLf & _
                            "       POS_TARIF = '" & ls_NomorHS & "', " & vbCrLf & _
                            "       KATEGORI_BARANG = '" & ls_KodeKategori & "', " & vbCrLf & _
                            "       KODE_SATUAN = '" & ls_KodeSatuan & "', " & vbCrLf & _
                            "       KODE_KEMASAN = '" & ls_KodeKemasan & "', " & vbCrLf & _
                            "       KODE_NEGARA_ASAL = '" & ls_KodeNegara & "', "
                
                sql = sql + "       KODE_FASILITAS_DOKUMEN = '" & ls_KodeFasilitas & "', " & vbCrLf & _
                            "       KODE_SKEMA_TARIF = '" & ls_KodeSkemaTarif & "', " & vbCrLf & _
                            "       JUMLAH_SATUAN = " & ls_JumlahSatuan & ", " & vbCrLf & _
                            "       JUMLAH_KEMASAN = " & ls_JumlahKemasan & ", " & vbCrLf & _
                            "       HARGA_SATUAN = " & ls_JumlahHarga & ", " & vbCrLf & _
                            "       FOB = " & ls_FOB & ", " & vbCrLf & _
                            "       HARGA_INVOICE = " & ls_HargaInvoice & ", " & vbCrLf & _
                            "       DISKON = " & ls_Diskon & ", " & vbCrLf & _
                            "       ASURANSI = " & ls_Asuransi & ", " & vbCrLf & _
                            "       CIF = " & ls_CIF & ", " & vbCrLf & _
                            "       CIF_RUPIAH = " & ls_CIFRupiah & " "
                
                sql = sql + "   WHERE SERI_BARANG = " & .TextMatrix(lirow, colHideNoSeri) & " AND ID_HEADER = " & ls_IDHeader & " " & vbCrLf & _
                            "  "
    
            End If
            
            DbMy.Execute sql
        End If
        
        rs1.Close
        rs2.Close
    Next
End With
End Sub

Private Sub up_SaveTPBHeaderMy()
Dim sql As String
Dim RS As New Recordset

On Error GoTo errHandler

sql = "Select * From TPB_Header WHERE NOMOR_AJU = '" & Replace(txtNoPengajuan, "-", "") & "'"
RS.Open sql, DbMy, adOpenDynamic, adLockOptimistic
    
If Not RS.EOF Then
    
    sql = " UPDATE TPB_Header " & vbCrLf & _
                    " SET   KODE_TUJUAN_TPB = '" & Left(cboTujuan, 1) & "', " & vbCrLf & _
                    "   NAMA_TTD = '" & txtPemberitahu & "', " & vbCrLf & _
                    "   JABATAN_TTD = '" & txtJabatan & "', " & vbCrLf & _
                    "   KOTA_TTD = '" & txtTempat & "', " & vbCrLf & _
                    "   KODE_KANTOR_BONGKAR = '" & txtKPBBCBongkar & "', " & vbCrLf & _
                    "   TANGGAL_TTD = '" & Format(dtpTanggal, "yyyy-MM-dd") & "', " & vbCrLf & _
                    "   KODE_KANTOR = '" & txtKPBBCPengawas & "', " & vbCrLf & _
                    "   NAMA_PEMASOK = '" & txtNamaPemasok & "', " & vbCrLf & _
                    "   ALAMAT_PEMASOK = '" & txtAlamatPemasok & "', " & vbCrLf & _
                    "   KODE_NEGARA_PEMASOK = '" & txtNegara & "', "
        
        sql = sql + "   KODE_ID_PENGUSAHA = '1', " & vbCrLf & _
                    "   ID_PENGUSAHA = '" & Replace(Replace(txtIdentitasImportir, ".", ""), "-", "") & "', " & vbCrLf & _
                    "   NAMA_PENGUSAHA = '" & txtNamaImportir & "', " & vbCrLf & _
                    "   NOMOR_IJIN_TPB = '" & txtNoIzin & "', " & vbCrLf & _
                    "   ALAMAT_PENGUSAHA = '" & txtAlamatImportir & "', " & vbCrLf & _
                    "   KODE_JENIS_API_PENGUSAHA = '" & Left(cboImportirAPI, 1) & "', " & vbCrLf & _
                    "   API_PENGUSAHA = '" & txtAPIImportir & "', " & vbCrLf & _
                    "   KODE_ID_PEMILIK = '" & Left(cboIDPemilik, 1) & "', " & vbCrLf & _
                    "   ID_PEMILIK = '" & Replace(Replace(txtIdentitasPemilik, ".", ""), "-", "") & "', " & vbCrLf & _
                    "   NAMA_PEMILIK = '" & txtNamaPemilik & "',  " & vbCrLf & _
                    "   ALAMAT_PEMILIK = '" & txtAlamatPemilik & "', " & vbCrLf & _
                    "   KODE_JENIS_API_PEMILIK = '" & Left(cboPemilikAPI, 1) & "', " & vbCrLf & _
                    "   API_PEMILIK = '" & txtAPIPemilik & "', " & vbCrLf & _
                    "   KODE_CARA_ANGKUT = '" & Left(cboCaraAngkut, 1) & "', " & vbCrLf & _
                    "   NAMA_PENGANGKUT = '" & txtNamaPengangkut & "', "
        
        sql = sql + "   NOMOR_VOY_FLIGHT = '" & txtVoyFlight & "', " & vbCrLf & _
                    "   KODE_BENDERA = '" & txtNegaraPengangkut & "', " & vbCrLf & _
                    "   KODE_PEL_MUAT = '" & txtPelabuhanMuat & "', " & vbCrLf & _
                    "   KODE_PEL_TRANSIT = '" & txtPelabuhanTransit & "', " & vbCrLf & _
                    "   KODE_PEL_BONGKAR = '" & txtPelabuhanBongkar & "', " & vbCrLf & _
                    "   POS_BC11 = '" & txtPos1 & "', " & vbCrLf & _
                    "   SERI = '0', " & vbCrLf & _
                    "   SUBPOS_BC11 = '" & txtPos2 & "', " & vbCrLf & _
                    "   SUBSUBPOS_BC11 = '" & txtPos3 & "', " & vbCrLf & _
                    "   NOMOR_BC11 = '" & txtNomorBC11 & "', " & vbCrLf & _
                    "   KODE_TPS = '" & txtKodePenimbunan & "', " & vbCrLf & _
                    "   KODE_VALUTA = '" & txtValuta & "', " & vbCrLf & _
                    "    "
        
        sql = sql + "   BRUTO = " & CDbl(txtBrutoBarang) & ", " & vbCrLf & _
                    "   NETTO = " & CDbl(txtNettoBarang) & ", " & vbCrLf & _
                    "   FOB = " & CDbl(txtFOB) & ", " & vbCrLf & _
                    "   FREIGHT = " & CDbl(txtFreight) & ", " & vbCrLf & _
                    "   HARGA_INVOICE = " & CDbl(txtNilaiCIF) & ", " & vbCrLf & _
                    "   ASURANSI = " & CDbl(txtAsuransiLNDN) & ", " & vbCrLf & _
                    "   CIF = " & CDbl(txtNilaiCIF) & ", " & vbCrLf & _
                    "   CIF_RUPIAH = " & CDbl(txtNilaiCIFRupiah) & ", " & vbCrLf & _
                    "   NDPBM = " & CDbl(txtNDPBM) & " " & vbCrLf & _
                    " WHERE NOMOR_AJU = '" & Replace(txtNoPengajuan, "-", "") & "' " & vbCrLf & _
                    "  "
          
Else
    sql = " Insert Into TPB_Header " & vbCrLf & _
                " (VERSI_MODUL,  " & vbCrLf & _
                " ID_MODUL,  " & vbCrLf & _
                " NOMOR_AJU,  " & vbCrLf & _
                " KODE_TUJUAN_TPB,  " & vbCrLf & _
                " NAMA_TTD,  " & vbCrLf & _
                " JABATAN_TTD,  " & vbCrLf & _
                " KOTA_TTD,  " & vbCrLf & _
                " KODE_KANTOR_BONGKAR,  " & vbCrLf & _
                " TANGGAL_TTD,  " & vbCrLf & _
                " KODE_KANTOR,  "
    
    sql = sql + " NAMA_PEMASOK,  " & vbCrLf & _
                " ALAMAT_PEMASOK, " & vbCrLf & _
                " KODE_NEGARA_PEMASOK,  " & vbCrLf & _
                " KODE_ID_PENGUSAHA,  " & vbCrLf & _
                " ID_PENGUSAHA,  " & vbCrLf & _
                " NAMA_PENGUSAHA, " & vbCrLf & _
                " NOMOR_IJIN_TPB, " & vbCrLf & _
                " ALAMAT_PENGUSAHA, " & vbCrLf & _
                " KODE_JENIS_API_PENGUSAHA, " & vbCrLf & _
                " API_PENGUSAHA, " & vbCrLf & _
                " KODE_ID_PEMILIK,  " & vbCrLf & _
                " ID_PEMILIK, " & vbCrLf & _
                " NAMA_PEMILIK, " & vbCrLf & _
                " ALAMAT_PEMILIK, "
    
    sql = sql + " KODE_JENIS_API_PEMILIK, " & vbCrLf & _
                " API_PEMILIK, " & vbCrLf & _
                " KODE_CARA_ANGKUT, " & vbCrLf & _
                " NAMA_PENGANGKUT, " & vbCrLf & _
                " NOMOR_VOY_FLIGHT, " & vbCrLf & _
                " KODE_BENDERA, " & vbCrLf & _
                " KODE_PEL_MUAT, " & vbCrLf & _
                " KODE_PEL_TRANSIT, " & vbCrLf & _
                " KODE_PEL_BONGKAR, " & vbCrLf & _
                " POS_BC11, " & vbCrLf & _
                " SERI, " & vbCrLf & _
                " SUBPOS_BC11, "
    
    sql = sql + " SUBSUBPOS_BC11, " & vbCrLf & _
                " NOMOR_BC11, " & vbCrLf & _
                " KODE_TPS, " & vbCrLf & _
                " BRUTO, " & vbCrLf & _
                " NETTO, " & vbCrLf & _
                " KODE_VALUTA, " & vbCrLf & _
                " FOB, " & vbCrLf & _
                " FREIGHT, " & vbCrLf & _
                " ASURANSI, " & vbCrLf & _
                " HARGA_INVOICE, " & vbCrLf & _
                " CIF, " & vbCrLf & _
                " CIF_RUPIAH, " & vbCrLf & _
                " NDPBM " & vbCrLf & _
                " ) " & vbCrLf & _
                " VALUES " & vbCrLf & _
                " ('3.1.8',  " & vbCrLf & _
                " '10372',  " & vbCrLf & _
                " '" & Replace(txtNoPengajuan, "-", "") & "',  " & vbCrLf & _
                " '" & Left(cboTujuan, 1) & "',  "
    
    sql = sql + " '" & txtPemberitahu & "',  " & vbCrLf & _
                " '" & txtJabatan & "',  " & vbCrLf & _
                " '" & txtTempat & "',  " & vbCrLf & _
                " '" & txtKPBBCBongkar & "',  " & vbCrLf & _
                " '" & Format(dtpTanggal, "yyyy-MM-dd") & "',  " & vbCrLf & _
                " '" & txtKPBBCPengawas & "', " & vbCrLf & _
                " '" & txtNamaPemasok & "',  " & vbCrLf & _
                " '" & txtAlamatPemasok & "',  " & vbCrLf & _
                " '" & txtNegara & "', " & vbCrLf & _
                " '1',  " & vbCrLf & _
                " '" & txtIdentitasImportir & "',  " & vbCrLf & _
                " '" & txtNamaImportir & "', "
    
    sql = sql + " '" & txtNoIzin & "', " & vbCrLf & _
                " '" & txtAlamatImportir & "', " & vbCrLf & _
                " '" & Left(cboImportirAPI, 1) & "', " & vbCrLf & _
                " '" & txtAPIImportir & "', " & vbCrLf & _
                " '" & Left(cboIDPemilik, 1) & "', " & vbCrLf & _
                " '" & txtIdentitasPemilik & "', " & vbCrLf & _
                " '" & txtNamaPemilik & "', " & vbCrLf & _
                " '" & txtAlamatPemilik & "', " & vbCrLf & _
                " '" & Left(cboPemilikAPI, 1) & "', " & vbCrLf & _
                " '" & txtAPIPemilik & "', " & vbCrLf & _
                " '" & Left(cboCaraAngkut, 1) & "', " & vbCrLf & _
                " '" & txtNamaPengangkut & "', "
    
    sql = sql + " '" & txtVoyFlight & "', " & vbCrLf & _
                " '" & txtNegaraPengangkut & "', " & vbCrLf & _
                " '" & txtPelabuhanMuat & "', " & vbCrLf & _
                " '" & txtPelabuhanTransit & "', " & vbCrLf & _
                " '" & txtPelabuhanBongkar & "', " & vbCrLf & _
                " '" & txtPos1 & "', " & vbCrLf & _
                " '0', " & vbCrLf & _
                " '" & txtPos2 & "', " & vbCrLf & _
                " '" & txtPos3 & "', " & vbCrLf & _
                " '" & txtNomorBC11 & "', " & vbCrLf & _
                " '" & txtKodePenimbunan & "', " & vbCrLf & _
                " " & CDbl(txtBrutoBarang) & ", " & vbCrLf & _
                " " & CDbl(txtNettoBarang) & ", "
               
    sql = sql + " '" & txtValuta & "', " & vbCrLf & _
                "  " & CDbl(txtFOB) & ", " & vbCrLf & _
                "  " & CDbl(txtFreight) & ", " & vbCrLf & _
                "  " & CDbl(txtAsuransiLNDN) & ", " & vbCrLf & _
                "  " & CDbl(txtNilaiCIF) & ", " & vbCrLf & _
                "  " & CDbl(txtNilaiCIF) & ", " & vbCrLf & _
                "  " & CDbl(txtNilaiCIFRupiah) & ", " & vbCrLf & _
                "  " & CDbl(txtNDPBM) & " " & vbCrLf & _
                " ) " & vbCrLf & _
                "  "
End If

DbMy.Execute sql


checkOKToMysql = True
Exit Sub
errHandler:
    LblErrMsg.Caption = err.Description
    checkOKToMysql = False
End Sub

Private Sub up_SaveTPBDokumenMy()
Dim rsHeader As New Recordset
Dim sql As String
Dim ls_IDHeader As String

sql = "Select * From TPB_Header WHERE NOMOR_AJU = '" & Replace(txtNoPengajuan, "-", "") & "'"
rsHeader.Open sql, DbMy, adOpenDynamic, adLockOptimistic

If Not rsHeader.EOF Then
    ls_IDHeader = rsHeader.Fields("ID")
End If
rsHeader.Close

Dim rs1 As New Recordset
Dim rs2 As New Recordset

sql = "Select * From Bea_Cukai_TPB_Dokumen WHERE NO_PENGAJUAN = '" & Replace(txtNoPengajuan, "-", "") & "'"
rs1.Open sql, Db, adOpenDynamic, adLockOptimistic

While Not rs1.EOF
    
    sql = "Select * From TPB_Dokumen WHERE ID_Header = '" & ls_IDHeader & "' AND SERI_DOKUMEN = " & rs1.Fields("SERI_DOKUMEN") & ""
    rs2.Open sql, DbMy, adOpenDynamic, adLockOptimistic
    
    If Not rs2.EOF Then
        sql = " Update TPB_Dokumen  " & vbCrLf & _
                    " Set Kode_Jenis_Dokumen = '" & rs1.Fields("Kode_Jenis_Dokumen") & "',  " & vbCrLf & _
                    "   Nomor_Dokumen = '" & rs1.Fields("Nomor_Dokumen") & "', " & vbCrLf & _
                    "   Tanggal_Dokumen = '" & rs1.Fields("Tanggal_Dokumen") & "', " & vbCrLf & _
                    "   Tipe_Dokumen = '" & rs1.Fields("Tipe_Dokumen") & "' " & vbCrLf & _
                    " Where Seri_Dokumen = " & rs1.Fields("SERI_DOKUMEN") & " AND ID_Header = " & ls_IDHeader & " " & vbCrLf & _
                    "  "
        
    Else
        sql = " Insert Into TPB_Dokumen " & vbCrLf & _
                    " (Seri_Dokumen, ID_Header, Kode_Jenis_Dokumen, Nomor_Dokumen, Tanggal_Dokumen, Tipe_Dokumen) " & vbCrLf & _
                    " Values " & vbCrLf & _
                    " (" & rs1.Fields("SERI_DOKUMEN") & ", " & ls_IDHeader & ", '" & rs1.Fields("Kode_Jenis_Dokumen") & "', '" & rs1.Fields("Nomor_Dokumen") & "', '" & rs1.Fields("Tanggal_Dokumen") & "', '" & rs1.Fields("Tipe_Dokumen") & "') " & vbCrLf & _
                    "  "
    End If
    
    DbMy.Execute sql
    
    
    rs2.Close
    
    rs1.MoveNext
Wend
rs1.Close

checkOKToMysql = True
Exit Sub
errHandler:
    LblErrMsg.Caption = err.Description
    checkOKToMysql = False
    
End Sub

Private Sub up_SaveTPBKemasanMy()
Dim rsHeader As New Recordset
Dim sql As String
Dim ls_IDHeader As String

On Error GoTo errHandler

sql = "Select * From TPB_Header WHERE NOMOR_AJU = '" & Replace(txtNoPengajuan, "-", "") & "'"
rsHeader.Open sql, DbMy, adOpenDynamic, adLockOptimistic

If Not rsHeader.EOF Then
    ls_IDHeader = rsHeader.Fields("ID")
End If
rsHeader.Close

Dim rs1 As New Recordset
Dim rs2 As New Recordset

sql = "Select * From Bea_Cukai_TPB_Kemasan WHERE NO_PENGAJUAN = '" & Replace(txtNoPengajuan, "-", "") & "'"
rs1.Open sql, Db, adOpenDynamic, adLockOptimistic

While Not rs1.EOF

    sql = "Select * From TPB_Kemasan WHERE ID_Header = " & ls_IDHeader & " AND KODE_JENIS_KEMASAN = '" & rs1.Fields("KODE_JENIS_KEMASAN") & "'"
    rs2.Open sql, DbMy, adOpenDynamic, adLockOptimistic
    
    If Not rs2.EOF Then
        sql = " UPDATE TPB_Kemasan " & vbCrLf & _
                    " SET JUMLAH_KEMASAN = " & rs1.Fields("JUMLAH_KEMASAN") & ", MERK_KEMASAN = '" & rs1.Fields("MERK_KEMASAN") & "' " & vbCrLf & _
                    " WHERE KODE_JENIS_KEMASAN = '" & rs1.Fields("KODE_JENIS_KEMASAN") & "' AND ID_HEADER = " & ls_IDHeader & " " & vbCrLf & _
                    "  " & vbCrLf & _
                    "  "
    Else
        sql = " INSERT INTO TPB_Kemasan " & vbCrLf & _
                    " (ID_HEADER,KODE_JENIS_KEMASAN,JUMLAH_KEMASAN,MERK_KEMASAN ) " & vbCrLf & _
                    " VALUES " & vbCrLf & _
                    " (" & ls_IDHeader & ", '" & rs1.Fields("KODE_JENIS_KEMASAN") & "', " & rs1.Fields("JUMLAH_KEMASAN") & ", '" & rs1.Fields("MERK_KEMASAN") & "') " & vbCrLf & _
                    "  " & vbCrLf & _
                    "  " & vbCrLf & _
                    "  "
    End If

    DbMy.Execute sql
    rs2.Close
rs1.MoveNext
Wend
rs1.Close

checkOKToMysql = True
Exit Sub
errHandler:
    LblErrMsg.Caption = err.Description
    checkOKToMysql = False
End Sub

Private Sub up_SaveTPBKontainerMy()
Dim rsHeader As New Recordset
Dim sql As String
Dim ls_IDHeader As String

On Error GoTo errHandler

sql = "Select * From TPB_Header WHERE NOMOR_AJU = '" & Replace(txtNoPengajuan, "-", "") & "'"
rsHeader.Open sql, DbMy, adOpenDynamic, adLockOptimistic

If Not rsHeader.EOF Then
    ls_IDHeader = rsHeader.Fields("ID")
End If
rsHeader.Close

Dim rs1 As New Recordset
Dim rs2 As New Recordset

sql = "Select * From Bea_Cukai_TPB_Kontainer WHERE NO_PENGAJUAN = '" & Replace(txtNoPengajuan, "-", "") & "'"
rs1.Open sql, Db, adOpenDynamic, adLockOptimistic

While Not rs1.EOF

    sql = "Select * From TPB_Kontainer WHERE ID_Header = " & ls_IDHeader & " AND NOMOR_KONTAINER = '" & rs1.Fields("NOMOR_KONTAINER") & "'"
    rs2.Open sql, DbMy, adOpenDynamic, adLockOptimistic

    If Not rs2.EOF Then
        sql = " UPDATE TPB_Kontainer " & vbCrLf & _
                    " SET KETERANGAN = '" & rs1.Fields("KETERANGAN") & "', KODE_TIPE_KONTAINER = '" & rs1.Fields("KODE_TIPE_KONTAINER") & "', KODE_UKURAN_KONTAINER = '" & rs1.Fields("KODE_UKURAN_KONTAINER") & "' " & vbCrLf & _
                    " WHERE NOMOR_KONTAINER = '" & rs1.Fields("NOMOR_KONTAINER") & "' AND ID_HEADER = " & ls_IDHeader & " " & vbCrLf & _
                    "  " & vbCrLf & _
                    "  "
    Else
        sql = " INSERT INTO TPB_Kontainer " & vbCrLf & _
                    " (ID_HEADER, NOMOR_KONTAINER, KODE_TIPE_KONTAINER, KODE_UKURAN_KONTAINER, KETERANGAN ) " & vbCrLf & _
                    " VALUES " & vbCrLf & _
                    " (" & ls_IDHeader & ", '" & rs1.Fields("NOMOR_KONTAINER") & "', '" & rs1.Fields("KODE_TIPE_KONTAINER") & "', '" & rs1.Fields("KODE_UKURAN_KONTAINER") & "', '" & rs1.Fields("KETERANGAN") & "') " & vbCrLf & _
                    "  "

    End If
    DbMy.Execute sql
    rs2.Close

rs1.MoveNext
Wend
rs1.Close

checkOKToMysql = True
Exit Sub
errHandler:
    LblErrMsg.Caption = err.Description
    checkOKToMysql = False
End Sub

Private Sub up_SaveTPBBarangTarifMy()
Dim rsHeader As New Recordset
Dim sql As String
Dim ls_IDHeader As String

On Error GoTo errHandler

sql = "Select * From TPB_Header WHERE NOMOR_AJU = '" & Replace(txtNoPengajuan, "-", "") & "'"
rsHeader.Open sql, DbMy, adOpenDynamic, adLockOptimistic

If Not rsHeader.EOF Then
    ls_IDHeader = rsHeader.Fields("ID")
End If
rsHeader.Close

Dim rs1 As New Recordset
Dim rs2 As New Recordset

sql = "DELETE FROM TPB_Barang_Tarif WHERE ID_HEADER = " & ls_IDHeader & ""
DbMy.Execute sql

sql = "Select * From Bea_Cukai_TPB_Barang_Tarif WHERE NO_PENGAJUAN = '" & Replace(txtNoPengajuan, "-", "") & "'"
rs1.Open sql, Db, adOpenDynamic, adLockOptimistic

While Not rs1.EOF

    sql = "Select * From TPB_Barang WHERE ID_HEADER = " & ls_IDHeader & " AND SERI_BARANG = " & rs1.Fields("SERI_BARANG") & ""
    rs2.Open sql, DbMy, adOpenDynamic, adLockOptimistic

    If Not rs2.EOF Then
    
        sql = " INSERT INTO TPB_Barang_Tarif " & vbCrLf & _
                    " (FLAG_BMT_SEMENTARA, JENIS_TARIF, JUMLAH_SATUAN,  " & vbCrLf & _
                    " KODE_FASILITAS, KODE_KOMODITI_CUKAI, KODE_SATUAN, " & vbCrLf & _
                    " KODE_SUB_KOMODITI_CUKAI, KODE_TARIF, NILAI_BAYAR, " & vbCrLf & _
                    " NILAI_FASILITAS, NILAI_SUDAH_DILUNASI, SERI_BARANG, " & vbCrLf & _
                    " TARIF, TARIF_FASILITAS, ID_BARANG, ID_HEADER " & vbCrLf & _
                    " ) " & vbCrLf & _
                    " VALUES " & vbCrLf & _
                    " ('" & rs1.Fields("FLAG_BMT_SEMENTARA") & "', '" & rs1.Fields("JENIS_TARIF") & "', '" & rs1.Fields("JUMLAH_SATUAN") & "',  " & vbCrLf & _
                    " '" & rs1.Fields("KODE_FASILITAS") & "', '" & rs1.Fields("KODE_KOMODITI_CUKAI") & "', '" & rs1.Fields("KODE_SATUAN") & "', " & vbCrLf & _
                    " '" & rs1.Fields("KODE_SUB_KOMODITI_CUKAI") & "', '" & rs1.Fields("KODE_TARIF") & "', " & rs1.Fields("NILAI_BAYAR") & ", "
        
        sql = sql + " " & rs1.Fields("NILAI_FASILITAS") & ", '" & rs1.Fields("NILAI_SUDAH_DILUNASI") & "', " & rs1.Fields("SERI_BARANG") & ", " & vbCrLf & _
                    " " & rs1.Fields("TARIF") & ", " & rs1.Fields("TARIF_FASILITAS") & ", " & rs2.Fields("ID") & ", " & ls_IDHeader & " " & vbCrLf & _
                    " ) " & vbCrLf & _
                    "  " & vbCrLf & _
                    "  "

                    
        DbMy.Execute sql
    End If
    rs2.Close
rs1.MoveNext
Wend
rs1.Close

checkOKToMysql = True
Exit Sub
errHandler:
    LblErrMsg.Caption = err.Description
    checkOKToMysql = False

End Sub

Private Sub up_SaveTPBBarangDokumenMy()
Dim rsHeader As New Recordset
Dim sql As String
Dim ls_IDHeader As String

On Error GoTo errHandler

sql = "Select * From TPB_Header WHERE NOMOR_AJU = '" & Replace(txtNoPengajuan, "-", "") & "'"
rsHeader.Open sql, DbMy, adOpenDynamic, adLockOptimistic

If Not rsHeader.EOF Then
    ls_IDHeader = rsHeader.Fields("ID")
End If
rsHeader.Close

Dim rs1 As New Recordset
Dim rs2 As New Recordset

sql = "DELETE FROM TPB_Barang_Dokumen WHERE ID_HEADER = " & ls_IDHeader & ""
DbMy.Execute sql

sql = "Select * From Bea_Cukai_TPB_Barang_Dokumen WHERE NO_PENGAJUAN = '" & Replace(txtNoPengajuan, "-", "") & "'"
rs1.Open sql, Db, adOpenDynamic, adLockOptimistic

While Not rs1.EOF

    sql = "Select * From TPB_Barang WHERE ID_HEADER = " & ls_IDHeader & " AND SERI_BARANG = " & rs1.Fields("NO_SERI") & ""
    rs2.Open sql, DbMy, adOpenDynamic, adLockOptimistic

    If Not rs2.EOF Then
    
        sql = " INSERT INTO TPB_Barang_Dokumen " & vbCrLf & _
                    " (ID_HEADER,SERI_DOKUMEN,ID_BARANG) " & vbCrLf & _
                    " VALUES  " & vbCrLf & _
                    " (" & ls_IDHeader & ", '" & rs1.Fields("SERI_DOKUMEN") & "', '" & rs2.Fields("ID") & "') " & vbCrLf & _
                    "  "
                    
        DbMy.Execute sql
    End If
    rs2.Close
rs1.MoveNext
Wend
rs1.Close

checkOKToMysql = True
Exit Sub
errHandler:
    LblErrMsg.Caption = err.Description
    checkOKToMysql = False
End Sub


Private Sub cmdAddBarang_Click()
    up_LoadDataBC23 txtNoPengajuan
    
    If checkAlreadyData = False Then
        LblErrMsg.Caption = "Please save the data first!"
        Exit Sub
    End If
    
    If txtNDPBM = "" Or CDbl(txtNDPBM) = 0 Then
        txtNDPBMInvoice.SetFocus
        LblErrMsg = "Please Input NDPBM!"
        SSTab1.Tab = 1
        Exit Sub
    End If
    
    frmBC23BrowseBarang.txtNoPengajuan = Replace(txtNoPengajuan, "-", "")
    frmBC23BrowseBarang.txtNoSeri = (gridBarang.Rows - 1) + 1
    frmBC23BrowseBarang.txtCIFFix = CDbl(txtNDPBM)
    If CDbl(txtFreight) = 0 Then
        frmBC23BrowseBarang.txtFreightFix = 0
    Else
        frmBC23BrowseBarang.txtFreightFix = CDbl(txtFOB) / CDbl(txtFreight)
    End If
    frmBC23BrowseBarang.cekSubmit = False
    frmBC23BrowseBarang.Show 1
    
    up_GridLoadBarang
End Sub

Private Sub cmdCancelKemasan_Click()
    txtJumlahKemasan = ""
    txtJenisKemasan = ""
    txtMerkKemasan = ""
    lblJenisKemasan.Caption = ""
    txtJenisKemasan.Enabled = True
    txtJumlahKemasan.SetFocus
    
End Sub

Private Sub cmdCancelKontainer_Click()
txtNomorKontainer1 = ""
txtNomorKontainer2 = ""
cboUkuranKontainer = ""
cboTipeKontainer = ""
txtKeteranganKontainer = ""
txtIDKontainer = ""
End Sub

Private Sub cmdDeleteKemasan_Click()
    If txtJenisKemasan = "" Then
        LblErrMsg = "Please select kemasan"
        Exit Sub
    End If
    If MsgBox("Are you sure want to delete?", vbYesNo + vbExclamation, "Delete") = vbYes Then
        up_DeleteKemasan
    End If
End Sub

Private Sub cmdDeleteKontainer_Click()
    If txtJenisKemasan = "" Then
        LblErrMsg = "Please select kemasan"
        Exit Sub
    End If
    If MsgBox("Are you sure want to delete?", vbYesNo + vbExclamation, "Delete") = vbYes Then
        up_DeleteKontainer
    End If
End Sub

Private Sub cmdDetailBarang_Click()
    up_LoadDataBC23 txtNoPengajuan
    
    If checkAlreadyData = False Then
        LblErrMsg.Caption = "Please save the data first!"
        Exit Sub
    End If
    
    If txtNDPBM = "" Or CDbl(txtNDPBM) = 0 Then
        txtNDPBMInvoice.SetFocus
        LblErrMsg = "Please Input NDPBM!"
        Exit Sub
    End If
    
Dim strkode As String
    If gridBarang.RowSel <> 0 Then Exit Sub
    strkode = gridBarang.TextMatrix(gridBarang.RowSel, colKodeBarang)
    frmBC23BrowseBarang.txtNoSeri = gridBarang.TextMatrix(gridBarang.RowSel, colHideNoSeri)
    frmBC23BrowseBarang.txtNoPengajuan = Replace(txtNoPengajuan, "-", "")
    frmBC23BrowseBarang.cmdDelete.Enabled = True
    frmBC23BrowseBarang.up_LoadDataBarang txtNoPengajuan, gridBarang.TextMatrix(gridBarang.RowSel, colHideNoSeri)
    frmBC23BrowseBarang.txtCIFFix = CDbl(txtNDPBM)
    
    If CDbl(txtFreight) = 0 Then
        frmBC23BrowseBarang.txtFreightFix = 0
    Else
        frmBC23BrowseBarang.txtFreightFix = CDbl(txtFOB) / CDbl(txtFreight)
    End If
    frmBC23BrowseBarang.CekData = True
    frmBC23BrowseBarang.cekSubmit = True
    frmBC23BrowseBarang.Show 1
    
    up_GridLoadBarang
End Sub

Private Sub cmdSaveHarga_Click()
    up_SaveHarga
End Sub

Private Sub cmdSaveKemasan_Click()
    up_SaveKemasan
End Sub


Private Sub cmdSaveKontainer_Click()
    up_SaveKontainer
End Sub



Private Sub Form_Activate()
    up_GridLoadDokumen
    up_GridLoadKemasan
    up_GridLoadKontainer
    up_GridLoadBarang
    
    up_GridLoadPungutan

End Sub

Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(Me)

'up_FillCombo

up_GridHeaderRespon
up_GridHeaderStatus

up_FillComboTujuan
up_FillComboAPI cboImportirAPI
up_FillComboAPI cboPemilikAPI
up_FillComboKodeID cboIDPemilik
up_FillComboBLAWB cboDokumenBLAWB
up_FillComboGeneral cboUkuranKontainer, "Bea_Cukai_Ukuran_Kontainer", "KODE_UKURAN_KONTAINER", "URAIAN_UKURAN_KONTAINER", 90, 110
up_FillComboGeneral cboTipeKontainer, "Bea_Cukai_Tipe_Kontainer", "KODE_TIPE_KONTAINER", "URAIAN_TIPE_KONTAINER", 60, 70
up_FillComboGeneral cboKodeHarga, "Bea_Cukai_Harga", "KODE_HARGA", "URAIAN_HARGA", 60, 350
up_FillComboGeneral cboAsuransi, "Bea_Cukai_Asuransi", "KODE_ASURANSI", "URAIAN_ASURANSI", 60, 110
cboAsuransi.ListIndex = 1

up_FillComboCaraAngkut

HakU = hakUpdate(Me.Name)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

'With Anchor1
'    .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
'    .DoInit
'End With

End Sub




Private Sub gridKemasan_Click()
    If gridKemasan.RowSel > 0 Then
        txtJenisKemasan.Enabled = False
        txtJenisKemasan = gridKemasan.TextMatrix(gridKemasan.RowSel, colKodeKemasan)
        txtJumlahKemasan = gridKemasan.TextMatrix(gridKemasan.RowSel, colJumlah)
        gb_LoadDataMaster "Bea_Cukai_Kemasan", "Uraian_Kemasan", lblJenisKemasan, "Where Kode_Kemasan = '" & txtJenisKemasan & "'"
        txtMerkKemasan = gridKemasan.TextMatrix(gridKemasan.RowSel, colNomorDokumen)
    End If
End Sub

Private Sub gridKontainer_Click()
    If gridKontainer.RowSel > 0 Then
'        txtJenisKemasan.Enabled = False
'        txtJenisKemasan = gridKemasan.TextMatrix(gridKemasan.RowSel, colKodeKemasan)
'        txtJumlahKemasan = gridKemasan.TextMatrix(gridKemasan.RowSel, colJumlah)
'        gb_LoadDataMaster "Bea_Cukai_Kemasan", "Uraian_Kemasan", lblJenisKemasan, "Where Kode_Kemasan = '" & txtJenisKemasan & "'"
'        txtMerkKemasan = gridKemasan.TextMatrix(gridKemasan.RowSel, colNomorDokumen)
        txtNomorKontainer1 = Left(gridKontainer.TextMatrix(gridKontainer.RowSel, colNomorKontainer), 4)
        txtNomorKontainer2 = Mid(gridKontainer.TextMatrix(gridKontainer.RowSel, colNomorKontainer), 5, 7)
        cboUkuranKontainer = gridKontainer.TextMatrix(gridKontainer.RowSel, colHideUkuran)
        cboTipeKontainer = gridKontainer.TextMatrix(gridKontainer.RowSel, colHideTipe)
        txtIDKontainer = gridKontainer.TextMatrix(gridKontainer.RowSel, colIDKontainer)
        txtKeteranganKontainer = gridKontainer.TextMatrix(gridKontainer.RowSel, colHideKeterangan)
    End If
End Sub

Private Sub txtAlamatPemasok_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtAlamatPemilik_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtAPIImportir_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtAPIPemilik_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtAsuransi_GotFocus()
txtAsuransi = CDbl(txtAsuransi)
End Sub

Private Sub txtAsuransi_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtAsuransi_LostFocus()
If cboKodeHarga <> "" Then
    If Trim(Split(cboKodeHarga, "-")(0)) = "CIF" Then
        txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
    Else
        
        txtCIF = Format(((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan) + CDbl(txtAsuransi) + CDbl(txtFreightPIB))) - CDbl(txtDiskon), "#,0.00")
        txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
        
        txtAsuransi = Format(txtAsuransi, "#,0.00")
    End If
End If
End Sub

Private Sub txtBiayaTambahan_GotFocus()
txtBiayaTambahan = CDbl(txtBiayaTambahan)
End Sub

Private Sub txtBiayaTambahan_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtBiayaTambahan_LostFocus()
If cboKodeHarga <> "" Then
    If Trim(Split(cboKodeHarga, "-")(0)) = "CIF" Then
        txtFOBPIB = "0.00"
        txtBiayaTambahan = Format(txtBiayaTambahan, "#,0.00")
        
        txtCIF = Format((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan)) - CDbl(txtDiskon), "#,0.00")
        txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
    Else
        txtFOBPIB = Format((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan)) - CDbl(txtDiskon), "#,0.00")
        txtBiayaTambahan = Format(txtBiayaTambahan, "#,0.00")
    End If
End If

End Sub

Private Sub txtBrutoBarang_GotFocus()
txtBrutoBarang = CDbl(txtBrutoBarang)
End Sub

Private Sub txtBrutoBarang_LostFocus()
txtBrutoBarang = Format(txtBrutoBarang, "#,0.0000")
End Sub

Private Sub txtCIF_Change()
    txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
End Sub

Private Sub txtDiskon_GotFocus()
txtDiskon = CDbl(txtDiskon)
End Sub

Private Sub txtDiskon_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtDiskon_LostFocus()
If cboKodeHarga <> "" Then
    If Trim(Split(cboKodeHarga, "-")(0)) = "CIF" Then
        txtFOBPIB = "0.00"
        txtDiskon = Format(txtDiskon, "#,0.00")
        
        txtCIF = Format((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan)) - CDbl(txtDiskon), "#,0.00")
        txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
    Else
        txtFOBPIB = Format((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan)) - CDbl(txtDiskon), "#,0.00")
        txtDiskon = Format(txtDiskon, "#,0.00")
    End If

End If

End Sub

Private Sub txtFOBPIB_Change()
    txtCIF = Format(CDbl(txtFOBPIB) + CDbl(txtAsuransi) + CDbl(txtFreightPIB), "#,0.00")
    txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
End Sub

Private Sub txtFreightPIB_GotFocus()
txtFreightPIB = CDbl(txtFreightPIB)
End Sub

Private Sub txtFreightPIB_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtFreightPIB_LostFocus()
If cboKodeHarga <> "" Then
    If Trim(Split(cboKodeHarga, "-")(0)) = "CIF" Then
        txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
        
    Else
        
        txtCIF = Format(((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan) + CDbl(txtAsuransi) + CDbl(txtFreightPIB))) - CDbl(txtDiskon), "#,0.00")
        txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
        
        txtFreightPIB = Format(txtFreightPIB, "#,0.00")
    End If
End If
End Sub

Private Sub txtHargaCNF_GotFocus()
    txtHargaCNF = CDbl(txtHargaCNF)
End Sub

Private Sub txtHargaCNF_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtHargaCNF_LostFocus()
If cboKodeHarga <> "" Then
    If Trim(Split(cboKodeHarga, "-")(0)) = "CIF" Then
        txtFOBPIB = "0.00"
        txtHargaCNF = Format(txtHargaCNF, "#,0.00")
        
        txtCIF = Format((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan)) - CDbl(txtDiskon), "#,0.00")
        txtCIFRp = Format(CDbl(txtCIF) * CDbl(txtNDPBMInvoice), "#,0.0000")
        
    Else
        txtFOBPIB = Format((CDbl(txtHargaCNF) + CDbl(txtBiayaTambahan)) - CDbl(txtDiskon), "#,0.00")
        txtHargaCNF = Format(txtHargaCNF, "#,0.00")
        
    End If
End If

End Sub

Private Sub txtIdentitasImportir_GotFocus()
    txtIdentitasImportir = Replace(Replace(txtIdentitasImportir, ".", ""), "-", "")
End Sub

Private Sub txtIdentitasImportir_LostFocus()
    If Len(txtIdentitasImportir) > 15 Then
        LblErrMsg.Caption = "Identitas/NPWP No maximum of 15 characters"
        txtIdentitasImportir.SetFocus
        Exit Sub
    End If
    txtIdentitasImportir = Left(txtIdentitasImportir.Text, 2) & "." & Mid(txtIdentitasImportir.Text, 3, 3) & "." & Mid(txtIdentitasImportir.Text, 6, 3) & "." & Mid(txtIdentitasImportir.Text, 9, 1) & "-" & Mid(txtIdentitasImportir.Text, 10, 3) & "." & Mid(txtIdentitasImportir.Text, 13, 3)
End Sub

Private Sub txtIdentitasPemilik_GotFocus()
    txtIdentitasPemilik = Replace(Replace(txtIdentitasPemilik, ".", ""), "-", "")
End Sub

Private Sub txtIdentitasPemilik_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
    
End Sub

Private Sub txtIdentitasPemilik_LostFocus()
'    If Len(txtIdentitasImportir) > 15 Then
'        LblErrMsg.Caption = "Identitas/NPWP No maximum of 15 characters"
'        txtIdentitasImportir.SetFocus
'        Exit Sub
'    End If
    Dim temp1, temp2, temp3, temp4, temp5, temp6 As String
    
    If cboIDPemilik.ListIndex = 0 Then
        temp1 = Left(Left(txtIdentitasPemilik.Text, 2) & "00", 2)
        temp2 = Left(Mid(txtIdentitasPemilik.Text, 3, 3) & "000", 3)
        temp3 = Left(Mid(txtIdentitasPemilik.Text, 6, 3) & "000", 3)
        temp4 = Left(Mid(txtIdentitasPemilik.Text, 9, 1) & "0", 1)
        temp5 = Left(Mid(txtIdentitasPemilik.Text, 10, 3) & "000", 3)
        temp6 = Left(Mid(txtIdentitasPemilik.Text, 13, 3) & "000", 3)
        
        txtIdentitasPemilik = temp1 & "." & temp2 & "." & temp3 & "." & temp4 & "-" & temp5 & "." & temp6
    End If

End Sub

Private Sub txtJenisKemasan_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtJenisKemasan_LostFocus()
gb_LoadDataMaster "Bea_Cukai_Kemasan", "Uraian_Kemasan", lblJenisKemasan, "Where Kode_Kemasan = '" & txtJenisKemasan & "'"
End Sub

Private Sub txtJumlahBarang_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtJumlahKemasan_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtKodePenimbunan_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtKodePenimbunan_LostFocus()
    gb_LoadDataMaster "Bea_Cukai_TPS", "Uraian_TPS", lblPenimbunan, "Where Kode_TPS = '" & txtKodePenimbunan & "' Order By ID Desc"
End Sub

Private Sub txtKPBBCBongkar_LostFocus()
    up_LoadKantorKPPBCBongkar txtKPBBCBongkar
End Sub

Private Sub txtKPBBCPengawas_LostFocus()
    up_LoadKantorPabean txtKPBBCPengawas
End Sub

Private Sub txtMerkKemasan_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNamaPemasok_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNamaPemilik_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNamaPengangkut_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNDPBMInvoice_GotFocus()
    txtNDPBMInvoice = CDbl(txtNDPBMInvoice)
End Sub

Private Sub txtNDPBMInvoice_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtNDPBMInvoice_LostFocus()
    txtCIFRp.Text = Format(CDbl(txtNDPBMInvoice) * CDbl(txtCIF), "#,0.0000")
    
    txtNDPBMInvoice = Format(txtNDPBMInvoice, "#,0.00")
End Sub

Private Sub txtNegara_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNegara_LostFocus()
    up_LoadNegara (txtNegara)
End Sub

Private Sub txtNegaraPengangkut_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNegaraPengangkut_LostFocus()
    gb_LoadDataMaster "Bea_Cukai_Negara", "Nama_Negara", lblNegaraPengangkut, "Where Kode_Negara = '" & txtNegaraPengangkut & "'"
End Sub

Private Sub txtNettoBarang_GotFocus()
txtNettoBarang = CDbl(txtNettoBarang)
End Sub

Private Sub txtNettoBarang_LostFocus()
txtNettoBarang = Format(txtNettoBarang, "#,0.0000")
End Sub

Private Sub txtNoDaftar_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtNomorKontainer1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtNomorKontainer2_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtPelabuhanBongkar_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPelabuhanBongkar_LostFocus()
    gb_LoadDataMaster "Bea_Cukai_Pelabuhan", "Uraian_Pelabuhan", lblPelabuhanBongkar, "Where Kode_Pelabuhan = '" & txtPelabuhanBongkar & "'"
End Sub

Private Sub txtPelabuhanMuat_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPelabuhanMuat_LostFocus()
    gb_LoadDataMaster "Bea_Cukai_Pelabuhan", "Uraian_Pelabuhan", lblPelabuhanMuat, "Where Kode_Pelabuhan = '" & txtPelabuhanMuat & "'"
End Sub

Private Sub txtPelabuhanTransit_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPelabuhanTransit_LostFocus()
    gb_LoadDataMaster "Bea_Cukai_Pelabuhan", "Uraian_Pelabuhan", lblPelabuhanTransit, "Where Kode_Pelabuhan = '" & txtPelabuhanTransit & "'"
End Sub

Private Sub txtValuta_Change()
    gb_LoadDataMaster "Bea_Cukai_Valuta", "Uraian_Valuta", lblValuta, "Where Kode_Valuta = '" & txtValuta & "'"
End Sub

Private Sub txtValutaInvoice_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtVoyFlight_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


