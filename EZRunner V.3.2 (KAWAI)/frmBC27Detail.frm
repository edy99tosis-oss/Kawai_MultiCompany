VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBC27Detail 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BC 2.7 Detail"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   16050
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBC27Detail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   16050
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "Syn&cronize"
      Height          =   375
      Index           =   2
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   78
      Tag             =   "FFTT*/"
      Top             =   10200
      Width           =   1140
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   1
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   52
      Tag             =   "FFTT*/"
      Top             =   10200
      Width           =   1140
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Back"
      Height          =   375
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   41
      Tag             =   "TFFT*/"
      Top             =   10200
      Width           =   1140
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   360
      TabIndex        =   39
      Tag             =   "TFTT*/"
      Top             =   9480
      Width           =   15330
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
         Left            =   120
         TabIndex        =   40
         Tag             =   "TFTF*/"
         Top             =   240
         Width           =   15045
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3855
      Left            =   360
      TabIndex        =   38
      Top             =   5400
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Barang"
      TabPicture(0)   =   "frmBC27Detail.frx":0E42
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdAddBarang"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdDetailBarang"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "gridBarang"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Harga"
      TabPicture(1)   =   "frmBC27Detail.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame13"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Dokumen"
      TabPicture(2)   =   "frmBC27Detail.frx":0E7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "btnBrowseDokumen"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "gridDokumen"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Kontainer"
      TabPicture(3)   =   "frmBC27Detail.frx":0E96
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label49"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "gridKontainer"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "txtJDataKontainer"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame10"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Kemasan"
      TabPicture(4)   =   "frmBC27Detail.frx":0EB2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame8"
      Tab(4).Control(1)=   "gridKemasan"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Respon"
      TabPicture(5)   =   "frmBC27Detail.frx":0ECE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label30"
      Tab(5).Control(1)=   "Label29"
      Tab(5).Control(2)=   "gridStatus"
      Tab(5).Control(3)=   "gridRespon"
      Tab(5).ControlCount=   4
      Begin VB.Frame Frame8 
         Height          =   3375
         Left            =   -66000
         TabIndex        =   107
         Tag             =   "TTTT*/"
         Top             =   360
         Width           =   5895
         Begin VB.CommandButton cmdCancelKemasan 
            BackColor       =   &H0080FFFF&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   113
            Tag             =   "FFTT*/"
            Top             =   2880
            Width           =   975
         End
         Begin VB.CommandButton cmdDeleteKemasan 
            BackColor       =   &H0080FFFF&
            Caption         =   "Delete"
            Height          =   375
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   112
            Tag             =   "FFTT*/"
            Top             =   2880
            Width           =   975
         End
         Begin VB.CommandButton cmdSaveKemasan 
            BackColor       =   &H0080FFFF&
            Caption         =   "Save"
            Height          =   375
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   111
            Tag             =   "FFTT*/"
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtMerkKemasan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   255
            TabIndex        =   110
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   3135
         End
         Begin VB.TextBox txtJenisKemasan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   109
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtJumlahKemasan 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   108
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah"
            Height          =   255
            Left            =   240
            TabIndex        =   117
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblJenisKemasan 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   2880
            TabIndex        =   116
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
            Height          =   255
            Left            =   240
            TabIndex        =   115
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   1335
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis"
            Height          =   255
            Left            =   240
            TabIndex        =   114
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   1335
         End
      End
      Begin VB.Frame Frame10 
         Height          =   2895
         Left            =   7680
         TabIndex        =   93
         Tag             =   "TTTT*/"
         Top             =   720
         Width           =   7335
         Begin VB.TextBox txtIDKontainer 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   100
            Tag             =   "TTFF*/"
            Top             =   1800
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancelKontainer 
            BackColor       =   &H0080FFFF&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   99
            Tag             =   "FFTT*/"
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdDeleteKontainer 
            BackColor       =   &H0080FFFF&
            Caption         =   "Delete"
            Height          =   375
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   98
            Tag             =   "FFTT*/"
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdSaveKontainer 
            BackColor       =   &H0080FFFF&
            Caption         =   "Save"
            Height          =   375
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   97
            Tag             =   "FFTT*/"
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox txtKeteranganKontainer 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   96
            Tag             =   "TTFF*/"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtNomorKontainer2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3240
            MaxLength       =   7
            TabIndex        =   95
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtNomorKontainer1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   94
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
            Height          =   255
            Left            =   240
            TabIndex        =   106
            Tag             =   "TTFF*/"
            Top             =   1320
            Width           =   1815
         End
         Begin MSForms.ComboBox cboTipeKontainer 
            Height          =   315
            Left            =   2040
            TabIndex        =   105
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
            Height          =   255
            Left            =   240
            TabIndex        =   104
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1695
         End
         Begin MSForms.ComboBox cboUkuranKontainer 
            Height          =   315
            Left            =   2040
            TabIndex        =   103
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
            Height          =   255
            Left            =   240
            TabIndex        =   102
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Kontainer"
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox txtJDataKontainer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   14160
         Locked          =   -1  'True
         TabIndex        =   90
         Tag             =   "FFTF*/"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton btnBrowseDokumen 
         BackColor       =   &H0080FFFF&
         Caption         =   "Browse"
         Height          =   375
         Left            =   -60960
         Style           =   1  'Graphical
         TabIndex        =   75
         Tag             =   "FFTT*/"
         Top             =   3360
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   61
         Tag             =   "TTFT*/"
         Top             =   360
         Width           =   6135
         Begin VB.TextBox txtFasilitasImpor2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5160
            TabIndex        =   73
            Tag             =   "TTFF*/"
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox txtFasilitasImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   72
            Tag             =   "TTFF*/"
            Top             =   1500
            Width           =   1335
         End
         Begin VB.TextBox txtKontrak 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   70
            Tag             =   "TTFF*/"
            Top             =   1110
            Width           =   2175
         End
         Begin VB.TextBox txtPackingList 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   68
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox txtInvoiceDokumen 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   66
            Tag             =   "TTFF*/"
            Top             =   330
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker dtpTglInvoice 
            Height          =   315
            Left            =   4440
            TabIndex        =   67
            Tag             =   "TTFF*/"
            Top             =   330
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
            Format          =   293601283
            CurrentDate     =   37798
         End
         Begin MSComCtl2.DTPicker dtpTglPackingList 
            Height          =   315
            Left            =   4440
            TabIndex        =   69
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
            Format          =   152174595
            CurrentDate     =   37798
         End
         Begin MSComCtl2.DTPicker dtpTglKontrak 
            Height          =   315
            Left            =   4440
            TabIndex        =   71
            Tag             =   "TTFF*/"
            Top             =   1110
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
            Format          =   152174595
            CurrentDate     =   37798
         End
         Begin MSComCtl2.DTPicker dtpTglFasilitasImpor 
            Height          =   315
            Left            =   3600
            TabIndex        =   74
            Tag             =   "TTFF*/"
            Top             =   1500
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
            Format          =   152174595
            CurrentDate     =   37798
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Fasilitas Impor"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Tag             =   "TTFF*/"
            Top             =   1530
            Width           =   1815
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Packing List"
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Tag             =   "TTFF*/"
            Top             =   750
            Width           =   1815
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Kontrak"
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Tag             =   "TTFF*/"
            Top             =   1140
            Width           =   1575
         End
      End
      Begin VB.Frame Frame13 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   53
         Tag             =   "TTFT*/"
         Top             =   360
         Width           =   5775
         Begin VB.TextBox txtNilaiCIF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   56
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txtValuta 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   55
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtHargaPenyerahan 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   54
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Nilai CIF"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Tag             =   "TTFF*/"
            Top             =   750
            Width           =   1575
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Valuta"
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Penyerahan"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Tag             =   "TTFF*/"
            Top             =   1110
            Width           =   1815
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   3600
            X2              =   5280
            Y1              =   660
            Y2              =   660
         End
         Begin VB.Label lblValuta 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   3600
            TabIndex        =   57
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdAddBarang 
         BackColor       =   &H0080FFFF&
         Caption         =   "Add"
         Height          =   375
         Left            =   -60960
         Style           =   1  'Graphical
         TabIndex        =   51
         Tag             =   "FFTT*/"
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdDetailBarang 
         BackColor       =   &H0080FFFF&
         Caption         =   "Detail"
         Height          =   375
         Left            =   -62040
         Style           =   1  'Graphical
         TabIndex        =   50
         Tag             =   "FFTT*/"
         Top             =   3360
         Width           =   975
      End
      Begin VB.Frame Frame11 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   42
         Tag             =   "TTFT*/"
         Top             =   360
         Width           =   5175
         Begin VB.TextBox txtVolume 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   119
            Tag             =   "TTFF*/"
            Text            =   "0.0000"
            Top             =   300
            Width           =   2175
         End
         Begin VB.TextBox txtBrutoBarang 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   45
            Tag             =   "TTFF*/"
            Text            =   "0.0000"
            Top             =   690
            Width           =   2175
         End
         Begin VB.TextBox txtNettoBarang 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   44
            Tag             =   "TTFF*/"
            Text            =   "0.0000"
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtJumlahBarang 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   43
            Tag             =   "TTFF*/"
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Volume (m3)"
            Height          =   255
            Left            =   240
            TabIndex        =   120
            Tag             =   "TTFF*/"
            Top             =   310
            Width           =   1815
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Bruto (Kg)"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Netto (Kg)"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Tag             =   "TTFF*/"
            Top             =   1110
            Width           =   1815
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Tag             =   "TTFF*/"
            Top             =   1470
            Width           =   1575
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid gridBarang 
         Height          =   2775
         Left            =   -69600
         TabIndex        =   49
         TabStop         =   0   'False
         Tag             =   "TTTT*/"
         Top             =   480
         Width           =   9645
         _cx             =   17013
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
      Begin VSFlex8Ctl.VSFlexGrid gridDokumen 
         Height          =   2775
         Left            =   -68400
         TabIndex        =   76
         TabStop         =   0   'False
         Tag             =   "TTTT*/"
         Top             =   480
         Width           =   8445
         _cx             =   14896
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
         Left            =   240
         TabIndex        =   91
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
      Begin VSFlex8Ctl.VSFlexGrid gridKemasan 
         Height          =   3255
         Left            =   -74640
         TabIndex        =   118
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
      Begin VSFlex8Ctl.VSFlexGrid gridRespon 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   137
         TabStop         =   0   'False
         Tag             =   "TTFT*/"
         Top             =   720
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
         Left            =   -67680
         TabIndex        =   138
         TabStop         =   0   'False
         Tag             =   "TTTT*/"
         Top             =   720
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
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   255
         Left            =   -67680
         TabIndex        =   140
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Respon"
         Height          =   255
         Left            =   -74760
         TabIndex        =   139
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data"
         Height          =   255
         Left            =   12960
         TabIndex        =   92
         Tag             =   "FFTF*/"
         Top             =   420
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2055
      Left            =   360
      TabIndex        =   21
      Top             =   3240
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   3625
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "TPB Asal Barang"
      TabPicture(0)   =   "frmBC27Detail.frx":0EEA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "TPB Tujuan Barang"
      TabPicture(1)   =   "frmBC27Detail.frx":0F06
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Pengangkutan"
      TabPicture(2)   =   "frmBC27Detail.frx":0F22
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Segel BC Asal"
      TabPicture(3)   =   "frmBC27Detail.frx":0F3E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame7 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   126
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   15135
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2280
            TabIndex        =   129
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2280
            TabIndex        =   128
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   4815
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2280
            MaxLength       =   100
            TabIndex        =   127
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   4815
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Segel"
            Height          =   255
            Left            =   240
            TabIndex        =   132
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis"
            Height          =   255
            Left            =   240
            TabIndex        =   131
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   975
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Catatan BC Tujuan"
            Height          =   255
            Left            =   240
            TabIndex        =   130
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   120
         TabIndex        =   121
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   15015
         Begin VB.TextBox txtJenisSarana 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3240
            TabIndex        =   123
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   3015
         End
         Begin VB.TextBox txtNoPolisi 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3240
            TabIndex        =   122
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "No Polisi"
            Height          =   255
            Left            =   240
            TabIndex        =   125
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Sarana Pengangkut Darat"
            Height          =   255
            Left            =   240
            TabIndex        =   124
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   3135
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   30
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   15015
         Begin VB.TextBox txtNoIzinPenerima 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   12120
            TabIndex        =   135
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2535
         End
         Begin VB.CommandButton cmdCopyPengusaha 
            BackColor       =   &H0080FFFF&
            Caption         =   "Copy Data Pengusaha"
            Height          =   375
            Left            =   12360
            Style           =   1  'Graphical
            TabIndex        =   77
            Tag             =   "FFTT*/"
            Top             =   990
            Width           =   2295
         End
         Begin VB.TextBox txtNPWPPenerima 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4320
            TabIndex        =   33
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtAlamatPenerima 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   32
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   4935
         End
         Begin VB.TextBox txtNamaPenerima 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   31
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   4935
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "No Izin"
            Height          =   255
            Left            =   11040
            TabIndex        =   136
            Tag             =   "TTFF*/"
            Top             =   285
            Width           =   975
         End
         Begin MSForms.ComboBox cboNPWPPenerima 
            Height          =   315
            Left            =   2040
            TabIndex        =   37
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
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   22
         Tag             =   "TFTF*/"
         Top             =   360
         Width           =   15015
         Begin VB.TextBox txtNoIzinPengusahaTPB 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   12120
            TabIndex        =   133
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtNamaPengusahaTPB 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   25
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   4935
         End
         Begin VB.TextBox txtAlamatPengusahaTPB 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   24
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   4935
         End
         Begin VB.TextBox txtNPWPPengusahaTPB 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4320
            TabIndex        =   23
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label66 
            BackStyle       =   0  'Transparent
            Caption         =   "No Izin"
            Height          =   255
            Left            =   11040
            TabIndex        =   134
            Tag             =   "TTFF*/"
            Top             =   285
            Width           =   975
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1575
         End
         Begin MSForms.ComboBox cboTipeNPWPPengusahaTPB 
            Height          =   315
            Left            =   2040
            TabIndex        =   26
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2175
            VariousPropertyBits=   746604575
            BackColor       =   -2147483648
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2415
      Left            =   360
      TabIndex        =   9
      Tag             =   "TFTF*/"
      Top             =   600
      Width           =   15360
      Begin VB.TextBox txtPLBAsal 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7440
         TabIndex        =   83
         Tag             =   "TTFF*/"
         Top             =   1530
         Width           =   1215
      End
      Begin VB.TextBox txtPLBTujuan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7440
         TabIndex        =   82
         Tag             =   "TTFF*/"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtKantorTujuan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   79
         Tag             =   "TTFF*/"
         Top             =   2000
         Width           =   1335
      End
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
         Height          =   375
         Index           =   3
         Left            =   13920
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtNoDaftar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "TTFF*/"
         Top             =   800
         Width           =   1815
      End
      Begin VB.TextBox txtKPBBCBongkar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   1600
         Width           =   1335
      End
      Begin VB.TextBox txtTempat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11760
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   300
         Width           =   1695
      End
      Begin VB.TextBox txtPemberitahu 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11760
         TabIndex        =   7
         Tag             =   "TTFF*/"
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtJabatan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11760
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   1125
         Width           =   3255
      End
      Begin MSMask.MaskEdBox txtNoPengajuan 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
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
         TabIndex        =   2
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
         Format          =   293142531
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   315
         Left            =   13560
         TabIndex        =   6
         Tag             =   "TTFF*/"
         Top             =   300
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
         Format          =   293142531
         CurrentDate     =   37798
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Asal"
         Height          =   195
         Left            =   6120
         TabIndex        =   89
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   855
      End
      Begin MSForms.ComboBox cboAsal 
         Height          =   315
         Left            =   7440
         TabIndex        =   88
         Tag             =   "FFFF*/"
         Top             =   360
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
      Begin MSForms.ComboBox cboPengiriman 
         Height          =   315
         Left            =   7440
         TabIndex        =   87
         Tag             =   "TTFF*/"
         Top             =   1125
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
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Pengiriman"
         Height          =   195
         Left            =   6120
         TabIndex        =   86
         Tag             =   "TTFF*/"
         Top             =   1185
         Width           =   975
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode PLB Asal"
         Height          =   195
         Left            =   6120
         TabIndex        =   85
         Tag             =   "TTFF*/"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label65 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode PLB Tujuan"
         Height          =   195
         Left            =   6120
         TabIndex        =   84
         Tag             =   "TTFF*/"
         Top             =   1995
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3360
         X2              =   6000
         Y1              =   2300
         Y2              =   2300
      End
      Begin VB.Label lblKPPBCKantorTujuan 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3360
         TabIndex        =   81
         Tag             =   "TTFF*/"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Kantor Tujuan"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Tag             =   "TTFF*/"
         Top             =   2000
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Pengajuan"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Tag             =   "TTFF*/"
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Kantor Asal"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Tag             =   "TTFF*/"
         Top             =   1630
         Width           =   1455
      End
      Begin MSForms.ComboBox cboTujuan 
         Height          =   315
         Left            =   7440
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   730
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
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tujuan"
         Height          =   255
         Left            =   6120
         TabIndex        =   15
         Tag             =   "TTFF*/"
         Top             =   750
         Width           =   855
      End
      Begin VB.Label lblKPPBCBongkar 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Tag             =   "TTFF*/"
         Top             =   1630
         Width           =   2535
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3360
         X2              =   6000
         Y1              =   1900
         Y2              =   1900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat/Tanggal"
         Height          =   195
         Index           =   2
         Left            =   10200
         TabIndex        =   13
         Tag             =   "TTFF*/"
         Top             =   390
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pemberitahu"
         Height          =   195
         Index           =   3
         Left            =   10200
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         Height          =   195
         Index           =   4
         Left            =   10200
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   1185
         Width           =   660
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BC 2.7 Detail"
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
      Left            =   480
      TabIndex        =   20
      Tag             =   "TTTF*/"
      Top             =   120
      Width           =   14535
   End
End
Attribute VB_Name = "frmBC27Detail"
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
Const colKodeBarang As Integer = 0
Const colNamaBarang As Integer = 1
Const colSatuan As Integer = 2
Const colVolume As Integer = 3
Const colHideNoSeri As Integer = 4
Const colCountBarang As Integer = 5

'-------------------------------------------
Const colJenisDokumen As Integer = 0
Const colNomorDokumen As Integer = 1
Const colTanggal As Integer = 2
Const colCountDokumen As Integer = 3


'-------------------------------------------
Const colNo As Integer = 0
Const colJumlah As Integer = 1
Const colKodeKemasan As Integer = 2
Const colNamaKemasan As Integer = 3
Const colMerkKemasan As Integer = 4
Const colCountKemasan As Integer = 5

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
Const colNoTarif As Integer = 0
Const colJenisPungutan As Integer = 1
Const colDitangguhkan As Integer = 2
Const colDibebaskan As Integer = 3
Const colTidakDipungut As Integer = 4
Const colCountPungutan As Integer = 5

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


'################################### Start Procedure ###############################################
Private Sub up_GridHeaderBarang()
    
    With gridBarang
        .ColS = colCountBarang
        .Rows = 1

        .TextMatrix(0, colKodeBarang) = "Kode"
        .TextMatrix(0, colNamaBarang) = "Uraian"
        .TextMatrix(0, colSatuan) = "Satuan"
        .TextMatrix(0, colVolume) = "Volume"

        .ColWidth(colKodeBarang) = 1200
        .ColWidth(colNamaBarang) = 2500
        .ColWidth(colSatuan) = 1000
        .ColWidth(colVolume) = 1500
        .ColWidth(colHideNoSeri) = 0
'        .ColFormat(colTanggal) = "dd MMM yyyy"
        .ColAlignment(colKodeBarang) = flexAlignLeftCenter
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
'      With gridPungutan
'        .ColS = colCountPungutan
'        .Rows = 1
'
'        .TextMatrix(0, colNo) = "No"
'        .TextMatrix(0, colJenisPungutan) = "Jenis Pungutan"
'        .TextMatrix(0, colDitangguhkan) = "Ditangguhkan"
'        .TextMatrix(0, colDibebaskan) = "Dibebaskan"
'        .TextMatrix(0, colTidakDipungut) = "Tidak Dipungut"
'
'        .ColWidth(colNo) = 500
'        .ColWidth(colJenisPungutan) = 2500
'        .ColWidth(colDitangguhkan) = 1700
'        .ColWidth(colDibebaskan) = 1700
'        .ColWidth(colTidakDipungut) = 1700
'
'        .ColFormat(colJenisPungutan) = "#,0.00"
'        .ColFormat(colDitangguhkan) = "#,0.00"
'        .ColFormat(colDibebaskan) = "#,0.00"
'        .ColFormat(colTidakDipungut) = "#,0.00"
'
'
'        .MergeCells = flexMergeRestrictRows
'        .WordWrap = True
'
'        .AllowUserResizing = flexResizeColumns
'
'    End With
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
    cmd.CommandText = "sp_BC27DetailBarang_Sel"

    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))

    Set RS = cmd.Execute

    With gridBarang
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1

            .TextMatrix(li_Row, colKodeBarang) = Trim(RS!Kode_Barang)
            .TextMatrix(li_Row, colNamaBarang) = Trim(RS!URAIAN)
            .TextMatrix(li_Row, colSatuan) = Trim(RS!URAIAN_SATUAN)
            .TextMatrix(li_Row, colVolume) = Format(RS!Volume, "#,0.00")
'            .TextMatrix(li_Row, ColQty) = Format(rs!JUMLAH_SATUAN, "#,0.00")
'            .TextMatrix(li_Row, colTotal) = Format(rs!Total, "#,0.00")
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
    cmd.CommandText = "sp_BC27TPBDokumenWithoutInvoice_Sel"

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
    cmd.CommandText = "sp_BC27LoadDataKemasan_Sel"

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
    cmd.CommandText = "sp_BC27LoadDataKontainer_Sel"

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

Private Sub up_LoadKantorKPPBCTujuan(pKode As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select * From Bea_Cukai_Kantor_Pabean Where Kode_Kantor = '" & pKode & "'"
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    lblKPPBCKantorTujuan.Caption = RS.Fields("Nama_Kantor")
Else
    lblKPPBCKantorTujuan.Caption = ""
End If
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

Private Sub up_FillComboAsal()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    sql = "Select * From Bea_Cukai_Jenis_TPB"
    
    Set RS = Db.Execute(sql)

    With cboAsal
        .clear
        .columnCount = 1
        .ColumnWidths = "50pt;225pt"
        .ListWidth = 275
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

Private Sub up_FillComboPengiriman()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    sql = "SELECT * FROM Bea_Cukai_Referensi_Tujuan_Pengiriman Where Kode_Dokumen='27'"
    
    Set RS = Db.Execute(sql)

'    Set rs = cmd.Execute

    With cboPengiriman
        .clear
        .columnCount = 1
        .ColumnWidths = "50pt;225pt"
        .ListWidth = 275
        .ListRows = 15

        i = 0
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS(2)) & " - " & IIf(IsNull(RS(2)), "", Trim(RS(3)))
            
            RS.MoveNext
            i = i + 1
            Loop
        .ListIndex = -1
    End With
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
    cmd.CommandText = "sp_BC27DetailKemasan_Upd"
    
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
        cmd.CommandText = "sp_BC27DetailKemasan_Ins"
            
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
    cmd.CommandText = "sp_BC27DetailKemasan_Del"
    
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

Private Sub up_DeleteKontainer()
    Dim cmd As ADODB.Command
    Dim prm1 As ADODB.Parameter
    Dim prm2 As ADODB.Parameter
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailKontainer_Del"
    
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("IDKontainer", adVarChar, adParamInput, 2, txtIDKontainer)
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
    cmd.CommandText = "sp_BC27DetailKontainer_Upd"
    
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
        cmd.CommandText = "sp_BC27DetailKontainer_Ins"
            
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
        up_SaveTPBKontainerMy
        up_SaveTPBBarangTarifMy
        up_SaveTPBBahanBakuMy
        up_SaveTPBBahanBakuTarifMy
        
        
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
Dim ls_CIF As Double
Dim ls_CIFRupiah As Double
Dim ls_HargaPenyerahan As Double
Dim ls_KodePerhitungan As String

Dim ls_KodeGuna As String
Dim ls_KondisiBarang As String
Dim ls_JangkaWaktuLebihBesar4Tahun As String


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
            ls_CIF = IIf(IsNull(rs1.Fields("CIF")), 0, rs1.Fields("CIF"))
            ls_CIFRupiah = IIf(IsNull(rs1.Fields("CIF_RUPIAH")), 0, rs1.Fields("CIF_RUPIAH"))
            
            ls_HargaPenyerahan = IIf(IsNull(rs1.Fields("HARGA_PENYERAHAN")), 0, rs1.Fields("HARGA_PENYERAHAN"))
            ls_KodePerhitungan = IIf(IsNull(rs1.Fields("KODE_PERHITUNGAN")), "", rs1.Fields("KODE_PERHITUNGAN"))
            
            ls_KodeGuna = IIf(IsNull(rs1.Fields("KODE_GUNA")), "", rs1.Fields("KODE_GUNA"))
            ls_KondisiBarang = IIf(IsNull(rs1.Fields("KONDISI_BARANG")), "", rs1.Fields("KONDISI_BARANG"))
            ls_JangkaWaktuLebihBesar4Tahun = IIf(IsNull(rs1.Fields("KODE_LEBIH_DARI4TAHUN")), "", rs1.Fields("KODE_LEBIH_DARI4TAHUN"))
            
            sql = "Select * From TPB_Barang WHERE SERI_BARANG = " & .TextMatrix(lirow, colHideNoSeri) & " AND ID_HEADER = " & ls_IDHeader & ""
            rs2.Open sql, DbMy, adOpenDynamic, adLockOptimistic
            'rsHeader.Open sql, DbMy, adOpenDynamic, adLockOptimistic
            
            If rs2.EOF Then
                sql = "     INSERT INTO TPB_Barang " & vbCrLf & _
                            "   (ID_HEADER, SERI_BARANG, KODE_BARANG, URAIAN, MERK, TIPE, SPESIFIKASI_LAIN, UKURAN, " & vbCrLf & _
                            "   POS_TARIF, KATEGORI_BARANG, KODE_SATUAN, KODE_KEMASAN,  " & vbCrLf & _
                            "   KODE_NEGARA_ASAL, KODE_FASILITAS_DOKUMEN, KODE_SKEMA_TARIF, " & vbCrLf & _
                            "   JUMLAH_SATUAN, JUMLAH_KEMASAN,  " & vbCrLf & _
                            "   CIF, CIF_RUPIAH, HARGA_PENYERAHAN, KODE_PERHITUNGAN, " & vbCrLf & _
                            "   KODE_GUNA, KONDISI_BARANG, KODE_LEBIH_DARI4TAHUN " & vbCrLf & _
                            "   ) " & vbCrLf & _
                            "   VALUES  " & vbCrLf & _
                            "   ('" & ls_IDHeader & "', '" & .TextMatrix(lirow, colHideNoSeri) & "', '" & ls_KodeBarang & "', '" & ls_Uraian & "', '" & ls_Merk & "', '" & ls_Tipe & "', '" & ls_SpesifikasiLain & "', '" & ls_Ukuran & "', " & vbCrLf & _
                            "   '" & ls_NomorHS & "', '" & ls_KodeKategori & "', '" & ls_KodeSatuan & "', '" & ls_KodeKemasan & "', " & vbCrLf & _
                            "   '" & ls_KodeNegara & "', '" & ls_KodeFasilitas & "', '" & ls_KodeSkemaTarif & "', "
                
                sql = sql + "   " & ls_JumlahSatuan & ", " & ls_JumlahKemasan & ", " & vbCrLf & _
                            "   " & ls_CIF & ", " & ls_CIFRupiah & ", " & ls_HargaPenyerahan & ", '" & ls_KodePerhitungan & "', " & vbCrLf & _
                            "   '" & ls_KodeGuna & "', '" & ls_KondisiBarang & "', '" & ls_JangkaWaktuLebihBesar4Tahun & "'" & vbCrLf & _
                            "   ) " & vbCrLf & _
                            "  "
    
            Else
                sql = "     UPDATE TPB_Barang " & vbCrLf & _
                            "   SET KODE_BARANG = '" & ls_KodeBarang & "', " & vbCrLf & _
                            "       URAIAN = '" & ls_Uraian & "', " & vbCrLf & _
                            "       MERK = '" & ls_Merk & "', " & vbCrLf & _
                            "       TIPE = '" & ls_Tipe & "', " & vbCrLf & _
                            "       SPESIFIKASI_LAIN = '" & ls_SpesifikasiLain & "', " & vbCrLf & _
                            "       POS_TARIF = '" & ls_NomorHS & "', "
                
                sql = sql + "       JUMLAH_SATUAN = " & ls_JumlahSatuan & ", " & vbCrLf & _
                            "       JUMLAH_KEMASAN = " & ls_JumlahKemasan & ", " & vbCrLf & _
                            "       CIF = " & ls_CIF & ", " & vbCrLf & _
                            "       HARGA_PENYERAHAN = " & ls_HargaPenyerahan & " " & vbCrLf & _
                            " "
                                            
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
'rs.Close
sql = "Select * From TPB_Header WHERE NOMOR_AJU = '" & Replace(txtNoPengajuan, "-", "") & "'"
RS.Open sql, DbMy, adOpenDynamic, adLockOptimistic
    
If Not RS.EOF Then
    
sql = " UPDATE TPB_Header " & vbCrLf & _
            " SET   KODE_TUJUAN_TPB = '" & Left(cboTujuan, 1) & "', " & vbCrLf & _
            "   KODE_JENIS_TPB = '" & Left(cboAsal, 1) & "', " & vbCrLf & _
            "   KODE_TUJUAN_PENGIRIMAN = '" & Left(cboPengiriman, 1) & "', " & vbCrLf & _
            "   NAMA_TTD = '" & txtPemberitahu & "', " & vbCrLf & _
            "   JABATAN_TTD = '" & txtJabatan & "', " & vbCrLf & _
            "   KOTA_TTD = '" & txtTempat & "', " & vbCrLf & _
            "   TANGGAL_TTD = '" & Format(dtpTanggal.Value, "yyyy-MM-dd") & "', " & vbCrLf & _
            "   KODE_KANTOR = '" & txtKPBBCBongkar & "', " & vbCrLf & _
            "   KODE_KANTOR_TUJUAN = '" & txtKantorTujuan & "', " & vbCrLf & _
            "   KODE_ID_PENGUSAHA = '" & Left(cboTipeNPWPPengusahaTPB, 1) & "', " & vbCrLf & _
            "   ID_PENGUSAHA = '" & txtNPWPPengusahaTPB & "', " & vbCrLf & _
            "   NAMA_PENGUSAHA = '" & txtNamaPengusahaTPB & "', " & vbCrLf & _
            "   NOMOR_IJIN_TPB = '" & txtNoIzinPengusahaTPB & "', "

sql = sql + "   ALAMAT_PENERIMA_BARANG = '" & txtAlamatPengusahaTPB & "', " & vbCrLf & _
            "   KODE_ID_PENERIMA_BARANG = '" & Left(cboNPWPPenerima, 1) & "', " & vbCrLf & _
            "   ID_PENERIMA_BARANG = '" & txtNPWPPenerima & "', " & vbCrLf & _
            "   NAMA_PENERIMA_BARANG = '" & txtNamaPenerima & "',  " & vbCrLf & _
            "   ALAMAT_PENERIMA_BARANG = '" & txtAlamatPenerima & "', " & vbCrLf & _
            "   NOMOR_IJIN_TPB_PENERIMA = '" & txtNoIzinPenerima & "', "

sql = sql + "   NAMA_PENGANGKUT = '" & txtJenisSarana & "', " & vbCrLf & _
            "   NOMOR_POLISI = '" & txtNoPolisi & "', " & vbCrLf & _
            "   VOlume = " & CDbl(txtVolume) & ", " & vbCrLf & _
            "   BRUTO = " & CDbl(txtBrutoBarang) & ", " & vbCrLf & _
            "   NETTO = " & CDbl(txtNettoBarang) & ", " & vbCrLf & _
            "   SERI = '0', "

sql = sql + "   KODE_VALUTA = '" & txtValuta & "', " & vbCrLf & _
            "   CIF = " & CDbl(txtNilaiCIF) & ", " & vbCrLf & _
            "   Harga_Penyerahan = " & CDbl(txtHargaPenyerahan) & " " & vbCrLf & _
            "  WHERE NOMOR_AJU = '" & Replace(txtNoPengajuan, "-", "") & "'"
                      
Else
sql = " Insert Into TPB_Header " & vbCrLf & _
            " (  " & vbCrLf & _
            " NOMOR_AJU, " & vbCrLf & _
            " VERSI_MODUL,  " & vbCrLf & _
            " ID_MODUL,  " & vbCrLf & _
            " KODE_TUJUAN_TPB,  " & vbCrLf & _
            " Kode_Jenis_TPB, " & vbCrLf & _
            " KODE_TUJUAN_PENGIRIMAN, " & vbCrLf & _
            " NAMA_TTD,  " & vbCrLf & _
            " JABATAN_TTD,  " & vbCrLf & _
            " KOTA_TTD,  " & vbCrLf & _
            " TANGGAL_TTD,  " & vbCrLf & _
            " KODE_KANTOR,  KODE_KANTOR_TUJUAN,"

sql = sql + " KODE_ID_PENGUSAHA, " & vbCrLf & _
            " ID_PENGUSAHA,  " & vbCrLf & _
            " NAMA_PENGUSAHA, " & vbCrLf & _
            " NOMOR_IJIN_TPB, " & vbCrLf & _
            " ALAMAT_PENGUSAHA, "

sql = sql + " KODE_ID_PENERIMA_BARANG, " & vbCrLf & _
            " ID_PENERIMA_BARANG, " & vbCrLf & _
            " NAMA_PENERIMA_BARANG, " & vbCrLf & _
            " ALAMAT_PENERIMA_BARANG, " & vbCrLf & _
            " NOMOR_IJIN_TPB_PENERIMA, "

sql = sql + " NAMA_PENGANGKUT, " & vbCrLf & _
            " NOMOR_POLISI, " & vbCrLf & _
            " VOLUME, " & vbCrLf & _
            " BRUTO, " & vbCrLf & _
            " NETTO, " & vbCrLf & _
            " SERI, " & vbCrLf & _
            " KODE_VALUTA, " & vbCrLf & _
            " CIF, " & vbCrLf & _
            " HARGA_PENYERAHAN" & vbCrLf & _
            " ) " & vbCrLf & _
            " VALUES " & vbCrLf & _
            " (  "

sql = sql + " '" & Replace(txtNoPengajuan, "-", "") & "', " & vbCrLf & _
            " '3.1.8',  " & vbCrLf & _
            " '10372',  " & vbCrLf & _
            " '" & Left(cboTujuan, 1) & "', '" & Left(cboAsal, 1) & "', " & vbCrLf & _
            " '" & Left(cboPengiriman, 1) & "',  " & vbCrLf & _
            " '" & txtPemberitahu & "',  " & vbCrLf & _
            " '" & txtJabatan & "',  " & vbCrLf & _
            " '" & txtTempat & "',  " & vbCrLf & _
            " '" & Format(dtpTanggal.Value, "yyyy-MM-dd") & "', " & vbCrLf & _
            " '" & txtKPBBCBongkar & "', '" & txtKantorTujuan & "',  "
            
sql = sql + " '1', " & vbCrLf & _
            " '" & txtNPWPPengusahaTPB & "', " & vbCrLf & _
            " '" & txtNamaPengusahaTPB & "', " & vbCrLf & _
            " '" & txtNoIzinPengusahaTPB & "', " & vbCrLf & _
            " '" & txtAlamatPengusahaTPB & "', "
            
sql = sql + " '" & Left(cboNPWPPenerima, 1) & "', " & vbCrLf & _
            " '" & txtNPWPPenerima & "', " & vbCrLf & _
            " '" & txtNamaPenerima & "', " & vbCrLf & _
            " '" & txtAlamatPenerima & "', " & vbCrLf & _
            " '" & txtNoIzinPenerima & "',  "

sql = sql + " '" & txtJenisSarana & "', " & vbCrLf & _
            " '" & txtNoPolisi & "', " & vbCrLf & _
            " " & CDbl(txtVolume) & ", " & vbCrLf & _
            " " & CDbl(txtBrutoBarang) & ", " & vbCrLf & _
            " " & CDbl(txtNettoBarang) & ", " & vbCrLf & _
            " '0', " & vbCrLf & _
            " '" & txtValuta & "', " & vbCrLf & _
            " " & CDbl(txtNilaiCIF) & ", " & vbCrLf & _
            " " & CDbl(txtHargaPenyerahan) & " " & vbCrLf & _
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

Private Sub up_SaveTPBBahanBakuMy()
Dim rsHeader As New Recordset
Dim sql As String
Dim ls_IDHeader As String

On Error GoTo errHandler

sql = "Select * From TPB_Header WHERE NOMOR_AJU = '" & Replace(txtNoPengajuan, "-", "") & "'"
rsHeader.Open sql, DbMy, adOpenDynamic, adLockOptimistic
'rs1.Close
If Not rsHeader.EOF Then
    ls_IDHeader = rsHeader.Fields("ID")
End If
rsHeader.Close

Dim rs1 As New Recordset
Dim rs2 As New Recordset

sql = "DELETE FROM TPB_Bahan_Baku WHERE ID_HEADER = " & ls_IDHeader & ""
DbMy.Execute sql
'rs1.Close
sql = "Select * From Bea_Cukai_TPB_Bahan_Baku WHERE NO_PENGAJUAN = '" & Replace(txtNoPengajuan, "-", "") & "'"
rs1.Open sql, Db, adOpenDynamic, adLockOptimistic

While Not rs1.EOF
'rs2.Close
    sql = "Select * From TPB_Barang WHERE ID_HEADER = " & ls_IDHeader & " AND SERI_BARANG = " & rs1.Fields("SERI_BARANG") & ""
    rs2.Open sql, DbMy, adOpenDynamic, adLockOptimistic

    If Not rs2.EOF Then
    
        sql = "     INSERT INTO TPB_Bahan_Baku " & vbCrLf & _
                    "   (  " & vbCrLf & _
                    "   KODE_BARANG,  " & vbCrLf & _
                    "   SERI_BARANG,  " & vbCrLf & _
                    "   KODE_ASAL_BAHAN_BAKU,  " & vbCrLf & _
                    "   SERI_BAHAN_BAKU,  " & vbCrLf & _
                    "   URAIAN,  " & vbCrLf & _
                    "   TIPE, Merk, Spesifikasi_Lain, SERI_BARANG_DOK_ASAL, " & vbCrLf & _
                    "   JENIS_SATUAN,  " & vbCrLf & _
                    "   JUMLAH_SATUAN,  " & vbCrLf & _
                    "   NOMOR_DAFTAR_DOK_ASAL,  "
        
        sql = sql + "   TANGGAL_DAFTAR_DOK_ASAL,  " & vbCrLf & _
                    "   POS_TARIF,  " & vbCrLf & _
                    "   CIF,  " & vbCrLf & _
                    "   KODE_KANTOR, " & vbCrLf & _
                    "   KODE_JENIS_DOK_ASAL, " & vbCrLf & _
                    "   NOMOR_AJU_DOK_ASAL, " & vbCrLf & _
                    "   HARGA_PENYERAHAN, "
        
        sql = sql + "   ID_BARANG, ID_HEADER " & vbCrLf & _
                    "   ) " & vbCrLf & _
                    "   VALUES " & vbCrLf & _
                    "   (  " & vbCrLf & _
                    "   '" & rs1.Fields("KODE_BARANG") & "',  " & vbCrLf & _
                    "   " & rs1.Fields("SERI_BARANG") & ",  " & vbCrLf & _
                    "   " & rs1.Fields("KODE_ASAL_BAHAN_BAKU") & ",  " & vbCrLf & _
                    "   " & rs1.Fields("SERI_BAHAN_BAKU") & ",  " & vbCrLf & _
                    "   '" & rs1.Fields("URAIAN") & "',  " & vbCrLf & _
                    "   '" & rs1.Fields("TIPE") & "','" & rs1.Fields("Merk") & "','" & rs1.Fields("Spesifikasi_Lain") & "',  "
        
        sql = sql + "   " & rs1.Fields("SERI_BARANG_DOK_ASAL") & ", " & vbCrLf & _
                    "   '" & rs1.Fields("JENIS_SATUAN") & "',  " & vbCrLf & _
                    "   " & rs1.Fields("JUMLAH_SATUAN") & ",  " & vbCrLf & _
                    "   '" & rs1.Fields("NOMOR_DAFTAR_DOK_ASAL") & "',  " & vbCrLf & _
                    "   '" & rs1.Fields("TANGGAL_DAFTAR_DOK_ASAL") & "',  " & vbCrLf & _
                    "   '" & rs1.Fields("POS_TARIF") & "',  " & vbCrLf & _
                    "   '" & rs1.Fields("CIF") & "',  " & vbCrLf & _
                    "   '" & rs1.Fields("KODE_KANTOR") & "', "
        
        sql = sql + "   '" & rs1.Fields("KODE_JENIS_DOK_ASAL") & "', " & vbCrLf & _
                    "   '" & rs1.Fields("NOMOR_AJU_DOK_ASAL") & "', " & vbCrLf & _
                    "   " & rs1.Fields("HARGA_PENYERAHAN") & ", " & vbCrLf & _
                    "   " & rs2.Fields("ID") & ", " & ls_IDHeader & " " & vbCrLf & _
                    "   ) " & vbCrLf & _
                    "  "
                    
        DbMy.Execute sql
    End If
    rs2.Close
rs1.MoveNext
Wend
'rs1.Close

checkOKToMysql = True
Exit Sub
errHandler:
    LblErrMsg.Caption = err.Description
    checkOKToMysql = False

End Sub

Private Sub up_SaveTPBBahanBakuTarifMy()
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

sql = "DELETE FROM tpb_bahan_baku_tarif WHERE ID_HEADER = " & ls_IDHeader & ""
DbMy.Execute sql

sql = "Select * From bea_cukai_tpb_bahan_baku_tarif WHERE NO_PENGAJUAN = '" & Replace(txtNoPengajuan, "-", "") & "'"
rs1.Open sql, Db, adOpenDynamic, adLockOptimistic

While Not rs1.EOF

    sql = "Select * From tpb_bahan_baku WHERE ID_HEADER = " & ls_IDHeader & " AND SERI_BARANG = " & rs1.Fields("SERI_BARANG") & " AND SERI_BAHAN_BAKU = " & rs1.Fields("SERI_BAHAN_BAKU") & ""
    rs2.Open sql, DbMy, adOpenDynamic, adLockOptimistic

    If Not rs2.EOF Then

        sql = " INSERT INTO tpb_bahan_baku_tarif " & vbCrLf & _
                    " (ID_HEADER, ID_BARANG, ID_BAHAN_BAKU, SERI_BARANG, SERI_BAHAN_BAKU,   " & vbCrLf & _
                    " JENIS_TARIF, KODE_FASILITAS, KODE_TARIF, NILAI_BAYAR,  " & vbCrLf & _
                    " NILAI_FASILITAS, TARIF,TARIF_FASILITAS, KODE_SATUAN,  " & vbCrLf & _
                    " JUMLAH_SATUAN, KODE_ASAL_BAHAN_BAKU, KODE_KOMODITI_CUKAI) " & vbCrLf & _
                    " VALUES " & vbCrLf & _
                    " (" & ls_IDHeader & "," & rs2.Fields("ID_BARANG") & "," & rs2.Fields("ID") & ", " & rs2.Fields("SERI_BARANG") & "," & rs2.Fields("SERI_BAHAN_BAKU") & ", " & vbCrLf & _
                    " '" & rs1.Fields("JENIS_TARIF") & "', '" & rs1.Fields("KODE_FASILITAS") & "', '" & rs1.Fields("KODE_TARIF") & "', " & rs1.Fields("NILAI_BAYAR") & ",  " & vbCrLf & _
                    " '" & rs1.Fields("NILAI_FASILITAS") & "', " & rs1.Fields("TARIF") & ", " & rs1.Fields("TARIF_FASILITAS") & ", '" & rs1.Fields("KODE_SATUAN") & "',  " & vbCrLf & _
                    " " & IIf(IsNull(rs1.Fields("JUMLAH_SATUAN")), 0, rs1.Fields("JUMLAH_SATUAN")) & ", " & rs2.Fields("KODE_ASAL_BAHAN_BAKU") & ", '" & rs1.Fields("KODE_KOMODITI_CUKAI") & "') " & vbCrLf & _
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

Public Sub up_LoadDataBC27(pNoPengajuan As String)
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim lsNoPengajuan As String
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27LoadData_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(pNoPengajuan, "-", ""))
    Set RS = cmd.Execute
    
    If Not RS.EOF Then
        lsNoPengajuan = IIf(IsNull(RS.Fields("NO_PENGAJUAN")), "", RS.Fields("NO_PENGAJUAN"))
        
        If lsNoPengajuan = "" Then
            checkAlreadyData = False
        Else
            checkAlreadyData = True
        End If
        
                
        'cboTujuan = IIf(IsNull(rs.Fields("TUJUANTPB")), "", rs.Fields("TUJUANTPB"))
        cboAsal = IIf(IsNull(RS.Fields("KODEJENISTPB")), "", RS.Fields("KODEJENISTPB"))
        cboTujuan = IIf(IsNull(RS.Fields("KODETUJUANTPB")), "", RS.Fields("KODETUJUANTPB"))
        cboPengiriman = IIf(IsNull(RS.Fields("KodeTujuanPengiriman")), "", RS.Fields("KodeTujuanPengiriman"))
        
        txtBrutoBarang = Format(IIf(IsNull(RS.Fields("BRUTO")), 0, RS.Fields("BRUTO")), "#,0.00")
        txtNettoBarang = Format(IIf(IsNull(RS.Fields("NETTO")), 0, RS.Fields("NETTO")), "#,0.00")
                        
        cboTipeNPWPPengusahaTPB = IIf(IsNull(RS.Fields("KODEIDPENGUSAHA")), "", RS.Fields("KODEIDPENGUSAHA"))
        txtNPWPPengusahaTPB = IIf(IsNull(RS.Fields("IDPENGUSAHA")), "", RS.Fields("IDPENGUSAHA"))
        txtNamaPengusahaTPB = IIf(IsNull(RS.Fields("NAMAPENGUSAHA")), "", RS.Fields("NAMAPENGUSAHA"))
        txtAlamatPengusahaTPB = IIf(IsNull(RS.Fields("ALAMATPENGUSAHA")), "", RS.Fields("ALAMATPENGUSAHA"))
        txtNoIzinPengusahaTPB = IIf(IsNull(RS.Fields("No_Izin")), "", RS.Fields("No_Izin"))

        cboNPWPPenerima = IIf(IsNull(RS.Fields("KODEIDPENERIMA")), "", RS.Fields("KODEIDPENERIMA"))
        txtNPWPPenerima = IIf(IsNull(RS.Fields("IDPENERIMA")), "", RS.Fields("IDPENERIMA"))
        txtNamaPenerima = IIf(IsNull(RS.Fields("NAMAPENERIMA")), "", RS.Fields("NAMAPENERIMA"))
        txtAlamatPenerima = IIf(IsNull(RS.Fields("ALAMATPENERIMA")), "", RS.Fields("ALAMATPENERIMA"))
        txtNoIzinPenerima = IIf(IsNull(RS.Fields("No_Izin_Penerima")), "", RS.Fields("No_Izin_Penerima"))
                
        txtJenisSarana = IIf(IsNull(RS.Fields("NAMAPENGANGKUT")), "", RS.Fields("NAMAPENGANGKUT"))
        txtNoPolisi = IIf(IsNull(RS.Fields("NOMORPOLISI")), "", RS.Fields("NOMORPOLISI"))
     
        '**** HARGA
        
        txtKPBBCBongkar.Text = IIf(IsNull(RS.Fields("KPPBC_BONGKAR")), "", RS.Fields("KPPBC_BONGKAR"))
        lblKPPBCBongkar.Caption = IIf(IsNull(RS.Fields("KANTOR_KPPBC_BONGKAR")), "", RS.Fields("KANTOR_KPPBC_BONGKAR"))
        txtKantorTujuan.Text = IIf(IsNull(RS.Fields("KODEKANTORTUJUAN")), "", RS.Fields("KODEKANTORTUJUAN"))
        'lblKPPBCKantorTujuan.Caption = IIf(IsNull(rs.Fields("KANTOR_KPPBC_BONGKAR")), "", rs.Fields("KANTOR_KPPBC_BONGKAR"))
        cboTujuan = IIf(IsNull(RS.Fields("TUJUANTPB")), "", RS.Fields("TUJUANTPB"))
        txtTempat.Text = Trim(RS.Fields("KOTA_TTD"))
        txtPemberitahu.Text = Trim(RS.Fields("NAMA_TTD"))
        txtJabatan.Text = Trim(RS.Fields("JABATAN_TTD"))
                
        txtInvoiceDokumen = Trim(RS.Fields("NomorDokumenInvoice"))
        dtpTglInvoice.Value = RS.Fields("TglDokumenInvoice")
        txtPackingList = Trim(RS.Fields("NomorDokumenPackingList"))
        dtpTglPackingList.Value = RS.Fields("TglDokumenPackingList")
        txtKontrak = Trim(RS.Fields("NomorDokumenKontrak"))
        dtpTglKontrak.Value = RS.Fields("TglDokumenKontrak")
        
       
        txtValuta = IIf(IsNull(RS.Fields("Kode_Valuta")), "", RS.Fields("Kode_Valuta"))
        lblValuta.Caption = IIf(IsNull(RS.Fields("URAIAN_Valuta")), "", RS.Fields("URAIAN_Valuta"))
        
        txtNilaiCIF = Format(IIf(IsNull(RS.Fields("CIF")), 0, RS.Fields("CIF")), "#,0.00")
        txtHargaPenyerahan = Format(IIf(IsNull(RS.Fields("HARGAPENYERAHAN")), 0, RS.Fields("HARGAPENYERAHAN")), "#,0.00")

        
    End If
    
End Sub

Private Sub up_SaveDetailBC27()
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
    
    Dim Y As Integer
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27Detail_Upd"
    
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("KodeKantorAsal", adVarChar, adParamInput, 200, txtKPBBCBongkar)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("KodeKantorTujuan", adVarChar, adParamInput, 200, txtKantorTujuan)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("KodeJenisTPB", adVarChar, adParamInput, 200, Left(cboAsal, 1))
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeTujuanTPB", adVarChar, adParamInput, 200, Left(cboTujuan, 1))
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("KodeTujuanPengiriman", adVarChar, adParamInput, 200, Left(cboPengiriman, 1))
    cmd.Parameters.append prm6
    Set prm7 = cmd.CreateParameter("KodeMuat", adVarChar, adParamInput, 200, txtKPBBCBongkar)
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("KodeBongkar", adVarChar, adParamInput, 200, txtKPBBCBongkar) ' Replace(Replace(txtNPWPPengusahaTPB, ".", ""), "-", ""))
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("NamaTTD", adVarChar, adParamInput, 200, txtPemberitahu)
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("JabatanTTD", adVarChar, adParamInput, 200, txtJabatan)
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("KotaTTD", adVarChar, adParamInput, 200, txtTempat)
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("TanggalTTD", adDate, adParamInput, , Format(dtpTanggal, "yyyy-MM-dd"))
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("IDPengusaha", adVarChar, adParamInput, 200, Replace(Replace(txtNPWPPengusahaTPB, ".", ""), "-", ""))
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("NamaPengusaha", adVarChar, adParamInput, 200, txtNamaPengusahaTPB) 'Left(cboNPWPPenerima, 1))
    cmd.Parameters.append prm14
    Set prm15 = cmd.CreateParameter("NomorIjinTPB", adVarChar, adParamInput, 200, txtNoIzinPengusahaTPB)
    cmd.Parameters.append prm15
    Set prm29 = cmd.CreateParameter("AlamatPengusaha", adVarChar, adParamInput, 200, txtAlamatPengusahaTPB)
    cmd.Parameters.append prm29
    Set prm16 = cmd.CreateParameter("KodeIDPenerimaBarang", adVarChar, adParamInput, 200, Left(cboNPWPPenerima, 1))
    cmd.Parameters.append prm16
    Set prm17 = cmd.CreateParameter("IDPenerimaBarang", adVarChar, adParamInput, 200, Replace(Replace(txtNPWPPenerima, ".", ""), "-", ""))
    cmd.Parameters.append prm17
    Set prm18 = cmd.CreateParameter("NamaPenerimaBarang", adVarChar, adParamInput, 200, txtNamaPenerima)
    cmd.Parameters.append prm18
    Set prm19 = cmd.CreateParameter("AlamatPenerimaBarang", adVarChar, adParamInput, 200, txtAlamatPenerima)
    cmd.Parameters.append prm19
    Set prm20 = cmd.CreateParameter("NoIzinPenerima", adVarChar, adParamInput, 200, txtNoIzinPenerima)
    cmd.Parameters.append prm20
    Set prm21 = cmd.CreateParameter("NamaPengangkut", adVarChar, adParamInput, 200, txtJenisSarana)
    cmd.Parameters.append prm21
    Set prm22 = cmd.CreateParameter("NoPolisi", adVarChar, adParamInput, 200, txtNoPolisi)
    cmd.Parameters.append prm22
    Set prm23 = cmd.CreateParameter("Volume", adDecimal, adParamInput, , CDbl(txtVolume))
    prm23.Precision = 38
    prm23.NumericScale = 4
    cmd.Parameters.append prm23
    Set prm24 = cmd.CreateParameter("Bruto", adDecimal, adParamInput, , CDbl(txtBrutoBarang))
    prm24.Precision = 38
    prm24.NumericScale = 4
    cmd.Parameters.append prm24
    Set prm25 = cmd.CreateParameter("Netto", adDecimal, adParamInput, , CDbl(txtNettoBarang))
    prm25.Precision = 38
    prm25.NumericScale = 4
    cmd.Parameters.append prm25
    Set prm26 = cmd.CreateParameter("KodeValuta", adVarChar, adParamInput, 200, txtValuta)
    cmd.Parameters.append prm26
    Set prm27 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , CDbl(txtNilaiCIF))
    prm27.Precision = 38
    prm27.NumericScale = 4
    cmd.Parameters.append prm27
    Set prm28 = cmd.CreateParameter("CIFRupiah", adDecimal, adParamInput, , CDbl(txtHargaPenyerahan))
    prm28.Precision = 38
    prm28.NumericScale = 4
    cmd.Parameters.append prm28
    
    cmd.Execute Y
        
    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC27Detail_Ins"
    
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(txtNoPengajuan, "-", ""))
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("KodeKantorAsal", adVarChar, adParamInput, 200, txtKPBBCBongkar)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("KodeKantorTujuan", adVarChar, adParamInput, 200, txtKantorTujuan)
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeJenisTPB", adVarChar, adParamInput, 200, Left(cboAsal, 1))
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTujuanTPB", adVarChar, adParamInput, 200, Left(cboTujuan, 1))
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("KodeTujuanPengiriman", adVarChar, adParamInput, 200, Left(cboPengiriman, 1))
        cmd.Parameters.append prm6
        Set prm7 = cmd.CreateParameter("KodeMuat", adVarChar, adParamInput, 200, txtKPBBCBongkar)
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("KodeBongkar", adVarChar, adParamInput, 200, txtKantorTujuan) ' Replace(Replace(txtNPWPPengusahaTPB, ".", ""), "-", ""))
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("NamaTTD", adVarChar, adParamInput, 200, txtPemberitahu)
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("JabatanTTD", adVarChar, adParamInput, 200, txtJabatan)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("KotaTTD", adVarChar, adParamInput, 200, txtTempat)
        cmd.Parameters.append prm11
        Set prm12 = cmd.CreateParameter("TanggalTTD", adDate, adParamInput, , Format(dtpTanggal, "yyyy-MM-dd"))
        cmd.Parameters.append prm12
        Set prm13 = cmd.CreateParameter("IDPengusaha", adVarChar, adParamInput, 200, Replace(Replace(txtNPWPPengusahaTPB, ".", ""), "-", ""))
        cmd.Parameters.append prm13
        Set prm14 = cmd.CreateParameter("NamaPengusaha", adVarChar, adParamInput, 200, txtNamaPengusahaTPB) 'Left(cboNPWPPenerima, 1))
        cmd.Parameters.append prm14
        Set prm15 = cmd.CreateParameter("NomorIjinTPB", adVarChar, adParamInput, 200, txtNoIzinPengusahaTPB)
        cmd.Parameters.append prm15
        Set prm29 = cmd.CreateParameter("AlamatPengusaha", adVarChar, adParamInput, 200, txtAlamatPengusahaTPB)
        cmd.Parameters.append prm29
        Set prm16 = cmd.CreateParameter("KodeIDPenerimaBarang", adVarChar, adParamInput, 200, Left(cboNPWPPenerima, 1))
        cmd.Parameters.append prm16
        Set prm17 = cmd.CreateParameter("IDPenerimaBarang", adVarChar, adParamInput, 200, Replace(Replace(txtNPWPPenerima, ".", ""), "-", ""))
        cmd.Parameters.append prm17
        Set prm18 = cmd.CreateParameter("NamaPenerimaBarang", adVarChar, adParamInput, 200, txtNPWPPenerima)
        cmd.Parameters.append prm18
        Set prm19 = cmd.CreateParameter("AlamatPenerimaBarang", adVarChar, adParamInput, 200, txtNPWPPenerima)
        cmd.Parameters.append prm19
        Set prm20 = cmd.CreateParameter("NomorIjinTPBPenerima", adVarChar, adParamInput, 200, txtNoIzinPenerima)
        cmd.Parameters.append prm20
        Set prm21 = cmd.CreateParameter("NamaPengangkut", adVarChar, adParamInput, 200, txtJenisSarana)
        cmd.Parameters.append prm21
        Set prm22 = cmd.CreateParameter("NoPolisi", adVarChar, adParamInput, 200, txtNoPolisi)
        cmd.Parameters.append prm22
        Set prm23 = cmd.CreateParameter("Volume", adDecimal, adParamInput, , CDbl(txtVolume))
        prm23.Precision = 38
        prm23.NumericScale = 4
        cmd.Parameters.append prm23
        Set prm24 = cmd.CreateParameter("Bruto", adDecimal, adParamInput, , CDbl(txtBrutoBarang))
        prm24.Precision = 38
        prm24.NumericScale = 4
        cmd.Parameters.append prm24
        Set prm25 = cmd.CreateParameter("Netto", adDecimal, adParamInput, , CDbl(txtNettoBarang))
        prm25.Precision = 38
        prm25.NumericScale = 4
        cmd.Parameters.append prm25
        Set prm26 = cmd.CreateParameter("KodeValuta", adVarChar, adParamInput, 200, txtValuta)
        cmd.Parameters.append prm26
        Set prm27 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , CDbl(txtNilaiCIF))
        prm27.Precision = 38
        prm27.NumericScale = 4
        cmd.Parameters.append prm27
        Set prm28 = cmd.CreateParameter("CIFRupiah", adDecimal, adParamInput, , CDbl(txtHargaPenyerahan))
        prm28.Precision = 38
        prm28.NumericScale = 4
        cmd.Parameters.append prm28
    
        cmd.Execute Y

    End If
    
    
    LblErrMsg = DisplayMsg(1101)

End Sub
'################################### End Procedure ###############################################

'################################### Start Function ###############################################
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
    ElseIf txtNPWPPengusahaTPB.Text = "" Then
        txtNPWPPengusahaTPB.SetFocus
        SSTab1.Tab = 0
        LblErrMsg = "Please Input NPWP Pengusaha TPB!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtNamaPengusahaTPB.Text = "" Then
        txtNamaPengusahaTPB.SetFocus
        SSTab1.Tab = 0
        LblErrMsg = "Please Input Nama Pengusaha TPB!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtAlamatPengusahaTPB.Text = "" Then
        txtAlamatPengusahaTPB.SetFocus
        SSTab1.Tab = 0
        LblErrMsg = "Please Input Alamat Pengusaha TPB!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtNoIzinPengusahaTPB.Text = "" Then
        txtNoIzinPengusahaTPB.SetFocus
        SSTab1.Tab = 0
        LblErrMsg = "Please Input No Izin Pengusaha TPB!"
        uf_ValidateInput = False
        Exit Function
'    ElseIf cboTipeAPIPengusahaTPB.Text = "" Then
'        cboTipeAPIPengusahaTPB.SetFocus
'        SSTab1.Tab = 0
'        LblErrMsg = "Please Input Tipe API Pengusaha TPB!"
'        uf_ValidateInput = False
'        Exit Function
'    ElseIf txtNomorAPIPengusahaTPB.Text = "" Then
'        txtNomorAPIPengusahaTPB.SetFocus
'        SSTab1.Tab = 0
'        LblErrMsg = "Please Input Nomor API Pengusaha TPB!"
'        uf_ValidateInput = False
'        Exit Function
    ElseIf cboNPWPPenerima.Text = "" Then
        cboNPWPPenerima.SetFocus
        SSTab1.Tab = 1
        LblErrMsg = "Please Input Tipe NPWP Pemilik!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtNPWPPenerima.Text = "" Then
        txtNPWPPenerima.SetFocus
        SSTab1.Tab = 1
        LblErrMsg = "Please Input NPWP Pemilik!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtNPWPPenerima.Text = "" Then
        txtNPWPPenerima.SetFocus
        SSTab1.Tab = 1
        LblErrMsg = "Please Input Nama Pemilik!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtNPWPPenerima.Text = "" Then
        txtNPWPPenerima.SetFocus
        SSTab1.Tab = 1
        LblErrMsg = "Please Input Alamat Pemilik!"
        uf_ValidateInput = False
        Exit Function
    ElseIf CDbl(txtBrutoBarang.Text) <= 0 Then
        txtBrutoBarang.SetFocus
        SSTab2.Tab = 0
        LblErrMsg = "Please Input Bruto!"
        uf_ValidateInput = False
        Exit Function
    ElseIf CDbl(txtNettoBarang.Text) <= 0 Then
        txtNettoBarang.SetFocus
        SSTab2.Tab = 0
        LblErrMsg = "Please Input Netto!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtValuta.Text = "" Then
        txtValuta.SetFocus
        SSTab2.Tab = 1
        LblErrMsg = "Please Input Valuta!"
        uf_ValidateInput = False
        Exit Function
    ElseIf CDbl(txtNilaiCIF.Text) <= 0 Then
        txtNilaiCIF.SetFocus
        SSTab2.Tab = 1
        LblErrMsg = "Please Input Nilai CIF!"
        uf_ValidateInput = False
        Exit Function
    ElseIf CDbl(txtHargaPenyerahan.Text) <= 0 Then
        txtHargaPenyerahan.SetFocus
        SSTab2.Tab = 1
        LblErrMsg = "Please Input Harga Penyerahan!"
        uf_ValidateInput = False
        Exit Function
    End If
    
    uf_ValidateInput = True
End Function


'################################### End Function ###############################################


Private Sub btnBrowseDokumen_Click()
    frmBC27BrowseDokumen.txtNoAju = Replace(txtNoPengajuan, "-", "")
    frmBC27BrowseDokumen.up_GridLoad
    frmBC27BrowseDokumen.Show 1
    up_GridLoadDokumen
    up_LoadDataBC27 txtNoPengajuan
End Sub

Private Sub cmdAction_Click(Index As Integer)
If Index = 0 Then
    frmBC27List.Show
    Unload Me
ElseIf Index = 1 Then
    If uf_ValidateInput = False Then Exit Sub
    
    Call up_SaveDetailBC27
ElseIf Index = 2 Then
    If MsgBox("Are you sure want to synchronize the data?", vbYesNo + vbExclamation, "Confirmation") = vbYes Then
        up_Syncronize
    End If
End If

End Sub

Private Sub cmdAddBarang_Click()
frmBC27BrowseBarang.txtNoPengajuan = Replace(txtNoPengajuan, "-", "")
frmBC27BrowseBarang.txtNoSeri = (gridBarang.Rows - 1) + 1
frmBC27BrowseBarang.cmdPrev.Enabled = False
frmBC27BrowseBarang.cmdNext.Enabled = False
'frmBC27BrowseBarang.gd_NDPBM = CDbl(txtNDPBM)
frmBC27BrowseBarang.Show 1
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

Private Sub cmdCopyPemilik_Click()
'cboTipeKontainer = cboNPWPPenerima
txtNPWPPenerima = txtNPWPPenerima
txtNamaPenerima = txtNPWPPenerima
txtAlamatPenerima = txtNPWPPenerima
'cboTipeAPIPenerima = cboTipeAPIPemilik
'txtNomorAPIPenerima = txtNomorAPIPemilik
End Sub

Private Sub cmdCopyPengusaha_Click()
cboNPWPPenerima = cboTipeNPWPPengusahaTPB
txtNPWPPenerima = txtNPWPPengusahaTPB
txtNPWPPenerima = txtNamaPengusahaTPB
txtNPWPPenerima = txtAlamatPengusahaTPB
'cboTipeAPIPemilik = cboTipeAPIPengusahaTPB
'txtNomorAPIPemilik = txtNomorAPIPengusahaTPB
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
    If txtIDKontainer = "" Then
        LblErrMsg = "Please select kontainer"
        Exit Sub
    End If
    If MsgBox("Are you sure want to delete?", vbYesNo + vbExclamation, "Delete") = vbYes Then
        up_DeleteKontainer
    End If
End Sub

Private Sub cmdDetailBarang_Click()
    If gridBarang.TextMatrix(gridBarang.RowSel, colHideNoSeri) = "" Then Exit Sub
    frmBC27BrowseBarang.txtNoSeri = gridBarang.TextMatrix(gridBarang.RowSel, colHideNoSeri)
    frmBC27BrowseBarang.txtNoPengajuan = Replace(txtNoPengajuan, "-", "")
    frmBC27BrowseBarang.cmdDelete.Enabled = True
    frmBC27BrowseBarang.up_LoadDataBarang txtNoPengajuan, gridBarang.TextMatrix(gridBarang.RowSel, colHideNoSeri)
    frmBC27BrowseBarang.up_LoadDataBahanBakuImpor txtNoPengajuan, gridBarang.TextMatrix(gridBarang.RowSel, colHideNoSeri), 1
    frmBC27BrowseBarang.up_LoadDataBahanBakuLokal txtNoPengajuan, gridBarang.TextMatrix(gridBarang.RowSel, colHideNoSeri), 1
    
    
    frmBC27BrowseBarang.CekData = True
    frmBC27BrowseBarang.cekSubmit = True
    frmBC27BrowseBarang.txtTotalItem = gridBarang.Rows - 1
    'frmBC27BrowseBarang.gd_NDPBM = CDbl(txtNDPBM)
    frmBC27BrowseBarang.Show 1
End Sub

Private Sub cmdSaveKemasan_Click()
up_SaveKemasan
End Sub

Private Sub cmdSaveKontainer_Click()
up_SaveKontainer
End Sub

Private Sub Form_Activate()
up_GridLoadBarang
up_GridLoadDokumen
up_GridLoadKemasan
up_GridLoadKontainer
up_GridHeaderStatus
up_GridHeaderRespon
End Sub

Private Sub Form_Load()
up_FillComboTujuan
up_FillComboAsal
up_FillComboPengiriman
up_FillComboKodeID cboTipeNPWPPengusahaTPB
up_FillComboKodeID cboNPWPPenerima

up_FillComboGeneral cboUkuranKontainer, "Bea_Cukai_Ukuran_Kontainer", "KODE_UKURAN_KONTAINER", "URAIAN_UKURAN_KONTAINER", 90, 110
up_FillComboGeneral cboTipeKontainer, "Bea_Cukai_Tipe_Kontainer", "KODE_TIPE_KONTAINER", "URAIAN_TIPE_KONTAINER", 60, 70

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
        txtNomorKontainer1 = Left(gridKontainer.TextMatrix(gridKontainer.RowSel, colNomorKontainer), 4)
        txtNomorKontainer2 = Mid(gridKontainer.TextMatrix(gridKontainer.RowSel, colNomorKontainer), 5, 7)
        cboUkuranKontainer = gridKontainer.TextMatrix(gridKontainer.RowSel, colHideUkuran)
        cboTipeKontainer = gridKontainer.TextMatrix(gridKontainer.RowSel, colHideTipe)
        txtIDKontainer = gridKontainer.TextMatrix(gridKontainer.RowSel, colIDKontainer)
        txtKeteranganKontainer = gridKontainer.TextMatrix(gridKontainer.RowSel, colHideKeterangan)
    End If
End Sub

'Private Sub txtNPWPPenerima_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub txtAlamatPengusahaTPB_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBrutoBarang_GotFocus()
If txtBrutoBarang = "" Then txtBrutoBarang = 0
txtBrutoBarang = CDbl(txtBrutoBarang)
End Sub

Private Sub txtBrutoBarang_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtBrutoBarang_LostFocus()
If txtBrutoBarang = "" Then txtBrutoBarang = 0
txtBrutoBarang = Format(CDbl(txtBrutoBarang), "#,0.0000")
End Sub

'if txtVolume

Private Sub txtFasilitasImpor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFasilitasImpor2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtHargaPenyerahan_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtHargaPenyerahan_LostFocus()
If txtHargaPenyerahan = "" Then txtHargaPenyerahan = 0
txtHargaPenyerahan = Format(CDbl(txtHargaPenyerahan), "#,0.00")
End Sub

Private Sub txtInvoiceDokumen_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtJabatan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtJenisKemasan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    frmBC27BrowseGeneral.gs_TableName = "Kemasan"
    frmBC27BrowseGeneral.Show 1
End If
End Sub

Private Sub txtJenisKemasan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtJenisKemasan_LostFocus()
gb_LoadDataMaster "Bea_Cukai_Kemasan", "Uraian_Kemasan", lblJenisKemasan, "Where Kode_Kemasan = '" & txtJenisKemasan & "'"
End Sub

Private Sub txtJumlahKemasan_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtKeteranganKontainer_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtKontrak_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtKPBBCBongkar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    frmBC27BrowseGeneral.gs_TableName = "Kantor Pabean"
    frmBC27BrowseGeneral.Show 1
End If
End Sub

Private Sub txtKantorTujuan_Change()
up_LoadKantorKPPBCTujuan txtKantorTujuan
End Sub

Private Sub txtKantorTujuan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    frmBC27BrowseGeneral.gs_TableName = "Kantor Pabean"
    frmBC27BrowseGeneral.Show 1
End If
End Sub

Private Sub txtKPBBCBongkar_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtKantorTujuan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtKPBBCBongkar_LostFocus()
    up_LoadKantorKPPBCBongkar txtKPBBCBongkar
End Sub

Private Sub txtKantorTujuan_LostFocus()
    up_LoadKantorKPPBCTujuan txtKantorTujuan
End Sub

Private Sub txtMerkKemasan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'Private Sub txtNPWPPenerima_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub txtNamaPengusahaTPB_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNDPBM_GotFocus()
    'txtNDPBM = CDbl(txtNDPBM)
End Sub

Private Sub txtNDPBM_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtNDPBM_LostFocus()
    'txtNDPBM = Format(CDbl(txtNDPBM), "#,0.0000")
End Sub

Private Sub txtNettoBarang_GotFocus()
If txtNettoBarang = "" Then txtNettoBarang = 0
txtNettoBarang = CDbl(txtNettoBarang)
End Sub

Private Sub txtNettoBarang_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtNettoBarang_LostFocus()
If txtNettoBarang = "" Then txtNettoBarang = 0
txtNettoBarang = Format(CDbl(txtNettoBarang), "#,0.0000")
End Sub

Private Sub txtNilaiCIF_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtNilaiCIF_LostFocus()
If txtNilaiCIF = "" Then txtNilaiCIF = 0
txtNilaiCIF = Format(CDbl(txtNilaiCIF), "#,0.00")
End Sub



Private Sub txtNIPER_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNoIzinPengusahaTPB_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNomorAPIPemilik_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNomorAPIPengusahaTPB_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNomorKontainer1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))


Select Case KeyAscii
    Case 32 To 64, 91 To 96, 123 To 126
        KeyAscii = 0
        Exit Sub
End Select

If KeyAscii = Asc("'") Then KeyAscii = 0
    
End Sub

Private Sub txtNomorKontainer2_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

'Private Sub txtNPWPPenerima_GotFocus()
'txtNPWPPenerima = Replace(Replace(txtNPWPPenerima, ".", ""), "-", "")
'End Sub

Private Sub txtNPWPPenerima_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNPWPPenerima_LostFocus()

If Left(cboNPWPPenerima, 1) = "1" Then
    If Len(txtNPWPPenerima) > 15 Then
        LblErrMsg.Caption = "Identitas/NPWP No maximum of 15 characters"
        txtNPWPPenerima.SetFocus
        Exit Sub
    End If
    txtNPWPPenerima = Left(txtNPWPPenerima.Text, 2) & "." & Mid(txtNPWPPenerima.Text, 3, 3) & "." & Mid(txtNPWPPenerima.Text, 6, 3) & "." & Mid(txtNPWPPenerima.Text, 9, 1) & "-" & Mid(txtNPWPPenerima.Text, 10, 3) & "." & Mid(txtNPWPPenerima.Text, 13, 3)
End If

End Sub

Private Sub txtNPWPPenerima_GotFocus()
txtNPWPPenerima = Replace(Replace(txtNPWPPenerima, ".", ""), "-", "")
End Sub

'Private Sub txtNPWPPenerima_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
'    If KeyAscii = Asc("'") Then KeyAscii = 0
'End Sub

'Private Sub txtNPWPPenerima_LostFocus()
'If Left(cboTipeNPWPPenerima, 1) = "1" Then
'    If Len(txtNPWPPenerima) > 15 Then
'        LblErrMsg.Caption = "Identitas/NPWP No maximum of 15 characters"
'        txtNPWPPenerima.SetFocus
'        Exit Sub
'    End If
'    txtNPWPPenerima = Left(txtNPWPPenerima.Text, 2) & "." & Mid(txtNPWPPenerima.Text, 3, 3) & "." & Mid(txtNPWPPenerima.Text, 6, 3) & "." & Mid(txtNPWPPenerima.Text, 9, 1) & "-" & Mid(txtNPWPPenerima.Text, 10, 3) & "." & Mid(txtNPWPPenerima.Text, 13, 3)
'End If
'End Sub

Private Sub txtNPWPPengusahaTPB_GotFocus()
txtNPWPPengusahaTPB = Replace(Replace(txtNPWPPengusahaTPB, ".", ""), "-", "")
End Sub

Private Sub txtNPWPPengusahaTPB_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtNPWPPengusahaTPB_LostFocus()
    If Len(txtNPWPPengusahaTPB) > 15 Then
        LblErrMsg.Caption = "Identitas/NPWP No maximum of 15 characters"
        txtNPWPPengusahaTPB.SetFocus
        Exit Sub
    End If
    txtNPWPPengusahaTPB = Left(txtNPWPPengusahaTPB.Text, 2) & "." & Mid(txtNPWPPengusahaTPB.Text, 3, 3) & "." & Mid(txtNPWPPengusahaTPB.Text, 6, 3) & "." & Mid(txtNPWPPengusahaTPB.Text, 9, 1) & "-" & Mid(txtNPWPPengusahaTPB.Text, 10, 3) & "." & Mid(txtNPWPPengusahaTPB.Text, 13, 3)

End Sub

Private Sub txtPackingList_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPemberitahu_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTempat_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtValuta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    frmBC27BrowseGeneral.gs_TableName = "Valuta"
    frmBC27BrowseGeneral.Show 1
End If
End Sub

Private Sub txtValuta_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtValuta_LostFocus()
gb_LoadDataMaster "Bea_Cukai_Valuta", "Uraian_Valuta", lblValuta, "Where Kode_Valuta = '" & txtValuta & "'"
End Sub
