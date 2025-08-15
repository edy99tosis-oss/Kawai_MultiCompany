VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBC27BrowseBarang 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Browse Barang"
   ClientHeight    =   10395
   ClientLeft      =   4755
   ClientTop       =   1920
   ClientWidth     =   13320
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBC27BrowseBarang.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10395
   ScaleWidth      =   13320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrev 
      BackColor       =   &H0080C0FF&
      Caption         =   "Prev"
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   179
      Top             =   9840
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H0080C0FF&
      Caption         =   "Next"
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   178
      Top             =   9840
      Width           =   975
   End
   Begin VB.TextBox txtNoPengajuan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   350
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   177
      Top             =   9840
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   120
      TabIndex        =   34
      Tag             =   "TFTT*/"
      Top             =   9120
      Width           =   12855
      Begin VB.Label LblerrMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LblErrMsg"
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
         Left            =   105
         TabIndex        =   35
         Tag             =   "TTFF*/"
         Top             =   195
         Width           =   12570
      End
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H0080FFFF&
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   9840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9840
      Width           =   975
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Submit"
      Height          =   375
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9840
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0080FFFF&
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   9840
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   15690
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      BackColor       =   16637923
      TabCaption(0)   =   "Data Detail Barang"
      TabPicture(0)   =   "frmBC27BrowseBarang.frx":0E42
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detail Bahan Baku Impor"
      TabPicture(1)   =   "frmBC27BrowseBarang.frx":0E5E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Detail Bahan Baku Lokal"
      TabPicture(2)   =   "frmBC27BrowseBarang.frx":0E7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame17"
      Tab(2).Control(1)=   "Frame16"
      Tab(2).Control(2)=   "Frame15"
      Tab(2).Control(3)=   "Frame14"
      Tab(2).Control(4)=   "Frame13"
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame17 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   164
         Top             =   5160
         Width           =   12615
         Begin VB.TextBox txtPPNFasilitasLokal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   350
            Left            =   3000
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   174
            Text            =   "0.00"
            Top             =   1200
            Width           =   3225
         End
         Begin VB.TextBox txtPPNBayarLokal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   350
            Left            =   3000
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   171
            Text            =   "0.00"
            Top             =   720
            Width           =   3225
         End
         Begin VB.TextBox txtPercentLokal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   5520
            MaxLength       =   10
            TabIndex        =   166
            Text            =   "0.00"
            Top             =   240
            Width           =   705
         End
         Begin VB.TextBox txtPPNLokal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   165
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rp"
            Height          =   195
            Index           =   57
            Left            =   2640
            TabIndex        =   176
            Top             =   1275
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PPN Fasilitas"
            Height          =   195
            Index           =   56
            Left            =   240
            TabIndex        =   175
            Top             =   1275
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rp"
            Height          =   195
            Index           =   55
            Left            =   2640
            TabIndex        =   173
            Top             =   798
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PPN Bayar"
            Height          =   195
            Index           =   54
            Left            =   240
            TabIndex        =   172
            Top             =   795
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Index           =   53
            Left            =   6360
            TabIndex        =   170
            Top             =   315
            Width           =   180
         End
         Begin MSForms.ComboBox cboJenisPPNLokal 
            Height          =   345
            Left            =   3000
            TabIndex        =   169
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2415
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4260;609"
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
            Caption         =   "%"
            Height          =   195
            Index           =   52
            Left            =   2715
            TabIndex        =   168
            Top             =   315
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PPN"
            Height          =   195
            Index           =   51
            Left            =   240
            TabIndex        =   167
            Top             =   315
            Width           =   330
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "HARGA && SATUAN"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   154
         Top             =   3960
         Width           =   12615
         Begin VB.TextBox txtHargaPerolahen 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   158
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtHargaPenyerahanLokal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            TabIndex        =   157
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtJumlahSatuanLokal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   156
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtSatuanLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9840
            TabIndex        =   155
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label73 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Perolehan"
            Height          =   255
            Left            =   240
            TabIndex        =   163
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label Label72 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Penyerahan"
            Height          =   255
            Left            =   4440
            TabIndex        =   162
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label70 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Satuan"
            Height          =   255
            Left            =   240
            TabIndex        =   161
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   1575
         End
         Begin VB.Label Label69 
            BackStyle       =   0  'Transparent
            Caption         =   "Satuan"
            Height          =   255
            Left            =   8760
            TabIndex        =   160
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1215
         End
         Begin VB.Line Line16 
            X1              =   10800
            X2              =   12480
            Y1              =   490
            Y2              =   490
         End
         Begin VB.Label lblSatuanLokal 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   10800
            TabIndex        =   159
            Tag             =   "TTFF*/"
            Top             =   280
            Width           =   1455
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "DATA BAHAN BAKU"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   139
         Top             =   2400
         Width           =   12615
         Begin VB.TextBox txtKodeBarangLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   146
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtNomorHSLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5880
            TabIndex        =   145
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtUraianBarangLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   144
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   10215
         End
         Begin VB.TextBox txtTipeLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   143
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtUkuranLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4560
            TabIndex        =   142
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtSpfLainLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7440
            TabIndex        =   141
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtMerkLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   10320
            TabIndex        =   140
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label66 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode"
            Height          =   255
            Left            =   240
            TabIndex        =   153
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label65 
            BackStyle       =   0  'Transparent
            Caption         =   "Nomor HS"
            Height          =   255
            Left            =   4440
            TabIndex        =   152
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label64 
            BackStyle       =   0  'Transparent
            Caption         =   "Uraian Barang"
            Height          =   255
            Left            =   240
            TabIndex        =   151
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   1575
         End
         Begin VB.Label Label63 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipe"
            Height          =   255
            Left            =   240
            TabIndex        =   150
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   1215
         End
         Begin VB.Label Label62 
            BackStyle       =   0  'Transparent
            Caption         =   "Ukuran"
            Height          =   255
            Left            =   3600
            TabIndex        =   149
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   855
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Spf Lain"
            Height          =   255
            Left            =   6480
            TabIndex        =   148
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   855
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "Merk"
            Height          =   255
            Left            =   9360
            TabIndex        =   147
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   855
         End
      End
      Begin VB.Frame Frame14 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   125
         Top             =   1200
         Width           =   12615
         Begin VB.TextBox txtKPPBCLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7080
            TabIndex        =   130
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtNoLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   129
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtDokumenAsalLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   128
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtNoAjuLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7080
            TabIndex        =   127
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtUrutKeLokal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   11160
            TabIndex        =   126
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   975
         End
         Begin MSComCtl2.DTPicker dtpTglLokal 
            Height          =   315
            Left            =   3360
            TabIndex        =   187
            Tag             =   "TTFF*/"
            Top             =   600
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
            Format          =   173408259
            CurrentDate     =   37798
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "No Aju"
            Height          =   255
            Left            =   5640
            TabIndex        =   138
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   1695
         End
         Begin VB.Line Line15 
            X1              =   8520
            X2              =   12000
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Label lblKPPBCLokal 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   8520
            TabIndex        =   137
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   3495
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "KPPBC Dok"
            Height          =   255
            Left            =   5640
            TabIndex        =   136
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "No / Tgl Dok"
            Height          =   255
            Index           =   48
            Left            =   240
            TabIndex        =   135
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   1695
         End
         Begin VB.Line Line14 
            X1              =   2880
            X2              =   4680
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Label lblDokAsalLokal 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   2880
            TabIndex        =   134
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "Dok Asal"
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            Height          =   255
            Left            =   3240
            TabIndex        =   132
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   135
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "Urut ke"
            Height          =   255
            Left            =   10080
            TabIndex        =   131
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   975
         End
      End
      Begin VB.Frame Frame13 
         Height          =   735
         Left            =   -74880
         TabIndex        =   120
         Top             =   480
         Width           =   12615
         Begin VB.CommandButton cmdNewLokal 
            BackColor       =   &H0080FFFF&
            Caption         =   "New"
            Height          =   375
            Left            =   11400
            Style           =   1  'Graphical
            TabIndex        =   184
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtNoSeriLokal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   1920
            TabIndex        =   122
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtTotalLokal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   3360
            TabIndex        =   121
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "Bahan Baku"
            Height          =   255
            Left            =   240
            TabIndex        =   124
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "dari"
            Height          =   255
            Left            =   2880
            TabIndex        =   123
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   495
         End
      End
      Begin VB.Frame Frame12 
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   120
         TabIndex        =   79
         Top             =   4680
         Width           =   12615
         Begin VB.TextBox txtSatuanTarifImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   5880
            MaxLength       =   10
            TabIndex        =   94
            Top             =   405
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtJumlahSatuanBMImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   93
            Top             =   840
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtTarifPersen5Impor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   5880
            MaxLength       =   10
            TabIndex        =   92
            Top             =   2085
            Width           =   705
         End
         Begin VB.TextBox txtPPHImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   91
            Top             =   2085
            Width           =   705
         End
         Begin VB.TextBox txtTarifPersen4Impor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   5880
            MaxLength       =   10
            TabIndex        =   90
            Top             =   1680
            Width           =   705
         End
         Begin VB.TextBox txtPPNBMImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   89
            Top             =   1680
            Width           =   705
         End
         Begin VB.TextBox txtTarifPersen3Impor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   5880
            MaxLength       =   10
            TabIndex        =   88
            Top             =   1260
            Width           =   705
         End
         Begin VB.TextBox txtPPNImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   87
            Top             =   1260
            Width           =   705
         End
         Begin VB.TextBox txtTarifPersen2Impor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   5880
            MaxLength       =   5
            TabIndex        =   86
            Top             =   840
            Width           =   705
         End
         Begin VB.TextBox txtTarifPersen1Impor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   85
            Top             =   405
            Width           =   945
         End
         Begin VB.CommandButton cmdBrowseTarifImpor 
            BackColor       =   &H0080FFFF&
            Caption         =   "O"
            Height          =   375
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtPersenCukaiImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   5880
            MaxLength       =   5
            TabIndex        =   83
            Top             =   3330
            Width           =   705
         End
         Begin VB.TextBox txtTarifImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   4590
            MaxLength       =   5
            TabIndex        =   82
            Top             =   2925
            Width           =   945
         End
         Begin VB.TextBox txtSatuanCukaiImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   5880
            MaxLength       =   5
            TabIndex        =   81
            Top             =   2925
            Width           =   705
         End
         Begin VB.TextBox txtJumlahTarifImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   80
            Top             =   3330
            Width           =   1000
         End
         Begin VB.Label Label2 
            Caption         =   "BM Bhn Baku"
            Height          =   255
            Left            =   120
            TabIndex        =   188
            Top             =   450
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Satuan"
            Height          =   195
            Index           =   47
            Left            =   120
            TabIndex        =   119
            Top             =   915
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Index           =   46
            Left            =   6720
            TabIndex        =   118
            Top             =   2160
            Width           =   180
         End
         Begin MSForms.ComboBox cboKeterangan5Impor 
            Height          =   345
            Left            =   2640
            TabIndex        =   117
            Tag             =   "TTFF*/"
            Top             =   2085
            Width           =   2895
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5106;609"
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
            Caption         =   "%"
            Height          =   195
            Index           =   45
            Left            =   2355
            TabIndex        =   116
            Top             =   2160
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PPh"
            Height          =   195
            Index           =   44
            Left            =   120
            TabIndex        =   115
            Top             =   2160
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Index           =   41
            Left            =   6720
            TabIndex        =   114
            Top             =   1755
            Width           =   180
         End
         Begin MSForms.ComboBox cboKeterangan4Impor 
            Height          =   345
            Left            =   2640
            TabIndex        =   113
            Tag             =   "TTFF*/"
            Top             =   1680
            Width           =   2895
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5106;609"
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
            Caption         =   "%"
            Height          =   195
            Index           =   40
            Left            =   2355
            TabIndex        =   112
            Top             =   1755
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PPnBM"
            Height          =   195
            Index           =   38
            Left            =   120
            TabIndex        =   111
            Top             =   1755
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Index           =   23
            Left            =   6720
            TabIndex        =   110
            Top             =   1335
            Width           =   180
         End
         Begin MSForms.ComboBox cboKeterangan3Impor 
            Height          =   345
            Left            =   2640
            TabIndex        =   109
            Tag             =   "TTFF*/"
            Top             =   1260
            Width           =   2895
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5106;609"
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
            Caption         =   "%"
            Height          =   195
            Index           =   22
            Left            =   2355
            TabIndex        =   108
            Top             =   1335
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PPN"
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   107
            Top             =   1320
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Index           =   20
            Left            =   6720
            TabIndex        =   106
            Top             =   915
            Width           =   180
         End
         Begin MSForms.ComboBox cboKeterangan2Impor 
            Height          =   345
            Left            =   2640
            TabIndex        =   105
            Tag             =   "TTFF*/"
            Top             =   840
            Width           =   2895
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5106;609"
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
            Caption         =   "%"
            Height          =   195
            Index           =   19
            Left            =   5640
            TabIndex        =   104
            Top             =   480
            Width           =   180
         End
         Begin MSForms.ComboBox cboKeterangan1Impor 
            Height          =   345
            Left            =   1560
            TabIndex        =   103
            Tag             =   "TTFF*/"
            Top             =   405
            Width           =   2415
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4260;609"
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
            Caption         =   "%"
            Height          =   195
            Index           =   18
            Left            =   6720
            TabIndex        =   102
            Top             =   3405
            Width           =   180
         End
         Begin MSForms.ComboBox cboKeteranganJenisTarif 
            Height          =   345
            Left            =   2640
            TabIndex        =   101
            Tag             =   "TTFF*/"
            Top             =   3330
            Width           =   2895
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5106;609"
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
            Caption         =   "/"
            Height          =   195
            Index           =   17
            Left            =   5640
            TabIndex        =   100
            Top             =   3000
            Width           =   75
         End
         Begin MSForms.ComboBox cboJenisTarifImpor 
            Height          =   345
            Left            =   1560
            TabIndex        =   99
            Tag             =   "TTFF*/"
            Top             =   2925
            Width           =   2895
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5106;609"
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
            Caption         =   "Jenis Tarif"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   98
            Top             =   3000
            Width           =   870
         End
         Begin MSForms.ComboBox cboCukaiImpor 
            Height          =   345
            Left            =   1560
            TabIndex        =   97
            Tag             =   "TTFF*/"
            Top             =   2520
            Width           =   3975
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "7011;609"
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
            Caption         =   "Cukai"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   96
            Top             =   2580
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   95
            Top             =   3405
            Width           =   600
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "HARGA && SATUAN"
         Height          =   1095
         Left            =   120
         TabIndex        =   69
         Top             =   3480
         Width           =   12615
         Begin VB.TextBox txtSatuanImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9720
            TabIndex        =   76
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtJumlahSatuanImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            TabIndex        =   74
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtHargaPenyerahanImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   72
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtHargaCIFImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   70
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblSatuanImpor 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   10680
            TabIndex        =   78
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Line Line10 
            X1              =   10680
            X2              =   12120
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "Satuan"
            Height          =   255
            Left            =   8160
            TabIndex        =   77
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Satuan"
            Height          =   255
            Left            =   4560
            TabIndex        =   75
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Penyerahan"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga CIF"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1575
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "DATA BAHAN BAKU"
         Height          =   1455
         Left            =   120
         TabIndex        =   54
         Top             =   2040
         Width           =   12615
         Begin VB.TextBox txtMerkImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   10320
            TabIndex        =   67
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtSpfLainImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7440
            TabIndex        =   65
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtUkuranImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4560
            TabIndex        =   63
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtTipeImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   61
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtUraianBarangImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   59
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   10215
         End
         Begin VB.TextBox txtNomorHSImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5880
            TabIndex        =   57
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtKodeBarangImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   55
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Merk"
            Height          =   255
            Left            =   9360
            TabIndex        =   68
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   855
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Spf Lain"
            Height          =   255
            Left            =   6480
            TabIndex        =   66
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   855
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Ukuran"
            Height          =   255
            Left            =   3600
            TabIndex        =   64
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   855
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipe"
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   1215
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Uraian Barang"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   1575
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Nomor HS"
            Height          =   255
            Left            =   4440
            TabIndex        =   58
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   -74760
         TabIndex        =   36
         Top             =   3600
         Width           =   12495
         Begin VB.TextBox txtJumlahBahanBaku 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9840
            TabIndex        =   37
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Bahan Baku"
            Height          =   255
            Left            =   7800
            TabIndex        =   38
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "SATUAN DAN HARGA"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   16
         Top             =   2160
         Width           =   12615
         Begin VB.TextBox txtHargaPenyerahan 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9840
            TabIndex        =   28
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   735
            Width           =   2655
         End
         Begin VB.TextBox txtNetto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   5520
            MaxLength       =   100
            TabIndex        =   21
            Top             =   360
            Width           =   1305
         End
         Begin VB.TextBox txtJenisSatuan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   20
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtNilaiCIF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   9840
            Locked          =   -1  'True
            TabIndex        =   19
            Tag             =   "TTFF*/"
            Text            =   "0.00"
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtJumlahSatuan 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   18
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtVolume 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5520
            TabIndex        =   17
            Tag             =   "TTFF*/"
            Top             =   735
            Width           =   1575
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Penyerahan Rp"
            Height          =   255
            Left            =   7800
            TabIndex        =   29
            Tag             =   "TTFF*/"
            Top             =   765
            Width           =   1935
         End
         Begin VB.Label lblJenisSatuan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   6
            Left            =   2640
            TabIndex        =   27
            Top             =   780
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Netto (Kgm)"
            Height          =   195
            Index           =   8
            Left            =   3840
            TabIndex        =   26
            Top             =   390
            Width           =   1050
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Satuan"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Tag             =   "TTFF*/"
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Nilai CIF"
            Height          =   255
            Left            =   7800
            TabIndex        =   24
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Satuan"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Tag             =   "TTFF*/"
            Top             =   390
            Width           =   1575
         End
         Begin VB.Line Line6 
            X1              =   2640
            X2              =   3600
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Volume (M3)"
            Height          =   255
            Left            =   3840
            TabIndex        =   22
            Tag             =   "TTFF*/"
            Top             =   765
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "DATA BARANG BC 2.7"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   12615
         Begin VB.TextBox txtNoSeri 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   9840
            TabIndex        =   190
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtTotalItem 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   11640
            TabIndex        =   189
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtMerk 
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   10920
            MaxLength       =   100
            TabIndex        =   14
            Top             =   1080
            Width           =   1545
         End
         Begin VB.TextBox txtSpfLain 
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   7920
            MaxLength       =   100
            TabIndex        =   12
            Top             =   1080
            Width           =   1545
         End
         Begin VB.TextBox txtUkuran 
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   4920
            MaxLength       =   100
            TabIndex        =   10
            Top             =   1080
            Width           =   1545
         End
         Begin VB.TextBox txtTipe 
            Appearance      =   0  'Flat
            Height          =   350
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   8
            Top             =   1080
            Width           =   1545
         End
         Begin VB.TextBox txtUraianBarang 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   6
            Tag             =   "TTFF*/"
            Top             =   720
            Width           =   10575
         End
         Begin VB.TextBox txtNomorHS 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5280
            TabIndex        =   4
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtKodeBarang 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   2
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Dari"
            Height          =   255
            Left            =   10920
            TabIndex        =   192
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Barang "
            Height          =   255
            Left            =   8880
            TabIndex        =   191
            Tag             =   "TTFF*/"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Merk"
            Height          =   195
            Index           =   3
            Left            =   9840
            TabIndex        =   15
            Top             =   1155
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spef Lain"
            Height          =   195
            Index           =   2
            Left            =   6840
            TabIndex        =   13
            Top             =   1155
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ukuran"
            Height          =   195
            Index           =   1
            Left            =   3840
            TabIndex        =   11
            Top             =   1155
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipe"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   9
            Top             =   1155
            Width           =   360
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Uraian Barang"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Tag             =   "TTFF*/"
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Nomor HS"
            Height          =   255
            Left            =   3840
            TabIndex        =   5
            Tag             =   "TTFF*/"
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Tag             =   "TTFF*/"
            Top             =   390
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1455
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   12615
         Begin VB.CommandButton cmdNewImpor 
            BackColor       =   &H0080FFFF&
            Caption         =   "New"
            Height          =   375
            Left            =   11400
            Style           =   1  'Graphical
            TabIndex        =   183
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtNoSeriImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   7080
            TabIndex        =   181
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtTotalImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   8640
            TabIndex        =   180
            Tag             =   "TTFF*/"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtUrutKeImpor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   11400
            TabIndex        =   52
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtNoAjuImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7080
            TabIndex        =   51
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox txtDokumenAsalImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   42
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtNoImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   41
            Tag             =   "TTFF*/"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtKPPBCImpor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7080
            TabIndex        =   40
            Tag             =   "TTFF*/"
            Top             =   240
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtpTglImpor 
            Height          =   315
            Left            =   3480
            TabIndex        =   186
            Tag             =   "TTFF*/"
            Top             =   600
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
            Format          =   174063619
            CurrentDate     =   37798
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "dari"
            Height          =   255
            Index           =   59
            Left            =   8160
            TabIndex        =   185
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bahan Baku"
            Height          =   255
            Index           =   58
            Left            =   5640
            TabIndex        =   182
            Tag             =   "TTFF*/"
            Top             =   990
            Width           =   1335
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Urut ke"
            Height          =   255
            Left            =   10560
            TabIndex        =   53
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   735
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            Height          =   255
            Left            =   3240
            TabIndex        =   50
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   135
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Dok Asal"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label lblDokAsalImpor 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   2880
            TabIndex        =   47
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1815
         End
         Begin VB.Line Line11 
            X1              =   2880
            X2              =   4680
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "No / Tgl Dok"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   46
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "KPPBC Dok"
            Height          =   255
            Left            =   5640
            TabIndex        =   45
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblKPPBCImpor 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   8520
            TabIndex        =   44
            Tag             =   "TTFF*/"
            Top             =   270
            Width           =   3855
         End
         Begin VB.Line Line9 
            X1              =   8520
            X2              =   12360
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "No Aju"
            Height          =   255
            Left            =   5640
            TabIndex        =   43
            Tag             =   "TTFF*/"
            Top             =   630
            Width           =   1695
         End
      End
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "dari"
      Height          =   255
      Left            =   3120
      TabIndex        =   49
      Tag             =   "TTFF*/"
      Top             =   2070
      Width           =   495
   End
End
Attribute VB_Name = "frmBC27BrowseBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsBarang As ADODB.Recordset
Dim rsbahanbakuimpor As ADODB.Recordset
Dim rsbahanbakulokal As ADODB.Recordset

Public cekSubmit As Boolean
Dim cekLoad As Boolean
Public CekData As Boolean
Public gd_NDPBM As Double

Public gd_HargaPenyerahan As Double

Public sql As String
Public RS As New ADODB.Recordset

'-------------------------------------------
Const colJenisDokumen As Integer = 0
Const colNomorDokumen As Integer = 1
Const colTanggal As Integer = 2
Const colcount As Integer = 3

'-------------------------------------------
Const colJenisDokumenImpor As Integer = 0
Const colNomorDokumenImpor As Integer = 1
Const colTanggalImpor As Integer = 2
Const colcountImpor As Integer = 3

'-------------------------------------------
Const colJenisDokumenLokal As Integer = 0
Const colNomorDokumenLokal As Integer = 1
Const colTanggalLokal As Integer = 2
Const colcountLokal As Integer = 3

Private Sub up_Clear()

LblerrMsg.Caption = ""
End Sub

Private Sub up_ClearImpor()
txtNoSeriImpor = ""
txtDokumenAsalImpor = ""
lblDokAsalImpor.Caption = ""
txtNoImpor = ""
dtpTglImpor.Value = Now
txtKPPBCImpor = ""
lblKPPBCImpor.Caption = ""
txtNoAjuImpor = ""
txtUrutKeImpor = ""
txtTotalImpor = ""
txtKodeBarangImpor = ""
txtNomorHSImpor = ""
txtUraianBarangImpor = ""
txtTipeImpor = ""
txtUkuranImpor = ""
txtSpfLainImpor = ""
txtMerkImpor = ""
txtHargaCIFImpor = ""
txtJumlahSatuanImpor = "0"
txtSatuanImpor = ""

cboKeterangan1Impor = ""
txtTarifPersen1Impor = ""
txtSatuanTarifImpor = ""
txtJumlahSatuanBMImpor = ""
cboKeterangan2Impor = ""
txtTarifPersen2Impor = ""
txtPPNImpor = ""
cboKeterangan3Impor = ""
txtTarifPersen3Impor = ""
txtPPNBMImpor = ""
cboKeterangan4Impor = ""
txtTarifPersen4Impor = ""
txtPPHImpor = ""
cboKeterangan5Impor = ""
txtTarifPersen5Impor = ""
cboCukaiImpor = ""
cboJenisTarifImpor = ""
txtTarifImpor = ""
txtSatuanCukaiImpor = ""
txtJumlahTarifImpor = ""
cboKeteranganJenisTarif = ""
txtPersenCukaiImpor = ""

Label1(14).Caption = ""
End Sub

Private Sub up_ClearLokal()
txtNoSeriLokal = ""
txtDokumenAsalLokal = ""
lblDokAsalLokal.Caption = ""
txtNoLokal = ""
dtpTglLokal.Value = Now
txtKPPBCLokal = ""
lblKPPBCLokal.Caption = ""
txtNoAjuLokal = ""
txtUrutKeLokal = ""
txtTotalLokal = ""
txtKodeBarangLokal = ""
txtNomorHSLokal = ""
txtUraianBarangLokal = ""
txtTipeLokal = ""
txtUkuranLokal = ""
txtSpfLainLokal = ""
txtMerkLokal = ""
txtHargaPerolahen = ""
txtJumlahSatuanLokal = "0"
txtHargaPenyerahanLokal = "0.00"
txtSatuanLokal = ""

txtPPNLokal = "0.00"
cboJenisPPNLokal = ""
txtPercentLokal = "0.00"
txtPPNBayarLokal = "0.00"
txtPPNFasilitasLokal = "0.00"
End Sub

Private Function uf_GetHargaPenyerahan(pNoAju As String, pNoSeriBarang As String) As Double
Dim sql As String
Dim RS As New Recordset

sql = "Select HARGA_PENYERAHAN From Bea_Cukai_TPB_Barang Where No_Pengajuan = '" & pNoAju & "' AND SERI_BARANG = " & pNoSeriBarang & ""
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    uf_GetHargaPenyerahan = RS.Fields("HARGA_PENYERAHAN")
Else
    uf_GetHargaPenyerahan = 0
End If
End Function

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

Private Sub up_FillComboGeneral(pcbo As MSForms.ComboBox, pTable As String, pField1 As String, pField2 As String, pColWidth2 As Integer, pListWidth As Integer)
Dim sql As String
Dim RS As New Recordset

    sql = "Select " & pField1 & ", " & pField2 & " From " & pTable & ""
    Set RS = Db.Execute(sql)

    With pcbo
        .clear
        .ColumnCount = 1
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

Private Sub up_GridHeaderDokumenLokal()
'    With gridDokumenLokal
'        .clear
'        .ColS = colcount
'        .Rows = 1
'
'        .TextMatrix(0, colJenisDokumenLokal) = "Jenis"
'        .TextMatrix(0, colNomorDokumenLokal) = "Nomor"
'        .TextMatrix(0, colTanggalLokal) = "Tanggal"
'
'        .ColWidth(colJenisDokumenLokal) = 1500
'        .ColWidth(colNomorDokumenLokal) = 1500
'        .ColWidth(colTanggalLokal) = 1200
'        .ColAlignment(colNomorDokumenLokal) = flexAlignLeftCenter
'
'        .ColFormat(colTanggalLokal) = "dd MMM yyyy"
'    End With
End Sub

Private Sub up_GridHeaderDokumenImpor()
'    With gridDokumenImpor
'        .clear
'        .ColS = colcount
'        .Rows = 1
'
'        .TextMatrix(0, colJenisDokumenImpor) = "Jenis"
'        .TextMatrix(0, colNomorDokumenImpor) = "Nomor"
'        .TextMatrix(0, colTanggalImpor) = "Tanggal"
'
'        .ColWidth(colJenisDokumenImpor) = 1500
'        .ColWidth(colNomorDokumenImpor) = 1500
'        .ColWidth(colTanggalImpor) = 1200
'        .ColAlignment(colNomorDokumenImpor) = flexAlignLeftCenter
'
'        .ColFormat(colTanggalImpor) = "dd MMM yyyy"
'    End With
End Sub

Private Sub up_GridHeaderDokumen()
'    With grid
'        .ColS = colcount
'        .Rows = 1
'
'        .TextMatrix(0, colJenisDokumen) = "Jenis"
'        .TextMatrix(0, colNomorDokumen) = "Nomor"
'        .TextMatrix(0, colTanggal) = "Tanggal"
'
'        .ColWidth(colJenisDokumen) = 1500
'        .ColWidth(colNomorDokumen) = 1500
'        .ColWidth(colTanggal) = 1200
'        .ColAlignment(colNomorDokumen) = flexAlignLeftCenter
'
'        .ColFormat(colTanggal) = "dd MMM yyyy"
'    End With
End Sub

Private Sub up_GenerateNomorSeriBahanBaku(pNoAju As String, pTipe As Integer, pNoSeriBarang As Integer)
Dim sql As String
Dim RS As New Recordset

sql = "Select NewSeriBahanBaku = ISNULL(Max(Seri_Bahan_Baku),0) + 1 From Bea_Cukai_TPB_Bahan_Baku Where NOMOR_AJU_DOK_ASAL = '" & pNoAju & "' AND KODE_ASAL_BAHAN_BAKU = " & pTipe & " AND SERI_BARANG = " & pNoSeriBarang & " "
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    If pTipe = 0 Then
        txtNoSeriImpor = RS.Fields("NewSeriBahanBaku")
    Else
        txtNoSeriLokal = RS.Fields("NewSeriBahanBaku")
    End If
    
End If
End Sub

Private Function uf_GetCountBahanBaku(pNoAju As String, pNoSeriBarang As Integer) As Integer
Dim sql As String
Dim RS As New Recordset

sql = "Select Count(*) As JData From Bea_Cukai_TPB_Bahan_Baku Where NOMOR_AJU_DOK_ASAL = '" & pNoAju & "' AND SERI_BARANG = " & pNoSeriBarang & " "
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    uf_GetCountBahanBaku = RS.Fields("JData")

Else
    uf_GetCountBahanBaku = 0
End If
End Function

Public Sub up_LoadDataBahanBakuImpor(pNoPengajuan As String, pNoSeriBarang As Integer, pNoSeriBahanBaku)
    LblerrMsg.Caption = ""
    
    Dim cmd As ADODB.Command
    Dim NomorHS As String
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27LoadDetailBahanBaku_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(pNoPengajuan, "-", ""))
    cmd.Parameters.append cmd.CreateParameter("NoSeriBarang", adInteger, adParamInput, 5, pNoSeriBarang)
    cmd.Parameters.append cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, pNoSeriBahanBaku)
    cmd.Parameters.append cmd.CreateParameter("KodeAsalBahanBaku", adInteger, adParamInput, 5, 0)
    
    Set rsbahanbakuimpor = cmd.Execute
    
    cekLoad = True

    If Not rsbahanbakuimpor.EOF Then
'        lblDokAsalImpor.Caption = IIf(IsNull(rsbahanbakuimpor.Fields("Uraian_Dokumen")), "", rsbahanbakuimpor.Fields("Uraian_Dokumen"))
'        txtDokumenAsalImpor = IIf(IsNull(rsbahanbakuimpor.Fields("Kode_Jenis_Dok_Asal")), "", rsbahanbakuimpor.Fields("Kode_Jenis_Dok_Asal"))
'        txtNoImpor = IIf(IsNull(rsbahanbakuimpor.Fields("Nomor_Daftar_Dok_Asal")), "", rsbahanbakuimpor.Fields("Nomor_Daftar_Dok_Asal"))
'        dtpTglImpor = rsbahanbakuimpor.Fields("Tanggal_Daftar_Dok_Asal")
'        txtKPPBCImpor = IIf(IsNull(rsbahanbakuimpor.Fields("kode_kantor")), "", rsbahanbakuimpor.Fields("kode_kantor"))
'        lblKPPBCImpor.Caption = IIf(IsNull(rsbahanbakuimpor.Fields("Nama_Kantor")), "", rsbahanbakuimpor.Fields("Nama_Kantor"))
'        txtNoAjuImpor = IIf(IsNull(rsbahanbakuimpor.Fields("NOMOR_AJU_DOK_ASAL")), "", rsbahanbakuimpor.Fields("NOMOR_AJU_DOK_ASAL"))
'        txtUrutKeImpor = IIf(IsNull(rsbahanbakuimpor.Fields("SERI_BARANG_DOK_ASAL")), "", rsbahanbakuimpor.Fields("SERI_BARANG_DOK_ASAL"))
'        txtNoSeriImpor = pNoSeriBahanBaku
'        txtKodeBarangImpor = IIf(IsNull(rsbahanbakuimpor.Fields("KODE_BARANG")), "", rsbahanbakuimpor.Fields("KODE_BARANG"))
'        txtNomorHSImpor = IIf(IsNull(rsbahanbakuimpor.Fields("POS_TARIF")), "", rsbahanbakuimpor.Fields("POS_TARIF"))
'        txtUraianBarangImpor = IIf(IsNull(rsbahanbakuimpor.Fields("Uraian_Barang")), "", rsbahanbakuimpor.Fields("Uraian_Barang"))
'        txtTipeImpor = IIf(IsNull(rsbahanbakuimpor.Fields("TIPE")), "", rsbahanbakuimpor.Fields("TIPE"))
'        txtUkuranImpor = IIf(IsNull(rsbahanbakuimpor.Fields("UKURAN")), "", rsbahanbakuimpor.Fields("UKURAN"))
'        txtSpfLainImpor = IIf(IsNull(rsbahanbakuimpor.Fields("SPESIFIKASI_LAIN")), "", rsbahanbakuimpor.Fields("SPESIFIKASI_LAIN"))
'        txtMerkImpor = IIf(IsNull(rsbahanbakuimpor.Fields("MERK")), "", rsbahanbakuimpor.Fields("MERK"))
'        txtHargaCIF = Format(IIf(IsNull(rsbahanbakuimpor.Fields("CIF")), 0, rsbahanbakuimpor.Fields("CIF")), "#,0.00")
'        txtJumlahSatuanImpor = IIf(IsNull(rsbahanbakuimpor.Fields("Jumlah_Satuan")), 0, rsbahanbakuimpor.Fields("Jumlah_Satuan"))
'        txtSatuanImpor = IIf(IsNull(rsbahanbakuimpor.Fields("JENIS_SATUAN")), "", rsbahanbakuimpor.Fields("JENIS_SATUAN"))
'        lblSatuanImpor.Caption = IIf(IsNull(rsbahanbakuimpor.Fields("Uraian_Satuan")), "", rsbahanbakuimpor.Fields("Uraian_Satuan"))
'
'        txtTotalImpor = IIf(IsNull(rsbahanbakuimpor.Fields("JData")), 0, rsbahanbakuimpor.Fields("JData"))

        lblDokAsalImpor.Caption = IIf(IsNull(rsbahanbakuimpor.Fields("Uraian_Dokumen")), "", rsbahanbakuimpor.Fields("Uraian_Dokumen"))
        'txtBMBrgJadi = txtTarifPersen1
        txtDokumenAsalImpor = IIf(IsNull(rsbahanbakuimpor.Fields("Kode_Jenis_Dok_Asal")), "", rsbahanbakuimpor.Fields("Kode_Jenis_Dok_Asal"))
        txtNoImpor = IIf(IsNull(rsbahanbakuimpor.Fields("Nomor_Daftar_Dok_Asal")), "", rsbahanbakuimpor.Fields("Nomor_Daftar_Dok_Asal"))
        dtpTglImpor = rsbahanbakuimpor.Fields("Tanggal_Daftar_Dok_Asal")
        txtKPPBCImpor = IIf(IsNull(rsbahanbakuimpor.Fields("kode_kantor")), "", rsbahanbakuimpor.Fields("kode_kantor"))
        lblKPPBCImpor.Caption = IIf(IsNull(rsbahanbakuimpor.Fields("Nama_Kantor")), "", rsbahanbakuimpor.Fields("Nama_Kantor"))
        txtNoAjuImpor = IIf(IsNull(rsbahanbakuimpor.Fields("NOMOR_AJU_DOK_ASAL")), "", rsbahanbakuimpor.Fields("NOMOR_AJU_DOK_ASAL"))
        txtUrutKeImpor = IIf(IsNull(rsbahanbakuimpor.Fields("SERI_BARANG_DOK_ASAL")), "", rsbahanbakuimpor.Fields("SERI_BARANG_DOK_ASAL"))
        txtNoSeriImpor = pNoSeriBahanBaku
        txtKodeBarangImpor = IIf(IsNull(rsbahanbakuimpor.Fields("KODE_BARANG")), "", rsbahanbakuimpor.Fields("KODE_BARANG"))
        txtNomorHSImpor = IIf(IsNull(rsbahanbakuimpor.Fields("POS_TARIF")), "", rsbahanbakuimpor.Fields("POS_TARIF"))
        txtUraianBarangImpor = IIf(IsNull(rsbahanbakuimpor.Fields("Uraian_Barang")), "", rsbahanbakuimpor.Fields("Uraian_Barang"))
        txtTipeImpor = IIf(IsNull(rsbahanbakuimpor.Fields("TIPE")), "", rsbahanbakuimpor.Fields("TIPE"))
        txtUkuranImpor = IIf(IsNull(rsbahanbakuimpor.Fields("UKURAN")), "", rsbahanbakuimpor.Fields("UKURAN"))
        txtSpfLainImpor = IIf(IsNull(rsbahanbakuimpor.Fields("SPESIFIKASI_LAIN")), "", rsbahanbakuimpor.Fields("SPESIFIKASI_LAIN"))
        txtMerkImpor = IIf(IsNull(rsbahanbakuimpor.Fields("MERK")), "", rsbahanbakuimpor.Fields("MERK"))
'        txtHargaCIFUSDImpor = Format(IIf(IsNull(rsbahanbakuimpor.Fields("CIF")), 0, rsbahanbakuimpor.Fields("CIF")), "#,0.00")
        'txtNDPBMImpor = Format(IIf(IsNull(rsbahanbakuimpor.Fields("NDPBM")), 0, rsbahanbakuimpor.Fields("NDPBM")), "#,0.00")
         txtHargaCIFImpor = Format(IIf(IsNull(rsbahanbakuimpor.Fields("CIF_Rupiah")), 0, rsbahanbakuimpor.Fields("CIF_Rupiah")), "#,0.00")
        'txtNettoImpor = Format(IIf(IsNull(rsbahanbakuimpor.Fields("Netto")), 0, rsbahanbakuimpor.Fields("Netto")), "#,0.00")
        txtJumlahSatuanImpor = IIf(IsNull(rsbahanbakuimpor.Fields("Jumlah_Satuan")), 0, rsbahanbakuimpor.Fields("Jumlah_Satuan"))
        txtSatuanImpor = IIf(IsNull(rsbahanbakuimpor.Fields("JENIS_SATUAN")), "", rsbahanbakuimpor.Fields("JENIS_SATUAN"))
        lblSatuanImpor.Caption = IIf(IsNull(rsbahanbakuimpor.Fields("Uraian_Satuan")), "", rsbahanbakuimpor.Fields("Uraian_Satuan"))
        
        'txtFasilitasImpor = IIf(IsNull(rsbahanbakuimpor.Fields("KODE_FASILITAS_DOKUMEN")), "", rsbahanbakuimpor.Fields("KODE_FASILITAS_DOKUMEN"))
        'Label1(13).Caption = IIf(IsNull(rsbahanbakuimpor.Fields("URAIAN_FASILITAS")), "", rsbahanbakuimpor.Fields("URAIAN_FASILITAS"))
        'txtSkemaTarifImpor = IIf(IsNull(rsbahanbakuimpor.Fields("KODE_SKEMA_TARIF")), "", rsbahanbakuimpor.Fields("KODE_SKEMA_TARIF"))
        'Label1(12).Caption = IIf(IsNull(rsbahanbakuimpor.Fields("Uraian_Skema")), "", rsbahanbakuimpor.Fields("Uraian_Skema"))
                
        txtTotalImpor = IIf(IsNull(rsbahanbakuimpor.Fields("JData")), 0, rsbahanbakuimpor.Fields("JData"))
        
'        cboJenisPungutanImpor = IIf(IsNull(rsbahanbakuimpor.fields("Uraian_Pungutan1")), "", rsbahanbakuimpor.fields("Uraian_Pungutan1"))
        cboKeterangan1Impor = IIf(IsNull(rsbahanbakuimpor.Fields("Uraian_Tarif1")), "", rsbahanbakuimpor.Fields("Uraian_Tarif1"))
        
        txtTarifPersen1Impor = IIf(IsNull(rsbahanbakuimpor.Fields("TARIF1")), 0, rsbahanbakuimpor.Fields("TARIF1"))
        cboKeterangan2Impor = IIf(IsNull(rsbahanbakuimpor.Fields("Uraian_Fasilitas1")), "", rsbahanbakuimpor.Fields("Uraian_Fasilitas1"))
        txtTarifPersen2Impor = IIf(IsNull(rsbahanbakuimpor.Fields("Tarif_Fasilitas1")), "", rsbahanbakuimpor.Fields("Tarif_Fasilitas1"))
        txtSatuanTarifImpor = IIf(IsNull(rsbahanbakuimpor.Fields("KODE_SATUAN_TARIF")), "", rsbahanbakuimpor.Fields("KODE_SATUAN_TARIF"))
        txtJumlahSatuanBMImpor = IIf(IsNull(rsbahanbakuimpor.Fields("JUMLAH_SATUAN_TARIF")), 0, rsbahanbakuimpor.Fields("JUMLAH_SATUAN_TARIF"))
        
        txtPPNImpor = IIf(IsNull(rsbahanbakuimpor.Fields("TARIF2")), 0, rsbahanbakuimpor.Fields("TARIF2"))
        cboKeterangan3Impor = IIf(IsNull(rsbahanbakuimpor.Fields("URAIAN_FASILITAS2")), "", rsbahanbakuimpor.Fields("URAIAN_FASILITAS2"))
        txtTarifPersen3Impor = IIf(IsNull(rsbahanbakuimpor.Fields("TARIF_FASILITAS2")), 0, rsbahanbakuimpor.Fields("TARIF_FASILITAS2"))
        
        txtPPNBMImpor = IIf(IsNull(rsbahanbakuimpor.Fields("TARIF3")), 0, rsbahanbakuimpor.Fields("TARIF3"))
        cboKeterangan4Impor = IIf(IsNull(rsbahanbakuimpor.Fields("URAIAN_FASILITAS3")), "", rsbahanbakuimpor.Fields("URAIAN_FASILITAS3"))
        txtTarifPersen4Impor = IIf(IsNull(rsbahanbakuimpor.Fields("TARIF_FASILITAS3")), 0, rsbahanbakuimpor.Fields("TARIF_FASILITAS3"))
        
        txtPPHImpor = IIf(IsNull(rsbahanbakuimpor.Fields("TARIF4")), 0, rsbahanbakuimpor.Fields("TARIF4"))
        cboKeterangan5Impor = IIf(IsNull(rsbahanbakuimpor.Fields("URAIAN_FASILITAS4")), "", rsbahanbakuimpor.Fields("URAIAN_FASILITAS4"))
        txtTarifPersen5Impor = IIf(IsNull(rsbahanbakuimpor.Fields("TARIF_FASILITAS4")), 0, rsbahanbakuimpor.Fields("TARIF_FASILITAS4"))
                
        cboCukaiImpor = IIf(IsNull(rsbahanbakuimpor.Fields("URAIAN_KOMODITI")), "", rsbahanbakuimpor.Fields("URAIAN_KOMODITI"))
        cboJenisTarifImpor = IIf(IsNull(rsbahanbakuimpor.Fields("URAIAN_TARIF_CUKAI")), "", rsbahanbakuimpor.Fields("URAIAN_TARIF_CUKAI"))
        txtTarifImpor = IIf(IsNull(rsbahanbakuimpor.Fields("TARIF_CUKAI")), 0, rsbahanbakuimpor.Fields("TARIF_CUKAI"))
        cboKeteranganJenisTarif = IIf(IsNull(rsbahanbakuimpor.Fields("URAIAN_FASILITAS_CUKAI")), "", rsbahanbakuimpor.Fields("URAIAN_FASILITAS_CUKAI"))
        txtPersenCukaiImpor = IIf(IsNull(rsbahanbakuimpor.Fields("TARIF_FASILITAS_CUKAI")), 0, rsbahanbakuimpor.Fields("TARIF_FASILITAS_CUKAI"))
        
        txtSatuanCukaiImpor = IIf(IsNull(rsbahanbakuimpor.Fields("KODE_SATUAN_CUKAI")), "", rsbahanbakuimpor.Fields("KODE_SATUAN_CUKAI"))
        txtJumlahTarifImpor = IIf(IsNull(rsbahanbakuimpor.Fields("JUMLAH_SATUAN_CUKAI")), 0, rsbahanbakuimpor.Fields("JUMLAH_SATUAN_CUKAI"))
    Else
        up_ClearImpor
    End If
End Sub

Public Sub up_LoadDataBahanBakuLokal(pNoPengajuan As String, pNoSeriBarang As Integer, pNoSeriBahanBaku)
    LblerrMsg.Caption = ""
    
    Dim cmd As ADODB.Command
    Dim NomorHS As String
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27LoadDetailBahanBaku_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(pNoPengajuan, "-", ""))
    cmd.Parameters.append cmd.CreateParameter("NoSeriBarang", adInteger, adParamInput, 5, pNoSeriBarang)
    cmd.Parameters.append cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, pNoSeriBahanBaku)
    cmd.Parameters.append cmd.CreateParameter("KodeAsalBahanBaku", adInteger, adParamInput, 5, 1)
    
    Set rsbahanbakulokal = cmd.Execute
    
    cekLoad = True
    
    If Not rsbahanbakulokal.EOF Then
        txtDokumenAsalLokal = IIf(IsNull(rsbahanbakulokal.Fields("Kode_Jenis_Dok_Asal")), "", rsbahanbakulokal.Fields("Kode_Jenis_Dok_Asal"))
        txtNoLokal = IIf(IsNull(rsbahanbakulokal.Fields("Nomor_Daftar_Dok_Asal")), "", rsbahanbakulokal.Fields("Nomor_Daftar_Dok_Asal"))
        dtpTglLokal = rsbahanbakulokal.Fields("Tanggal_Daftar_Dok_Asal")
        txtKPPBCLokal = IIf(IsNull(rsbahanbakulokal.Fields("kode_kantor")), "", rsbahanbakulokal.Fields("kode_kantor"))
        lblKPPBCLokal.Caption = IIf(IsNull(rsbahanbakulokal.Fields("Nama_Kantor")), "", rsbahanbakulokal.Fields("Nama_Kantor"))
        txtNoAjuLokal = IIf(IsNull(rsbahanbakulokal.Fields("NOMOR_AJU_DOK_ASAL")), "", rsbahanbakulokal.Fields("NOMOR_AJU_DOK_ASAL"))
        txtUrutKeLokal = IIf(IsNull(rsbahanbakulokal.Fields("SERI_BARANG_DOK_ASAL")), "", rsbahanbakulokal.Fields("SERI_BARANG_DOK_ASAL"))
        txtNoSeriLokal = pNoSeriBahanBaku
        txtKodeBarangLokal = IIf(IsNull(rsbahanbakulokal.Fields("SERI_BARANG_DOK_ASAL")), "", rsbahanbakulokal.Fields("SERI_BARANG_DOK_ASAL"))
        txtNomorHSLokal = IIf(IsNull(rsbahanbakulokal.Fields("POS_TARIF")), "", rsbahanbakulokal.Fields("POS_TARIF"))
        txtUraianBarangLokal = IIf(IsNull(rsbahanbakulokal.Fields("Uraian_Barang")), "", rsbahanbakulokal.Fields("Uraian_Barang"))
        txtTipeLokal = IIf(IsNull(rsbahanbakulokal.Fields("TIPE")), "", rsbahanbakulokal.Fields("TIPE"))
        txtUkuranLokal = IIf(IsNull(rsbahanbakulokal.Fields("UKURAN")), "", rsbahanbakulokal.Fields("UKURAN"))
        txtSpfLainLokal = IIf(IsNull(rsbahanbakulokal.Fields("SPESIFIKASI_LAIN")), "", rsbahanbakulokal.Fields("SPESIFIKASI_LAIN"))
        txtMerkLokal = IIf(IsNull(rsbahanbakulokal.Fields("MERK")), "", rsbahanbakulokal.Fields("MERK"))
        txtHargaPerolahen = Format(IIf(IsNull(rsbahanbakulokal.Fields("CIF")), 0, rsbahanbakulokal.Fields("CIF")), "#,0.00")
        txtHargaPenyerahanLokal = Format(IIf(IsNull(rsbahanbakulokal.Fields("Harga_Penyerahan")), 0, rsbahanbakulokal.Fields("Harga_Penyerahan")), "#,0.00")
        txtJumlahSatuanLokal = IIf(IsNull(rsbahanbakulokal.Fields("Jumlah_Satuan")), 0, rsbahanbakulokal.Fields("Jumlah_Satuan"))
        txtSatuanLokal = IIf(IsNull(rsbahanbakulokal.Fields("JENIS_SATUAN")), "", rsbahanbakulokal.Fields("JENIS_SATUAN"))
        lblSatuanLokal.Caption = IIf(IsNull(rsbahanbakulokal.Fields("Uraian_Satuan")), "", rsbahanbakulokal.Fields("Uraian_Satuan"))
        
        txtTotalLokal = IIf(IsNull(rsbahanbakulokal.Fields("JData")), 0, rsbahanbakulokal.Fields("JData"))
        
        txtPPNBayarLokal = Format(IIf(IsNull(rsbahanbakulokal.Fields("NILAI_BAYAR")), 0, rsbahanbakulokal.Fields("NILAI_BAYAR")), "#,0.00")
        txtPPNFasilitasLokal = Format(IIf(IsNull(rsbahanbakulokal.Fields("NILAI_FASILITAS")), 0, rsbahanbakulokal.Fields("NILAI_FASILITAS")), "#,0.00")
                        
    Else
        up_ClearLokal
    End If
End Sub

Public Sub up_LoadDataBarang(pNoPengajuan As String, pNoSeri As Integer)
    LblerrMsg.Caption = ""
    
    Dim cmd As ADODB.Command
    Dim NomorHS As String
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27LoadDetailBarang_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(pNoPengajuan, "-", ""))
    cmd.Parameters.append cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, pNoSeri)
    
    Set rsBarang = cmd.Execute
    
    cekLoad = True
    
    If Not rsBarang.EOF Then
        cekSubmit = True
        
        NomorHS = IIf(IsNull(rsBarang.Fields("POS_TARIF")), "", rsBarang.Fields("POS_TARIF"))
        txtNoSeri = IIf(IsNull(rsBarang.Fields("SERI_BARANG")), 0, rsBarang.Fields("SERI_BARANG"))
        txtTotalItem = IIf(IsNull(rsBarang.Fields("JData")), 0, rsBarang.Fields("JData"))
        
        txtNomorHS = Replace(NomorHS, ".", "")
        txtNomorHS = Mid(txtNomorHS.Text, 1, 10)
        If txtNomorHS <> "" Then
            txtNomorHS = Left(txtNomorHS.Text, 4) & "." & Mid(txtNomorHS.Text, 5, 2) & "." & Mid(txtNomorHS.Text, 7, 2) & "." & Mid(txtNomorHS.Text, 9, 2)
        End If
        
        txtKodeBarang = IIf(IsNull(rsBarang.Fields("Kode_Barang")), "", rsBarang.Fields("Kode_Barang"))
        txtUraianBarang = IIf(IsNull(rsBarang.Fields("Uraian_Barang")), "", rsBarang.Fields("Uraian_Barang"))
        txtMerk = IIf(IsNull(rsBarang.Fields("Merk")), "", rsBarang.Fields("Merk"))
        txtTipe = IIf(IsNull(rsBarang.Fields("Tipe")), "", rsBarang.Fields("Tipe"))
        txtUkuran = IIf(IsNull(rsBarang.Fields("UKURAN")), "", rsBarang.Fields("UKURAN"))
        txtSpfLain = IIf(IsNull(rsBarang.Fields("SPESIFIKASI_LAIN")), "", rsBarang.Fields("SPESIFIKASI_LAIN"))
        
        txtJumlahSatuan = Format(IIf(IsNull(rsBarang.Fields("JUMLAH_SATUAN")), 0, rsBarang.Fields("JUMLAH_SATUAN")), "#,0.00")
        txtJenisSatuan = IIf(IsNull(rsBarang.Fields("KODE_SATUAN")), "", rsBarang.Fields("KODE_SATUAN"))
        lblJenisSatuan(6).Caption = IIf(IsNull(rsBarang.Fields("URAIAN_SATUAN")), "", rsBarang.Fields("URAIAN_SATUAN"))
        txtNetto = IIf(IsNull(rsBarang.Fields("NETTO")), 0, rsBarang.Fields("NETTO"))
        
        txtVolume = Format(IIf(IsNull(rsBarang.Fields("Volume")), 0, rsBarang.Fields("Volume")), "#,0.00")
        txtNilaiCIF = Format(IIf(IsNull(rsBarang.Fields("CIF")), 0, rsBarang.Fields("CIF")), "#,0.00")
        txtHargaPenyerahan = Format(IIf(IsNull(rsBarang.Fields("HARGA_PENYERAHAN")), 0, rsBarang.Fields("HARGA_PENYERAHAN")), "#,0.00")
        
        txtJumlahBahanBaku = IIf(IsNull(rsBarang.Fields("JUMLAH_BAHAN_BAKU")), 0, rsBarang.Fields("JUMLAH_BAHAN_BAKU"))
    End If
    
    cekLoad = False
End Sub

Private Sub up_DeleteBarang()
Dim cmd As ADODB.Command
    
Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC27DetailBarang_Del"

cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append cmd.CreateParameter("NoSeri", adInteger, adParamInput, , txtNoSeri)

cmd.Execute

LblerrMsg.Caption = DisplayMsg(1201)

'DoEvents
'
'Unload Me
End Sub

Private Sub up_DeleteBahanBaku(pAsalBahanBaku As Integer, pNoSeriBahanBaku As Integer)
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
    
Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC27DetailBahanBaku_Del"

cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append cmd.CreateParameter("NoSeri", adInteger, adParamInput, , txtNoSeri)
cmd.Parameters.append cmd.CreateParameter("KodeAsalBahanBaku", adInteger, adParamInput, , pAsalBahanBaku)
cmd.Parameters.append cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, , pNoSeriBahanBaku)

cmd.Execute

LblerrMsg.Caption = DisplayMsg(1201)
End Sub

Private Sub up_SaveDataBahanBakuImpor()
Dim cmd As ADODB.Command
Dim i As Integer

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

Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC27DetailBahanBakuImpor_Upd"
    
Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append prm1
Set prm2 = cmd.CreateParameter("KodeBarang", adVarChar, adParamInput, 15, txtKodeBarangImpor)
cmd.Parameters.append prm2
Set prm3 = cmd.CreateParameter("NoSeriBarang", adInteger, adParamInput, 3, txtNoSeri)
cmd.Parameters.append prm3
Set prm4 = cmd.CreateParameter("NoSeriImpor", adInteger, adParamInput, 3, txtNoSeriImpor)
cmd.Parameters.append prm4
Set prm5 = cmd.CreateParameter("KodeAsalBahanBaku", adInteger, adParamInput, 3, 0)
cmd.Parameters.append prm5
Set prm6 = cmd.CreateParameter("Uraian", adVarChar, adParamInput, 255, txtUraianBarangImpor)
cmd.Parameters.append prm6
Set prm7 = cmd.CreateParameter("Tipe", adVarChar, adParamInput, 255, txtTipeImpor)
cmd.Parameters.append prm7
Set prm8 = cmd.CreateParameter("SpefLain", adVarChar, adParamInput, 255, txtSpfLainImpor)
cmd.Parameters.append prm8
Set prm9 = cmd.CreateParameter("Merk", adVarChar, adParamInput, 255, txtMerkImpor)
cmd.Parameters.append prm9
Set prm10 = cmd.CreateParameter("NoSeriAsal", adInteger, adParamInput, 3, txtUrutKeImpor)
cmd.Parameters.append prm10
Set prm11 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 255, txtSatuanImpor)
cmd.Parameters.append prm11
Set prm12 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , txtJumlahSatuanImpor)
prm12.Precision = 38
prm12.NumericScale = 4
cmd.Parameters.append prm12
Set prm13 = cmd.CreateParameter("NomorDokAsal", adVarChar, adParamInput, 255, txtNoImpor)
cmd.Parameters.append prm13
Set prm14 = cmd.CreateParameter("TanggalDokAsal", adDate, adParamInput, , Format(dtpTglImpor, "yyyy-MM-dd"))
cmd.Parameters.append prm14
Set prm15 = cmd.CreateParameter("PosTarif", adVarChar, adParamInput, 255, txtNomorHSImpor)
cmd.Parameters.append prm15
Set prm16 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , txtHargaCIFImpor)
prm16.Precision = 38
prm16.NumericScale = 2
cmd.Parameters.append prm16
Set prm17 = cmd.CreateParameter("KodeKantor", adVarChar, adParamInput, 255, txtKPPBCImpor)
cmd.Parameters.append prm17
Set prm18 = cmd.CreateParameter("JenisDok", adVarChar, adParamInput, 255, txtDokumenAsalImpor)
cmd.Parameters.append prm18
Set prm19 = cmd.CreateParameter("NomorAjuAsal", adVarChar, adParamInput, 255, txtNoAjuImpor)
cmd.Parameters.append prm19
Set prm20 = cmd.CreateParameter("HargaPenyerahan", adDecimal, adParamInput, , gd_HargaPenyerahan)
prm20.Precision = 38
prm20.NumericScale = 2
cmd.Parameters.append prm20
Set prm21 = cmd.CreateParameter("Ukuran", adVarChar, adParamInput, 255, txtUkuranImpor)
cmd.Parameters.append prm21

cmd.Execute i

If i = 0 Then
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuImpor_Ins"

    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("KodeBarang", adVarChar, adParamInput, 15, txtKodeBarangImpor)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBarang", adInteger, adParamInput, 3, txtNoSeri)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("NoSeriImpor", adInteger, adParamInput, 3, txtNoSeriImpor)
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeAsalBahanBaku", adInteger, adParamInput, 3, 0)
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("Uraian", adVarChar, adParamInput, 255, txtUraianBarangImpor)
    cmd.Parameters.append prm6
    Set prm7 = cmd.CreateParameter("Tipe", adVarChar, adParamInput, 255, txtTipeImpor)
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("SpefLain", adVarChar, adParamInput, 255, txtSpfLainImpor)
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("Merk", adVarChar, adParamInput, 255, txtMerkImpor)
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("NoSeriAsal", adInteger, adParamInput, 3, txtUrutKeImpor)
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 255, txtSatuanImpor)
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , txtJumlahSatuanImpor)
    prm12.Precision = 38
    prm12.NumericScale = 4
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("NomorDokAsal", adVarChar, adParamInput, 255, txtNoImpor)
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("TanggalDokAsal", adDate, adParamInput, , Format(dtpTglImpor, "yyyy-MM-dd"))
    cmd.Parameters.append prm14
    Set prm15 = cmd.CreateParameter("PosTarif", adVarChar, adParamInput, 255, txtNomorHSImpor)
    cmd.Parameters.append prm15
    Set prm16 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , txtHargaCIFImpor)
    prm16.Precision = 38
    prm16.NumericScale = 2
    cmd.Parameters.append prm16
    Set prm17 = cmd.CreateParameter("KodeKantor", adVarChar, adParamInput, 255, txtKPPBCImpor)
    cmd.Parameters.append prm17
    Set prm18 = cmd.CreateParameter("JenisDok", adVarChar, adParamInput, 255, txtDokumenAsalImpor)
    cmd.Parameters.append prm18
    Set prm19 = cmd.CreateParameter("NomorAjuAsal", adVarChar, adParamInput, 255, txtNoAjuImpor)
    cmd.Parameters.append prm19
    Set prm20 = cmd.CreateParameter("HargaPenyerahan", adDecimal, adParamInput, , gd_HargaPenyerahan)
    prm20.Precision = 38
    prm20.NumericScale = 2
    cmd.Parameters.append prm20
    Set prm21 = cmd.CreateParameter("Ukuran", adVarChar, adParamInput, 255, txtUkuranImpor)
    cmd.Parameters.append prm21

    cmd.Execute
End If



    '####################### BM ########################
    
    Dim TarifBM As Double
    Dim TarifPPN As Double
    Dim TarifPPBBM As Double
    Dim TarifPPH As Double
    Dim TarifKomoditi As Double
        
    Dim NilaiBayar As Double
    Dim NilaiFasilitas As Double
    
    'DELETE BM
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuBeaMasukTambahan_Del"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeriBrg", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriImpor)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, "0")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BM")
    cmd.Parameters.append prm5
        
    cmd.Execute
    
    'INSERT BM
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuBeaMasukTambahan_Ins"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeriBrg", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriImpor)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 20, "BM")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Trim(Split(cboKeterangan2Impor, "-")(0)))
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Trim(Split(cboKeterangan1Impor, "-")(0)))
    cmd.Parameters.append prm6
    
    Set prm7 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
    prm7.Precision = 38
    prm7.NumericScale = 2
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
    prm8.Precision = 38
    prm8.NumericScale = 2
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtTarifPersen1Impor))
    prm9.Precision = 38
    prm9.NumericScale = 2
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtTarifPersen2Impor))
    prm10.Precision = 38
    prm10.NumericScale = 2
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, txtSatuanTarifImpor)
    cmd.Parameters.append prm11
    If txtJumlahSatuanBMImpor = "" Then txtJumlahSatuanBMImpor = 0
    Set prm12 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , CDbl(txtJumlahSatuanBMImpor))
    prm12.Precision = 38
    prm12.NumericScale = 4
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("KodeKomoditi", adVarChar, adParamInput, 50, Null)
    cmd.Parameters.append prm14
    
    cmd.Execute
    
    '####################### BM ########################
    
    
    
    '####################### PPN ########################
    'DELETE PPN
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuBeaMasukTambahan_Del"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeriBrg", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriImpor)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, "0")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "PPN")
    cmd.Parameters.append prm5
        
    cmd.Execute
    
    'INSERT PPN
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuBeaMasukTambahan_Ins"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeriBrg", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriImpor)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 20, "PPN")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboKeterangan3Impor, 1))
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm6
    
    TarifPPN = TarifBM * (CDbl(txtPPNImpor) / 100)
    
    If Left(cboKeterangan3Impor, 1) = 0 Then
        NilaiBayar = TarifPPN
        NilaiFasilitas = 0
    Else
        NilaiBayar = 0
        NilaiFasilitas = TarifPPN
    End If
    
    Set prm7 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
    prm7.Precision = 38
    prm7.NumericScale = 2
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
    prm8.Precision = 38
    prm8.NumericScale = 2
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtPPNImpor))
    prm9.Precision = 38
    prm9.NumericScale = 2
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtTarifPersen3Impor))
    prm10.Precision = 38
    prm10.NumericScale = 2
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm11
    If txtJumlahSatuanBMImpor = "" Then txtJumlahSatuanBMImpor = 0
    Set prm12 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
    prm12.Precision = 38
    prm12.NumericScale = 4
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("KodeKomoditi", adVarChar, adParamInput, 50, Null)
    cmd.Parameters.append prm14
    
    cmd.Execute
        
    '####################### PPN ########################
    
    
    '####################### PPN BM ########################
    'DELETE PPN BM
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuBeaMasukTambahan_Del"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeriBrg", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriImpor)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, "0")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "PPNBM")
    cmd.Parameters.append prm5
        
    cmd.Execute
    
    'INSERT PPN BM
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuBeaMasukTambahan_Ins"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeriBrg", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriImpor)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 20, "PPNBM")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboKeterangan4Impor, 1))
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm6
    
    If txtPPNBMImpor = "" Then txtPPNBMImpor = 0
    TarifPPBBM = TarifBM * (CDbl(txtPPNBMImpor) / 100)
    
    If Left(cboKeterangan4Impor, 1) = 0 Then
        NilaiBayar = TarifPPBBM
        NilaiFasilitas = 0
    Else
        NilaiBayar = 0
        NilaiFasilitas = TarifPPBBM
    End If
    
    Set prm7 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
    prm7.Precision = 38
    prm7.NumericScale = 2
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
    prm8.Precision = 38
    prm8.NumericScale = 2
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtPPNBMImpor))
    prm9.Precision = 38
    prm9.NumericScale = 2
    cmd.Parameters.append prm9
    If txtTarifPersen4Impor = "" Then txtTarifPersen4Impor = 0
    Set prm10 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtTarifPersen4Impor))
    prm10.Precision = 38
    prm10.NumericScale = 2
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm11
    If txtJumlahSatuanBMImpor = "" Then txtJumlahSatuanBMImpor = 0
    Set prm12 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
    prm12.Precision = 38
    prm12.NumericScale = 4
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("KodeKomoditi", adVarChar, adParamInput, 50, Null)
    cmd.Parameters.append prm14
    
    cmd.Execute
        
    '####################### PPN BM ########################
    
    '####################### PPH ########################
    'DELETE PPH
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuBeaMasukTambahan_Del"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeriBrg", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriImpor)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, "0")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "PPH")
    cmd.Parameters.append prm5
        
    cmd.Execute
    
    'INSERT PPH
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuBeaMasukTambahan_Ins"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeriBrg", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriImpor)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 20, "PPH")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboKeterangan5Impor, 1))
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm6
    
    TarifPPH = TarifBM * (CDbl(txtPPHImpor) / 100)
    
    If Left(cboKeterangan5Impor, 1) = 0 Then
        NilaiBayar = TarifPPH
        NilaiFasilitas = 0
    Else
        NilaiBayar = 0
        NilaiFasilitas = TarifPPH
    End If
    
    Set prm7 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
    prm7.Precision = 38
    prm7.NumericScale = 2
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
    prm8.Precision = 38
    prm8.NumericScale = 2
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtPPHImpor))
    prm9.Precision = 38
    prm9.NumericScale = 2
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtTarifPersen5Impor))
    prm10.Precision = 38
    prm10.NumericScale = 2
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
    prm12.Precision = 38
    prm12.NumericScale = 4
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("KodeKomoditi", adVarChar, adParamInput, 50, Null)
    cmd.Parameters.append prm14
    
    cmd.Execute
        
    '####################### PPH ########################
    
    '####################### KOMODITI ########################
    'DELETE KOMODITI
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuBeaMasukTambahan_Del"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeriBrg", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriImpor)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, "0")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "CUKAI")
    cmd.Parameters.append prm5
        
    cmd.Execute
    
    'INSERT KOMODITI
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuBeaMasukTambahan_Ins"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeriBrg", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriImpor)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 20, "CUKAI")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboKeteranganJenisTarif, 1))
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Left(cboJenisTarifImpor, 1))
    cmd.Parameters.append prm6
    
    If txtTarifImpor = "" Then txtTarifImpor = 0
    TarifKomoditi = TarifBM * (CDbl(txtTarifImpor) / 100)
    
    If Left(cboKeteranganJenisTarif, 1) = 0 Then
        NilaiBayar = TarifPPH
        NilaiFasilitas = 0
    Else
        NilaiBayar = 0
        NilaiFasilitas = TarifPPH
    End If
    
    Set prm7 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
    prm7.Precision = 38
    prm7.NumericScale = 2
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
    prm8.Precision = 38
    prm8.NumericScale = 2
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtTarifImpor))
    prm9.Precision = 38
    prm9.NumericScale = 2
    cmd.Parameters.append prm9
    If txtPersenCukaiImpor = "" Then txtPersenCukaiImpor = 0
    Set prm10 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtPersenCukaiImpor))
    prm10.Precision = 38
    prm10.NumericScale = 2
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, txtSatuanCukaiImpor)
    cmd.Parameters.append prm11
    If txtJumlahTarifImpor = "" Then txtJumlahTarifImpor = 0
    Set prm12 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , CDbl(txtJumlahTarifImpor))
    prm12.Precision = 38
    prm12.NumericScale = 4
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("KodeKomoditi", adVarChar, adParamInput, 50, Left(cboCukaiImpor, 1))
    cmd.Parameters.append prm14
    
    cmd.Execute
        
    '####################### KOMODITI ########################
    
    up_LoadDataBahanBakuImpor txtNoPengajuan, txtNoSeri, txtNoSeriImpor
    'up_GridLoadDokumenImpor
    
    up_LoadDataBarang txtNoPengajuan, txtNoSeri
    'up_GridLoadDokumen
    
    If i = 0 Then
    '    txtKodeBarang.Enabled = False
        LblerrMsg = DisplayMsg(1000)
    Else
        LblerrMsg = DisplayMsg(1101)
    End If
    
End Sub

Private Sub up_SaveDataBahanBakuLokal()
Dim cmd As ADODB.Command
Dim i As Integer

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

Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC27DetailBahanBakuLokal_Upd"
    
Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append prm1
Set prm2 = cmd.CreateParameter("KodeBarang", adVarChar, adParamInput, 15, txtKodeBarangLokal)
cmd.Parameters.append prm2
Set prm3 = cmd.CreateParameter("NoSeriBarang", adInteger, adParamInput, 3, txtNoSeri)
cmd.Parameters.append prm3
Set prm4 = cmd.CreateParameter("NoSeriLokal", adInteger, adParamInput, 3, txtNoSeriLokal)
cmd.Parameters.append prm4
Set prm5 = cmd.CreateParameter("KodeAsalBahanBaku", adInteger, adParamInput, 3, 1)
cmd.Parameters.append prm5
Set prm6 = cmd.CreateParameter("Uraian", adVarChar, adParamInput, 255, txtUraianBarangLokal)
cmd.Parameters.append prm6
Set prm7 = cmd.CreateParameter("Tipe", adVarChar, adParamInput, 255, txtTipeLokal)
cmd.Parameters.append prm7
Set prm8 = cmd.CreateParameter("SpefLain", adVarChar, adParamInput, 255, txtSpfLainLokal)
cmd.Parameters.append prm8
Set prm9 = cmd.CreateParameter("Merk", adVarChar, adParamInput, 255, txtMerkLokal)
cmd.Parameters.append prm9
Set prm10 = cmd.CreateParameter("NoSeriAsal", adInteger, adParamInput, 3, txtUrutKeLokal)
cmd.Parameters.append prm10
Set prm11 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 255, txtSatuanLokal)
cmd.Parameters.append prm11
Set prm12 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , txtJumlahSatuanLokal)
prm12.Precision = 38
prm12.NumericScale = 4
cmd.Parameters.append prm12
Set prm13 = cmd.CreateParameter("NomorDokAsal", adVarChar, adParamInput, 255, txtNoLokal)
cmd.Parameters.append prm13
Set prm14 = cmd.CreateParameter("TanggalDokAsal", adDate, adParamInput, , Format(dtpTglLokal, "yyyy-MM-dd"))
cmd.Parameters.append prm14
Set prm15 = cmd.CreateParameter("PosTarif", adVarChar, adParamInput, 255, txtNomorHSLokal)
cmd.Parameters.append prm15
Set prm16 = cmd.CreateParameter("HargaPenyerahan", adDecimal, adParamInput, , txtHargaPenyerahanLokal)
prm16.Precision = 38
prm16.NumericScale = 4
cmd.Parameters.append prm16
Set prm17 = cmd.CreateParameter("KodeKantor", adVarChar, adParamInput, 255, txtKPPBCLokal)
cmd.Parameters.append prm17
Set prm18 = cmd.CreateParameter("JenisDok", adVarChar, adParamInput, 255, txtDokumenAsalLokal)
cmd.Parameters.append prm18
Set prm19 = cmd.CreateParameter("NomorAjuAsal", adVarChar, adParamInput, 255, txtNoAjuLokal)
cmd.Parameters.append prm19
Set prm19 = cmd.CreateParameter("HargaPenyerahan", adDecimal, adParamInput, , gd_HargaPenyerahan)
prm19.Precision = 38
prm19.NumericScale = 2
cmd.Parameters.append prm19
'Set prm20 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , txtHargaCIF)
'prm20.Precision = 38
'prm20.NumericScale = 2
'cmd.Parameters.append prm20
Set prm21 = cmd.CreateParameter("Ukuran", adVarChar, adParamInput, 255, txtUkuranLokal)
cmd.Parameters.append prm21

cmd.Execute i

If i = 0 Then
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuLokal_Ins"

    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("KodeBarang", adVarChar, adParamInput, 15, txtKodeBarangLokal)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBarang", adInteger, adParamInput, 3, txtNoSeri)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("NoSeriLokal", adInteger, adParamInput, 3, txtNoSeriLokal)
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeAsalBahanBaku", adInteger, adParamInput, 3, 1)
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("Uraian", adVarChar, adParamInput, 255, txtUraianBarangLokal)
    cmd.Parameters.append prm6
    Set prm7 = cmd.CreateParameter("Tipe", adVarChar, adParamInput, 255, txtTipeLokal)
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("SpefLain", adVarChar, adParamInput, 255, txtSpfLainLokal)
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("Merk", adVarChar, adParamInput, 255, txtMerkLokal)
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("NoSeriAsal", adInteger, adParamInput, 3, txtUrutKeLokal)
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 255, txtSatuanLokal)
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , txtJumlahSatuanLokal)
    prm12.Precision = 38
    prm12.NumericScale = 4
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("NomorDokAsal", adVarChar, adParamInput, 255, txtNoLokal)
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("TanggalDokAsal", adDate, adParamInput, , Format(dtpTglLokal, "yyyy-MM-dd"))
    cmd.Parameters.append prm14
    Set prm15 = cmd.CreateParameter("PosTarif", adVarChar, adParamInput, 255, txtNomorHSLokal)
    cmd.Parameters.append prm15
    Set prm16 = cmd.CreateParameter("HargaPenyerahan", adDecimal, adParamInput, , txtHargaPenyerahanLokal)
    prm16.Precision = 38
    prm16.NumericScale = 4
    cmd.Parameters.append prm16
    Set prm17 = cmd.CreateParameter("KodeKantor", adVarChar, adParamInput, 255, txtKPPBCLokal)
    cmd.Parameters.append prm17
    Set prm18 = cmd.CreateParameter("JenisDok", adVarChar, adParamInput, 255, txtDokumenAsalLokal)
    cmd.Parameters.append prm18
    Set prm19 = cmd.CreateParameter("NomorAjuAsal", adVarChar, adParamInput, 255, txtNoAjuLokal)
    cmd.Parameters.append prm19
    Set prm19 = cmd.CreateParameter("HargaPenyerahan", adDecimal, adParamInput, , gd_HargaPenyerahan)
    prm19.Precision = 38
    prm19.NumericScale = 2
    cmd.Parameters.append prm19
    Set prm20 = cmd.CreateParameter("Ukuran", adVarChar, adParamInput, 255, txtUkuranLokal)
    cmd.Parameters.append prm20


    cmd.Execute
End If

    '#######################  ########################
    Dim NilaiBayar As Double
    Dim NilaiFasilitas As Double
    
    NilaiBayar = 0
    NilaiFasilitas = 0
    
    'DELETE
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuBeaMasukTambahan_Del"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeriBrg", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriLokal)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 1)
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "PPN")
    cmd.Parameters.append prm5
        
    cmd.Execute
    
    'INSERT
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBahanBakuBeaMasukTambahan_Ins"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeriBrg", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriLokal)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 20, "PPN")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboJenisPPNLokal, 2))
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm6
    
    If Left(cboJenisPPNLokal, 1) = "0" Then
        NilaiBayar = CDbl(txtHargaPenyerahanLokal) * (CDbl(txtPPNLokal) / 100)
        NilaiFasilitas = 0
    ElseIf Left(cboJenisPPNLokal, 1) = "4" Then
        NilaiBayar = 0
        NilaiFasilitas = CDbl(txtHargaPenyerahanLokal) * (CDbl(txtPPNLokal) / 100)
    End If
    
    Set prm7 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
    prm7.Precision = 38
    prm7.NumericScale = 2
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
    prm8.Precision = 38
    prm8.NumericScale = 2
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtPPNLokal))
    prm9.Precision = 38
    prm9.NumericScale = 2
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtPercentLokal))
    prm10.Precision = 38
    prm10.NumericScale = 2
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm11
    If txtJumlahSatuanBMImpor = "" Then txtJumlahSatuanBMImpor = 0
    Set prm12 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
    prm12.Precision = 38
    prm12.NumericScale = 4
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 1)
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("KodeKomoditi", adVarChar, adParamInput, 50, Null)
    cmd.Parameters.append prm14
    
    cmd.Execute
    
    '#######################  ########################
    
    up_LoadDataBahanBakuLokal txtNoPengajuan, txtNoSeri, txtNoSeriLokal
    'up_GridLoadDokumenLokal
    
    If i = 0 Then
    '    txtKodeBarang.Enabled = False
        LblerrMsg = DisplayMsg(1000)
    Else
        LblerrMsg = DisplayMsg(1101)
    End If

End Sub

Private Sub up_SaveDataBarang()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim Y As Integer
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
    
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC27DetailBarang_Upd"
        
    'If txtID = "" Then txtID = "0"
    
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("KodeBarang", adVarChar, adParamInput, 15, txtKodeBarang)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 3, txtNoSeri)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("NamaBarang", adVarChar, adParamInput, 255, txtUraianBarang)
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("PosTarif", adVarChar, adParamInput, 50, Replace(txtNomorHS, ".", ""))
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("Tipe", adVarChar, adParamInput, 255, txtTipe)
    cmd.Parameters.append prm6
    Set prm7 = cmd.CreateParameter("Ukuran", adVarChar, adParamInput, 255, txtUkuran)
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("SpesifikasiLain", adVarChar, adParamInput, 255, txtSpfLain)
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("Merk", adVarChar, adParamInput, 255, txtMerk)
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , txtJumlahSatuan)
    prm10.Precision = 38
    prm10.NumericScale = 4
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 10, txtJenisSatuan)
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , txtNilaiCIF)
    prm12.Precision = 38
    prm12.NumericScale = 2
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("Netto", adDecimal, adParamInput, , txtNetto)
    prm13.Precision = 38
    prm13.NumericScale = 2
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("HargaPenyerahan", adDecimal, adParamInput, , txtHargaPenyerahan)
    prm14.Precision = 38
    prm14.NumericScale = 2
    cmd.Parameters.append prm14
    Set prm15 = cmd.CreateParameter("Volume", adDecimal, adParamInput, , txtVolume)
    prm15.Precision = 38
    prm15.NumericScale = 2
    cmd.Parameters.append prm15
    Set prm16 = cmd.CreateParameter("JumlahBahanBaku", adDecimal, adParamInput, , txtJumlahBahanBaku)
    prm16.Precision = 38
    prm16.NumericScale = 2
    cmd.Parameters.append prm16
    
    cmd.Execute Y
    
    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC27DetailBarang_Ins"
        
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("KodeBarang", adVarChar, adParamInput, 15, txtKodeBarang)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 3, txtNoSeri)
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("NamaBarang", adVarChar, adParamInput, 255, txtUraianBarang)
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("PosTarif", adVarChar, adParamInput, 50, Replace(txtNomorHS, ".", ""))
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("Tipe", adVarChar, adParamInput, 255, txtTipe)
    cmd.Parameters.append prm6
    Set prm7 = cmd.CreateParameter("Ukuran", adVarChar, adParamInput, 255, txtUkuran)
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("SpesifikasiLain", adVarChar, adParamInput, 255, txtSpfLain)
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("Merk", adVarChar, adParamInput, 255, txtMerk)
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , txtJumlahSatuan)
    prm10.Precision = 38
    prm10.NumericScale = 4
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 10, txtJenisSatuan)
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , txtNilaiCIF)
    prm12.Precision = 38
    prm12.NumericScale = 2
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("Netto", adDecimal, adParamInput, , txtNetto)
    prm13.Precision = 38
    prm13.NumericScale = 2
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("HargaPenyerahan", adDecimal, adParamInput, , txtHargaPenyerahan)
    prm14.Precision = 38
    prm14.NumericScale = 2
    cmd.Parameters.append prm14
    Set prm15 = cmd.CreateParameter("Volume", adDecimal, adParamInput, , txtVolume)
    prm15.Precision = 38
    prm15.NumericScale = 2
    cmd.Parameters.append prm15
    Set prm16 = cmd.CreateParameter("JumlahBahanBaku", adDecimal, adParamInput, , txtJumlahBahanBaku)
    prm16.Precision = 38
    prm16.NumericScale = 2
    cmd.Parameters.append prm16
    
        cmd.Execute
    End If
    
    '####################### TOTAL PUNGUTAN ########################
    
    up_LoadDataBarang txtNoPengajuan, txtNoSeri
        
    If Y = 0 Then
        txtKodeBarang.Enabled = False
        LblerrMsg = DisplayMsg(1000)
    Else
        LblerrMsg = DisplayMsg(1101)
    End If

    cekSubmit = True
End Sub

Private Function uf_ValidateInputBarang() As Boolean
    SSTab1.Tab = 0
   If txtKodeBarang = "" Then
        txtKodeBarang.SetFocus
        LblerrMsg = "Please Input Kode Barang!"
        uf_ValidateInputBarang = False
        Exit Function
    ElseIf txtUraianBarang = "" Then
        txtUraianBarang.SetFocus
        LblerrMsg = "Please Input Uraian Barang!"
        uf_ValidateInputBarang = False
        Exit Function
    ElseIf txtJumlahSatuan.Text = "" Then
        txtJumlahSatuan.SetFocus
        LblerrMsg = "Please Input Jumlah Satuan!"
        uf_ValidateInputBarang = False
        Exit Function
    ElseIf txtNetto.Text = "" Then
        txtNetto.SetFocus
        LblerrMsg = "Please Input Netto!"
        uf_ValidateInputBarang = False
        Exit Function
    ElseIf txtVolume.Text = "" Then
        txtVolume.SetFocus
        LblerrMsg = "Please Input Volume!"
        uf_ValidateInputBarang = False
        Exit Function
    ElseIf txtHargaPenyerahan.Text = "" Then
        txtHargaPenyerahan.SetFocus
        LblerrMsg = "Please Input Harga Penyerahan!"
        uf_ValidateInputBarang = False
        Exit Function
    ElseIf txtJumlahBahanBaku.Text = "" Then
        txtJumlahBahanBaku.SetFocus
        LblerrMsg = "Please Input Jumlah Bahan Baku!"
        uf_ValidateInputBarang = False
        Exit Function
    End If
    
    uf_ValidateInputBarang = True
End Function

Private Function uf_ValidateInputBahanBakuImpor() As Boolean
    SSTab1.Tab = 1
    If txtNoSeriImpor = "" Then
        cmdNewImpor.SetFocus
        LblerrMsg = "Please press New Button first!"
        uf_ValidateInputBahanBakuImpor = False
        Exit Function
    ElseIf txtDokumenAsalImpor = "" Then
        txtDokumenAsalImpor.SetFocus
        LblerrMsg = "Please input Dokumen Asal!"
        uf_ValidateInputBahanBakuImpor = False
        Exit Function
    ElseIf txtDokumenAsalImpor = "" Then
        txtDokumenAsalImpor.SetFocus
        LblerrMsg = "Please input Dokumen Asal!"
        uf_ValidateInputBahanBakuImpor = False
        Exit Function
    ElseIf txtNoImpor = "" Then
        txtNoImpor.SetFocus
        LblerrMsg = "Please input Nomor Dokumen!"
        uf_ValidateInputBahanBakuImpor = False
        Exit Function
    ElseIf txtKPPBCImpor = "" Then
        txtKPPBCImpor.SetFocus
        LblerrMsg = "Please input KPPBC Dokumen!"
        uf_ValidateInputBahanBakuImpor = False
        Exit Function
    ElseIf txtNoAjuImpor = "" Then
        txtNoAjuImpor.SetFocus
        LblerrMsg = "Please input Nomor Aju!"
        uf_ValidateInputBahanBakuImpor = False
        Exit Function
    ElseIf txtUrutKeImpor = "" Then
        txtUrutKeImpor.SetFocus
        LblerrMsg = "Please input Nomor Seri Dokumen Asal!"
        uf_ValidateInputBahanBakuImpor = False
        Exit Function
    ElseIf txtKodeBarangImpor = "" Then
        txtKodeBarangImpor.SetFocus
        LblerrMsg = "Please input Kode Barang Impor!"
        uf_ValidateInputBahanBakuImpor = False
        Exit Function
    ElseIf txtNomorHSImpor = "" Then
        txtNomorHSImpor.SetFocus
        LblerrMsg = "Please input Nomor HS!"
        uf_ValidateInputBahanBakuImpor = False
        Exit Function
    ElseIf txtUraianBarangImpor = "" Then
        txtUraianBarangImpor.SetFocus
        LblerrMsg = "Please input Uraian Barang!"
        uf_ValidateInputBahanBakuImpor = False
        Exit Function
    ElseIf txtJumlahSatuanImpor = "" Then
        txtJumlahSatuanImpor.SetFocus
        LblerrMsg = "Please input Jumlah Satuan!"
        uf_ValidateInputBahanBakuImpor = False
        Exit Function
    ElseIf txtSatuanImpor = "" Then
        txtSatuanImpor.SetFocus
        LblerrMsg = "Please input Satuan!"
        uf_ValidateInputBahanBakuImpor = False
        Exit Function
    End If
    
    uf_ValidateInputBahanBakuImpor = True
End Function

Private Function uf_ValidateInputBahanBakuLokal() As Boolean
    SSTab1.Tab = 2
    If txtNoSeriLokal = "" Then
        cmdNewLokal.SetFocus
        LblerrMsg = "Please press New Button first!"
        uf_ValidateInputBahanBakuLokal = False
        Exit Function
    ElseIf txtDokumenAsalLokal = "" Then
        txtDokumenAsalLokal.SetFocus
        LblerrMsg = "Please input Dokumen Asal!"
        uf_ValidateInputBahanBakuLokal = False
        Exit Function
    ElseIf txtDokumenAsalLokal = "" Then
        txtDokumenAsalLokal.SetFocus
        LblerrMsg = "Please input Dokumen Asal!"
        uf_ValidateInputBahanBakuLokal = False
        Exit Function
    ElseIf txtNoLokal = "" Then
        txtNoLokal.SetFocus
        LblerrMsg = "Please input Nomor Dokumen!"
        uf_ValidateInputBahanBakuLokal = False
        Exit Function
    ElseIf txtKPPBCLokal = "" Then
        txtKPPBCLokal.SetFocus
        LblerrMsg = "Please input KPPBC Dokumen!"
        uf_ValidateInputBahanBakuLokal = False
        Exit Function
    ElseIf txtNoAjuLokal = "" Then
        txtNoAjuLokal.SetFocus
        LblerrMsg = "Please input Nomor Aju!"
        uf_ValidateInputBahanBakuLokal = False
        Exit Function
    ElseIf txtUrutKeLokal = "" Then
        txtUrutKeLokal.SetFocus
        LblerrMsg = "Please input Nomor Seri Dokumen Asal!"
        uf_ValidateInputBahanBakuLokal = False
        Exit Function
    ElseIf txtJumlahSatuanLokal = "" Then
        txtJumlahSatuanLokal.SetFocus
        LblerrMsg = "Please input Jumlah Satuan!"
        uf_ValidateInputBahanBakuLokal = False
        Exit Function
    ElseIf txtSatuanLokal = "" Then
        txtSatuanLokal.SetFocus
        LblerrMsg = "Please input Satuan!"
        uf_ValidateInputBahanBakuLokal = False
        Exit Function
    End If
    
    uf_ValidateInputBahanBakuLokal = True
End Function

Private Sub cboJenisPPNLokal_Change()
If Left(cboJenisPPNLokal, 1) = "0" Then
   txtPPNBayarLokal = Format(CDbl(txtHargaPenyerahanLokal) * (CDbl(txtPPNLokal) / 100), "#,0.00")
   txtPPNFasilitasLokal = 0
ElseIf Left(cboJenisPPNLokal, 1) = "4" Then
   txtPPNFasilitasLokal = Format(CDbl(txtHargaPenyerahanLokal) * (CDbl(txtPPNLokal) / 100), "#,0.00")
   txtPPNBayarLokal = 0
   
Else
   txtPPNBayarLokal = "0.00"
   txtPPNFasilitasLokal = "0.00"
End If
End Sub

Private Sub cboJenisTarif_Change()
'If cboJenisTarif <> "" Then
'    If Trim(Split(cboJenisTarif, "-")(0)) = "2" Then
'        Label1(37).Visible = True
'        txtJumlahCukai.Visible = True
'        cboKeterangan.Left = 2640
'        txtPersenCukai.Left = 5880
'        Label1(42).Caption = "/"
'        txtSatuanCukai.Visible = True
'        Label1(43).Left = 5640
'        txtPersenCukai.Width = 705
'    Else
'        Label1(37).Visible = False
'        txtJumlahCukai.Visible = False
'        cboKeterangan.Left = 1560
'        txtPersenCukai.Left = 4590
'        Label1(42).Caption = "%"
'        txtSatuanCukai.Visible = False
'        Label1(43).Left = 5640
'        txtPersenCukai.Width = 945
'    End If
'
'End If

End Sub

Private Sub cboJenisTarifImpor_Change()
If cboJenisTarifImpor <> "" Then
    If Trim(Split(cboJenisTarifImpor, "-")(0)) = "2" Then
        Label1(14).Visible = True
        txtJumlahTarifImpor.Visible = True
        cboKeteranganJenisTarif.Left = 2640
        txtPersenCukaiImpor.Left = 5880
        Label1(17).Caption = "/"
        txtSatuanCukaiImpor.Visible = True
        Label1(18).Left = 5640
        txtPersenCukaiImpor.Width = 705
    Else
        Label1(14).Visible = False
        txtJumlahTarifImpor.Visible = False
        cboKeteranganJenisTarif.Left = 1560
        txtPersenCukaiImpor.Left = 4590
        Label1(17).Caption = "%"
        txtSatuanCukaiImpor.Visible = False
        Label1(18).Left = 5640
        txtPersenCukaiImpor.Width = 945
    End If

End If
End Sub

Private Sub cboKeterangan1_Change()
'        If cboKeterangan1 <> "" Then
'            If Trim(Split(cboKeterangan1, "-")(0)) = "2" Then
'                txtJumlahSpesifik.Visible = True
'                txtSatuanTarif.Visible = True
'                Label1(39).Visible = True
'                Label1(24).Caption = "/"
'            Else
'                txtJumlahSpesifik.Visible = False
'                txtSatuanTarif.Visible = False
'                Label1(39).Visible = False
'                Label1(24).Caption = "%"
'            End If
'        End If

        
        If cekLoad = False Then
            'txtTarifPersen2.Text = "100.00"
'            txtTarifPersen5.Text = "100.00"
'            txtPPh.Text = "7.5"
'            cboKeterangan2.ListIndex = 0
'            cboKeterangan5.ListIndex = 0
        End If
End Sub

Private Sub cboKeterangan1Impor_Change()
        If cboKeterangan1Impor <> "" Then
            If Trim(Split(cboKeterangan1Impor, "-")(0)) = "2" Then
                txtJumlahSatuanBMImpor.Visible = True
                txtSatuanTarifImpor.Visible = True
                Label1(47).Visible = True
                Label1(19).Caption = "/"
            Else
                txtJumlahSatuanBMImpor.Visible = False
                txtSatuanTarifImpor.Visible = False
                Label1(47).Visible = False
                Label1(19).Caption = "%"
            End If
        End If
End Sub

Private Sub cmdBrowseDokImpor_Click()
    If txtTotalImpor = "" Then
        LblerrMsg.Caption = "Please save the data first!"
        Exit Sub
    End If
    frmBC27BrowseBarangDokumen.txtNoPengajuan = Replace(txtNoPengajuan, "-", "")
    frmBC27BrowseBarangDokumen.txtNoSeri = txtNoSeri
    frmBC27BrowseBarangDokumen.txtKodeBarang = txtKodeBarang
    frmBC27BrowseBarangDokumen.txtNoSeriBahanBaku = txtNoSeriImpor
    frmBC27BrowseBarangDokumen.txtKodeAsalBahanBaku = 0
    frmBC27BrowseBarangDokumen.Show 1
End Sub

Private Sub cmdBrowseDokumen_Click()
   If cekSubmit = False Then
        LblerrMsg.Caption = "Please save the data first!"
        Exit Sub
    End If
    frmBC27BrowseBarangDokumen.txtNoPengajuan = Replace(txtNoPengajuan, "-", "")
    frmBC27BrowseBarangDokumen.txtNoSeri = txtNoSeri
    frmBC27BrowseBarangDokumen.txtKodeBarang = txtKodeBarang
    frmBC27BrowseBarangDokumen.Show 1
End Sub

Private Sub cmdBrowseDokumenLokal_Click()
   If txtTotalLokal = "" Then
        LblerrMsg.Caption = "Please save the data first!"
        Exit Sub
    End If
    frmBC27BrowseBarangDokumen.txtNoPengajuan = Replace(txtNoPengajuan, "-", "")
    frmBC27BrowseBarangDokumen.txtNoSeri = txtNoSeri
    frmBC27BrowseBarangDokumen.txtKodeBarang = txtKodeBarang
    frmBC27BrowseBarangDokumen.txtNoSeriBahanBaku = txtNoSeriLokal
    frmBC27BrowseBarangDokumen.txtKodeAsalBahanBaku = 1
    frmBC27BrowseBarangDokumen.Show 1

End Sub

Private Sub cmdBrowseTarif_Click()
If cekSubmit = False Then
    LblerrMsg.Caption = "Please save the data first!"
    Exit Sub
End If

If CekData = False Then
    frmBC27BrowseBeaMasukTambahan.txtNoPengajuan = txtNoPengajuan
    frmBC27BrowseBeaMasukTambahan.txtNoSeri = txtNoSeri
    frmBC27BrowseBeaMasukTambahan.txtNomorHS = txtNomorHS
    frmBC27BrowseBeaMasukTambahan.txtUraianBarang = txtUraianBarang
    frmBC27BrowseBeaMasukTambahan.txtCIF = txtNilaiCIF
    'frmBC27BrowseBeaMasukTambahan.txtCIFRupiah = txtCIFRupiah
    'frmBC27BrowseBeaMasukTambahan.cboJenisTarifBM = cboKeterangan1
    'If CDbl(txtTarifPersen1) > 0 Then
    'frmBC27BrowseBeaMasukTambahan.txtBesarTarif = Format(CDbl(txtTarifPersen1), "#,0.00")
    'End If
'    frmBC27BrowseBeaMasukTambahan.cboTarifFasilitas = cboKeterangan2
'    If CDbl(txtTarifPersen2) > 0 Then
'    frmBC27BrowseBeaMasukTambahan.txtTarifFasilitas = Format(CDbl(txtTarifPersen2), "#,0.00")
'    End If
'    If CDbl(txtTarifPersen1) > 0 Then
'        frmBC27BrowseBeaMasukTambahan.txtBMFasilitas = Format((CDbl(txtTarifPersen1) / 100) * CDbl(txtCIFRupiah), "#,0.00")
'    End If
    
    frmBC27BrowseBeaMasukTambahan.Show 1
Else
    frmBC27BrowseBeaMasukTambahan.txtNoPengajuan = txtNoPengajuan
    frmBC27BrowseBeaMasukTambahan.txtNoSeri = txtNoSeri
    frmBC27BrowseBeaMasukTambahan.up_LoadData txtNoPengajuan, txtNoSeri
    frmBC27BrowseBeaMasukTambahan.Show 1
End If
End Sub

Private Sub cmdBrowseTarifImpor_Click()
If cekSubmit = False Then
    LblerrMsg.Caption = "Please save the data first!"
    Exit Sub
End If

If CekData = False Then
    frmBC27BrowseBeaMasukTambahan.txtNoPengajuan = txtNoPengajuan
    frmBC27BrowseBeaMasukTambahan.txtNoSeri = txtNoSeri
    frmBC27BrowseBeaMasukTambahan.txtNoSeriBahanBaku = txtNoSeriImpor
    frmBC27BrowseBeaMasukTambahan.txtNomorHS = txtNomorHSImpor
    frmBC27BrowseBeaMasukTambahan.txtUraianBarang = txtUraianBarangImpor
    frmBC27BrowseBeaMasukTambahan.txtCIF = txtHargaCIFImpor
    frmBC27BrowseBeaMasukTambahan.cboJenisTarifBM = cboKeterangan1Impor
    If CDbl(txtTarifPersen1Impor) > 0 Then
        frmBC27BrowseBeaMasukTambahan.txtBesarTarif = Format(CDbl(txtTarifPersen1Impor), "#,0.00")
    End If
    frmBC27BrowseBeaMasukTambahan.cboTarifFasilitas = cboKeterangan2Impor
    If CDbl(txtTarifPersen2Impor) > 0 Then
        frmBC27BrowseBeaMasukTambahan.txtTarifFasilitas = Format(CDbl(txtTarifPersen2Impor), "#,0.00")
    End If
    If CDbl(txtTarifPersen1Impor) > 0 Then
            frmBC27BrowseBeaMasukTambahan.txtBMBayar = 0
            frmBC27BrowseBeaMasukTambahan.txtBMFasilitas = 0
    End If
    
    frmBC27BrowseBeaMasukTambahan.up_LoadDataBahanBaku txtNoPengajuan, txtNoSeri, txtNoSeriImpor, 0
    
    
    frmBC27BrowseBeaMasukTambahan.Show 1
Else
    frmBC27BrowseBeaMasukTambahan.txtNoPengajuan = txtNoPengajuan
    frmBC27BrowseBeaMasukTambahan.txtNoSeri = txtNoSeri

    frmBC27BrowseBeaMasukTambahan.up_LoadDataBahanBaku txtNoPengajuan, txtNoSeri, txtNoSeriImpor, 0
    frmBC27BrowseBeaMasukTambahan.Show 1
End If
End Sub

Private Sub cmdCancel_Click()
If SSTab1.Tab = 0 Then
    up_Clear
ElseIf SSTab1.Tab = 1 Then
    up_ClearImpor
ElseIf SSTab1.Tab = 2 Then
    up_ClearLokal
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub



Private Sub cmdDelete_Click()
If SSTab1.Tab = 0 Then
    If txtNoSeri = "" Then Exit Sub
    
    If MsgBox("Are you sure want to delete?", vbYesNo + vbExclamation, "Delete") = vbYes Then
        up_DeleteBarang
        up_LoadDataBarang txtNoPengajuan, 1
    End If
ElseIf SSTab1.Tab = 1 Then
    If txtNoSeriImpor = "" Then Exit Sub
    If MsgBox("Are you sure want to delete?", vbYesNo + vbExclamation, "Delete") = vbYes Then
         up_DeleteBahanBaku 0, txtNoSeriImpor
         up_LoadDataBahanBakuImpor txtNoPengajuan, txtNoSeri, 1
    End If
      
ElseIf SSTab1.Tab = 2 Then
    If txtNoSeriLokal = "" Then Exit Sub
    If MsgBox("Are you sure want to delete?", vbYesNo + vbExclamation, "Delete") = vbYes Then
         up_DeleteBahanBaku 1, txtNoSeriLokal
         up_LoadDataBahanBakuLokal txtNoPengajuan, txtNoSeri, 1
    End If
End If
End Sub

Private Sub cmdNewImpor_Click()
If txtTotalItem = "" Then
    LblerrMsg = "Please save the product first!"
    Exit Sub
End If

up_ClearImpor
up_GenerateNomorSeriBahanBaku txtNoPengajuan, 0, txtNoSeri

Dim iData As Integer

iData = uf_GetCountBahanBaku(txtNoPengajuan, txtNoSeri)

If txtJumlahBahanBaku = iData Then
    LblerrMsg = "Please add Jumlah Bahan Baku!"
    txtJumlahBahanBaku.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If

If MsgBox("Are you sure want to get the data from BC 2.3 Document?", vbQuestion & vbYesNo, "Confirmation") = vbYes Then
    Me.MousePointer = vbHourglass
    FrmBC27LoadBarang.Show 1
    Me.MousePointer = vbDefault
    
    txtPPNImpor = 10
    txtPPHImpor = 7.5
'
    If txtHargaCIFImpor = "" Then txtHargaCIFImpor = 0
End If
End Sub

Private Sub cmdNewLokal_Click()
If txtTotalItem = "" Then
    LblerrMsg = "Please save the product first!"
    Exit Sub
End If

up_ClearLokal
up_GenerateNomorSeriBahanBaku txtNoPengajuan, 1, txtNoSeri

Dim iData As Integer
iData = uf_GetCountBahanBaku(txtNoPengajuan, txtNoSeri)

If txtJumlahBahanBaku = iData Then
    LblerrMsg = "Please add Jumlah Bahan Baku!"
    txtJumlahBahanBaku.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If

If MsgBox("Are you sure want to get data from the last document?", vbQuestion & vbYesNo, "Confirmation") = vbYes Then
    Me.MousePointer = vbHourglass
    FrmBC27LoadBarangLokal.Show 1
    Me.MousePointer = vbDefault
    
End If
End Sub

Private Sub cmdNext_Click()
Dim ino As Integer


If SSTab1.Tab = 0 Then
    ino = txtNoSeri + 1
    If ino > txtTotalItem Then Exit Sub
    up_LoadDataBarang txtNoPengajuan, ino
    'up_GridLoadDokumen
    
    up_LoadDataBahanBakuImpor txtNoPengajuan, ino, 1
    'up_GridLoadDokumenImpor
    
    up_LoadDataBahanBakuLokal txtNoPengajuan, ino, 1
    'up_GridLoadDokumenLokal
ElseIf SSTab1.Tab = 1 Then
    If txtTotalImpor = "" Then Exit Sub
    ino = txtNoSeriImpor + 1
    If ino > txtTotalImpor Then Exit Sub
    up_LoadDataBahanBakuImpor txtNoPengajuan, txtNoSeri, ino
    'up_GridLoadDokumenImpor
ElseIf SSTab1.Tab = 2 Then
    If txtTotalLokal = "" Then Exit Sub
    ino = txtNoSeriLokal + 1
    If ino > txtTotalLokal Then Exit Sub
    up_LoadDataBahanBakuLokal txtNoPengajuan, txtNoSeri, ino
    'up_GridLoadDokumenLokal
End If
End Sub

Private Sub cmdNilaiBMImpor_Click()
frmBC27BrowseBarangTarifFasilitas.txtNoPengajuan = txtNoPengajuan
frmBC27BrowseBarangTarifFasilitas.txtNoSeri = txtNoSeri
frmBC27BrowseBarangTarifFasilitas.txtNoSeriBahan = txtNoSeriImpor
frmBC27BrowseBarangTarifFasilitas.up_GridLoadBahanBaku
frmBC27BrowseBarangTarifFasilitas.Show 1
End Sub

Private Sub cmdPrev_Click()
Dim ino As Integer

If SSTab1.Tab = 0 Then
    ino = txtNoSeri - 1
    If ino < 1 Then Exit Sub
    up_LoadDataBarang txtNoPengajuan, ino
    'up_GridLoadDokumen
    
    up_LoadDataBahanBakuImpor txtNoPengajuan, ino, 1
    'up_GridLoadDokumenImpor
    
    up_LoadDataBahanBakuLokal txtNoPengajuan, ino, 1
    'up_GridLoadDokumenLokal
ElseIf SSTab1.Tab = 1 Then
    If txtTotalImpor = "" Then Exit Sub
    ino = txtNoSeriImpor - 1
    If ino > txtTotalImpor Or ino < 1 Then Exit Sub
    up_LoadDataBahanBakuImpor txtNoPengajuan, txtNoSeri, ino
    'up_GridLoadDokumenImpor
ElseIf SSTab1.Tab = 2 Then
    If txtTotalLokal = "" Then Exit Sub
    ino = txtNoSeriLokal - 1
    If ino > txtTotalLokal Or ino < 1 Then Exit Sub
    up_LoadDataBahanBakuLokal txtNoPengajuan, txtNoSeri, ino
    'up_GridLoadDokumenLokal
End If
End Sub

Private Sub CmdSubmit_Click()
If SSTab1.Tab = 0 Then
    If uf_ValidateInputBarang = False Then Exit Sub
    up_SaveDataBarang
ElseIf SSTab1.Tab = 1 Then
    If txtTotalItem = "" Then
        LblerrMsg = "Please save the product first!"
        Exit Sub
    ElseIf txtJumlahBahanBaku = "" Then
        txtJumlahBahanBaku.SetFocus
        LblerrMsg = "Please set Jumlah Bahan Baku!"
        SSTab1.Tab = 0
        Exit Sub
    End If
    If uf_ValidateInputBahanBakuImpor = False Then Exit Sub
    
    gd_HargaPenyerahan = uf_GetHargaPenyerahan(txtNoPengajuan, txtNoSeri)
    up_SaveDataBahanBakuImpor
    gd_HargaPenyerahan = 0
ElseIf SSTab1.Tab = 2 Then
    If txtTotalItem = "" Then
        LblerrMsg = "Please save the product first!"
        Exit Sub
    ElseIf txtJumlahBahanBaku = "" Then
        txtJumlahBahanBaku.SetFocus
        LblerrMsg = "Please set Jumlah Bahan Baku!"
        SSTab1.Tab = 0
        Exit Sub
    End If
    If uf_ValidateInputBahanBakuLokal = False Then Exit Sub
    
    gd_HargaPenyerahan = uf_GetHargaPenyerahan(txtNoPengajuan, txtNoSeri)
    up_SaveDataBahanBakuLokal
    gd_HargaPenyerahan = 0
End If

End Sub

Private Sub cmdTarifFasilitas_Click()
frmBC27BrowseBarangTarifFasilitas.txtNoPengajuan = txtNoPengajuan
frmBC27BrowseBarangTarifFasilitas.txtNoSeri = txtNoSeri
frmBC27BrowseBarangTarifFasilitas.up_GridLoadTarif
frmBC27BrowseBarangTarifFasilitas.Show 1
End Sub

Private Sub Form_Activate()
'up_GridLoadDokumen
'up_GridLoadDokumenImpor
'up_GridLoadDokumenLokal

End Sub

Private Sub Form_Load()
    'up_FillComboPerhitungan
    up_Clear
    up_ClearImpor
    up_ClearLokal
    
    up_FillComboGeneral cboKeterangan1Impor, "Bea_Cukai_Jenis_Tarif Where ID In (5,6) ", "KODE_JENIS_TARIF", "URAIAN_JENIS_TARIF", 60, 150
    up_FillComboGeneral cboKeterangan2Impor, "Bea_Cukai_Tarif_Fasilitas Where KODE_FASILITAS IN ('0','6')", "KODE_FASILITAS", "URAIAN_Fasilitas", 60, 150
    up_FillComboGeneral cboKeterangan3Impor, "Bea_Cukai_Tarif_Fasilitas Where KODE_FASILITAS IN ('0','6')", "KODE_FASILITAS", "URAIAN_Fasilitas", 60, 150
    up_FillComboGeneral cboKeterangan4Impor, "Bea_Cukai_Tarif_Fasilitas Where KODE_FASILITAS IN ('0','6')", "KODE_FASILITAS", "URAIAN_Fasilitas", 60, 150
    up_FillComboGeneral cboKeterangan5Impor, "Bea_Cukai_Tarif_Fasilitas Where KODE_FASILITAS IN ('0','6')", "KODE_FASILITAS", "URAIAN_Fasilitas", 60, 150
    
    up_FillComboGeneral cboCukaiImpor, "Bea_Cukai_Komoditi Where ID In (5,6,7) ", "KODE_KOMODITI", "URAIAN_KOMODITI", 100, 200
    up_FillComboGeneral cboJenisTarifImpor, "Bea_Cukai_Jenis_Tarif Where ID In (5,6,7) ", "KODE_JENIS_TARIF", "URAIAN_JENIS_TARIF", 60, 150
    
    up_FillComboGeneral cboKeteranganJenisTarif, "Bea_Cukai_Tarif_Fasilitas Where KODE_FASILITAS IN ('0','6') ", "KODE_FASILITAS", "URAIAN_Fasilitas", 60, 150
    
    up_FillComboGeneral cboJenisPPNLokal, "Bea_Cukai_Tarif_Fasilitas Where KODE_FASILITAS IN ('0','4') ", "KODE_FASILITAS", "URAIAN_Fasilitas", 60, 150
End Sub

Private Sub txtBMBrgJadi_GotFocus()
'txtBMBrgJadi = CDbl(txtBMBrgJadi)
End Sub

Private Sub txtBMBrgJadi_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtDokumenAsalImpor_Change()
sql = "Select Uraian_Dokumen From Bea_Cukai_Dokumen Where Kode_Dokumen =  '" & txtDokumenAsalImpor.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lblDokAsalImpor.Caption = Trim(RS("Uraian_Dokumen"))
    Else
        lblDokAsalImpor.Caption = ""
        Exit Sub
    End If
End Sub

Private Sub txtDokumenAsalImpor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    frmBC27BrowseGeneral.gs_TableName = "Dokumen Asal"
    frmBC27BrowseGeneral.Show 1
End If
End Sub

Private Sub txtDokumenAsalImpor_LostFocus()
gb_LoadDataMaster "(Select Kode_Dokumen, Uraian_Dokumen From Bea_Cukai_Dokumen Where Kode_Dokumen in (16,23,27,52) Union All Select 99, 'LAINNYA') a", "Uraian_Dokumen", lblDokAsalImpor, "Where Kode_Dokumen = '" & txtDokumenAsalImpor & "'"
End Sub

Private Sub txtDokumenAsalLokal_Change()
sql = "Select Uraian_Dokumen From Bea_Cukai_Dokumen Where Kode_Dokumen =  '" & txtDokumenAsalLokal.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lblDokAsalLokal.Caption = Trim(RS("Uraian_Dokumen"))
    Else
        lblDokAsalLokal.Caption = ""
        Exit Sub
    End If
End Sub

Private Sub txtDokumenAsalLokal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFasilitas_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFasilitasImpor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtHargaCIF_LostFocus()
'If txtHargaCIF = "" Then txtHargaCIF = 0
'txtCIFRupiahImpor = Format(CDbl(txtHargaCIF) * CDbl(txtNDPBMImpor), "#,0.00")
'txtHargaCIF = Format(CDbl(txtHargaCIF), "#,0.00")
End Sub

Private Sub txtHargaPerolahen_GotFocus()
If txtHargaPerolahen = "" Then txtHargaPerolahen = 0
txtHargaPerolahen = CDbl(txtHargaPerolahen)
End Sub

Private Sub txtHargaPerolahen_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtHargaPerolahen_LostFocus()
'txtCIFRupiahLokal = Format(CDbl(txtHargaCIFUSDLokal) * CDbl(txtNDPBMLokal), "#,0.00")
txtHargaPerolahen = Format(CDbl(txtHargaPerolahen), "#,0.00")
End Sub

Private Sub txtHargaPenyerahan_GotFocus()
If txtHargaPenyerahan = "" Then txtHargaPenyerahan = 0
txtHargaPenyerahan = CDbl(txtHargaPenyerahan)
End Sub

Private Sub txtHargaPenyerahan_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtHargaPenyerahan_LostFocus()
txtHargaPenyerahan = Format(CDbl(txtHargaPenyerahan), "#,0.00")
End Sub

Private Sub txtJenisKemasan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtJenisSatuan_Change()
sql = "Select Uraian_Satuan From Bea_Cukai_Satuan Where Kode_Satuan =  '" & txtJenisSatuan.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lblJenisSatuan(6).Caption = Trim(RS("Uraian_Satuan"))
    Else
        lblJenisSatuan(6).Caption = ""
        Exit Sub
    End If
End Sub

Private Sub txtJenisSatuan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtJumlahBahanBaku_GotFocus()
If txtJumlahBahanBaku = "" Then txtJumlahBahanBaku = 0
txtJumlahBahanBaku = CDbl(txtJumlahBahanBaku)
End Sub

Private Sub txtJumlahBahanBaku_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtJumlahCukai_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtJumlahKemasan_GotFocus()
'If txtJumlahKemasan = "" Then txtJumlahKemasan = 0
'txtJumlahKemasan = CDbl(txtJumlahKemasan)
End Sub

Private Sub txtJumlahKemasan_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtJumlahSatuan_GotFocus()
If txtJumlahSatuan = "" Then txtJumlahSatuan = 0
txtJumlahSatuan = CDbl(txtJumlahSatuan)
End Sub

Private Sub txtJumlahSatuan_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtJumlahSatuanImpor_GotFocus()
If txtJumlahSatuanImpor = "" Then txtJumlahSatuanImpor = 0
txtJumlahSatuanImpor = CDbl(txtJumlahSatuanImpor)
End Sub

Private Sub txtJumlahSatuanImpor_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtJumlahSatuanImpor_LostFocus()
If txtJumlahSatuanImpor = "" Then txtJumlahSatuanImpor = 0
txtJumlahSatuanImpor = Format(CDbl(txtJumlahSatuanImpor), "#,0")
End Sub

Private Sub txtJumlahSatuanLokal_GotFocus()
If txtJumlahSatuanLokal = "" Then txtJumlahSatuanLokal = 0
txtJumlahSatuanLokal = CDbl(txtJumlahSatuanLokal)
End Sub

Private Sub txtJumlahSatuanLokal_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtJumlahSatuanLokal_LostFocus()
If txtJumlahSatuanLokal = "" Then txtJumlahSatuanLokal = 0
txtJumlahSatuanLokal = Format(CDbl(txtJumlahSatuanLokal), "#,0")
End Sub

Private Sub txtJumlahSpesifik_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtKategoriBarang_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    frmBC27BrowseGeneral.gs_TableName = "Kategori"
    frmBC27BrowseGeneral.Show 1
End If
End Sub

Private Sub txtKategoriBarang_LostFocus()
'gb_LoadDataMaster "Bea_Cukai_Kategori_BarangBC25", "Uraian_Kategori", lblKategori, "Where Kode_Kategori = '" & txtKategoriBarang & "'"
End Sub

Private Sub txtKodeBarang_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    frmBC27BrowseGeneral.gs_TableName = "Barang"
    frmBC27BrowseGeneral.Show 1
End If
End Sub

Private Sub txtKodeBarang_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtKodeBarangImpor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    frmBC27BrowseGeneral.gs_TableName = "Barang Impor"
    frmBC27BrowseGeneral.Show 1
End If

End Sub

Private Sub txtKodeBarangLokal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtKondisiBarang_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    frmBC27BrowseGeneral.gs_TableName = "Kondisi"
    frmBC27BrowseGeneral.Show 1
End If
End Sub

Private Sub txtKondisiBarang_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtKondisiBarang_LostFocus()
'gb_LoadDataMaster "Bea_Cukai_Kondisi_Barang", "Uraian_Kondisi", lblKondisiBarang, "Where Kode_Kondisi = '" & txtKondisiBarang & "'"

End Sub

Private Sub txtKPPBCImpor_Change()
sql = "SELECT Nama_Kantor FROM Bea_Cukai_Kantor_pabean Where Kode_Kantor = '" & txtKPPBCImpor.Text & "'"
        Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lblKPPBCImpor.Caption = Trim(RS("Nama_Kantor"))
    Else
        lblKPPBCImpor.Caption = ""
        Exit Sub
    End If
End Sub

Private Sub txtKPPBCImpor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    frmBC27BrowseGeneral.gs_TableName = "KPPBC Impor"
    frmBC27BrowseGeneral.Show 1
End If
End Sub

Private Sub txtKPPBCImpor_LostFocus()
'gb_LoadDataMaster "Bea_Cukai_Kantor_Pabean", "Nama_Kantor", lblPenggunaan, "Where Kode_Kantor = '" & txtKPPBCImpor & "'"
End Sub

Private Sub txtKPPBCLokal_Change()
 sql = "SELECT Nama_Kantor FROM Bea_Cukai_Kantor_pabean Where Kode_Kantor = '" & txtKPPBCLokal.Text & "'"
        Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lblKPPBCLokal.Caption = Trim(RS("Nama_Kantor"))
    Else
        lblKPPBCLokal.Caption = ""
        Exit Sub
    End If
End Sub

Private Sub txtKPPBCLokal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    frmBC27BrowseGeneral.gs_TableName = "KPPBC Lokal"
    frmBC27BrowseGeneral.Show 1
End If
End Sub

Private Sub txtKPPBCLokal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMerk_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMerkImpor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMerkLokal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNDPBMImpor_LostFocus()
'If txtNDPBMImpor = "" Then txtNDPBMImpor = 0
'txtCIFRupiahImpor = Format(CDbl(txtHargaCIF) * CDbl(txtNDPBMImpor), "#,0.00")
'txtNDPBMImpor = Format(CDbl(txtNDPBMImpor), "#,0.00")
End Sub

Private Sub txtHargaPenyerahanLokal_GotFocus()
If txtHargaPenyerahanLokal = "" Then txtHargaPenyerahanLokal = 0
txtHargaPenyerahanLokal = CDbl(txtHargaPenyerahanLokal)
End Sub

Private Sub txtHargaPenyerahanLokal_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtHargaPenyerahanLokal_LostFocus()
If txtHargaPenyerahanLokal = "" Then txtHargaPenyerahanLokal = 0
'txtCIFRupiahLokal = Format(CDbl(txtHargaCIFUSDLokal) * CDbl(txtNDPBMLokal), "#,0.00")
txtHargaPenyerahanLokal = Format(CDbl(txtHargaPenyerahanLokal), "#,0.00")
End Sub

Private Sub txtNegaraAsal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNetto_GotFocus()
If txtNetto = "" Then txtNetto = 0
txtNetto = CDbl(txtNetto)
End Sub

Private Sub txtNetto_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


Private Sub txtNettoImpor_GotFocus()
'If txtNettoImpor = "" Then txtNettoImpor = 0
'txtNettoImpor = CDbl(txtNettoImpor)
End Sub

Private Sub txtNettoImpor_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtNettoImpor_LostFocus()
'If txtNettoImpor = "" Then txtNettoImpor = 0
'txtNettoImpor = Format(CDbl(txtNettoImpor), "#,0.00")
End Sub

Private Sub txtNoAjuImpor_GotFocus()
txtNoAjuImpor = Replace(Replace(txtNoAjuImpor, "-", ""), ".", "")
End Sub

Private Sub txtNoAjuImpor_LostFocus()
txtNoAjuImpor = Left(txtNoAjuImpor.Text, 2) & "." & Mid(txtNoAjuImpor.Text, 3, 3) & "." & Mid(txtNoAjuImpor.Text, 6, 3) & "." & Mid(txtNoAjuImpor.Text, 9, 1) & "-" & Mid(txtNoAjuImpor.Text, 10, 3) & "." & Mid(txtNoAjuImpor.Text, 13, 3)
End Sub

Private Sub txtNoAjuLokal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNoImpor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNoLokal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNomorHS_GotFocus()
txtNomorHS = Replace(txtNomorHS, ".", "")
End Sub

Private Sub txtNomorHS_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNomorHS_LostFocus()
    txtNomorHS = Replace(txtNomorHS, ".", "")
    txtNomorHS = Mid(txtNomorHS.Text, 1, 10)
    If txtNomorHS <> "" Then
        txtNomorHS = Left(txtNomorHS.Text, 4) & "." & Mid(txtNomorHS.Text, 5, 2) & "." & Mid(txtNomorHS.Text, 7, 2) & "." & Mid(txtNomorHS.Text, 9, 2)
    End If
End Sub


Private Sub txtNomorHSImpor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNomorHSLokal_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtPenggunaan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    frmBC27BrowseGeneral.gs_TableName = "Penggunaan"
    frmBC27BrowseGeneral.Show 1
End If
End Sub

Private Sub txtPenggunaan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtPenggunaan_LostFocus()
'gb_LoadDataMaster "Bea_Cukai_Kode_Guna", "Uraian_Guna", lblPenggunaan, "Where Kode_Guna = '" & txtPenggunaan & "'"
End Sub

Private Sub txtPPh_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtPPn_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtPPNBm_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtPPNImpor_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtPPNLokal_LostFocus()
If Left(cboJenisPPNLokal, 1) = "0" Then
   txtPPNBayarLokal = Format(CDbl(txtHargaPenyerahanLokal) * (CDbl(txtPPNLokal) / 100), "#,0.00")
   txtPPNFasilitasLokal = 0
ElseIf Left(cboJenisPPNLokal, 1) = "4" Then
   txtPPNFasilitasLokal = Format(CDbl(txtHargaPenyerahanLokal) * (CDbl(txtPPNLokal) / 100), "#,0.00")
   txtPPNBayarLokal = 0
Else
   txtPPNBayarLokal = "0.00"
   txtPPNFasilitasLokal = "0.00"
End If
End Sub

Private Sub txtSatuanImpor_Change()
sql = "Select Uraian_Satuan From Bea_Cukai_Satuan Where Kode_Satuan =  '" & txtSatuanImpor.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lblSatuanImpor.Caption = Trim(RS("Uraian_Satuan"))
    Else
        lblSatuanImpor.Caption = ""
        Exit Sub
    End If
End Sub

Private Sub txtSatuanImpor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtSatuanLokal_Change()
sql = "Select Uraian_Satuan From Bea_Cukai_Satuan Where Kode_Satuan =  '" & txtSatuanLokal.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        lblSatuanLokal.Caption = Trim(RS("Uraian_Satuan"))
    Else
        lblSatuanLokal.Caption = ""
        Exit Sub
    End If
End Sub

Private Sub txtSatuanLokal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSkemaTarif_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSkemaTarifImpor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSpfLain_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSpfLainImpor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSpfLainLokal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTarif_KeyDown(KeyCode As Integer, Shift As Integer)
'    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
'    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtTarifPersen1_GotFocus()
'If txtTarifPersen1 = "" Then txtTarifPersen1 = 0
'txtTarifPersen1 = CDbl(txtTarifPersen1)
End Sub

Private Sub txtTarifPersen1_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


Private Sub txtTarifPersen1_LostFocus()
'If txtTarifPersen1 = "" Then txtTarifPersen1 = 0
'txtTarifPersen1 = Format(CDbl(txtTarifPersen1), "#,0.00")
End Sub

Private Sub txtTarifPersen1Impor_Change()
If txtTarifPersen1Impor = "" Then txtTarifPersen1Impor = 0
txtTarifPersen1Impor = CDbl(txtTarifPersen1Impor)
End Sub

Private Sub txtTarifPersen1Impor_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtTarifPersen1Impor_LostFocus()
If txtTarifPersen1Impor = "" Then txtTarifPersen1Impor = 0
txtTarifPersen1Impor = Format(CDbl(txtTarifPersen1Impor), "#,0")
End Sub

Private Sub txtTarifPersen2_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtTarifPersen2_LostFocus()
'If txtTarifPersen2 = "" Then txtTarifPersen2 = 0
'txtTarifPersen2 = Format(CDbl(txtTarifPersen2), "#,0.00")
End Sub

Private Sub txtTarifPersen3_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtTarifPersen3_LostFocus()
'If txtTarifPersen3 = "" Then txtTarifPersen3 = 0
'txtTarifPersen3 = Format(CDbl(txtTarifPersen3), "#,0.00")
End Sub

Private Sub txtTarifPersen4_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtTarifPersen4_LostFocus()
'If txtTarifPersen4 = "" Then txtTarifPersen4 = 0
'txtTarifPersen4 = Format(CDbl(txtTarifPersen4), "#,0.00")
End Sub

Private Sub txtTarifPersen5_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtTarifPersen5_LostFocus()
'If txtTarifPersen5 = "" Then txtTarifPersen5 = 0
'txtTarifPersen5 = Format(CDbl(txtTarifPersen5), "#,0.00")
End Sub

Private Sub txtTipe_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtTipeImpor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTipeLokal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUkuran_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUkuranImpor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUkuranLokal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUraianBarang_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUraianBarangLokal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtvolume_GotFocus()
If txtVolume = "" Then txtVolume = 0

txtVolume = CDbl(txtVolume)
End Sub

Private Sub txtVolume_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

