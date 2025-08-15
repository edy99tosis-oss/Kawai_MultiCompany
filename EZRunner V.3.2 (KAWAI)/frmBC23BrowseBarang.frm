VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmBC23BrowseBarang 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Barang BC 23"
   ClientHeight    =   11010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13230
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBC23BrowseBarang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   13230
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFreightFix 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   350
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   115
      Top             =   10200
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.TextBox txtCIFFix 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   350
      Left            =   7920
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   114
      Top             =   10200
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.TextBox txtNoPengajuan 
      Appearance      =   0  'Flat
      Height          =   350
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   111
      Top             =   10200
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FDDFE3&
      Caption         =   "CUKAI"
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
      Left            =   240
      TabIndex        =   93
      Top             =   7920
      Width           =   7335
      Begin VB.TextBox txtPersenCukai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   4560
         MaxLength       =   5
         TabIndex        =   47
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txtTarif 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   3720
         MaxLength       =   5
         TabIndex        =   43
         Top             =   600
         Width           =   705
      End
      Begin VB.TextBox txtSatuanCukai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   4800
         MaxLength       =   5
         TabIndex        =   44
         Top             =   600
         Width           =   705
      End
      Begin VB.TextBox txtJumlahCukai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   45
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   43
         Left            =   5400
         TabIndex        =   101
         Top             =   1035
         Width           =   180
      End
      Begin MSForms.ComboBox cboKeterangan 
         Height          =   345
         Left            =   2520
         TabIndex        =   46
         Tag             =   "TTFF*/"
         Top             =   960
         Width           =   1935
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3413;617"
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
         Index           =   42
         Left            =   4560
         TabIndex        =   100
         Top             =   675
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         Height          =   195
         Index           =   37
         Left            =   120
         TabIndex        =   96
         Top             =   1035
         Width           =   600
      End
      Begin MSForms.ComboBox cboJenisTarif 
         Height          =   345
         Left            =   1680
         TabIndex        =   42
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   1935
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3413;609"
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
         Index           =   36
         Left            =   120
         TabIndex        =   95
         Top             =   660
         Width           =   870
      End
      Begin MSForms.ComboBox cboKomoditi 
         Height          =   345
         Left            =   1680
         TabIndex        =   41
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   1935
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3413;609"
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
         Caption         =   "Komoditi"
         Height          =   195
         Index           =   35
         Left            =   120
         TabIndex        =   94
         Top             =   300
         Width           =   750
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   240
      TabIndex        =   91
      Tag             =   "TFTT*/"
      Top             =   9480
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
         TabIndex        =   92
         Tag             =   "TTFF*/"
         Top             =   195
         Width           =   12570
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FDDFE3&
      Height          =   4575
      Left            =   7680
      TabIndex        =   79
      Top             =   4800
      Width           =   5415
      Begin VB.CommandButton cmdBrowseDokumen 
         BackColor       =   &H0080FFFF&
         Caption         =   "Browse"
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtSkema 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   240
         MaxLength       =   5
         TabIndex        =   40
         Top             =   1320
         Width           =   705
      End
      Begin VB.TextBox txtFasilitas 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   240
         MaxLength       =   5
         TabIndex        =   39
         Top             =   480
         Width           =   705
      End
      Begin VSFlex8Ctl.VSFlexGrid grid 
         Height          =   1935
         Left            =   240
         TabIndex        =   99
         TabStop         =   0   'False
         Tag             =   "TTTT*/"
         Top             =   1800
         Width           =   4965
         _cx             =   8758
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
      Begin VB.Label lblSkema 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   255
         Left            =   1080
         TabIndex        =   106
         Top             =   1365
         Width           =   4215
      End
      Begin VB.Label lblFasilitas 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   255
         Left            =   1080
         TabIndex        =   105
         Top             =   525
         Width           =   4215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SKEMA TARIF"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   40
         Left            =   240
         TabIndex        =   98
         Top             =   1000
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FASILITAS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   38
         Left            =   240
         TabIndex        =   97
         Top             =   140
         Width           =   1035
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   3015
      Left            =   240
      TabIndex        =   78
      Top             =   4800
      Width           =   7335
      Begin VB.TextBox txtSatuanTarif 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   26
         Top             =   647
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtJumlahSpesifik 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   27
         Top             =   1080
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdTarifFasilitas 
         BackColor       =   &H0080FFFF&
         Caption         =   "Tarif && Fasilitas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtTarifPersen5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   38
         Top             =   2520
         Width           =   705
      End
      Begin VB.TextBox txtPPh 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   36
         Top             =   2520
         Width           =   705
      End
      Begin VB.TextBox txtTarifPersen4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   35
         Top             =   2040
         Width           =   705
      End
      Begin VB.TextBox txtPPNBm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   33
         Top             =   2040
         Width           =   705
      End
      Begin VB.TextBox txtTarifPersen3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   32
         Top             =   1560
         Width           =   705
      End
      Begin VB.TextBox txtPPN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   30
         Top             =   1560
         Width           =   705
      End
      Begin VB.TextBox txtTarifPersen2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   5880
         MaxLength       =   5
         TabIndex        =   29
         Top             =   1080
         Width           =   705
      End
      Begin VB.TextBox txtTarifPersen1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   25
         Top             =   647
         Width           =   705
      End
      Begin VB.CommandButton cmdBrowseTarif 
         BackColor       =   &H0080FFFF&
         Caption         =   "O"
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   635
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Satuan"
         Height          =   195
         Index           =   39
         Left            =   120
         TabIndex        =   116
         Top             =   1155
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   34
         Left            =   6720
         TabIndex        =   90
         Top             =   2598
         Width           =   180
      End
      Begin MSForms.ComboBox cboKeterangan5 
         Height          =   345
         Left            =   2640
         TabIndex        =   37
         Tag             =   "TTFF*/"
         Top             =   2520
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
         Index           =   33
         Left            =   2355
         TabIndex        =   89
         Top             =   2598
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPh"
         Height          =   195
         Index           =   32
         Left            =   120
         TabIndex        =   88
         Top             =   2595
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   31
         Left            =   6720
         TabIndex        =   87
         Top             =   2115
         Width           =   180
      End
      Begin MSForms.ComboBox cboKeterangan4 
         Height          =   345
         Left            =   2640
         TabIndex        =   34
         Tag             =   "TTFF*/"
         Top             =   2040
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
         Index           =   30
         Left            =   2355
         TabIndex        =   86
         Top             =   2118
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPnBM"
         Height          =   195
         Index           =   29
         Left            =   120
         TabIndex        =   85
         Top             =   2115
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   28
         Left            =   6720
         TabIndex        =   84
         Top             =   1635
         Width           =   180
      End
      Begin MSForms.ComboBox cboKeterangan3 
         Height          =   345
         Left            =   2640
         TabIndex        =   31
         Tag             =   "TTFF*/"
         Top             =   1560
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
         Index           =   27
         Left            =   2355
         TabIndex        =   83
         Top             =   1638
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPN"
         Height          =   195
         Index           =   26
         Left            =   120
         TabIndex        =   82
         Top             =   1635
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   25
         Left            =   6720
         TabIndex        =   81
         Top             =   1155
         Width           =   180
      End
      Begin MSForms.ComboBox cboKeterangan2 
         Height          =   345
         Left            =   2640
         TabIndex        =   28
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   2895
         VariousPropertyBits=   746604575
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
         Index           =   24
         Left            =   5640
         TabIndex        =   80
         Top             =   720
         Width           =   180
      End
      Begin MSForms.ComboBox cboKeterangan1 
         Height          =   345
         Left            =   2640
         TabIndex        =   24
         Tag             =   "TTFF*/"
         Top             =   645
         Width           =   2055
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3625;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboJenisPungutan 
         Height          =   345
         Left            =   120
         TabIndex        =   22
         Tag             =   "TTFF*/"
         Top             =   650
         Width           =   1935
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3413;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Caption         =   "  KEMASAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   7680
      TabIndex        =   63
      Top             =   2040
      Width           =   5415
      Begin VB.TextBox txtNegara 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   240
         MaxLength       =   5
         TabIndex        =   21
         Top             =   2160
         Width           =   705
      End
      Begin VB.TextBox txtNetto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1200
         Width           =   1185
      End
      Begin VB.TextBox txtJenis 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   19
         Top             =   720
         Width           =   1185
      End
      Begin VB.TextBox txtJumlah 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label lblNegara 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   255
         Left            =   1080
         TabIndex        =   104
         Top             =   2208
         Width           =   3015
      End
      Begin VB.Label lblNetto 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   255
         Left            =   2520
         TabIndex        =   103
         Top             =   1245
         Width           =   495
      End
      Begin VB.Label lblJenis 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   255
         Left            =   2520
         TabIndex        =   102
         Top             =   768
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NEGARA ASAL BARANG"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   240
         TabIndex        =   77
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Netto"
         Height          =   195
         Index           =   19
         Left            =   240
         TabIndex        =   76
         Top             =   1275
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   75
         Top             =   795
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   74
         Top             =   315
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Caption         =   "HARGA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   62
      Top             =   2040
      Width           =   7335
      Begin VB.TextBox txtCIF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   17
         Top             =   2160
         Width           =   1785
      End
      Begin VB.TextBox txtHargaCIF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   16
         Top             =   1680
         Width           =   1785
      End
      Begin VB.TextBox txtAsuransi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   15
         Top             =   1200
         Width           =   1785
      End
      Begin VB.TextBox txtFreight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   14
         Top             =   720
         Width           =   1785
      End
      Begin VB.TextBox txtHargaDetil 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   13
         Top             =   240
         Width           =   1785
      End
      Begin VB.TextBox txtHargaSatuan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   12
         Top             =   2160
         Width           =   1785
      End
      Begin VB.TextBox txtSatuan 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1680
         Width           =   705
      End
      Begin VB.TextBox txtJumlahSatuan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1200
         Width           =   1785
      End
      Begin VB.TextBox txtBTDiskon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   9
         Top             =   720
         Width           =   1785
      End
      Begin VB.TextBox txtTotalDetilFOB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   8
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label lblSatuan 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   255
         Left            =   2640
         TabIndex        =   110
         Top             =   1728
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CIF Rp"
         Height          =   195
         Index           =   16
         Left            =   3960
         TabIndex        =   73
         Top             =   2235
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga CIF"
         Height          =   195
         Index           =   15
         Left            =   3960
         TabIndex        =   72
         Top             =   1755
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asuransi"
         Height          =   195
         Index           =   14
         Left            =   3960
         TabIndex        =   71
         Top             =   1275
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Freight"
         Height          =   195
         Index           =   13
         Left            =   3960
         TabIndex        =   70
         Top             =   795
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Detil"
         Height          =   195
         Index           =   12
         Left            =   3960
         TabIndex        =   69
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Satuan"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   68
         Top             =   2235
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   67
         Top             =   1755
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Satuan"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   66
         Top             =   1275
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BT-Diskon"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   65
         Top             =   795
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total / Detil (FOB)"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   64
         Top             =   315
         Width           =   1560
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0080FFFF&
      Caption         =   "Close"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   10200
      Width           =   975
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Submit"
      Height          =   375
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   10200
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   10200
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H0080FFFF&
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   10200
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1815
      Left            =   240
      TabIndex        =   53
      Top             =   120
      Width           =   12855
      Begin VB.TextBox txtNoSeri 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   11850
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtSpfLain 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   11040
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1320
         Width           =   1545
      End
      Begin VB.TextBox txtUkuran 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   7800
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1320
         Width           =   1545
      End
      Begin VB.TextBox txtTipe 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   4920
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1320
         Width           =   1545
      End
      Begin VB.TextBox txtMerk 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1320
         Width           =   1545
      End
      Begin VB.TextBox txtNomorHS 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   5400
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   2265
      End
      Begin VB.TextBox txtKategoriBarang 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   3
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txtUraianBarang 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   2
         Top             =   600
         Width           =   9345
      End
      Begin VB.TextBox txtKodeBarang 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spf Lain"
         Height          =   195
         Index           =   23
         Left            =   10080
         TabIndex        =   109
         Top             =   1398
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ukuran"
         Height          =   195
         Index           =   21
         Left            =   6840
         TabIndex        =   108
         Top             =   1398
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe"
         Height          =   195
         Index           =   20
         Left            =   3960
         TabIndex        =   107
         Top             =   1398
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS :"
         Height          =   195
         Index           =   4
         Left            =   10080
         TabIndex        =   61
         Top             =   210
         Width           =   825
      End
      Begin VB.Label LblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "LENGKAP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11040
         TabIndex        =   60
         Top             =   210
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Merk"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   59
         Top             =   1398
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uraian Barang"
         Height          =   195
         Index           =   6
         Left            =   2640
         TabIndex        =   58
         Top             =   1038
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor HS"
         Height          =   195
         Index           =   3
         Left            =   4200
         TabIndex        =   57
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori Barang"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   56
         Top             =   1038
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uraian Barang"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   678
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   318
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmBC23BrowseBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public cekSubmit As Boolean
Dim cekLoad As Boolean
Public CekData As Boolean

'-------------------------------------------
Const colJenisDokumen As Integer = 0
Const colNomorDokumen As Integer = 1
Const colTanggal As Integer = 2
Const colcount As Integer = 3

Private Sub up_Clear()
LblerrMsg.Caption = ""
lblFasilitas.Caption = ""
lblJenis.Caption = ""
lblNegara.Caption = ""
lblNetto.Caption = ""
lblSatuan.Caption = ""
lblSkema.Caption = ""
Label1(6).Caption = ""
LblStatus.Caption = ""

txtTotalDetilFOB = "0.00"
txtBTDiskon = "0.00"
txtJumlahSatuan = "0.00"
txtHargaCIF = "0.00"
txtAsuransi = "0.00"
txtHargaDetil = "0.00"
txtCIF = "0.00"

'cboJenisPungutan.ListIndex = 0

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

Private Sub up_GridHeader()
    With Grid
        .ColS = colcount
        .Rows = 1
        
        .TextMatrix(0, colJenisDokumen) = "Jenis"
        .TextMatrix(0, colNomorDokumen) = "Nomor"
        .TextMatrix(0, colTanggal) = "Tanggal"
        
        .ColWidth(colJenisDokumen) = 1500
        .ColWidth(colNomorDokumen) = 1500
        .ColWidth(colTanggal) = 1200
        .ColAlignment(colNomorDokumen) = flexAlignLeftCenter
        
        .ColFormat(colTanggal) = "dd MMM yyyy"
    End With
End Sub

Private Sub up_LoadKategoriBarang(pKode As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select * From Bea_Cukai_Kategori_Barang Where KODE_KATEGORI = '" & pKode & "'"
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    Label1(6).Caption = RS.Fields("URAIAN_KATEGORI")
Else
    Label1(6).Caption = ""
End If
End Sub

Private Sub up_LoadSatuan(pKode As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select * From Bea_Cukai_Satuan Where Kode_Satuan = '" & pKode & "'"
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    lblSatuan.Caption = RS.Fields("Uraian_Satuan")
Else
    lblSatuan.Caption = ""
End If
End Sub

Private Sub up_LoadJenisKemasan(pKode As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select * From Bea_Cukai_Kemasan Where Kode_Kemasan = '" & pKode & "'"
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    lblJenis.Caption = RS.Fields("Uraian_Kemasan")
Else
    lblJenis.Caption = ""
End If
End Sub

Private Sub up_LoadNegara(pKode As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select * From Bea_Cukai_Negara Where Kode_Negara = '" & pKode & "'"
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    lblNegara.Caption = RS.Fields("Nama_Negara")
Else
    lblNegara.Caption = ""
End If
End Sub

Private Sub up_LoadFasilitas(pKode As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select * From Bea_Cukai_Fasilitas Where Kode_Fasilitas = '" & pKode & "'"
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    lblFasilitas.Caption = RS.Fields("Uraian_Fasilitas")
Else
    lblFasilitas.Caption = ""
End If
End Sub

Private Sub up_LoadSkemaTarif(pKode As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select * From Bea_Cukai_Skema_Tarif Where Kode_Skema = '" & pKode & "'"
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    lblSkema.Caption = RS.Fields("Uraian_Skema")
Else
    lblSkema.Caption = ""
End If
End Sub

Public Sub up_LoadDataBarang(pNoPengajuan As String, pNoSeri As Integer)
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim NomorHS As String
        
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23LoadDetailBarang_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(pNoPengajuan, "-", ""))
    cmd.Parameters.append cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, pNoSeri)
    
    Set RS = cmd.Execute
    
    cekLoad = True
    
    If Not RS.EOF Then
        cekSubmit = True
        
        NomorHS = IIf(IsNull(RS.Fields("POS_TARIF")), "", RS.Fields("POS_TARIF"))
        
        txtNomorHS = Replace(NomorHS, ".", "")
        txtNomorHS = Mid(txtNomorHS.Text, 1, 10)
        If txtNomorHS <> "" Then
            txtNomorHS = Left(txtNomorHS.Text, 4) & "." & Mid(txtNomorHS.Text, 5, 2) & "." & Mid(txtNomorHS.Text, 7, 2) & "." & Mid(txtNomorHS.Text, 9, 2)
        End If
        
        txtKodeBarang = IIf(IsNull(RS.Fields("Kode_Barang")), "", RS.Fields("Kode_Barang"))
        txtUraianBarang = IIf(IsNull(RS.Fields("Uraian_Barang")), "", RS.Fields("Uraian_Barang"))
'        txtNomorHS = IIf(IsNull(rs.fields("POS_TARIF")), "", rs.fields("POS_TARIF"))
        txtKategoriBarang = IIf(IsNull(RS.Fields("KATEGORI_BARANG")), "", RS.Fields("KATEGORI_BARANG"))
        Label1(6).Caption = IIf(IsNull(RS.Fields("URAIAN_KATEGORI")), "", RS.Fields("URAIAN_KATEGORI"))
        txtMerk = IIf(IsNull(RS.Fields("Merk")), "", RS.Fields("Merk"))
        txtTipe = IIf(IsNull(RS.Fields("Tipe")), "", RS.Fields("Tipe"))
        txtUkuran = IIf(IsNull(RS.Fields("UKURAN")), "", RS.Fields("UKURAN"))
        txtSpfLain = IIf(IsNull(RS.Fields("SPESIFIKASI_LAIN")), "", RS.Fields("SPESIFIKASI_LAIN"))
        
        txtTotalDetilFOB = Format(IIf(IsNull(RS.Fields("FOB")), 0, RS.Fields("FOB")), "#,0.00")
        txtBTDiskon = Format(IIf(IsNull(RS.Fields("DISKON")), 0, RS.Fields("DISKON")), "#,0.00")
        txtJumlahSatuan = Format(IIf(IsNull(RS.Fields("JUMLAH_SATUAN")), 0, RS.Fields("JUMLAH_SATUAN")), "#,0.00")
        txtSatuan = IIf(IsNull(RS.Fields("KODE_SATUAN")), "", RS.Fields("KODE_SATUAN"))
        lblSatuan.Caption = IIf(IsNull(RS.Fields("URAIAN_SATUAN")), "", RS.Fields("URAIAN_SATUAN"))
        txtHargaSatuan = IIf(IsNull(RS.Fields("HARGA_SATUAN")), 0, RS.Fields("HARGA_SATUAN"))
        
        txtHargaDetil = Format(IIf(IsNull(RS.Fields("HARGA_INVOICE")), 0, RS.Fields("HARGA_INVOICE")), "#,0.00")
        txtFreight = Format(IIf(IsNull(RS.Fields("FREIGHT")), 0, RS.Fields("FREIGHT")), "#,0.00")
        txtAsuransi = Format(IIf(IsNull(RS.Fields("ASURANSI")), 0, RS.Fields("ASURANSI")), "#,0.00")
        txtHargaCIF = Format(IIf(IsNull(RS.Fields("CIF")), 0, RS.Fields("CIF")), "#,0.00")
        txtCIF = Format(IIf(IsNull(RS.Fields("CIF_RUPIAH")), 0, RS.Fields("CIF_RUPIAH")), "#,0.0000")
        
        txtJumlah = Format(IIf(IsNull(RS.Fields("JUMLAH_KEMASAN")), 0, RS.Fields("JUMLAH_KEMASAN")), "#,0.00")
        txtJenis = IIf(IsNull(RS.Fields("KODE_KEMASAN")), "", RS.Fields("KODE_KEMASAN"))
        txtNetto = Format(IIf(IsNull(RS.Fields("NETTO")), 0, RS.Fields("NETTO")), "#,0.00")
        txtNegara = IIf(IsNull(RS.Fields("KODE_NEGARA_ASAL")), "", RS.Fields("KODE_NEGARA_ASAL"))
        lblNegara.Caption = IIf(IsNull(RS.Fields("Nama_Negara")), "", RS.Fields("Nama_Negara"))
        
        txtFasilitas = IIf(IsNull(RS.Fields("KODE_FASILITAS_DOKUMEN")), "", RS.Fields("KODE_FASILITAS_DOKUMEN"))
        lblFasilitas.Caption = IIf(IsNull(RS.Fields("URAIAN_FASILITAS")), "", RS.Fields("URAIAN_FASILITAS"))
        txtSkema = IIf(IsNull(RS.Fields("KODE_SKEMA_TARIF")), "", RS.Fields("KODE_SKEMA_TARIF"))
        lblSkema.Caption = IIf(IsNull(RS.Fields("URAIAN_SKEMA")), "", RS.Fields("URAIAN_SKEMA"))
                
        cboJenisPungutan = IIf(IsNull(RS.Fields("URAIAN_PUNGUTAN1")), "", RS.Fields("URAIAN_PUNGUTAN1"))
        cboKeterangan1 = IIf(IsNull(RS.Fields("URAIAN_TARIF1")), "", RS.Fields("URAIAN_TARIF1"))
        cboKeterangan2 = IIf(IsNull(RS.Fields("URAIAN_FASILITAS1")), "", RS.Fields("URAIAN_FASILITAS1"))
        txtTarifPersen1 = IIf(IsNull(RS.Fields("TARIF1")), 0, RS.Fields("TARIF1"))
        txtTarifPersen2 = IIf(IsNull(RS.Fields("TARIF_FASILITAS1")), 0, RS.Fields("TARIF_FASILITAS1"))
        
        txtJumlahSpesifik = IIf(IsNull(RS.Fields("JUMLAH_SATUAN_TARIF")), 0, RS.Fields("JUMLAH_SATUAN_TARIF"))
        txtSatuanTarif = IIf(IsNull(RS.Fields("KODE_SATUAN_TARIF")), "", RS.Fields("KODE_SATUAN_TARIF"))
        
        txtPPN = IIf(IsNull(RS.Fields("TARIF2")), 0, RS.Fields("TARIF2"))
        cboKeterangan3 = IIf(IsNull(RS.Fields("URAIAN_FASILITAS2")), "", RS.Fields("URAIAN_FASILITAS2"))
        txtTarifPersen3 = IIf(IsNull(RS.Fields("TARIF_FASILITAS2")), 0, RS.Fields("TARIF_FASILITAS2"))
        
        txtPPNBm = IIf(IsNull(RS.Fields("TARIF3")), 0, RS.Fields("TARIF3"))
        cboKeterangan4 = IIf(IsNull(RS.Fields("URAIAN_FASILITAS3")), "", RS.Fields("URAIAN_FASILITAS3"))
        txtTarifPersen4 = IIf(IsNull(RS.Fields("TARIF_FASILITAS3")), 0, RS.Fields("TARIF_FASILITAS3"))
        
        txtPPh = IIf(IsNull(RS.Fields("TARIF4")), 0, RS.Fields("TARIF4"))
        cboKeterangan5 = IIf(IsNull(RS.Fields("URAIAN_FASILITAS4")), "", RS.Fields("URAIAN_FASILITAS4"))
        txtTarifPersen5 = IIf(IsNull(RS.Fields("TARIF_FASILITAS4")), 0, RS.Fields("TARIF_FASILITAS4"))
        
        cboKomoditi = IIf(IsNull(RS.Fields("URAIAN_KOMODITI")), "", RS.Fields("URAIAN_KOMODITI"))
        cboJenisTarif = IIf(IsNull(RS.Fields("URAIAN_TARIF_CUKAI")), "", RS.Fields("URAIAN_TARIF_CUKAI"))
        cboKeterangan = IIf(IsNull(RS.Fields("URAIAN_FASILITAS_CUKAI")), "", RS.Fields("URAIAN_FASILITAS_CUKAI"))
        txtTarif = IIf(IsNull(RS.Fields("TARIF_CUKAI")), 0, RS.Fields("TARIF_CUKAI"))
        txtPersenCukai = IIf(IsNull(RS.Fields("TARIF_FASILITAS_CUKAI")), 0, RS.Fields("TARIF_FASILITAS_CUKAI"))
        txtSatuanCukai = IIf(IsNull(RS.Fields("KODE_SATUAN_CUKAI")), "", RS.Fields("KODE_SATUAN_CUKAI"))
        txtJumlahCukai = IIf(IsNull(RS.Fields("JUMLAH_SATUAN_CUKAI")), 0, RS.Fields("JUMLAH_SATUAN_CUKAI"))
        
    End If
    
    cekLoad = False
End Sub

Private Sub up_GridLoad()
Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer

    up_GridHeader
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailBarangDokumenPerBarang_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan.Text)
    cmd.Parameters.append cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri.Text)
    
    Set RS = cmd.Execute
     
    With Grid
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1

            .TextMatrix(li_Row, colJenisDokumen) = Trim(RS!Uraian_Dokumen)
            .TextMatrix(li_Row, colNomorDokumen) = Trim(RS!Nomor_Dokumen)
            .TextMatrix(li_Row, colTanggal) = Format(RS!Tanggal_Dokumen, "dd MMM yyyy")

            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End With
    
End Sub

Private Function uf_ValidateInput() As Boolean
    If txtKodeBarang = "" Then
        txtKodeBarang.SetFocus
        LblerrMsg = "Please Input Kode Barang!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtUraianBarang = "" Then
        txtUraianBarang.SetFocus
        LblerrMsg = "Please Input Uraian Barang!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtKategoriBarang.Text = "" Then
        txtKategoriBarang.SetFocus
        LblerrMsg = "Please Input Kategori Barang!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtNomorHS.Text = "" Then
        txtNomorHS.SetFocus
        LblerrMsg = "Please Input Nomor HS!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtMerk.Text = "" Then
        txtMerk.SetFocus
        LblerrMsg = "Please Input Merk!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtTipe.Text = "" Then
        txtTipe.SetFocus
        LblerrMsg = "Please Input Tipe!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtUkuran.Text = "" Then
        txtUkuran.SetFocus
        LblerrMsg = "Please Input Ukuran!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtSpfLain.Text = "" Then
        txtSpfLain.SetFocus
        LblerrMsg = "Please Input Spesifikasi Lain!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtTotalDetilFOB.Text = "" Or CDbl(txtTotalDetilFOB.Text) = 0 Then
        txtTotalDetilFOB.SetFocus
        LblerrMsg = "Please Input Total / Detil (FOB)!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtJumlahSatuan.Text = "" Or CDbl(txtJumlahSatuan.Text) = 0 Then
        txtJumlahSatuan.SetFocus
        LblerrMsg = "Please Input Jumlah Satuan!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtSatuan.Text = "" Then
        txtSatuan.SetFocus
        LblerrMsg = "Please Input Satuan!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtJumlah.Text = "" Or CDbl(txtJumlah.Text) = 0 Then
        txtJumlah.SetFocus
        LblerrMsg = "Please Input Jumlah Kemasan!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtJenis.Text = "" Then
        txtJenis.SetFocus
        LblerrMsg = "Please Input Jenis Kemasan!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtNetto.Text = "" Or CDbl(txtNetto.Text) = 0 Then
        txtNetto.SetFocus
        LblerrMsg = "Please Input Netto!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtNegara.Text = "" Then
        txtNegara.SetFocus
        LblerrMsg = "Please Input Negara Asal!"
        uf_ValidateInput = False
        Exit Function
    ElseIf cboJenisPungutan.Text = "" Then
        cboJenisPungutan.SetFocus
        LblerrMsg = "Please Input Tarif BM!"
        uf_ValidateInput = False
        Exit Function
    ElseIf cboKeterangan1.Text = "" Then
        cboKeterangan1.SetFocus
        LblerrMsg = "Please Input Tarif BM Advolorum!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtTarifPersen1.Text = "" Then
        txtTarifPersen1.SetFocus
        LblerrMsg = "Please Input Persentase Tarif BM Advolorum!"
        uf_ValidateInput = False
        Exit Function
    ElseIf cboJenisTarif <> "" Then
        If Trim(Split(cboKeterangan1, "-")(0)) = "2" Then
            If txtJumlahSpesifik.Text = "" Then
                txtJumlahSpesifik.SetFocus
                LblerrMsg = "Please Jumlah Satuan!"
                uf_ValidateInput = False
                Exit Function
            ElseIf txtSatuanTarif.Text = "" Then
                txtSatuanTarif.SetFocus
                LblerrMsg = "Please Input Satuan!"
                uf_ValidateInput = False
                Exit Function
            End If
        End If
    ElseIf cboJenisTarif <> "" Then
        If Trim(Split(cboKeterangan1, "-")(0)) = "2" Then
            If txtJumlahCukai.Text = "" Then
                txtJumlahCukai.SetFocus
                LblerrMsg = "Please Jumlah Cukai!"
                uf_ValidateInput = False
                Exit Function
            ElseIf txtSatuanCukai.Text = "" Then
                txtSatuanCukai.SetFocus
                LblerrMsg = "Please Input Satuan Cukai!"
                uf_ValidateInput = False
                Exit Function
            End If
        End If
    End If
    
    uf_ValidateInput = True
End Function

Private Sub up_Delete()
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
    
Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC23DetailBarang_Del"

cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append cmd.CreateParameter("NoSeri", adInteger, adParamInput, , txtNoSeri)

cmd.Execute

LblerrMsg.Caption = DisplayMsg(1201)

DoEvents

Unload Me
End Sub

Private Sub up_SaveData()
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
    cmd.CommandText = "sp_BC23DetailBarang_Upd"
    
'    If txtID = "" Then txtID = "0"
    
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
    Set prm6 = cmd.CreateParameter("Kategori", adVarChar, adParamInput, 50, txtKategoriBarang)
    cmd.Parameters.append prm6
    Set prm7 = cmd.CreateParameter("Tipe", adVarChar, adParamInput, 255, txtTipe)
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("Ukuran", adVarChar, adParamInput, 255, txtUkuran)
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("SpesifikasiLain", adVarChar, adParamInput, 255, txtSpfLain)
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("Merk", adVarChar, adParamInput, 255, txtMerk)
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("FOB", adDecimal, adParamInput, , txtTotalDetilFOB)
    prm11.Precision = 38
    prm11.NumericScale = 2
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("Diskon", adDecimal, adParamInput, , txtBTDiskon)
    prm12.Precision = 38
    prm12.NumericScale = 2
    cmd.Parameters.append prm12
    Set prm13 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , txtJumlahSatuan)
    prm13.Precision = 38
    prm13.NumericScale = 4
    cmd.Parameters.append prm13
    Set prm14 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 10, txtSatuan)
    cmd.Parameters.append prm14
    Set prm15 = cmd.CreateParameter("HargaSatuan", adDecimal, adParamInput, , txtHargaSatuan)
    prm15.Precision = 38
    prm15.NumericScale = 4
    cmd.Parameters.append prm15
    Set prm16 = cmd.CreateParameter("HargaInvoice", adDecimal, adParamInput, , txtHargaDetil)
    prm16.Precision = 38
    prm16.NumericScale = 2
    cmd.Parameters.append prm16
    Set prm17 = cmd.CreateParameter("Freight", adDecimal, adParamInput, , txtFreight)
    prm17.Precision = 38
    prm17.NumericScale = 2
    cmd.Parameters.append prm17
    Set prm18 = cmd.CreateParameter("Asuransi", adDecimal, adParamInput, , txtAsuransi)
    prm18.Precision = 38
    prm18.NumericScale = 2
    cmd.Parameters.append prm18
    Set prm19 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , txtHargaCIF)
    prm19.Precision = 38
    prm19.NumericScale = 2
    cmd.Parameters.append prm19
    Set prm20 = cmd.CreateParameter("CIFRupiah", adDecimal, adParamInput, , txtCIF)
    prm20.Precision = 38
    prm20.NumericScale = 2
    cmd.Parameters.append prm20
    Set prm21 = cmd.CreateParameter("JumlahKemasan", adInteger, adParamInput, 10, txtJumlah)
    cmd.Parameters.append prm21
    Set prm22 = cmd.CreateParameter("KodeKemasan", adVarChar, adParamInput, 50, txtJenis)
    cmd.Parameters.append prm22
    Set prm23 = cmd.CreateParameter("Netto", adDecimal, adParamInput, , txtNetto)
    prm23.Precision = 38
    prm23.NumericScale = 2
    cmd.Parameters.append prm23
    Set prm24 = cmd.CreateParameter("KodeNegara", adVarChar, adParamInput, 15, txtNegara)
    cmd.Parameters.append prm24
    Set prm25 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 15, txtFasilitas)
    cmd.Parameters.append prm25
    Set prm26 = cmd.CreateParameter("KodeSkema", adVarChar, adParamInput, 15, txtSkema)
    cmd.Parameters.append prm26
    
    cmd.Execute Y
    
    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC23DetailBarang_Ins"
        
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
        Set prm6 = cmd.CreateParameter("Kategori", adVarChar, adParamInput, 50, txtKategoriBarang)
        cmd.Parameters.append prm6
        Set prm7 = cmd.CreateParameter("Tipe", adVarChar, adParamInput, 255, txtTipe)
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Ukuran", adVarChar, adParamInput, 255, txtUkuran)
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("SpesifikasiLain", adVarChar, adParamInput, 255, txtSpfLain)
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("Merk", adVarChar, adParamInput, 255, txtMerk)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("FOB", adDecimal, adParamInput, , txtTotalDetilFOB)
        prm11.Precision = 38
        prm11.NumericScale = 2
        cmd.Parameters.append prm11
        Set prm12 = cmd.CreateParameter("Diskon", adDecimal, adParamInput, , txtBTDiskon)
        prm12.Precision = 38
        prm12.NumericScale = 2
        cmd.Parameters.append prm12
        Set prm13 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , txtJumlahSatuan)
        prm13.Precision = 38
        prm13.NumericScale = 4
        cmd.Parameters.append prm13
        Set prm14 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 10, txtSatuan)
        cmd.Parameters.append prm14
        Set prm15 = cmd.CreateParameter("HargaSatuan", adDecimal, adParamInput, , txtHargaSatuan)
        prm15.Precision = 38
        prm15.NumericScale = 4
        cmd.Parameters.append prm15
        Set prm16 = cmd.CreateParameter("HargaInvoice", adDecimal, adParamInput, , txtHargaDetil)
        prm16.Precision = 38
        prm16.NumericScale = 2
        cmd.Parameters.append prm16
        Set prm17 = cmd.CreateParameter("Freight", adDecimal, adParamInput, , txtFreight)
        prm17.Precision = 38
        prm17.NumericScale = 2
        cmd.Parameters.append prm17
        Set prm18 = cmd.CreateParameter("Asuransi", adDecimal, adParamInput, , txtAsuransi)
        prm18.Precision = 38
        prm18.NumericScale = 2
        cmd.Parameters.append prm18
        Set prm19 = cmd.CreateParameter("CIF", adDecimal, adParamInput, , txtHargaCIF)
        prm19.Precision = 38
        prm19.NumericScale = 2
        cmd.Parameters.append prm19
        Set prm20 = cmd.CreateParameter("CIFRupiah", adDecimal, adParamInput, , txtCIF)
        prm20.Precision = 38
        prm20.NumericScale = 2
        cmd.Parameters.append prm20
        Set prm21 = cmd.CreateParameter("JumlahKemasan", adInteger, adParamInput, 10, txtJumlah)
        cmd.Parameters.append prm21
        Set prm22 = cmd.CreateParameter("KodeKemasan", adVarChar, adParamInput, 50, txtJenis)
        cmd.Parameters.append prm22
        Set prm23 = cmd.CreateParameter("Netto", adDecimal, adParamInput, , txtNetto)
        prm23.Precision = 38
        prm23.NumericScale = 2
        cmd.Parameters.append prm23
        Set prm24 = cmd.CreateParameter("KodeNegara", adVarChar, adParamInput, 15, txtNegara)
        cmd.Parameters.append prm24
        Set prm25 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 15, txtFasilitas)
        cmd.Parameters.append prm25
        Set prm26 = cmd.CreateParameter("KodeSkema", adVarChar, adParamInput, 15, txtSkema)
        cmd.Parameters.append prm26
    
        cmd.Execute
    End If

    
    Dim R As Integer
    
    '####################### BM ########################
    
    'DELETE BM/BMKITE
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailBeaMasukTambahan_Del"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, Trim(Split(cboJenisPungutan, "-")(0)))
    cmd.Parameters.append prm3
        
    cmd.Execute
    
    'INSERT BM/BMKITE
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailBeaMasukTambahan_Ins"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 20, Trim(Split(cboJenisPungutan, "-")(0)))
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Trim(Split(cboKeterangan2, "-")(0)))
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Trim(Split(cboKeterangan1, "-")(0)))
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
    prm6.Precision = 38
    prm6.NumericScale = 2
    cmd.Parameters.append prm6
    Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , CDbl(txtTarifPersen1 / 100) * CDbl(txtCIF))
    prm7.Precision = 38
    prm7.NumericScale = 2
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtTarifPersen1))
    prm8.Precision = 38
    prm8.NumericScale = 2
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtTarifPersen2))
    prm9.Precision = 38
    prm9.NumericScale = 2
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, txtSatuanTarif)
    cmd.Parameters.append prm10
    If txtJumlahSpesifik = "" Then txtJumlahSpesifik = 0
    Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , CDbl(txtJumlahSpesifik))
    prm11.Precision = 38
    prm11.NumericScale = 4
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, Null)
    cmd.Parameters.append prm12
        
    cmd.Execute
    
'    'SIMPAN DATA KE TABLE TPB PUNGUTAN
'    Set cmd = New ADODB.Command
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandTimeout = 0
'    cmd.ActiveConnection = Db
'    cmd.CommandText = "sp_BC23DetailPungutan_Upd"
'
'
'    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
'    cmd.Parameters.append prm1
'    Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
'    cmd.Parameters.append prm2
'    Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 20, Trim(Split(cboJenisPungutan, "-")(0)))
'    cmd.Parameters.append prm3
    
    
    
    '####################### BM ########################
    

    '####################### PPN ########################

    Dim TarifPPN As Double
    Dim TarifBM As Double

    R = 0

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailBeaMasukTambahan_Upd"

    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "PPN")
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Trim(Split(cboKeterangan3, "-")(0)))
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
    prm6.Precision = 38
    prm6.NumericScale = 2
    cmd.Parameters.append prm6

    If Trim(Split(cboKeterangan1, "-")(0)) = "2" Then
        TarifBM = CDbl(txtCIF) + (CDbl(txtTarifPersen1) * CDbl(txtJumlahSpesifik))
    Else
        TarifBM = CDbl(txtCIF) + (CDbl(txtTarifPersen1 / 100) * CDbl(txtCIF))
    End If
    TarifPPN = TarifBM * (CDbl(txtPPN) / 100)

    Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , TarifPPN)
    prm7.Precision = 38
    prm7.NumericScale = 2
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtPPN))
    prm8.Precision = 38
    prm8.NumericScale = 2
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtTarifPersen3))
    prm9.Precision = 38
    prm9.NumericScale = 2
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
    prm11.Precision = 38
    prm11.NumericScale = 4
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, Null)
    cmd.Parameters.append prm12

    cmd.Execute R

    If R = 0 Then

        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC23DetailBeaMasukTambahan_Ins"

        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "PPN")
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Trim(Split(cboKeterangan3, "-")(0)))
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
        prm6.Precision = 38
        prm6.NumericScale = 2
        cmd.Parameters.append prm6
        
        If Trim(Split(cboKeterangan1, "-")(0)) = "2" Then
            TarifBM = CDbl(txtCIF) + (CDbl(txtTarifPersen1) * CDbl(txtJumlahSpesifik))
        Else
            TarifBM = CDbl(txtCIF) + (CDbl(txtTarifPersen1 / 100) * CDbl(txtCIF))
        End If
        
        TarifPPN = TarifBM * (CDbl(txtPPN) / 100)


        Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , TarifPPN)
        prm7.Precision = 38
        prm7.NumericScale = 2
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtPPN))
        prm8.Precision = 38
        prm8.NumericScale = 2
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtTarifPersen3))
        prm9.Precision = 38
        prm9.NumericScale = 2
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm10

        Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
        prm11.Precision = 38
        prm11.NumericScale = 4
        cmd.Parameters.append prm11
        Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, Null)
        cmd.Parameters.append prm12

        cmd.Execute
    End If
'
'    '####################### PPN ########################


'    '####################### PPNBM ########################
    Dim TarifPPNBM As Double

    R = 0

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailBeaMasukTambahan_Upd"

    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "PPNBM")
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Trim(Split(cboKeterangan3, "-")(0)))
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
    prm6.Precision = 38
    prm6.NumericScale = 2
    cmd.Parameters.append prm6


    If Trim(Split(cboKeterangan1, "-")(0)) = "2" Then
        TarifBM = CDbl(txtCIF) + (CDbl(txtTarifPersen1) * CDbl(txtJumlahSpesifik))
    Else
        TarifBM = CDbl(txtCIF) + (CDbl(txtTarifPersen1 / 100) * CDbl(txtCIF))
    End If
    If txtPPNBm = "" Then txtPPNBm = 0
    TarifPPNBM = TarifBM * (CDbl(txtPPNBm) / 100)


    Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , TarifPPNBM)
    prm7.Precision = 38
    prm7.NumericScale = 2
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtPPNBm))
    prm8.Precision = 38
    prm8.NumericScale = 2
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtTarifPersen3))
    prm9.Precision = 38
    prm9.NumericScale = 2
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
    prm11.Precision = 38
    prm11.NumericScale = 4
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, Null)
    cmd.Parameters.append prm12

    cmd.Execute R

    If R = 0 Then

        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC23DetailBeaMasukTambahan_Ins"

        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "PPNBM")
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Trim(Split(cboKeterangan3, "-")(0)))
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
        prm6.Precision = 38
        prm6.NumericScale = 2
        cmd.Parameters.append prm6


        If Trim(Split(cboKeterangan1, "-")(0)) = "2" Then
            TarifBM = CDbl(txtCIF) + (CDbl(txtTarifPersen1) * CDbl(txtJumlahSpesifik))
        Else
            TarifBM = CDbl(txtCIF) + (CDbl(txtTarifPersen1 / 100) * CDbl(txtCIF))
        End If
        TarifPPNBM = TarifBM * (CDbl(txtPPNBm) / 100)


        Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , TarifPPNBM)
        prm7.Precision = 38
        prm7.NumericScale = 2
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtPPNBm))
        prm8.Precision = 38
        prm8.NumericScale = 2
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtTarifPersen3))
        prm9.Precision = 38
        prm9.NumericScale = 2
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm10

        Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
        prm11.Precision = 38
        prm11.NumericScale = 4
        cmd.Parameters.append prm11
        Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, Null)
        cmd.Parameters.append prm12

        cmd.Execute
    End If
'
'    '####################### PPNBM ########################
'
'  '####################### PPH ########################
    Dim TarifPPH As Double

    R = 0

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailBeaMasukTambahan_Upd"

    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "PPH")
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Trim(Split(cboKeterangan3, "-")(0)))
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm5
    Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
    prm6.Precision = 38
    prm6.NumericScale = 2
    cmd.Parameters.append prm6


    If Trim(Split(cboKeterangan1, "-")(0)) = "2" Then
        TarifBM = CDbl(txtCIF) + (CDbl(txtTarifPersen1) * CDbl(txtJumlahSpesifik))
    Else
        TarifBM = CDbl(txtCIF) + (CDbl(txtTarifPersen1 / 100) * CDbl(txtCIF))
    End If
    TarifPPH = TarifBM * (CDbl(txtPPh) / 100)


    Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , TarifPPH)
    prm7.Precision = 38
    prm7.NumericScale = 2
    cmd.Parameters.append prm7
    Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtPPh))
    prm8.Precision = 38
    prm8.NumericScale = 2
    cmd.Parameters.append prm8
    Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtTarifPersen3))
    prm9.Precision = 38
    prm9.NumericScale = 2
    cmd.Parameters.append prm9
    Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
    cmd.Parameters.append prm10
    Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
    prm11.Precision = 38
    prm11.NumericScale = 4
    cmd.Parameters.append prm11
    Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, Null)
    cmd.Parameters.append prm12
    
    cmd.Execute R

    If R = 0 Then

        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC23DetailBeaMasukTambahan_Ins"

        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "PPH")
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Trim(Split(cboKeterangan3, "-")(0)))
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
        prm6.Precision = 38
        prm6.NumericScale = 2
        cmd.Parameters.append prm6


        If Trim(Split(cboKeterangan1, "-")(0)) = "2" Then
            TarifBM = CDbl(txtCIF) + (CDbl(txtTarifPersen1) * CDbl(txtJumlahSpesifik))
        Else
            TarifBM = CDbl(txtCIF) + (CDbl(txtTarifPersen1 / 100) * CDbl(txtCIF))
        End If
        TarifPPH = TarifBM * (CDbl(txtPPh) / 100)


        Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , TarifPPH)
        prm7.Precision = 38
        prm7.NumericScale = 2
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtPPh))
        prm8.Precision = 38
        prm8.NumericScale = 2
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtTarifPersen3))
        prm9.Precision = 38
        prm9.NumericScale = 2
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
        prm11.Precision = 38
        prm11.NumericScale = 4
        cmd.Parameters.append prm11
        Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, Null)
        cmd.Parameters.append prm12

        cmd.Execute
    End If
    
    '####################### PPH ########################
    
    'DELETE KOMODITI
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailBeaMasukTambahan_Del"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "CUKAI")
    cmd.Parameters.append prm3
        
    cmd.Execute
    
    'INSERT KOMODITI
    
'  '####################### KOMODITI ########################
    Dim TarifKomoditi As Double
    
    If cboKeterangan.ListIndex > -1 Then
    
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC23DetailBeaMasukTambahanCukai_Ins"
    
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "CUKAI")
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Trim(Split(cboKeterangan, "-")(0)))
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Left(cboJenisTarif, 1))
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
        prm6.Precision = 38
        prm6.NumericScale = 2
        cmd.Parameters.append prm6
    
        
        If Trim(Split(cboJenisTarif, "-")(0)) = "2" Then
            TarifKomoditi = (CDbl(txtTarif) * CDbl(txtJumlahCukai))
        Else
            TarifKomoditi = CDbl(txtTarif / 100) * CDbl(txtCIF)
        End If
        
    '    TarifKomoditi = TarifBM * (CDbl(txtPPh) / 100)
    
        
        Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , TarifKomoditi)
        prm7.Precision = 38
        prm7.NumericScale = 2
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtTarif))
        prm8.Precision = 38
        prm8.NumericScale = 2
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , CDbl(txtPersenCukai))
        prm9.Precision = 38
        prm9.NumericScale = 2
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, txtSatuanCukai)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , txtJumlahCukai)
        prm11.Precision = 38
        prm11.NumericScale = 4
        cmd.Parameters.append prm11
        Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, Null)
        cmd.Parameters.append prm12
        Set prm13 = cmd.CreateParameter("KodeKomoditi", adVarChar, adParamInput, 1, Left(cboKomoditi, 1))
        cmd.Parameters.append prm13
        
        cmd.Execute
    
    End If
    
    
    
    
    '####################### KOMODITI ########################
    
    '####################### TOTAL PUNGUTAN ########################
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailTotalPungutan_Ins"
    
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
        
    cmd.Execute
    
    '####################### TOTAL PUNGUTAN ########################
    
    up_GridLoad
    
    If Y = 0 Then
        txtKodeBarang.Enabled = False
        LblerrMsg = DisplayMsg(1000)
    Else
        LblerrMsg = DisplayMsg(1101)
    End If

    cekSubmit = True
End Sub


Private Sub cboJenisTarif_Change()
    If Trim(Split(cboJenisTarif, "-")(0)) = "2" Then
        Label1(37).Visible = True
        txtJumlahCukai.Visible = True
        cboKeterangan.Left = 2520
        txtPersenCukai.Left = 4560
        Label1(42).Caption = "/"
        txtSatuanCukai.Visible = True
        Label1(43).Left = 5400
    Else
        Label1(37).Visible = False
        txtJumlahCukai.Visible = False
        cboKeterangan.Left = 1680
        txtPersenCukai.Left = 3720
        Label1(42).Caption = "%"
        txtSatuanCukai.Visible = False
        Label1(43).Left = 4560
    End If
End Sub

Private Sub cboKeterangan1_Change()

'    If cekLoad = False Then
        If Trim(Split(cboKeterangan1, "-")(0)) = "2" Then
            txtJumlahSpesifik.Visible = True
            txtSatuanTarif.Visible = True
            Label1(39).Visible = True
            Label1(24).Caption = "/"
        Else
            txtJumlahSpesifik.Visible = False
            txtSatuanTarif.Visible = False
            Label1(39).Visible = False
            Label1(24).Caption = "%"
        End If
        
        If cekLoad = False Then
            txtTarifPersen2.Text = "100.00"
            txtTarifPersen5.Text = "100.00"
            txtPPh.Text = "2.5"
            cboKeterangan2.ListIndex = 0
            cboKeterangan5.ListIndex = 0
        End If

End Sub

Private Sub cboKeterangan3_Change()
If cekLoad = False Then txtTarifPersen3 = "100.00"
End Sub

Private Sub cboKeterangan4_Change()
If cekLoad = False Then txtTarifPersen4 = "100.00"
End Sub

Private Sub cmdBrowseDokumen_Click()
    If cekSubmit = False Then
        LblerrMsg.Caption = "Please save the data first!"
        Exit Sub
    End If
    frmBC23BrowseBarangDokumen.txtNoPengajuan = Replace(txtNoPengajuan, "-", "")
    frmBC23BrowseBarangDokumen.txtNoSeri = txtNoSeri
    frmBC23BrowseBarangDokumen.txtKodeBarang = txtKodeBarang
    frmBC23BrowseBarangDokumen.Show 1
End Sub

Private Sub cmdBrowseTarif_Click()
If cekSubmit = False Then
    LblerrMsg.Caption = "Please save the data first!"
    Exit Sub
End If

If CekData = False Then
    frmBC23BrowseBeaMasukTambahan.txtNoPengajuan = txtNoPengajuan
    frmBC23BrowseBeaMasukTambahan.txtNoSeri = txtNoSeri
    frmBC23BrowseBeaMasukTambahan.txtNomorHS = txtNomorHS
    frmBC23BrowseBeaMasukTambahan.txtUraianBarang = txtUraianBarang
    frmBC23BrowseBeaMasukTambahan.txtCIF = txtHargaCIF
    frmBC23BrowseBeaMasukTambahan.txtCIFRupiah = txtCIF
    frmBC23BrowseBeaMasukTambahan.cboJenisTarifBM = cboKeterangan1
    If CDbl(txtTarifPersen1) > 0 Then
    frmBC23BrowseBeaMasukTambahan.txtBesarTarif = Format(CDbl(txtTarifPersen1), "#,0.00")
    End If
    frmBC23BrowseBeaMasukTambahan.cboTarifFasilitas = cboKeterangan2
    If CDbl(txtTarifPersen2) > 0 Then
    frmBC23BrowseBeaMasukTambahan.txtTarifFasilitas = Format(CDbl(txtTarifPersen2), "#,0.00")
    End If
    If CDbl(txtTarifPersen1) > 0 Then
        frmBC23BrowseBeaMasukTambahan.txtBMFasilitas = Format((CDbl(txtTarifPersen1) / 100) * CDbl(txtCIF), "#,0.00")
    End If
    
    frmBC23BrowseBeaMasukTambahan.Show 1
Else
    frmBC23BrowseBeaMasukTambahan.txtNoPengajuan = txtNoPengajuan
    frmBC23BrowseBeaMasukTambahan.txtNoSeri = txtNoSeri
    frmBC23BrowseBeaMasukTambahan.up_LoadData txtNoPengajuan, txtNoSeri
    frmBC23BrowseBeaMasukTambahan.Show 1
End If
End Sub

Private Sub cmdCancel_Click()
    If cekSubmit = False Then
        up_Clear
    Else
        'up_LoadDataBarang
    End If
    
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure want to delete?", vbYesNo + vbExclamation, "Delete") = vbYes Then
        up_Delete
    End If
End Sub

Private Sub CmdSubmit_Click()
    If uf_ValidateInput = False Then Exit Sub
    up_SaveData
End Sub

Private Sub cmdTarifFasilitas_Click()
frmBC23BrowseBarangTarifFasilitas.txtNoPengajuan = txtNoPengajuan
frmBC23BrowseBarangTarifFasilitas.txNoSeri = txtNoSeri
frmBC23BrowseBarangTarifFasilitas.Show 1
End Sub

Private Sub Form_Activate()
    up_GridLoad
End Sub

Private Sub Form_Load()
    
    up_FillComboGeneral cboJenisPungutan, "Bea_Cukai_Pungutan Where ID In (1,9) ", "KODE_PUNGUTAN", "URAIAN_PUNGUTAN", 60, 200
    cboJenisPungutan.ListIndex = 0
    
    up_FillComboGeneral cboKeterangan1, "Bea_Cukai_Jenis_Tarif Where ID In (5,6) ", "KODE_JENIS_TARIF", "URAIAN_JENIS_TARIF", 60, 150
    
    up_FillComboGeneral cboKeterangan2, "Bea_Cukai_Tarif_Fasilitas Where KODE_FASILITAS = 2 ", "KODE_FASILITAS", "URAIAN_Fasilitas", 60, 150
    up_FillComboGeneral cboKeterangan3, "Bea_Cukai_Tarif_Fasilitas Where KODE_FASILITAS = 5 ", "KODE_FASILITAS", "URAIAN_Fasilitas", 60, 150
    up_FillComboGeneral cboKeterangan4, "Bea_Cukai_Tarif_Fasilitas Where KODE_FASILITAS = 5 ", "KODE_FASILITAS", "URAIAN_Fasilitas", 60, 150
    up_FillComboGeneral cboKeterangan5, "Bea_Cukai_Tarif_Fasilitas Where KODE_FASILITAS = 5 ", "KODE_FASILITAS", "URAIAN_Fasilitas", 60, 150
    
    up_FillComboGeneral cboKomoditi, "Bea_Cukai_Komoditi Where ID In (5,6,7) ", "KODE_KOMODITI", "URAIAN_KOMODITI", 60, 150
    up_FillComboGeneral cboJenisTarif, "Bea_Cukai_Jenis_Tarif Where ID In (5,6,7) ", "KODE_JENIS_TARIF", "URAIAN_JENIS_TARIF", 60, 150
    
    up_FillComboGeneral cboKeterangan, "Bea_Cukai_Tarif_Fasilitas Where KODE_FASILITAS = 4 ", "KODE_FASILITAS", "URAIAN_Fasilitas", 60, 150
    
    up_Clear
    
    
End Sub

Private Sub txtAsuransi_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtBTDiskon_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtCIF_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtFasilitas_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFasilitas_LostFocus()
    up_LoadFasilitas txtFasilitas
End Sub

Private Sub txtFreight_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtHargaCIF_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtHargaDetil_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtHargaSatuan_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtJenis_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtJenis_LostFocus()
    up_LoadJenisKemasan txtJenis
End Sub

Private Sub txtJumlah_GotFocus()
    If txtJumlah = "" Then txtJumlah = 0
    txtJumlah = CDbl(txtJumlah)
End Sub

Private Sub txtJumlah_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtJumlah_LostFocus()
    txtJumlah = Format(txtJumlah, "#,0.00")

End Sub

Private Sub txtJumlahCukai_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtJumlahSatuan_GotFocus()
    If txtJumlahSatuan = "" Then txtJumlahSatuan = 0
    txtJumlahSatuan = CDbl(txtJumlahSatuan)
End Sub

Private Sub txtJumlahSatuan_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtJumlahSatuan_LostFocus()
    If txtJumlahSatuan = "" Then txtJumlahSatuan = 0
    txtJumlahSatuan = Format(txtJumlahSatuan, "#,0.0000")
    
    If CDbl(txtJumlahSatuan) = 0 Then
        txtHargaSatuan = "0.00"
    Else
        txtHargaSatuan = Format(txtTotalDetilFOB / txtJumlahSatuan, "#,0.00")
    End If
    
End Sub

Private Sub txtJumlahSpesifik_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtJumlahSpesifik_LostFocus()
txtJumlahSpesifik = Format(CDbl(txtJumlahSpesifik), "#,0.00")
End Sub

Private Sub txtKategoriBarang_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtKategoriBarang_LostFocus()
    up_LoadKategoriBarang txtKategoriBarang
End Sub

Private Sub txtKodeBarang_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMerk_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNegara_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNegara_LostFocus()
up_LoadNegara txtNegara
End Sub

Private Sub txtNetto_GotFocus()
  If txtNetto = "" Then txtNetto = 0
    txtNetto = CDbl(txtNetto)
End Sub

Private Sub txtNetto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtNetto_LostFocus()
    txtNetto = Format(txtNetto, "#,0.00")
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



Private Sub txtPersenCukai_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtPPh_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtPPn_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtPPN_LostFocus()
If txtPPN = "" Then txtPPN = 0
txtPPN = Format(CDbl(txtPPN), "#,0.00")
End Sub

Private Sub txtPPNBm_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtPPNBm_LostFocus()
txtPPNBm = Format(CDbl(txtPPNBm), "#,0.00")
End Sub

Private Sub txtSatuan_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSatuan_LostFocus()
up_LoadSatuan txtSatuan
End Sub

Private Sub txtSatuanCukai_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSkema_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSkema_LostFocus()
up_LoadSkemaTarif txtSkema
End Sub

Private Sub txtSpfLain_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTarif_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtTarifPersen1_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtTarifPersen2_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtTarifPersen2_LostFocus()
    txtTarifPersen2 = Format(CDbl(txtTarifPersen2), "#,0.00")
End Sub

Private Sub txtTarifPersen3_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtTarifPersen4_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtTarifPersen5_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtTipe_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTotalDetilFOB_GotFocus()
    txtTotalDetilFOB = CDbl(txtTotalDetilFOB)
'    txtCIF = CDbl(txtCIFFix) * CDbl(txtHargaCIF)
End Sub

Private Sub txtTotalDetilFOB_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtTotalDetilFOB_LostFocus()
    If txtTotalDetilFOB = "" Then txtTotalDetilFOB = 0
    
    txtTotalDetilFOB = Format(txtTotalDetilFOB, "#,0.00")
    
    txtHargaDetil = txtTotalDetilFOB
    If CDbl(txtFreightFix) = 0 Then
        txtFreight = "0.00"
    Else
        txtFreight = Format(Round(txtTotalDetilFOB / txtFreightFix, 2), "#,0.00")
    End If
    
    
    txtHargaCIF = Format(CDbl(txtHargaDetil) + CDbl(txtFreight) + CDbl(txtAsuransi), "#,0.00")
    txtCIF = Format(CDbl(txtHargaCIF) * CDbl(txtCIFFix), "#,0.00")
    
End Sub

Private Sub txtUkuran_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUraianBarang_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
