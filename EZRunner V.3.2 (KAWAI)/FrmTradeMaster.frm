VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTradeMaster 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Trade Master"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTradeMaster.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10770
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8715
      Left            =   600
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   720
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   15372
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16637923
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmTradeMaster.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label23"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label20"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Grid"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtlocation(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtlocation(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Others"
      TabPicture(1)   =   "FrmTradeMaster.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).Control(2)=   "Frame7"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame7 
         Caption         =   "BC Information"
         Height          =   1875
         Left            =   -74760
         TabIndex        =   162
         Top             =   5200
         Width           =   5715
         Begin VB.TextBox txtNoIzin 
            Height          =   285
            Index           =   0
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   169
            Top             =   1000
            Width           =   3615
         End
         Begin VB.TextBox txtKodeKPPBC 
            Height          =   285
            Index           =   10
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   166
            Top             =   650
            Width           =   3615
         End
         Begin MSComCtl2.DTPicker dtNoIzin 
            Height          =   330
            Left            =   1800
            TabIndex        =   170
            Top             =   1380
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
            Format          =   134086659
            CurrentDate     =   37860
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   25
            Left            =   1440
            TabIndex        =   174
            Top             =   1420
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   24
            Left            =   1440
            TabIndex        =   173
            Top             =   990
            Width           =   105
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal "
            Height          =   195
            Index           =   11
            Left            =   150
            TabIndex        =   172
            Top             =   1450
            Width           =   1125
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "No Izin"
            Height          =   195
            Index           =   10
            Left            =   150
            TabIndex        =   171
            Top             =   1000
            Width           =   1125
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode KPPBC"
            Height          =   195
            Index           =   9
            Left            =   150
            TabIndex        =   168
            Top             =   650
            Width           =   1125
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   23
            Left            =   1440
            TabIndex        =   167
            Top             =   650
            Width           =   105
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "BC Type"
            Height          =   195
            Index           =   8
            Left            =   150
            TabIndex        =   165
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   22
            Left            =   1440
            TabIndex        =   164
            Top             =   240
            Width           =   105
         End
         Begin MSForms.ComboBox cboBCType 
            Height          =   315
            Left            =   1800
            TabIndex        =   163
            Top             =   240
            Width           =   1215
            VariousPropertyBits=   746604571
            MaxLength       =   10
            DisplayStyle    =   3
            Size            =   "2143;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8115
         Left            =   150
         TabIndex        =   84
         Top             =   360
         Width           =   6915
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   4500
            MaxLength       =   15
            TabIndex        =   160
            Text            =   "AAAAAA"
            Top             =   180
            Width           =   2115
         End
         Begin VB.TextBox txtinsurance 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3030
            TabIndex        =   152
            TabStop         =   0   'False
            Top             =   6870
            Width           =   3075
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   18
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   21
            Text            =   "AAA"
            Top             =   6480
            Width           =   525
         End
         Begin VB.CheckBox cekinvto 
            Caption         =   "Check1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   6090
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   2040
            MaxLength       =   20
            TabIndex        =   5
            Text            =   "AAAAAAAAAAAAAAAAAAAA"
            Top             =   1710
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   5
            Left            =   2040
            MaxLength       =   200
            TabIndex        =   7
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   2460
            Width           =   4665
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   6
            Left            =   2040
            MaxLength       =   200
            TabIndex        =   8
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   2820
            Width           =   4665
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   7
            Left            =   2040
            MaxLength       =   100
            TabIndex        =   9
            Text            =   "AAAAAAAAAAAAAAAAAAAA"
            Top             =   3225
            Width           =   4665
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   8
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   10
            Text            =   "AAAAAAAAAA"
            Top             =   3585
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   14
            Text            =   "AAAAAAAAAAAAAAA"
            Top             =   4830
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   11
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   17
            Text            =   "AA"
            Top             =   5640
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   12
            Left            =   5100
            MaxLength       =   2
            TabIndex        =   19
            Text            =   "AA"
            Top             =   5640
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   4
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   6
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   2100
            Width           =   4665
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   4410
            MaxLength       =   50
            TabIndex        =   15
            Text            =   "AAAAAAAAAAAAAAA"
            Top             =   4830
            Width           =   2085
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   0
            Text            =   "AAAAAA"
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   2
            Left            =   2040
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1320
            Width           =   4665
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   20
            Left            =   2040
            MaxLength       =   100
            TabIndex        =   11
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   3990
            Width           =   4665
         End
         Begin VB.TextBox txtname 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3420
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   6090
            Width           =   3075
         End
         Begin VB.TextBox txtregion 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3750
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   5250
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "SAP Code"
            Height          =   195
            Index           =   1
            Left            =   3480
            TabIndex        =   161
            Top             =   225
            Width           =   1485
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "NG Cls"
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   159
            Top             =   7710
            Width           =   885
         End
         Begin MSForms.ComboBox cboNGCls 
            Height          =   315
            Left            =   2040
            TabIndex        =   158
            Top             =   7680
            Width           =   855
            VariousPropertyBits=   746604571
            MaxLength       =   6
            DisplayStyle    =   7
            Size            =   "1508;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   21
            Left            =   1830
            TabIndex        =   157
            Top             =   7680
            Width           =   105
         End
         Begin VB.Label Label54 
            Caption         =   ":"
            Height          =   315
            Left            =   1830
            TabIndex        =   154
            Top             =   990
            Width           =   135
         End
         Begin MSForms.ComboBox cmbbox_warehouse 
            Height          =   330
            Left            =   2040
            TabIndex        =   3
            Top             =   930
            Width           =   1635
            VariousPropertyBits=   746604571
            MaxLength       =   10
            DisplayStyle    =   3
            Size            =   "2884;582"
            ShowDropButtonWhen=   2
            Value           =   "cmbbox_warehouse"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Warehouse Code"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   153
            Top             =   990
            Width           =   1785
         End
         Begin VB.Line Line8 
            X1              =   3030
            X2              =   6090
            Y1              =   7140
            Y2              =   7140
         End
         Begin VB.Label lblEPTE 
            Height          =   285
            Left            =   5700
            TabIndex        =   151
            Top             =   4050
            Visible         =   0   'False
            Width           =   1065
         End
         Begin MSForms.ComboBox cboEPTE 
            Height          =   315
            Left            =   5100
            TabIndex        =   13
            Top             =   4410
            Width           =   945
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "1667;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label45 
            Caption         =   ":"
            Height          =   255
            Index           =   1
            Left            =   4920
            TabIndex        =   150
            Top             =   4440
            Width           =   105
         End
         Begin MSForms.Label Label56 
            Height          =   255
            Left            =   3900
            TabIndex        =   149
            Top             =   4440
            Width           =   495
            Caption         =   "EPTE"
            Size            =   "873;450"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   2
            Left            =   1830
            TabIndex        =   130
            Top             =   6840
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   1
            Left            =   1830
            TabIndex        =   129
            Top             =   7320
            Width           =   105
         End
         Begin VB.Label Label55 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   128
            Top             =   6090
            Width           =   105
         End
         Begin VB.Label Label53 
            Caption         =   ":"
            Height          =   285
            Left            =   1830
            TabIndex        =   127
            Top             =   6480
            Width           =   105
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   ":"
            Height          =   195
            Left            =   4920
            TabIndex        =   126
            Top             =   5685
            Width           =   75
         End
         Begin VB.Label Label51 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   125
            Top             =   5655
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   0
            Left            =   1830
            TabIndex        =   124
            Top             =   5250
            Width           =   105
         End
         Begin VB.Label Label49 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   123
            Top             =   4860
            Width           =   105
         End
         Begin VB.Label Label48 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   122
            Top             =   4020
            Width           =   105
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Days"
            Height          =   195
            Index           =   6
            Left            =   2640
            TabIndex        =   121
            Top             =   6540
            Width           =   585
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Pay "
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   120
            Top             =   6510
            Width           =   1485
         End
         Begin VB.Label lblpocls 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   2970
            TabIndex        =   119
            Top             =   6870
            Visible         =   0   'False
            Width           =   645
         End
         Begin MSForms.ComboBox cbopocls 
            Height          =   315
            Left            =   2040
            TabIndex        =   22
            Top             =   7290
            Width           =   855
            VariousPropertyBits=   746604571
            MaxLength       =   6
            DisplayStyle    =   7
            Size            =   "1508;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "PO Cls"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   118
            Top             =   7350
            Width           =   1485
         End
         Begin VB.Label lblcountry 
            BackStyle       =   0  'Transparent
            Height          =   315
            Left            =   3720
            TabIndex        =   117
            Top             =   3960
            Visible         =   0   'False
            Width           =   555
         End
         Begin MSForms.ComboBox cbotradecode 
            Height          =   315
            Left            =   2040
            TabIndex        =   20
            Top             =   6060
            Width           =   1215
            VariousPropertyBits=   746604571
            MaxLength       =   10
            DisplayStyle    =   3
            Size            =   "2143;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cbocountry 
            Height          =   315
            Left            =   2040
            TabIndex        =   12
            Top             =   4410
            Width           =   1575
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "2778;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice to  "
            Height          =   195
            Index           =   3
            Left            =   390
            TabIndex        =   116
            Top             =   6120
            Width           =   1200
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Country Cls"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   115
            Top             =   4440
            Width           =   1485
         End
         Begin MSForms.ComboBox cbotradecls 
            Height          =   315
            Left            =   2040
            TabIndex        =   1
            Top             =   555
            Width           =   855
            VariousPropertyBits=   746604571
            MaxLength       =   1
            DisplayStyle    =   7
            Size            =   "1508;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Line Line1 
            X1              =   3030
            X2              =   4500
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label lbltradecls 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   3030
            TabIndex        =   114
            Top             =   585
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   113
            Top             =   2115
            Width           =   1485
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            Height          =   195
            Left            =   4080
            TabIndex        =   112
            Top             =   4890
            Width           =   255
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Pay Day  "
            Height          =   195
            Index           =   0
            Left            =   3900
            TabIndex        =   111
            Top             =   5685
            Width           =   1485
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Closing Day "
            Height          =   195
            Left            =   120
            TabIndex        =   110
            Top             =   5685
            Width           =   1485
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone / Fax "
            Height          =   195
            Left            =   120
            TabIndex        =   109
            Top             =   4860
            Width           =   1485
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Postal Code"
            Height          =   195
            Left            =   120
            TabIndex        =   108
            Top             =   3630
            Width           =   1485
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "City "
            Height          =   195
            Left            =   120
            TabIndex        =   107
            Top             =   3255
            Width           =   1485
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 2 "
            Height          =   195
            Left            =   120
            TabIndex        =   106
            Top             =   2865
            Width           =   1485
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 1"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   2490
            Width           =   1485
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Trade ABBR "
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   104
            Top             =   1755
            Width           =   1485
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Trade Cls"
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   600
            Width           =   1485
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Trade Code"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   102
            Top             =   225
            Width           =   1485
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Trade Name"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   101
            Top             =   1350
            Width           =   1485
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Insurance Covered  :"
            Height          =   195
            Left            =   90
            TabIndex        =   100
            Top             =   6870
            Width           =   1665
         End
         Begin MSForms.Label Label22 
            Height          =   195
            Left            =   5070
            TabIndex        =   99
            Top             =   600
            Width           =   1575
            Caption         =   "Affiliate Company"
            Size            =   "2778;344"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox cekaff 
            Height          =   255
            Left            =   4710
            TabIndex        =   2
            Top             =   570
            Width           =   255
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "450;450"
            Value           =   "0"
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Region "
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   5250
            Width           =   1485
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            Height          =   315
            Left            =   120
            TabIndex        =   97
            Top             =   4020
            Width           =   1485
         End
         Begin MSForms.ComboBox cboregion 
            Height          =   315
            Left            =   2040
            TabIndex        =   16
            Top             =   5220
            Width           =   1575
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   7
            Size            =   "2778;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Line Line2 
            X1              =   3420
            X2              =   6480
            Y1              =   6390
            Y2              =   6390
         End
         Begin MSForms.ComboBox cboinsurance 
            Height          =   315
            Left            =   2040
            TabIndex        =   23
            Top             =   6840
            Width           =   855
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "1508;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Line Line5 
            X1              =   3750
            X2              =   6450
            Y1              =   5520
            Y2              =   5520
         End
         Begin VB.Label lbl90 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   96
            Top             =   240
            Width           =   105
         End
         Begin VB.Label Label39 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   95
            Top             =   570
            Width           =   105
         End
         Begin VB.Label Label40 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   94
            Top             =   1350
            Width           =   105
         End
         Begin VB.Label Label41 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   93
            Top             =   1725
            Width           =   105
         End
         Begin VB.Label Label42 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   92
            Top             =   2100
            Width           =   105
         End
         Begin VB.Label Label43 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   91
            Top             =   2460
            Width           =   105
         End
         Begin VB.Label Label44 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   90
            Top             =   2850
            Width           =   105
         End
         Begin VB.Label Label45 
            Caption         =   ":"
            Height          =   255
            Index           =   0
            Left            =   1830
            TabIndex        =   89
            Top             =   3240
            Width           =   105
         End
         Begin VB.Label Label46 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   88
            Top             =   3600
            Width           =   105
         End
         Begin VB.Label Label47 
            Caption         =   ":"
            Height          =   255
            Left            =   1830
            TabIndex        =   87
            Top             =   4410
            Width           =   105
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "PO Case Mark"
         Height          =   2325
         Left            =   -74790
         TabIndex        =   70
         Top             =   510
         Width           =   5685
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   4
            Left            =   1140
            MaxLength       =   25
            TabIndex        =   39
            Top             =   1830
            Width           =   4275
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   3
            Left            =   1140
            MaxLength       =   25
            TabIndex        =   38
            Top             =   1470
            Width           =   4275
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   2
            Left            =   1140
            MaxLength       =   25
            TabIndex        =   37
            Top             =   1080
            Width           =   4275
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   1
            Left            =   1140
            MaxLength       =   25
            TabIndex        =   36
            Top             =   660
            Width           =   4275
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   0
            Left            =   1140
            MaxLength       =   25
            TabIndex        =   35
            Top             =   270
            Width           =   4275
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   15
            Left            =   900
            TabIndex        =   143
            Top             =   1845
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   14
            Left            =   900
            TabIndex        =   142
            Top             =   1485
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   13
            Left            =   900
            TabIndex        =   141
            Top             =   1095
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   12
            Left            =   900
            TabIndex        =   140
            Top             =   675
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   11
            Left            =   900
            TabIndex        =   139
            Top             =   285
            Width           =   105
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Line5 "
            Height          =   285
            Left            =   150
            TabIndex        =   80
            Top             =   1830
            Width           =   615
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Line4"
            Height          =   255
            Left            =   150
            TabIndex        =   79
            Top             =   1485
            Width           =   615
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Line3"
            Height          =   285
            Left            =   150
            TabIndex        =   78
            Top             =   1080
            Width           =   585
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Line2"
            Height          =   255
            Left            =   150
            TabIndex        =   77
            Top             =   675
            Width           =   705
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Line1"
            Height          =   285
            Left            =   150
            TabIndex        =   76
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "PO Marking"
         Height          =   2235
         Left            =   -74790
         TabIndex        =   69
         Top             =   2880
         Width           =   5715
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   9
            Left            =   1140
            MaxLength       =   25
            TabIndex        =   44
            Top             =   1770
            Width           =   4275
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   8
            Left            =   1140
            MaxLength       =   25
            TabIndex        =   43
            Top             =   1410
            Width           =   4275
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   7
            Left            =   1140
            MaxLength       =   25
            TabIndex        =   42
            Top             =   1050
            Width           =   4275
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   6
            Left            =   1140
            MaxLength       =   25
            TabIndex        =   41
            Top             =   660
            Width           =   4275
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   5
            Left            =   1140
            MaxLength       =   25
            TabIndex        =   40
            Top             =   270
            Width           =   4275
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   20
            Left            =   900
            TabIndex        =   148
            Top             =   1785
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   19
            Left            =   900
            TabIndex        =   147
            Top             =   1425
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   18
            Left            =   900
            TabIndex        =   146
            Top             =   1065
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   17
            Left            =   900
            TabIndex        =   145
            Top             =   675
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   16
            Left            =   900
            TabIndex        =   144
            Top             =   285
            Width           =   105
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Line5"
            Height          =   285
            Left            =   150
            TabIndex        =   75
            Top             =   1770
            Width           =   675
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Line4"
            Height          =   255
            Left            =   150
            TabIndex        =   74
            Top             =   1425
            Width           =   645
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Line3"
            Height          =   285
            Left            =   150
            TabIndex        =   73
            Top             =   1050
            Width           =   525
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Line2 "
            Height          =   255
            Left            =   150
            TabIndex        =   72
            Top             =   675
            Width           =   525
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Line1"
            Height          =   285
            Left            =   150
            TabIndex        =   71
            Top             =   270
            Width           =   585
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1665
         Left            =   7230
         TabIndex        =   65
         Top             =   4080
         Width           =   6765
         Begin VB.TextBox txtPOPayment 
            Height          =   345
            Left            =   6015
            MaxLength       =   3
            TabIndex        =   32
            Text            =   "AAA"
            Top             =   1080
            Width           =   480
         End
         Begin VB.TextBox txttransport 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2955
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   1110
            Width           =   1425
         End
         Begin VB.TextBox txtprice 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2955
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   210
            Width           =   3375
         End
         Begin VB.TextBox txtpayment 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2955
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   660
            Width           =   3345
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Day"
            Height          =   195
            Index           =   1
            Left            =   4650
            TabIndex        =   156
            Top             =   1155
            Width           =   1155
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   ":"
            Height          =   195
            Left            =   5865
            TabIndex        =   155
            Top             =   1155
            Width           =   75
         End
         Begin VB.Line Line7 
            X1              =   2955
            X2              =   4395
            Y1              =   1410
            Y2              =   1410
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   10
            Left            =   1530
            TabIndex        =   138
            Top             =   1170
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   9
            Left            =   1530
            TabIndex        =   137
            Top             =   690
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   8
            Left            =   1530
            TabIndex        =   136
            Top             =   270
            Width           =   105
         End
         Begin MSForms.ComboBox cbotransport 
            Height          =   345
            Left            =   1740
            TabIndex        =   31
            Top             =   1080
            Width           =   1005
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   7
            Size            =   "1773;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cbopayment 
            Height          =   345
            Left            =   1740
            TabIndex        =   30
            Top             =   630
            Width           =   1005
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   7
            Size            =   "1773;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboprice 
            Height          =   345
            Left            =   1740
            TabIndex        =   29
            Top             =   210
            Width           =   1005
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   7
            Size            =   "1773;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Line Line4 
            X1              =   2955
            X2              =   6510
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line3 
            X1              =   2955
            X2              =   6510
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Transportation"
            Height          =   285
            Left            =   150
            TabIndex        =   68
            Top             =   1170
            Width           =   1305
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Terms"
            Height          =   285
            Left            =   150
            TabIndex        =   67
            Top             =   690
            Width           =   1350
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Price Condition"
            Height          =   285
            Left            =   150
            TabIndex        =   66
            Top             =   240
            Width           =   1350
         End
      End
      Begin VB.TextBox txtlocation 
         Height          =   285
         Index           =   0
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   33
         Text            =   "AAAAAA"
         Top             =   8055
         Width           =   975
      End
      Begin VB.TextBox txtlocation 
         Height          =   285
         Index           =   1
         Left            =   8400
         MaxLength       =   25
         TabIndex        =   34
         Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
         Top             =   8040
         Width           =   3255
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   7230
         TabIndex        =   55
         Top             =   390
         Width           =   6765
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   17
            Left            =   1710
            MaxLength       =   50
            TabIndex        =   175
            Text            =   "AAAAAAAAAAAAAAAAAAAA"
            Top             =   3270
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   1710
            MaxLength       =   20
            TabIndex        =   24
            Text            =   "AAAAAAAAAAAAAAAAAAAA"
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Height          =   645
            Index           =   14
            Left            =   1710
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Text            =   "FrmTradeMaster.frx":0E7A
            Top             =   705
            Width           =   3855
         End
         Begin VB.TextBox Text1 
            Height          =   855
            Index           =   15
            Left            =   1710
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Text            =   "FrmTradeMaster.frx":0E99
            Top             =   1485
            Width           =   3855
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   16
            Left            =   1710
            MaxLength       =   50
            TabIndex        =   27
            Text            =   "AAAAAAAAAAAAAAAAAAAA"
            Top             =   2490
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   19
            Left            =   1710
            MaxLength       =   50
            TabIndex        =   28
            Text            =   "AAAAAAAAAAAAAAAAAAAA"
            Top             =   2910
            Width           =   2535
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   26
            Left            =   1500
            TabIndex        =   177
            Top             =   3270
            Width           =   105
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "NITKU"
            Height          =   195
            Left            =   120
            TabIndex        =   176
            Top             =   3270
            Width           =   1215
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   7
            Left            =   1500
            TabIndex        =   135
            Top             =   2910
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   6
            Left            =   1500
            TabIndex        =   134
            Top             =   2490
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   5
            Left            =   1500
            TabIndex        =   133
            Top             =   1500
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   4
            Left            =   1500
            TabIndex        =   132
            Top             =   750
            Width           =   105
         End
         Begin VB.Label Label50 
            Caption         =   ":"
            Height          =   255
            Index           =   3
            Left            =   1500
            TabIndex        =   131
            Top             =   240
            Width           =   105
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP No"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP Name"
            Height          =   435
            Left            =   120
            TabIndex        =   59
            Top             =   750
            Width           =   1215
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP Address"
            Height          =   555
            Left            =   120
            TabIndex        =   58
            Top             =   1500
            Width           =   1290
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP City"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   2490
            Width           =   1215
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "NPPKP No "
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   2910
            Width           =   1215
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   1125
         Left            =   7230
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   6240
         Width           =   6780
         _cx             =   11959
         _cy             =   1984
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
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
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
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   195
         Left            =   7320
         TabIndex        =   63
         Top             =   7635
         Width           =   1095
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Location Name"
         Height          =   195
         Left            =   8400
         TabIndex        =   64
         Top             =   7635
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00A6D2FF&
         BackStyle       =   1  'Opaque
         Height          =   330
         Index           =   1
         Left            =   7200
         Top             =   7560
         Width           =   6795
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Delivery Place"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7230
         TabIndex        =   62
         Top             =   5880
         Width           =   6765
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0080C0FF&
         Height          =   675
         Index           =   0
         Left            =   7200
         Top             =   7800
         Width           =   6795
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   1
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   10200
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clear Grid"
      Height          =   375
      Index           =   2
      Left            =   11100
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   10230
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Index           =   0
      Left            =   12330
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   10230
      Width           =   1140
   End
   Begin VB.Frame Frame3 
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
      Height          =   555
      Left            =   600
      TabIndex        =   51
      Top             =   9480
      Width           =   14175
      Begin VB.Label LblErrMsg 
         Alignment       =   2  'Center
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
         Left            =   105
         TabIndex        =   53
         Top             =   195
         Width           =   13770
      End
   End
   Begin VB.CommandButton cmdinquiry 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Inquiry"
      Height          =   375
      Left            =   1770
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   10230
      Width           =   1140
   End
   Begin VB.CommandButton cmdsubmenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   540
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   10230
      Width           =   1140
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13020
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Master"
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
      Left            =   390
      TabIndex        =   50
      Top             =   240
      Width           =   14205
   End
End
Attribute VB_Name = "FrmTradeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS As New ADODB.Recordset
Dim rsGrid As New ADODB.Recordset

'Dim rsRegion As New ADODB.Recordset
Dim ubahgrid As Boolean
'
Dim sql As String, sqlGrid As String ', sqlregion As String
Public ubahtrade As Boolean

Dim bteColSelect As Byte
Dim bteColCode As Byte
Dim bteColName As Byte

Function headerGrid()
    
    bteColSelect = 0
    bteColCode = 1
    bteColName = 2
    
    With Grid
        .clear
        .Rows = 1
        .ColS = 3
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColCode) = "Code"
        .TextMatrix(0, bteColName) = "Location Name"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColCode) = 1000
        .ColWidth(bteColName) = 3000
        
        .Cell(flexcpAlignment, 0, 0, 0, bteColName) = flexAlignCenterCenter
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColCode) = flexAlignCenterCenter
        .ColAlignment(bteColName) = flexAlignLeftCenter
        
        .EditMaxLength = 1
        
    End With
    
End Function

Sub Kosong()
  Text1(0).Text = ""
  Text1(1).Text = ""
  Text1(0).BackColor = vbWhite
  Text1(0).DataChanged = False
  cbotradecls.ListIndex = -1
  cbotradecls.BackColor = vbWhite
  lbltradecls.Caption = ""
  lbltradecls.DataChanged = False
  cmbbox_warehouse = ""
    
  For i = 2 To 20
   If i <> 18 Then
    Text1(i).Text = ""
    Text1(i).BackColor = vbWhite
    Text1(i).DataChanged = False
   End If
  Next i
  Text1(0).Enabled = True
  
  For i = 0 To 9
    Text2(i).Text = ""
    Text2(i).BackColor = vbWhite
    Text2(i).DataChanged = False
  Next i
  
  LblErrMsg.Caption = ""
  
  txtPOPayment.Text = ""
  
  cbocountry.ListIndex = 0
  lblcountry.Caption = ""
  cbocountry.BackColor = vbWhite
  lblcountry.DataChanged = False
  
  cbopocls.ListIndex = 1
  lblpocls.Caption = ""
  cbopocls.BackColor = vbWhite
  lblpocls.DataChanged = False
    
  cboEPTE.ListIndex = 1
  lblEPTE.Caption = ""
  cboEPTE.BackColor = vbWhite
  lblEPTE.DataChanged = False
    
  cboinsurance.ListIndex = -1
  txtinsurance.Text = ""
  cboinsurance.BackColor = vbWhite
  txtinsurance.DataChanged = False
      
  cboregion.ListIndex = -1
  cboregion.BackColor = vbWhite
  txtregion.Text = ""
  txtregion.DataChanged = False
        
  cbopayment.ListIndex = -1
  cbopayment.BackColor = vbWhite
  txtpayment.Text = ""
  txtpayment.DataChanged = False
        
  cboprice.ListIndex = -1
  cboprice.BackColor = vbWhite
  txtprice.Text = ""
  txtprice.DataChanged = False
        
  cbotransport.ListIndex = -1
  cbotransport.BackColor = vbWhite
  txttransport.Text = ""
  txttransport.DataChanged = False
  
  cboBCType.ListIndex = -1
  cboBCType.BackColor = vbWhite
  cboBCType.Text = ""
  cboBCType.DataChanged = False
        
  cekaff.Value = 0
  cekinvto.Value = 0
  cbotradecode.ListIndex = -1
  txtName.Text = ""
  cbotradecode.BackColor = vbWhite
  txtName.DataChanged = False
  Call cekinvto_Click
  txtKodeKPPBC(10).Text = ""
  txtNoIzin(0).Text = ""
  dtNoIzin.Value = Now
  cboNGCls.ListIndex = -1
    
  ubahtrade = False
  kosonggrid
  headerGrid
End Sub

Sub kosonggrid()
    kosongColGrid
    ubahgrid = False
    txtlocation(0).Text = ""
    txtlocation(0).locked = False
    txtlocation(0).BackColor = vbWhite
    txtlocation(0).DataChanged = False
    txtlocation(1).Text = ""
    txtlocation(1).BackColor = vbWhite
    txtlocation(1).DataChanged = False
End Sub


Private Sub cmbbox_warehouse_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'If KeyCode = vbKeyDelete Then lbl_warehouse.Caption = ""
'If KeyCode = vbKeyBack Then lbl_warehouse.Caption = ""

If KeyCode = 13 Then
    If Trim(cmbbox_warehouse) = "" Then Exit Sub
    If cmbbox_warehouse.ListCount > 0 Then
        For i = 0 To cmbbox_warehouse.ListCount - 1
            If UCase(Trim(cmbbox_warehouse.Text)) = UCase(Trim(cmbbox_warehouse.List(i, 0))) Then
                  cmbbox_warehouse = Trim(cmbbox_warehouse.List(i, 0))
                  'lbl_warehouse.Caption = Trim(cmbbox_warehouse.List(i, 1))
                  LblErrMsg.Caption = ""
                Exit For
            Else
              'lbl_warehouse.Caption = ""
              LblErrMsg.Caption = DisplayMsg("4023")
              cmbbox_warehouse.SetFocus
            End If
        Next
    Else
        If Trim(cmbbox_warehouse.Text) <> "" Then
            'lbl_warehouse.Caption = ""
            LblErrMsg.Caption = DisplayMsg("4023")
            cmbbox_warehouse.SetFocus
        End If
    End If
End If

End Sub

Private Sub cmbbox_warehouse_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
If KeyAscii = 34 Then KeyAscii = 0 ' kutip 2
End Sub
Sub adtocombo()

'=====================Setting Combo WareHouse=====================
cmbbox_warehouse.clear
cmbbox_warehouse.columnCount = 2
cmbbox_warehouse.TextColumn = 1
i = 0

Dim rs_warehouse_master As New ADODB.Recordset
If rs_warehouse_master.State <> adStateClosed Then rs_warehouse_master.Close
rs_warehouse_master.CursorLocation = adUseClient
rs_warehouse_master.Open "select * from warehouse_master", Db, adOpenKeyset, adLockOptimistic
If rs_warehouse_master.EOF = False Or rs_warehouse_master.BOF = False Then
    rs_warehouse_master.MoveFirst
    While rs_warehouse_master.EOF = False
        cmbbox_warehouse.AddItem ""
        cmbbox_warehouse.List(i, 0) = rs_warehouse_master!wh_code
        cmbbox_warehouse.List(i, 1) = rs_warehouse_master!WH_Name
        rs_warehouse_master.MoveNext
        i = i + 1
    Wend
    
    cmbbox_warehouse.ColumnWidths = "50 pt; 120 pt"
    cmbbox_warehouse.ListWidth = 170
    cmbbox_warehouse.ListIndex = 0
    cmbbox_warehouse.Text = ""
End If
If rs_warehouse_master.State = 1 Then rs_warehouse_master.Close
'==============================================================

    With cbotradecls
        .clear
        .columnCount = 2
        .ColumnWidths = "25pt; 100pt"
        .ListWidth = 125
        .ListRows = 5
        
        .AddItem ""
        .List(0, 0) = 0
        .List(0, 1) = "Own Information"
        .AddItem ""
        .List(1, 0) = 1
        .List(1, 1) = "Internal"
        .AddItem ""
        .List(2, 0) = 2
        .List(2, 1) = "External"
        .AddItem ""
        .List(3, 0) = 3
        .List(3, 1) = "Sub Contractor"
        .AddItem ""
        .List(4, 0) = 4
        .List(4, 1) = "Consignee"
        
        'Add for KAWAI
        .AddItem ""
        .List(5, 0) = 5
        .List(5, 1) = "Forwarder"

    End With
    
    With cbocountry
        .clear
        .AddItem "Domestic"
        .AddItem "Overseas"
    End With
    
    With cboEPTE
        .clear
        .columnCount = 2
        .ColumnWidths = "30pt;0pt"
        .ListWidth = 30
        
        .AddItem
        .List(0, 0) = "Yes"
        .List(0, 1) = 1
        .AddItem
        .List(1, 0) = "No"
        .List(1, 1) = 0
    End With
    
    With cboNGCls
        .clear
        .columnCount = 2
        .ColumnWidths = "30pt;0pt"
        .ListWidth = 30
        
        .AddItem
        .List(0, 0) = "Yes"
        .List(0, 1) = 1
        .AddItem
        .List(1, 0) = "No"
        .List(1, 1) = 0
    End With
    
    cboNGCls.Enabled = False
'    With cboinsurance
'        .clear
'        .ColumnCount = 2
'        .ColumnWidths = "30pt;0pt"
'        .ListWidth = 30
'
'        .AddItem
'        .List(0, 0) = "Yes"
'        .List(0, 1) = 1
'        .AddItem
'        .List(1, 0) = "No"
'        .List(1, 1) = 0
'    End With
            
    '*******************************************
    'If Not (rsRegion.BOF And rsRegion.EOF) Then
     'Dim i As Integer
     'i = 0
     'With cboregion
      '.clear
      '.ColumnCount = 2
      '.ColumnWidths = "50pt;150pt"
      '.ListWidth = 200
      '.ListRows = 5
        
      'Do While Not rsRegion.EOF
       '.AddItem
       '.List(i, 0) = Trim(rsRegion!Region_Cls)
       '.List(i, 1) = Trim(rsRegion!Description)
       'i = i + 1
       'rsRegion.MoveNext
      'Loop
     'End With
    'End If
    '********************************************
    Call up_FillCombo(cboregion, "Region_Cls")
    cboregion.ListWidth = 150
    cboregion.ColumnWidths = "20 pt;130 pt"
    
    Call up_FillCombo(cbopayment, "PaymentTerm_Cls")
    cbopayment.ListWidth = 150
    cbopayment.ColumnWidths = "20 pt;130 pt"
    
    Call up_FillCombo(cboprice, "PriceCondition_Cls")
    cboprice.ListWidth = 150
    cboprice.ColumnWidths = "20 pt;130 pt"
    
    Call up_FillCombo(cbotransport, "Transportation_Cls")
    cbotransport.ListWidth = 150
    cbotransport.ColumnWidths = "20 pt;130 pt"
        
    Call up_FillCombo(cboinsurance, "Insurance_Cls")
    cboinsurance.ListWidth = 150
    cboinsurance.ColumnWidths = "20 pt;130 pt"
        
    RS.filter = ""
    RS.Requery
    If Not (RS.BOF And RS.EOF) Then
    'Dim i As Integer
    i = 0
    With cbotradecode
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;300pt"
        .ListWidth = 350
        .ListRows = 15
        
        Do While Not RS.EOF
            .AddItem ""
            .List(i, 0) = Trim(RS!Trade_Code)
            .List(i, 1) = Trim(RS!trade_name & "")
            i = i + 1
            RS.MoveNext
        Loop
    End With
    End If
    
    With cbopocls
        .clear
        
        .columnCount = 2
        .ColumnWidths = "30pt;0pt"
        .ListWidth = 30
        
        .AddItem
        .List(0, 0) = "Yes"
        .List(0, 1) = 1
        .AddItem
        .List(1, 0) = "No"
        .List(1, 1) = 0
    End With
    
    
End Sub

Sub BrowseGrid()
    rsGrid.filter = ""
    rsGrid.Requery
    rsGrid.filter = "Trade_Code='" & Text1(0).Text & "' "
    i = 1
    With Grid
    Do While Not rsGrid.EOF
        .Rows = .Rows + 1
        .TextMatrix(i, bteColCode) = Trim(rsGrid("Location_Code"))
        .TextMatrix(i, bteColName) = Trim(rsGrid("Location_Name"))
        .Cell(flexcpBackColor, i, bteColSelect) = &HFFFFFF
        rsGrid.MoveNext
        i = i + 1
    Loop
    End With
End Sub

Sub Browse()
  RS.filter = "Trade_Code='" & Text1(0).Text & "' "
  If RS.BOF And RS.EOF Then
    Text1(0).Enabled = True
    cbotradecls.SetFocus
  Else
    Text1(0).Text = Trim(RS(0))
    Text1(0).BackColor = vbWhite
    cbotradecls.Text = Trim(RS(1))
    cbotradecls.BackColor = vbWhite
    lbltradecls.DataChanged = False
   
   cmbbox_warehouse = IIf(IsNull(RS!Subcon_WH_Code), "", Trim(RS!Subcon_WH_Code))
    Text1(1).Text = IIf(IsNull(RS!SAP_Code), "", Trim(RS!SAP_Code))
    Text1(2).Text = IIf(IsNull(RS!trade_name), "", Trim(RS!trade_name))
    Text1(3).Text = IIf(IsNull(RS!trade_abbr), "", Trim(RS!trade_abbr))
    Text1(4).Text = IIf(IsNull(RS!contact_person), "", Trim(RS!contact_person))
    Text1(5).Text = IIf(IsNull(RS!address1), "", Trim(RS!address1))
    Text1(6).Text = IIf(IsNull(RS!address2), "", Trim(RS!address2))
    Text1(7).Text = IIf(IsNull(RS!City), "", Trim(RS!City))
    Text1(8).Text = IIf(IsNull(RS!postal_code), "", Trim(RS!postal_code))
    Text1(9).Text = IIf(IsNull(RS!Telephone), "", Trim(RS!Telephone))
    Text1(10).Text = IIf(IsNull(RS!fax), "", Trim(RS!fax))
    Text1(11).Text = IIf(IsNull(RS!Closing_Day), "", Trim(RS!Closing_Day))
    Text1(12).Text = IIf(IsNull(RS!Pay_Day), "", Trim(RS!Pay_Day))
    Text1(13).Text = IIf(IsNull(RS!NPWP_No), "", Trim(RS!NPWP_No))
    Text1(14).Text = IIf(IsNull(RS!NPWP_Name), "", Trim(RS!NPWP_Name))
    Text1(15).Text = IIf(IsNull(RS!NPWP_Address), "", Trim(RS!NPWP_Address))
    Text1(16).Text = IIf(IsNull(RS!NPWP_City), "", Trim(RS!NPWP_City))
    Text1(18).Text = IIf(IsNull(RS!InvoicePay_Days), "", Trim(RS!InvoicePay_Days))
    Text1(19).Text = IIf(IsNull(RS!NPPKP_No), "", Trim(RS!NPPKP_No))
    Text1(20).Text = IIf(IsNull(RS!Country), "", Trim(RS!Country))
    Text1(17).Text = IIf(IsNull(RS!NITKU), "", Trim(RS!NITKU))
    
    Text2(0).Text = IIf(IsNull(RS!POCaseMark1), "", Trim(RS!POCaseMark1))
    Text2(1).Text = IIf(IsNull(RS!POCaseMark2), "", Trim(RS!POCaseMark2))
    Text2(2).Text = IIf(IsNull(RS!POCaseMark3), "", Trim(RS!POCaseMark3))
    Text2(3).Text = IIf(IsNull(RS!POCaseMark4), "", Trim(RS!POCaseMark4))
    Text2(4).Text = IIf(IsNull(RS!POCaseMark5), "", Trim(RS!POCaseMark5))
    Text2(5).Text = IIf(IsNull(RS!POMarking1), "", Trim(RS!POMarking1))
    Text2(6).Text = IIf(IsNull(RS!POMarking2), "", Trim(RS!POMarking2))
    Text2(7).Text = IIf(IsNull(RS!POMarking3), "", Trim(RS!POMarking3))
    Text2(8).Text = IIf(IsNull(RS!POMarking4), "", Trim(RS!POMarking4))
    Text2(9).Text = IIf(IsNull(RS!POMarking5), "", Trim(RS!POMarking5))
    
    txtPOPayment.Text = IIf(IsNull(RS!POPayment_Day), "", Trim(RS!POPayment_Day))
    
    If RS!NG_Cls = "1" Then
        cboNGCls.ListIndex = 0
    ElseIf RS!NG_Cls = "0" Then
        cboNGCls.ListIndex = 1
    ElseIf IsNull(RS!NG_Cls) Then
        cboNGCls.ListIndex = -1
    End If
    
    txtKodeKPPBC(10).Text = IIf(IsNull(RS!CODE_KPPBC), "", Trim(RS!CODE_KPPBC))
    txtNoIzin(0).Text = IIf(IsNull(RS!No_Izin), "", Trim(RS!No_Izin))
    dtNoIzin.Value = IIf(IsNull(RS!NoIzin_Date), Now, (RS!NoIzin_Date))
    
    For i = 2 To 20
     If i <> 17 Then
      Text1(i).BackColor = vbWhite
      Text1(i).DataChanged = False
     End If
    Next i
    
    For i = 0 To 9
     Text2(i).BackColor = vbWhite
     Text2(i).DataChanged = False
    Next i
    
    If RS!country_cls = "1" Then
        cbocountry.ListIndex = 1
    Else
        cbocountry.ListIndex = 0
    End If
    cbocountry.BackColor = vbWhite
    lblcountry.DataChanged = False
    
    If Trim(RS!Invoice_To) <> "" And Not IsNull(RS!Invoice_To) Then
        cekinvto.Value = 1
        For i = 0 To cbotradecode.ListCount - 1
            If cbotradecode.List(i, 0) = Trim(RS!Invoice_To) Then
                cbotradecode.ListIndex = i
                Exit For
            End If
        Next i
    Else
        cekinvto.Value = 0
        cbotradecode.ListIndex = -1
        txtName.Text = ""
    End If
    cbotradecode.BackColor = vbWhite
    txtName.DataChanged = False
    
    If Trim(RS!affiliate_cls) <> "0" And Not IsNull(RS!affiliate_cls) Then
     cekaff.Value = 1
    Else
     cekaff.Value = 0
    End If
    
    If Trim(RS!Region_Cls) <> "" And Not IsNull(RS!Region_Cls) Then
        For i = 0 To cboregion.ListCount - 1
            If cboregion.List(i, 0) = Trim(RS!Region_Cls) Then
                cboregion.ListIndex = i
                Exit For
            End If
        Next i
    Else
       cboregion.ListIndex = -1
       txtregion.Text = ""
    End If
    cboregion.BackColor = vbWhite
    txtregion.DataChanged = False
    
    If Trim(RS!Type_BC) <> "" And Not IsNull(RS!Type_BC) Then
        For i = 0 To cboBCType.ListCount - 1
            If cboBCType.List(i, 0) = Trim(RS!Type_BC) Then
                cboBCType.ListIndex = i
                Exit For
            End If
        Next i
    Else
       cboBCType.ListIndex = -1
       cboBCType.Text = ""
    End If
    cboBCType.BackColor = vbWhite
    
    If Trim(RS!POPayment_Terms) <> "" And Not IsNull(RS!POPayment_Terms) Then
        For i = 0 To cboregion.ListCount - 1
            If cbopayment.List(i, 0) = Trim(RS!POPayment_Terms) Then
                cbopayment.ListIndex = i
                Exit For
            End If
        Next i
    Else
       cbopayment.ListIndex = -1
       txtpayment.Text = ""
    End If
    cbopayment.BackColor = vbWhite
    txtpayment.DataChanged = False
    
    If Trim(RS!Price_Condition) <> "" And Not IsNull(RS!Price_Condition) Then
        For i = 0 To cboprice.ListCount - 1
            If cboprice.List(i, 0) = Trim(RS!Price_Condition) Then
                cboprice.ListIndex = i
                Exit For
            End If
        Next i
    Else
       cboprice.ListIndex = -1
       txtprice.Text = ""
    End If
    cboprice.BackColor = vbWhite
    txtprice.DataChanged = False
    
    If Trim(RS!Transportation_Cls) <> "" And Not IsNull(RS!Transportation_Cls) Then
        For i = 0 To cbotransport.ListCount - 1
            If cbotransport.List(i, 0) = Trim(RS!Transportation_Cls) Then
                cbotransport.ListIndex = i
                Exit For
            End If
        Next i
    Else
       cbotransport.ListIndex = -1
       txttransport.Text = ""
    End If
    cbotransport.BackColor = vbWhite
    txttransport.DataChanged = False
    
    If RS!PO_Cls = 1 Then
        cbopocls.ListIndex = 0
    Else
        cbopocls.ListIndex = 1
    End If
    cbopocls.BackColor = vbWhite
    lblpocls.DataChanged = False
    
    If RS!Epte_cls = 1 Then
        cboEPTE.ListIndex = 0
    Else
        cboEPTE.ListIndex = 1
    End If
    cboEPTE.BackColor = vbWhite
    lblEPTE.DataChanged = False
    
'    If rs!insurance_cls = 1 Then
'        cboinsurance.ListIndex = 0
'    Else
'        cboinsurance.ListIndex = 1
'    End If
'    cboinsurance.BackColor = vbWhite
'    lblinsurance.DataChanged = False
    
    If Trim(RS!Insurance_Cls) <> "" And Not IsNull(RS!Insurance_Cls) Then
        For i = 0 To cboinsurance.ListCount - 1
            If cboinsurance.List(i, 0) = Trim(RS!Insurance_Cls) Then
                cboinsurance.ListIndex = i
                Exit For
            End If
        Next i
    Else
       cboinsurance.ListIndex = -1
       txtinsurance.Text = ""
    End If
    cboinsurance.BackColor = vbWhite
    txtinsurance.DataChanged = False
    
    ubahtrade = True
    Text1(0).DataChanged = False
    Text1(0).Enabled = False
    headerGrid
    BrowseGrid
  End If
  
End Sub

Private Sub cbocountry_Click()
    If cbocountry.ListIndex <> -1 Then
        lblcountry.Caption = cbocountry.Text
    End If
End Sub

Private Sub cbocountry_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cbocountry_Click
End Sub

Private Sub cboEPTE_Click()
If cboEPTE.ListIndex <> -1 Then
  lblEPTE.Caption = cboEPTE.Column(1)
End If
End Sub

Private Sub cboEPTE_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then cboEPTE_Click
End Sub

Private Sub cboinsurance_Click()
 If cboinsurance.ListIndex <> -1 Then
  txtinsurance.Text = cboinsurance.Column(1)
 End If
End Sub

Private Sub cboinsurance_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode = 13 Then
   For i = 0 To cboinsurance.ListCount - 1
       If cboinsurance.Text = cboinsurance.List(i, 0) Then
           txtinsurance.Text = cboinsurance.List(i, 1)
           Exit For
       Else
           txtinsurance.Text = ""
       End If
   Next i
    cboinsurance_Click
 End If
End Sub
Private Sub cbopocls_Click()
    If cbopocls.ListIndex <> -1 Then
        lblpocls.Caption = cbopocls.Column(1)
    End If
End Sub

Private Sub cbopocls_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cbopocls_Click
End Sub
Private Sub cbopayment_Click()
 If cbopayment.ListIndex <> -1 Then
    txtpayment.Text = cbopayment.Column(1)
 End If
End Sub
Private Sub cbopayment_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode = 13 Then
   For i = 0 To cbopayment.ListCount - 1
       If cbopayment.Text = cbopayment.List(i, 0) Then
           txtpayment.Text = cbopayment.List(i, 1)
           Exit For
       Else
           txtpayment.Text = ""
       End If
   Next i
    cbopayment_Click
 End If
End Sub
Private Sub cboprice_Click()
 If cboprice.ListIndex <> -1 Then
    txtprice.Text = cboprice.Column(1)
 End If
End Sub

Private Sub cboprice_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
   For i = 0 To cboprice.ListCount - 1
       If cboprice.Text = cboprice.List(i, 0) Then
           txtprice.Text = cboprice.List(i, 1)
           Exit For
       Else
           txtprice.Text = ""
       End If
   Next i
    cboprice_Click
 End If
End Sub
Private Sub cbotransport_Click()
If cbotransport.ListIndex <> -1 Then
    txttransport.Text = cbotransport.Column(1)
 End If
End Sub

Private Sub cbotransport_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode = 13 Then
   For i = 0 To cbotransport.ListCount - 1
       If cbotransport.Text = cbotransport.List(i, 0) Then
           txttransport.Text = cbotransport.List(i, 1)
           Exit For
       Else
           txttransport.Text = ""
       End If
   Next i
    cbotransport_Click
 End If
End Sub

Private Sub cboregion_Click()
 If cboregion.ListIndex <> -1 Then
    txtregion.Text = cboregion.Column(1)
 End If
End Sub

Private Sub cboregion_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode = 13 Then
   For i = 0 To cboregion.ListCount - 1
       If cboregion.Text = cboregion.List(i, 0) Then
           txtregion.Text = cboregion.List(i, 1)
           Exit For
       Else
           txtregion.Text = ""
       End If
   Next i
    cboregion_Click
 End If
End Sub

Private Sub cbotradecls_Change()
    cbotradecls_Click
End Sub

Private Sub cbotradecls_Click()
    If cbotradecls.ListIndex <> -1 Then
        lbltradecls.Caption = cbotradecls.Column(1)
'        If cbotradecls.ListIndex = 1 Then
            cboNGCls.Enabled = True
'        Else
'            cboNGCls.Enabled = False
'        End If
    End If
End Sub

Private Sub cbotradecls_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cbotradecls_Click
End Sub

Private Sub cbotradecode_Click()
    If cbotradecode.ListIndex <> -1 Then
        txtName.Text = cbotradecode.Column(1)
    End If
End Sub

Private Sub cbotradecode_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    For i = 0 To cbotradecode.ListCount - 1
        If cbotradecode.Text = cbotradecode.List(i, 0) Then
            txtName.Text = cbotradecode.Column(1)
            Exit For
        Else
            txtName.Text = ""
        End If
    Next i
    cbotradecls_Click
End If
End Sub
Private Sub cekinvto_Click()
    If cekinvto.Value = 0 Then
        cbotradecode.Enabled = False
        txtName.Enabled = False
    Else
        cbotradecode.Enabled = True
        txtName.Enabled = True
    End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
  sql = "select * from Trade_Master"
  If RS.State <> adStateClosed Then RS.Close
  RS.Open sql, Db, adOpenKeyset, adLockOptimistic
  
  sqlGrid = "select * from Delivery_place"
  If rsGrid.State <> adStateClosed Then rsGrid.Close
  rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic

  'sqlregion = "select * from Region_Cls"
  'If rsRegion.State <> adStateClosed Then rsRegion.Close
  'rsRegion.Open sqlregion, Db, adOpenKeyset, adLockOptimistic
  
  Call settingcombo
  adtocombo
  CtrlMenu1.FormName = Me.Name
  Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
  Kosong
End Sub

Private Sub Text1_Change(Index As Integer)
Dim valid As Boolean
valid = True
If Index = 11 Then
    If (Val(Text1(11).Text) > 31) Then
        LblErrMsg.Caption = DisplayMsg(1022)
        Text1(11).SetFocus
        valid = False
    End If
ElseIf Index = 12 Then
  If (Val(Text1(12).Text) > 31) Then
    LblErrMsg.Caption = DisplayMsg(1022)
    Text1(12).SetFocus
    valid = False
    End If
End If
If valid Then LblErrMsg.Caption = ""
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    SSTab1.Tab = 0
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
  
  If Index = 11 Or Index = 12 Or Index = 18 Then
     If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
       KeyAscii = 0
     End If
     If KeyAscii = Asc(".") Then KeyAscii = 0
  End If
  
  If KeyAscii = Asc("'") Then
    KeyAscii = 0
  End If
    
  If KeyAscii = 13 Then
    If Index = 0 Then
        Browse
    Else
      SendKeys vbTab
    End If
  End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
  If Index = 0 Then
    If (Text1(0).Text <> "") Then Browse
  End If
  
  If Index = 11 Then
    If (Val(Text1(11).Text) > 31) Then Text1(11).SetFocus
  ElseIf Index = 12 Then
      If (Val(Text1(12).Text) > 31) Then Text1(12).SetFocus
  End If
End Sub

Function enter()
    Call Text1_KeyPress(0, 13)
End Function

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim TextGrid As String
With Grid
    TextGrid = Grid.Text
    If TextGrid = "S" Then
       txtlocation(0).Text = Trim(.TextMatrix(Row, bteColCode))
       txtlocation(0).locked = True
       txtlocation(0).BackColor = vbWhite
       txtlocation(0).DataChanged = False
       txtlocation(1).Text = Trim(.TextMatrix(Row, bteColName))
       txtlocation(1).BackColor = vbWhite
       txtlocation(1).DataChanged = False
       ubahgrid = True
       Call kosongColGrid
    ElseIf TextGrid = "D" Then
       Call kosongColGrid("S")
    End If
    .TextMatrix(Row, Col) = TextGrid
End With
End Sub

Private Sub kosongColGrid(Optional Kolom As String)
    Dim i As Integer
    With Grid
        .Col = bteColSelect
        If Kolom <> "" Then
           For i = 1 To .Rows - 1
              If .Text = Kolom Then .Text = ""
              If .TextMatrix(i, bteColSelect) <> "D" Then .TextMatrix(i, bteColSelect) = ""
           Next i
        Else
           For i = 1 To .Rows - 1
              If .TextMatrix(i, bteColSelect) <> "" Then .TextMatrix(i, bteColSelect) = ""
           Next i
        End If
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  If Grid.Col <> bteColSelect Then Cancel = True
End Sub

Private Sub Grid_Click()
  With Grid
    If .Row = 1 And .Col <> bteColSelect Then
      If .ColSort(.Col) = flexSortStringAscending Then
        .ColSort(.Col) = flexSortStringDescending
      Else
        .ColSort(.Col) = flexSortStringAscending
      End If
      .Sort = .ColSort(.Col)
    End If
  End With
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
  If Grid.Col = bteColSelect Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> Asc("S") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
  End If
End Sub

Private Sub command3_Click(Index As Integer)
Dim sql1 As String
Dim rs1 As New Recordset
Dim hapus As Boolean
Dim tanya

hapus = False
  Select Case Index
  Case 0: Kosong
          Text1(0).SetFocus
          
  Case 1:   If hakUpdate(Me.Name) = 0 Then _
            LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub

            Dim rs2 As New ADODB.Recordset
'            If rs2.State = 1 Then rs2.Close
'            rs2.CursorLocation = adUseClient
'            rs2.Open "select * from warehouse_master where wh_code='" & Trim(Text1(0)) & "'", Db, adOpenKeyset, adLockOptimistic
'            If rs2.EOF = False Then
'                LblErrMsg = DisplayMsg(3004)
'                If rs2.State = 1 Then rs2.Close
'                Me.MousePointer = vbDefault: Exit Sub
'            End If
'            If rs2.State = 1 Then rs2.Close

          If Text1(0).Text = "" Then
            Text1(0).SetFocus
            LblErrMsg = DisplayMsg(1016)
            Exit Sub
          ElseIf cbotradecls.ListIndex = -1 Then
            cbotradecls.SetFocus
            LblErrMsg = DisplayMsg(1017)
            Exit Sub
          ElseIf Text1(2).Text = "" Then
            Text1(2).SetFocus
            LblErrMsg = DisplayMsg(1018)
            Exit Sub
          ElseIf Text1(3).Text = "" Then
            Text1(3).SetFocus
            LblErrMsg = DisplayMsg(1019)
            Exit Sub
          Else
                    
            If Trim(cmbbox_warehouse) <> "" Then
                    rs2.CursorLocation = adUseClient
                    rs2.Open "select * from warehouse_master where wh_code='" & Trim(cmbbox_warehouse) & "'", Db, adOpenKeyset, adLockOptimistic
                    If rs2.EOF = False Then
                    
                    Else
                        LblErrMsg = DisplayMsg(4023)
                        If rs2.State = 1 Then rs2.Close
                        Me.MousePointer = vbDefault: Exit Sub
                    End If
                    If rs2.State = 1 Then rs2.Close
            End If
         '
          If cboregion.Text <> "" Then
           cboregion.MatchEntry = 1
           cboregion.Text = cboregion.Text
           If cboregion.MatchFound = False Then
            LblErrMsg = DisplayMsg(8045)
            cboregion.SetFocus
            cboregion.MatchEntry = 2
            Exit Sub
           End If
           cboregion.MatchEntry = 2
          End If
                    
          If cbopayment.Text <> "" Then
           cbopayment.MatchEntry = 1
           cbopayment.Text = cbopayment.Text
           If cbopayment.MatchFound = False Then
            LblErrMsg = DisplayMsg(8050)
            cbopayment.SetFocus
            cbopayment.MatchEntry = 2
            Exit Sub
           End If
           cbopayment.MatchEntry = 2
          End If
                    
           If cboprice.Text <> "" Then
           cboprice.MatchEntry = 1
           cboprice.Text = cboprice.Text
           If cboprice.MatchFound = False Then
            LblErrMsg = DisplayMsg(8051)
            cboprice.SetFocus
            cboprice.MatchEntry = 2
            Exit Sub
           End If
           cboprice.MatchEntry = 2
          End If
                    
          If cbotransport.Text <> "" Then
           cbotransport.MatchEntry = 1
           cbotransport.Text = cbotransport.Text
           If cbotransport.MatchFound = False Then
            LblErrMsg = DisplayMsg(8059)
            cbotransport.SetFocus
            cbotransport.MatchEntry = 2
            Exit Sub
           End If
           cbotransport.MatchEntry = 2
          End If
          
          If cbotradecode.Text <> "" Then
            cbotradecode.MatchEntry = 1
            cbotradecode.Text = cbotradecode.Text
            If cbotradecode.MatchFound = False Then
                LblErrMsg = DisplayMsg(4013)
                cbotradecode.SetFocus
                cbotradecode.MatchEntry = 2
                Exit Sub
            End If
            cbotradecode.MatchEntry = 2
          End If
          
          If cbotradecls.ListIndex = 1 Then
            If cboNGCls.ListIndex = -1 Then
                LblErrMsg.Caption = DisplayMsg("0079")
                cboNGCls.SetFocus
                Exit Sub
            Else
                LblErrMsg.Caption = ""
            End If
          End If
                    
            RS.filter = "Trade_Code='" & Text1(0).Text & "' "
            
            If RS.BOF And RS.EOF Then RS.AddNew
              RS(0) = Text1(0).Text
              Text1(0).DataChanged = False
              Text1(0).BackColor = vbWhite
              RS(1) = cbotradecls.Text
              cbotradecls.BackColor = vbWhite
              lbltradecls.DataChanged = False
            
            If Trim(cmbbox_warehouse) = "" Then
                RS!Subcon_WH_Code = Null
            Else
                RS!Subcon_WH_Code = Trim(cmbbox_warehouse)
            End If
            RS!trade_name = Text1(2).Text
            RS!SAP_Code = Text1(1).Text
            RS!trade_abbr = Text1(3).Text
            RS!contact_person = Text1(4).Text
            RS!address1 = Text1(5).Text
            RS!address2 = Text1(6).Text
            RS!City = Text1(7).Text
            RS!postal_code = Text1(8).Text
            RS!Telephone = Text1(9).Text
            RS!fax = Text1(10).Text
            RS!Closing_Day = Text1(11).Text
            RS!Pay_Day = Text1(12).Text
            RS!NPWP_No = Text1(13).Text
            RS!NPWP_Name = Text1(14).Text
            RS!NPWP_Address = Text1(15).Text
            RS!NPWP_City = Text1(16).Text
            RS!InvoicePay_Days = IIf(IsNumeric(Text1(18).Text) = False, 0, Text1(18).Text)
            RS!NPPKP_No = Text1(19).Text
            RS!Country = Text1(20).Text
            RS!NITKU = Text1(17).Text
            
            RS!POCaseMark1 = Text2(0).Text
            RS!POCaseMark2 = Text2(1).Text
            RS!POCaseMark3 = Text2(2).Text
            RS!POCaseMark4 = Text2(3).Text
            RS!POCaseMark5 = Text2(4).Text
            RS!POMarking1 = Text2(5).Text
            RS!POMarking2 = Text2(6).Text
            RS!POMarking3 = Text2(7).Text
            RS!POMarking4 = Text2(8).Text
            RS!POMarking5 = Text2(9).Text
                        
            RS!POPayment_Day = IIf(IsNumeric(txtPOPayment.Text) = False, 0, txtPOPayment.Text)
                                    
            For i = 2 To 20
             If i <> 17 Then
              Text1(i).DataChanged = False
              Text1(i).BackColor = vbWhite
             End If
            Next i
            
            For i = 0 To 9
              Text2(i).DataChanged = False
              Text2(i).BackColor = vbWhite
            Next i
            
            If cbocountry.Text = "Domestic" Then
                RS("country_cls") = "0"
            ElseIf cbocountry.Text = "Overseas" Then
                RS("country_cls") = "1"
            End If
            cbocountry.BackColor = vbWhite
            lblcountry.DataChanged = False
            
            If cboregion.Text <> "" Then
             RS("region_cls") = cboregion.Text
            Else
             RS("region_cls") = Null
            End If
            cboregion.BackColor = vbWhite
            txtregion.DataChanged = False
            
            If cbopayment.Text <> "" Then
             RS("POPayment_Terms") = cbopayment.Text
            Else
             RS("POPayment_Terms") = Null
            End If
            cbopayment.BackColor = vbWhite
            txtpayment.DataChanged = False
            
            If cboprice.Text <> "" Then
             RS("Price_Condition") = cboprice.Text
            Else
             RS("Price_Condition") = Null
            End If
            cboprice.BackColor = vbWhite
            txtprice.DataChanged = False
            
            If cbotransport.Text <> "" Then
             RS("Transportation_Cls") = cbotransport.Text
            Else
             RS("Transportation_Cls") = Null
            End If
            cbotransport.BackColor = vbWhite
            txttransport.DataChanged = False
            
            If cekinvto.Value = 1 Then
                RS("invoice_to") = cbotradecode.Text
            Else
                RS("invoice_to") = Null
            End If
            cbotradecode.BackColor = vbWhite
            txtName.DataChanged = False
            
            If cekaff.Value = True Then
             RS("Affiliate_Cls") = 1
            Else
             RS("Affiliate_Cls") = 0
            End If
            
            If cboBCType.Text <> "" Then
             RS("Type_BC") = cboBCType.Text
            Else
             RS("Type_BC") = Null
            End If
            cboBCType.BackColor = vbWhite
            'txtpayment.DataChanged = False
            
            RS("Epte_cls") = cboEPTE.Column(1)
            cboEPTE.BackColor = vbWhite
            lblEPTE.DataChanged = False
            
            If cboinsurance.Text <> "" Then
             RS("Insurance_Cls") = cboinsurance.Text
            Else
             RS("Insurance_Cls") = Null
            End If
            cboinsurance.BackColor = vbWhite
            txtinsurance.DataChanged = False
            
            RS("po_cls") = cbopocls.Column(1)
            cbopocls.BackColor = vbWhite
            lblpocls.DataChanged = False
            
            If cboNGCls.ListIndex <> -1 Then
                RS("NG_Cls") = cboNGCls.Column(1)
            Else
                RS("NG_Cls") = "0"
            End If
            
            If txtKodeKPPBC(10).Text <> "" Then
                RS("CODE_KPPBC") = txtKodeKPPBC(10).Text
            Else
                RS("CODE_KPPBC") = Null
            End If
            
            If txtNoIzin(0).Text <> "" Then
                RS("No_Izin") = txtNoIzin(0).Text
            Else
                RS("No_Izin") = Null
            End If
            
            If Format(dtNoIzin.Value, "dd-MMM-yyyy") <> Format(Now, "dd-MMM-yyyy") Then
                Dim pDate As Date
                pDate = Format(dtNoIzin.Value, "dd-MMM-yyyy")
                
                RS("NoIzin_Date") = pDate
            Else
                RS("NoIzin_Date") = Null
            End If
            
            RS("last_update") = Now
            RS("last_user") = userLogin
            RS.update
            
            With Grid
                For i = 1 To .Rows - 1
                  If .TextMatrix(i, bteColSelect) = "D" Then
                    If IsEmpty(tanya) Then tanya = MsgBox("Do You Really Want To Delete This Data ?", vbQuestion & vbYesNo, "Confirmation")
                    If tanya = vbYes Then
                        sql1 = "select * from OrderEntry_Master where Location_Code = '" & .TextMatrix(i, bteColCode) & "' " & _
                               "and cust_code='" & Text1(0).Text & "' "
                        Set rs1 = Db.Execute(sql1)
                        If Not (rs1.BOF And rs1.EOF) Then
                          LblErrMsg.Caption = DisplayMsg(1204) '"Can't Delete This Record, Record is used in table Order Entry Master"
                          .Row = i
                          .SetFocus
                          Exit Sub
                        Else
                            sql1 = "delete from Delivery_Place where Trade_Code='" & Text1(0).Text & "' and " & _
                                  "location_code='" & .TextMatrix(i, bteColCode) & "' "
                            Db.Execute sql1
                        
                            hapus = True
                        End If
                    Else
                        Exit For
                    End If
                  End If
                Next i
                
                If (hapus) Then headerGrid: BrowseGrid: LblErrMsg = DisplayMsg(1201): Exit Sub
            End With
            
            If txtlocation(1).Text <> "" Then
                If txtlocation(0).Text = "" Then LblErrMsg = DisplayMsg(1031): txtlocation(0).SetFocus: Exit Sub
            End If
            
            If txtlocation(0).Text <> "" Then
                If txtlocation(1).Text = "" Then LblErrMsg = DisplayMsg(1032): txtlocation(1).SetFocus: Exit Sub
            
                If ubahgrid = False Then
                    rsGrid.filter = "Trade_Code='" & Text1(0).Text & "' and Location_code='" & txtlocation(0) & "' "
                    If Not (rsGrid.EOF And rsGrid.BOF) Then
                        LblErrMsg = DisplayMsg(1023): txtlocation(0).SetFocus: Exit Sub
                    Else
                        rsGrid.AddNew
                    End If
                Else
                    rsGrid.filter = "Trade_Code='" & RS("Trade_Code") & "' and Location_Code='" & txtlocation(0) & "' "
                End If
                
                rsGrid("Trade_Code") = Text1(0).Text
                rsGrid("Location_Code") = txtlocation(0).Text
                rsGrid("Location_name") = txtlocation(1).Text
                rsGrid("last_update") = Now
                rsGrid("last_user") = userLogin
                rsGrid.update
                
            rsGrid.Requery
            rsGrid.filter = ""
            
            kosonggrid
            headerGrid
            BrowseGrid
                
            End If
            
            LblErrMsg = DisplayMsg(IIf((ubahtrade = False), 1000, 1101))
            ubahtrade = True
            Text1(0).Enabled = False
          End If
          
    Case 2: kosonggrid
    
 End Select
  
End Sub

Private Sub cmdinquiry_Click()
  If hakAkses("FrmTradeMasterInquiry") = 0 Then LblErrMsg = DisplayMsg(3007): Exit Sub
  If CancelButton Then Exit Sub
  FrmTradeMasterInquiry.cmdSubMenu.Caption = "&Back"
  FrmTradeMasterInquiry.Show
  Unload Me
End Sub

Private Sub CmdSubMenu_Click()
  If CancelButton Then Exit Sub
  frmMainMenu.Show
  Unload Me
End Sub

Private Function CancelButton() As Boolean
Dim i As Integer

   If Text1(0).DataChanged = True Then
        Text1(0).BackColor = vbRed
        CancelButton = True
    Else
        Text1(0).BackColor = vbWhite
    End If

   If lbltradecls.DataChanged = True Then
        cbotradecls.BackColor = vbRed
        CancelButton = True
    Else
        cbotradecls.BackColor = vbWhite
    End If
   
   If lblcountry.DataChanged = True Then
        cbocountry.BackColor = vbRed
        CancelButton = True
    Else
        cbocountry.BackColor = vbWhite
    End If
    
    If txtName.DataChanged = True Then
        cbotradecode.BackColor = vbRed
        CancelButton = True
    Else
        cbotradecode.BackColor = vbWhite
    End If
    
    If lblpocls.DataChanged = True Then
        cbopocls.BackColor = vbRed
        CancelButton = True
    Else
        cbopocls.BackColor = vbWhite
    End If
    
    If lblEPTE.DataChanged = True Then
        cboEPTE.BackColor = vbRed
        CancelButton = True
    Else
        cboEPTE.BackColor = vbWhite
    End If
    
    If txtinsurance.DataChanged = True Then
        cboinsurance.BackColor = vbRed
        CancelButton = True
    Else
        cboinsurance.BackColor = vbWhite
    End If
   
   If txtregion.DataChanged = True Then
    cboregion.BackColor = vbRed
    CancelButton = True
   Else
    cboregion.BackColor = vbWhite
   End If
   
   If txtprice.DataChanged = True Then
    cboprice.BackColor = vbRed
    CancelButton = True
   Else
    cboprice.BackColor = vbWhite
   End If
   
   If txtpayment.DataChanged = True Then
    cbopayment.BackColor = vbRed
    CancelButton = True
   Else
    cbopayment.BackColor = vbWhite
   End If
   
   If txttransport.DataChanged = True Then
    cbotransport.BackColor = vbRed
    CancelButton = True
   Else
    cbotransport.BackColor = vbWhite
   End If
   
  For i = 0 To 1
   If txtlocation(i).DataChanged = True Then
        txtlocation(i).BackColor = vbRed
        CancelButton = True
    Else
        txtlocation(i).BackColor = vbWhite
    End If
  Next i

For i = 2 To 20
   If i <> 17 Then
    If Text1(i).DataChanged = True Then
     Text1(i).BackColor = vbRed
     CancelButton = True
    Else
     Text1(i).BackColor = vbWhite
    End If
   End If
Next i

For i = 0 To 9
   If Text2(i).DataChanged = True Then
     Text2(i).BackColor = vbRed
     CancelButton = True
   Else
     Text2(i).BackColor = vbWhite
   End If
Next i

If CancelButton Then LblErrMsg.Caption = DisplayMsg(1049) '"Please Submit first!"

End Function

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblErrMsg.Caption = ErrMsg
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RS.State <> adStateClosed Then RS.Close
    If rsGrid.State <> adStateClosed Then rsGrid.Close
End Sub

Private Sub text2_GotFocus(Index As Integer)
    SSTab1.Tab = 1
End Sub

Private Sub SSTab1_GotFocus()
    If SSTab1.Tab = 0 Then
        If Text1(0).Enabled Then Text1(0).SetFocus Else cbotradecls.SetFocus
    Else
        Text2(0).SetFocus
    End If
End Sub

Private Sub txtKodeKPPBC_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Keychar As String
    If KeyAscii > 31 Then
    Keychar = Chr(KeyAscii)
        If Not IsNumeric(Keychar) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtlocation_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub settingcombo()
Dim ls_sql As String
Dim rs_combo As New ADODB.Recordset
Dim i As Integer


cboBCType.columnCount = 1
cboBCType.clear

ls_sql = "select BC_type from BC_master"
rs_combo.Open ls_sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
i = 0

Do While Not rs_combo.EOF
cboBCType.AddItem rs_combo("BC_type")
rs_combo.MoveNext
Loop

cboBCType.ColumnWidths = "90"
cboBCType.ListWidth = 90
cboBCType.ListRows = 7


End Sub

