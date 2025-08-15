VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPackingCreate 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Packing List Create"
   ClientHeight    =   10725
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   15240
   Icon            =   "FrmPackingCreate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10725
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdserialnumber 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print Serial Number"
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
      TabIndex        =   98
      Top             =   10170
      Width           =   1980
   End
   Begin VB.CommandButton CmdStuffing 
      BackColor       =   &H0080FFFF&
      Caption         =   "Stuffing &Report"
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
      Left            =   8670
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   10170
      Width           =   1590
   End
   Begin VB.TextBox TxtRemarks 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8460
      MaxLength       =   50
      TabIndex        =   30
      Top             =   7710
      Width           =   5325
   End
   Begin VB.TextBox TxtTTCtnQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   6495
      Locked          =   -1  'True
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   9030
      Width           =   2085
   End
   Begin VB.TextBox TxtPacNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   9030
      Width           =   1890
   End
   Begin VB.TextBox TxtTTAm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   2175
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   9030
      Width           =   2085
   End
   Begin VB.TextBox TxtTTQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   4335
      Locked          =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   9030
      Width           =   2085
   End
   Begin VB.TextBox TxtTTW 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   8655
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   9030
      Width           =   2085
   End
   Begin VB.TextBox TxtTTWG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   10815
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   9030
      Width           =   2085
   End
   Begin VB.TextBox TxtTTV 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   12975
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   9030
      Width           =   2085
   End
   Begin VB.TextBox TxtCaseMark 
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
      Index           =   4
      Left            =   8460
      MaxLength       =   100
      TabIndex        =   29
      Top             =   7710
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.TextBox TxtCaseMark 
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
      Index           =   3
      Left            =   4545
      MaxLength       =   100
      TabIndex        =   28
      Top             =   8085
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.TextBox TxtCaseMark 
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
      Index           =   2
      Left            =   4545
      MaxLength       =   100
      TabIndex        =   27
      Top             =   7710
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.TextBox TxtCaseMark 
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
      Index           =   1
      Left            =   1575
      MaxLength       =   100
      TabIndex        =   26
      Top             =   8085
      Width           =   5865
   End
   Begin VB.TextBox TxtCaseMark 
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
      Index           =   0
      Left            =   1575
      MaxLength       =   100
      TabIndex        =   25
      Top             =   7710
      Width           =   5865
   End
   Begin VB.CommandButton CmdPreview 
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
      Left            =   10350
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   10170
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancel 
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
      Left            =   11589
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   10170
      Width           =   1140
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
      Left            =   14025
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   10170
      Width           =   1140
   End
   Begin VB.CommandButton CmdClear 
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
      Left            =   12807
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   10170
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   1875
      Left            =   150
      TabIndex        =   49
      Top             =   2070
      Width           =   15045
      Begin VB.TextBox TxtForwarder 
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
         Left            =   13035
         MaxLength       =   50
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   210
         Width           =   1920
      End
      Begin VB.TextBox TxtCountry 
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
         Left            =   13035
         MaxLength       =   25
         TabIndex        =   20
         Top             =   585
         Width           =   1920
      End
      Begin VB.CommandButton CmdCreate 
         BackColor       =   &H0080FFFF&
         Caption         =   "Create"
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
         TabIndex        =   22
         Top             =   1410
         Width           =   1140
      End
      Begin VB.TextBox TxtFinal 
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
         Left            =   13035
         MaxLength       =   25
         TabIndex        =   21
         Top             =   990
         Width           =   1920
      End
      Begin VB.TextBox TxtTo 
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
         Left            =   9435
         MaxLength       =   25
         TabIndex        =   17
         Top             =   990
         Width           =   1695
      End
      Begin VB.TextBox TxtFrom 
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
         Left            =   9435
         MaxLength       =   25
         TabIndex        =   16
         Top             =   615
         Width           =   1695
      End
      Begin VB.TextBox TxtMotherVessel 
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
         Left            =   1515
         MaxLength       =   25
         TabIndex        =   10
         Top             =   1365
         Width           =   1605
      End
      Begin VB.TextBox TxtVessel 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2835
         MaxLength       =   25
         TabIndex        =   9
         Top             =   990
         Width           =   2145
      End
      Begin VB.ComboBox cboStatus 
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
         ItemData        =   "FrmPackingCreate.frx":0E42
         Left            =   135
         List            =   "FrmPackingCreate.frx":0E4C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker DtPacking 
         Height          =   315
         Left            =   6480
         TabIndex        =   11
         Top             =   240
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   335020035
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DtEtd 
         Height          =   315
         Left            =   6480
         TabIndex        =   12
         Top             =   615
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   335020035
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DtStuffing 
         Height          =   315
         Left            =   9435
         TabIndex        =   15
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   335020035
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DtEta 
         Height          =   315
         Left            =   6480
         TabIndex        =   13
         Top             =   990
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   335020035
         CurrentDate     =   37798
      End
      Begin VB.TextBox TxtDay 
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
         Left            =   13035
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   210
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.TextBox TxtPaymentTerm 
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
         Left            =   13035
         MaxLength       =   50
         TabIndex        =   84
         Top             =   210
         Visible         =   0   'False
         Width           =   1920
      End
      Begin MSForms.ComboBox cboPlaceofDestination 
         Height          =   315
         Left            =   11730
         TabIndex        =   101
         Top             =   1365
         Width           =   780
         VariousPropertyBits=   612386843
         DisplayStyle    =   7
         Size            =   "1376;556"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Place of Destination"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   36
         Left            =   10620
         TabIndex        =   100
         Top             =   1365
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPlaceofDestination 
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
         Left            =   12660
         TabIndex        =   99
         Top             =   1380
         Width           =   960
      End
      Begin VB.Line Line5 
         Index           =   2
         X1              =   12540
         X2              =   13710
         Y1              =   1665
         Y2              =   1665
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forwarder"
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
         Left            =   11460
         TabIndex        =   94
         Top             =   300
         Width           =   870
      End
      Begin MSForms.ComboBox cboTrans 
         Height          =   315
         Left            =   2850
         TabIndex        =   8
         Top             =   630
         Width           =   2130
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "3757;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line5 
         Index           =   1
         X1              =   9150
         X2              =   10530
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Label lblpackingtype 
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
         Left            =   9210
         TabIndex        =   88
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Type"
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
         Left            =   7050
         TabIndex        =   87
         Top             =   1425
         Width           =   1140
      End
      Begin MSForms.ComboBox cbopackingtype 
         Height          =   315
         Left            =   8310
         TabIndex        =   18
         Top             =   1365
         Width           =   780
         VariousPropertyBits=   612386843
         DisplayStyle    =   7
         Size            =   "1376;556"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboPaymentTerm 
         Height          =   315
         Left            =   4680
         TabIndex        =   14
         Top             =   1365
         Width           =   810
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "1429;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Term"
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
         Index           =   14
         Left            =   3300
         TabIndex        =   86
         Top             =   1425
         Width           =   1260
      End
      Begin VB.Label lblTerm 
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
         Left            =   5655
         TabIndex        =   85
         Top             =   1395
         Width           =   1140
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stuffing Date"
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
         Left            =   8175
         TabIndex        =   81
         Top             =   300
         Width           =   1125
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   5580
         X2              =   6930
         Y1              =   1650
         Y2              =   1650
      End
      Begin MSForms.ComboBox cboPaymentCode 
         Height          =   315
         Left            =   13035
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   210
         Visible         =   0   'False
         Width           =   1920
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   7
         Size            =   "3387;556"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblFix 
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
         Left            =   13035
         TabIndex        =   61
         Top             =   277
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final Destination"
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
         Left            =   11460
         TabIndex        =   60
         Top             =   1050
         Width           =   1410
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   8175
         TabIndex        =   59
         Top             =   1050
         Width           =   210
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BL No"
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
         Index           =   16
         Left            =   11460
         TabIndex        =   58
         Top             =   675
         Width           =   495
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Port"
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
         Left            =   8175
         TabIndex        =   57
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E.T.A"
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
         Left            =   5160
         TabIndex        =   56
         Top             =   1050
         Width           =   450
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E.T.D"
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
         Left            =   5160
         TabIndex        =   55
         Top             =   675
         Width           =   465
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Date"
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
         Left            =   5160
         TabIndex        =   54
         Top             =   300
         Width           =   1125
      End
      Begin MSForms.ComboBox cboPacking 
         Height          =   315
         Left            =   2850
         TabIndex        =   7
         Top             =   240
         Width           =   2130
         VariousPropertyBits=   612386843
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "3757;556"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mother Vessel"
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
         Left            =   165
         TabIndex        =   53
         Top             =   1425
         Width           =   1200
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Feeder Vessel"
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
         Left            =   1515
         TabIndex        =   52
         Top             =   1050
         Width           =   1200
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transportation"
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
         Left            =   1515
         TabIndex        =   51
         Top             =   675
         Width           =   1245
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing No"
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
         Left            =   1515
         TabIndex        =   50
         Top             =   300
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1125
      Left            =   150
      TabIndex        =   44
      Top             =   870
      Width           =   15045
      Begin VB.TextBox TxtTitle 
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
         Left            =   9885
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   630
         Width           =   4890
      End
      Begin VB.TextBox lblNotify 
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
         Height          =   240
         Left            =   11505
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2070
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox lblCust 
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
         Height          =   240
         Left            =   3090
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   270
         Width           =   4935
      End
      Begin MSComCtl2.DTPicker DtDel1 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   645
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
         Format          =   335151107
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DtDel2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   3510
         TabIndex        =   2
         Top             =   645
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
         Format          =   335151107
         CurrentDate     =   37798
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title Description"
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
         Left            =   8400
         TabIndex        =   95
         Top             =   690
         Width           =   1380
      End
      Begin VB.Line Line6 
         Visible         =   0   'False
         X1              =   11235
         X2              =   14475
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Label lblWH 
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
         Left            =   11235
         TabIndex        =   90
         Top             =   1215
         Width           =   3255
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warehouse"
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
         Left            =   8565
         TabIndex        =   89
         Top             =   1215
         Visible         =   0   'False
         Width           =   960
      End
      Begin MSForms.ComboBox cboWH 
         Height          =   315
         Left            =   9660
         TabIndex        =   33
         Top             =   1200
         Visible         =   0   'False
         Width           =   1500
         VariousPropertyBits=   612386843
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "2646;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         X1              =   11505
         X2              =   14745
         Y1              =   555
         Y2              =   555
      End
      Begin MSForms.ComboBox cboDelPlace 
         Height          =   315
         Left            =   9885
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   255
         Width           =   1500
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee"
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
         Left            =   8400
         TabIndex        =   83
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lblDelPlace 
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
         Height          =   240
         Left            =   11475
         TabIndex        =   82
         Top             =   285
         Width           =   3255
      End
      Begin MSForms.ComboBox cboPONo 
         Height          =   315
         Left            =   5925
         TabIndex        =   3
         Top             =   630
         Width           =   2160
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "3810;556"
         ColumnCount     =   2
         ListRows        =   7
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do. No"
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
         Left            =   5190
         TabIndex        =   62
         Top             =   720
         Width           =   585
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notify"
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
         Left            =   8610
         TabIndex        =   48
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSForms.ComboBox cboNotify 
         Height          =   255
         Left            =   9570
         TabIndex        =   32
         Top             =   1200
         Visible         =   0   'False
         Width           =   1500
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;450"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line4 
         X1              =   11235
         X2              =   14475
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer CD"
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
         Left            =   165
         TabIndex        =   47
         Top             =   315
         Width           =   1170
      End
      Begin VB.Line Line2 
         X1              =   3120
         X2              =   8040
         Y1              =   555
         Y2              =   555
      End
      Begin MSForms.ComboBox cboCust 
         Height          =   315
         Left            =   1500
         TabIndex        =   0
         Top             =   255
         Width           =   1500
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   3180
         TabIndex        =   46
         Top             =   705
         Width           =   165
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
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
         Left            =   165
         TabIndex        =   45
         Top             =   705
         Width           =   1185
      End
   End
   Begin VB.CommandButton CmdSubMenu 
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   10170
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   120
      TabIndex        =   40
      Top             =   9510
      Width           =   15045
      Begin VB.Label lblErrMsg 
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
         TabIndex        =   42
         Top             =   195
         Width           =   14805
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   3180
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4020
      Width           =   15045
      _cx             =   26538
      _cy             =   5609
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
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13260
      TabIndex        =   93
      Top             =   90
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   7620
      TabIndex        =   96
      Top             =   7725
      Width           =   765
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Ctn Qty"
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
      Left            =   6967
      TabIndex        =   92
      Top             =   8685
      Width           =   1140
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Information"
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
      Left            =   210
      TabIndex        =   80
      Top             =   7320
      Width           =   1395
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   330
      Left            =   120
      Top             =   7260
      Width           =   15045
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00A6D2FF&
      Height          =   915
      Left            =   120
      Top             =   7560
      Width           =   15045
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   495
      Left            =   120
      Top             =   8940
      Width           =   15045
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Packing No."
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
      Left            =   225
      TabIndex        =   79
      Top             =   8685
      Width           =   1005
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   2880
      TabIndex        =   78
      Top             =   8685
      Width           =   660
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty"
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
      Left            =   4980
      TabIndex        =   77
      Top             =   8685
      Width           =   780
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Weight (Net)"
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
      Index           =   28
      Left            =   8910
      TabIndex        =   76
      Top             =   8685
      Width           =   1560
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Weight (Gross)"
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
      Index           =   29
      Left            =   10965
      TabIndex        =   75
      Top             =   8685
      Width           =   1770
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Volume"
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
      Left            =   13455
      TabIndex        =   74
      Top             =   8685
      Width           =   1125
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line5 "
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
      Left            =   7620
      TabIndex        =   67
      Top             =   7725
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line4"
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
      Left            =   3945
      TabIndex        =   66
      Top             =   8130
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line3"
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
      Left            =   3945
      TabIndex        =   65
      Top             =   7755
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title 3"
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
      Left            =   270
      TabIndex        =   64
      Top             =   8130
      Width           =   525
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title 2"
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
      Left            =   270
      TabIndex        =   63
      Top             =   7755
      Width           =   525
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Packing List Create"
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
      Left            =   150
      TabIndex        =   43
      Top             =   270
      Width           =   15045
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   330
      Left            =   120
      Top             =   8625
      Width           =   15045
   End
End
Attribute VB_Name = "FrmPackingCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim ColCtr As Long
Dim ColAsk As Long
Dim ColContainerNo As Long
Dim ColSealNo As Long
Dim ColDrySize As Long
Dim ColOrder As Long
Dim ColProd As Long
Dim ColDesc As Long
Dim ColQty As Long
Dim colrem As Long
Dim ColUnit As Long
Dim ColSerialFrom As Long
Dim ColSerialTo As Long
Dim ColWeight As Long
Dim ColWGros As Long
Dim ColVol As Long
Dim ColPacking As Long
Dim ColCartonNo As Long
Dim ColLength As Long
Dim ColWidth As Long
Dim ColThickness  As Long
Dim ColDelDate As Long
Dim ColCurr As Long
Dim ColPrice As Long
Dim ColAmount As Long
Dim ColCounter As Long
Dim ColCurrT As Long
Dim ColUnitT As Long
Dim ColRemT As Long
Dim ColPos As Long
Dim ColPip As Long
Dim ColMaker As Long
Dim ColSeq As Long
Dim ColQtyPerCarton As Long
Dim ColNetWeight As Long
Dim ColGrossWeight As Long
Dim ColPo As Long
Dim ColpoSeqNo As Long

Dim Q As Double
Dim CartonQty As Double, VolumePacking As Double, NetWeightPacking As Double, GrossWeightPacking As Double
Dim Volume As Double, NetWeight As Double, GrossWeight As Double, NetSisa As Double, GrossSisa  As Double
Dim rsno As New ADODB.Recordset
Dim sqlno As String
Dim listPODate As String

Private Sub Header()

    ColCtr = 0
    ColAsk = 1
    ColContainerNo = 2
    'ColSealNo = 3
    ColDrySize = 3
    ColOrder = 4
    ColMaker = 5
    ColProd = 6
    ColDesc = 7
    ColQty = 8
    colrem = 9
    ColUnit = 10
    ColSerialFrom = 11
    ColSerialTo = 12
    ColPacking = 11 + 2
    ColCartonNo = 12 + 2

    ColWeight = 13 + 2
    ColWGros = 14 + 2
    ColLength = 15 + 2
    ColWidth = 16 + 2
    ColThickness = 17 + 2
    ColVol = 18 + 2
    
    ColDelDate = 19 + 2
    ColCurr = 20 + 2
    ColPrice = 21 + 2
    ColAmount = 22 + 2
    ColCounter = 23 + 2
    ColCurrT = 24 + 2
    ColUnitT = 25 + 2
    ColRemT = 26 + 2
    ColPos = 27 + 2
    ColPip = 28 + 2
    ColSeq = 29 + 2
    ColQtyPerCarton = 30 + 2
    ColNetWeight = 31 + 2
    ColGrossWeight = 32 + 2
    ColPo = 33 + 2
    ColpoSeqNo = 34 + 2
    
    With grid
        
        .clear
        .ColS = 35 + 2
        .Rows = 2
        
        .ColWidth(ColCtr) = 300
        .ColWidth(ColAsk) = 300
        .ColWidth(ColContainerNo) = 2000
        .ColWidth(ColDrySize) = 1200
        .ColWidth(ColOrder) = 1200
        .ColWidth(ColMaker) = 1200
        .ColWidth(ColDesc) = 2000
        .ColWidth(ColQty) = 900
        .ColWidth(colrem) = 1000
        .ColWidth(ColUnit) = 500
        .ColWidth(ColSerialFrom) = 1100
        .ColWidth(ColSerialTo) = 1100
        .ColWidth(ColWeight) = 900
        .ColWidth(ColWGros) = 900
        .ColWidth(ColVol) = 900
        .ColWidth(ColPacking) = 800
        .ColWidth(ColCartonNo) = 800
        .ColWidth(ColLength) = 950
        .ColWidth(ColWidth) = 950
        .ColWidth(ColThickness) = 950
        .ColWidth(ColDelDate) = 1250
        .ColWidth(ColCurr) = 600
        .ColWidth(ColPrice) = 1200
        .ColWidth(ColAmount) = 1200
        .ColWidth(ColSeq) = 1000
        .ColWidth(ColQtyPerCarton) = 1000
        .ColWidth(ColNetWeight) = 1000
        .ColWidth(ColGrossWeight) = 1000
        .ColWidth(ColPo) = 1000
        .ColWidth(ColpoSeqNo) = 1000
        
        For i = 0 To 1
            .TextMatrix(i, ColCtr) = " "
            .TextMatrix(i, ColAsk) = "C"
'            .TextMatrix(i, ColContainerNo) = "Container/Seal No."
'            .TextMatrix(i, ColDrySize) = "Dry Size Of Container"
            '.TextMatrix(i, ColOrder) = "Order No. (Ref No.)"
            
            .TextMatrix(i, ColContainerNo) = "Container No."
            .TextMatrix(i, ColDrySize) = "Seal No"
            .TextMatrix(i, ColOrder) = "Delivery No. (Ref No.)"
            .TextMatrix(i, ColMaker) = "Maker Item Code"
            .TextMatrix(i, ColProd) = "Prod Code"
            .TextMatrix(i, ColDesc) = "Desc"
            .TextMatrix(i, ColQty) = "Qty"
            .TextMatrix(i, colrem) = "Remaining Qty"
            .TextMatrix(i, ColUnit) = "Unit"
            .TextMatrix(i, ColSerialFrom) = "Serial From"
            .TextMatrix(i, ColSerialTo) = "Serial To"
            .TextMatrix(i, ColWeight) = "Weight Net"
            .TextMatrix(i, ColWGros) = "Weight Gross"
            .TextMatrix(i, ColVol) = "Volume (M3)"
            .TextMatrix(i, ColPacking) = "Qty / Carton"
            .TextMatrix(i, ColCartonNo) = "Carton No"
            .TextMatrix(i, ColLength) = "Length (mm)"
            .TextMatrix(i, ColWidth) = "Width (mm)"
            .TextMatrix(i, ColThickness) = "Thickness (mm)"
            .TextMatrix(i, ColDelDate) = "Delivery Date"
            .TextMatrix(i, ColCurr) = "Curr"
            .TextMatrix(i, ColPrice) = "Price"
            .TextMatrix(i, ColAmount) = "Amount"
            .TextMatrix(i, ColQtyPerCarton) = "Qty / Ctn"
            .TextMatrix(i, ColNetWeight) = "Net"
            .TextMatrix(i, ColGrossWeight) = "Gross"
            .TextMatrix(i, ColPo) = "PO"
            .TextMatrix(i, ColpoSeqNo) = "PoSeqNo"
        Next
        
        .ColHidden(ColProd) = True
        .ColHidden(ColCounter) = True
        .ColHidden(ColCurrT) = True
        .ColHidden(ColUnitT) = True
        .ColHidden(ColRemT) = True
        .ColHidden(ColPos) = True
        .ColHidden(ColPip) = True
        .ColHidden(ColSeq) = True
        .ColHidden(ColQtyPerCarton) = True
        .ColHidden(ColNetWeight) = True
        .ColHidden(ColGrossWeight) = True
        .ColHidden(ColPo) = True
        .ColHidden(ColpoSeqNo) = True
        
        If hakPrice(Me.Name) = 0 Then
            .ColHidden(ColCurr) = True
            .ColHidden(ColPrice) = True
            .ColHidden(ColAmount) = True
        Else
            .ColHidden(ColCurr) = False
            .ColHidden(ColPrice) = False
            .ColHidden(ColAmount) = False
        End If
        
        .Cell(flexcpAlignment, 0, ColCtr, 1, .ColS - 1) = flexAlignCenterCenter
        .MergeRow(0) = True
        .MergeRow(1) = True
        For i = 0 To .ColS - 1
            .MergeCol(i) = True
        Next i
        .MergeCells = flexMergeFixedOnly
        
        .ColAlignment(ColDrySize) = flexAlignLeftCenter
    
    End With

End Sub
Sub AdCboTrans()
'TRANS
Dim rsisi As New ADODB.Recordset
    With CboTrans
        .clear
        .columnCount = 2
        .AddItem ""
        .List(0, 0) = ""
        .List(0, 1) = "00"
        sql = "select * from Transportation_Cls"
        If rsisi.State = 1 Then rsisi.Close
        rsisi.Open sql, Db, adOpenKeyset, adLockOptimistic
        If Not rsisi.EOF And Not rsisi.BOF Then
            rsisi.MoveFirst
            While Not rsisi.EOF
                .AddItem
                .List(.ListCount - 1, 0) = Trim(rsisi.Fields("Description"))
                .List(.ListCount - 1, 1) = rsisi.Fields("Transportation_Cls")
                rsisi.MoveNext
            Wend
        End If
        .ListWidth = 125
        .ColumnWidths = "125pt;0pt"
        .ListIndex = 0
    End With
    
End Sub
Private Sub IsiCombo()
    
    Dim i As Long
    Dim rsisi As New ADODB.Recordset

    'CUST n NOTIFY
    cboCust.clear
    cboCust.columnCount = 2
    cboCust.ColumnWidths = "70 pt;200pt"
    cboCust.ListWidth = 270
    cboCust.ListRows = 15
    
    cboNotify.clear
    cboNotify.columnCount = 2
    cboNotify.ColumnWidths = "70 pt;200 pt"
    cboNotify.ListWidth = 270
    cboNotify.ListRows = 15
    
    cboDelPlace.clear
    cboDelPlace.columnCount = 3
    cboDelPlace.ColumnWidths = "70pt;200pt;0pt"
    cboDelPlace.ListWidth = 270
    cboDelPlace.ListRows = 15
    
  ' --Batas--
  
    If rsisi.State = 1 Then rsisi.Close
    sql = "select trade_code, trade_name, country from trade_master where (trade_cls = 2) " 'and (Country_Cls='1')"
    
    rsisi.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not rsisi.EOF And Not rsisi.BOF Then
        i = 0
        rsisi.MoveFirst
        While Not rsisi.EOF
                    
            cboCust.AddItem
            cboCust.List(i, 0) = Trim(rsisi.Fields("trade_code"))
            cboCust.List(i, 1) = rsisi.Fields("trade_name")
            
            cboNotify.AddItem
            cboNotify.List(i, 0) = Trim(rsisi.Fields("trade_code"))
            cboNotify.List(i, 1) = rsisi.Fields("trade_name")
            
            rsisi.MoveNext
            i = i + 1
        Wend
    End If
    
    
    If rsisi.State = 1 Then rsisi.Close
    sql = "select trade_code, trade_name, country from trade_master where (trade_cls = 4) " 'and (Country_Cls='1')"
    
    rsisi.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not rsisi.EOF And Not rsisi.BOF Then
        i = 0
        rsisi.MoveFirst
        While Not rsisi.EOF
                    
            cboDelPlace.AddItem
            cboDelPlace.List(i, 0) = Trim(rsisi.Fields("trade_code"))
            cboDelPlace.List(i, 1) = rsisi.Fields("trade_name")
            cboDelPlace.List(i, 2) = rsisi.Fields("country")
            
            rsisi.MoveNext
            i = i + 1
        Wend
    End If
   

    AdCboTrans
    'PO PAYMENT CODE
    With cboPaymentCode
        .clear
        .columnCount = 2
        .AddItem ""
        .List(0, 0) = "00"
        .List(0, 1) = ""
        sql = "select * from paymentcode_cls"
        If rsisi.State = 1 Then rsisi.Close
        rsisi.Open sql, Db, adOpenKeyset, adLockOptimistic
        If Not rsisi.EOF And Not rsisi.BOF Then
            rsisi.MoveFirst
            While Not rsisi.EOF
                .AddItem
                .List(.ListCount - 1, 0) = rsisi.Fields("paymentcode_cls")
                .List(.ListCount - 1, 1) = Trim(rsisi.Fields("Description"))
                rsisi.MoveNext
            Wend
        End If
        .ListWidth = 125
        .ColumnWidths = "25pt;100pt"
        .ListIndex = 0
    End With
    
    'PO PAYMENT TERM
    With cboPaymentTerm
        .clear
        .columnCount = 2
        .AddItem ""
        .List(0, 0) = "00"
        .List(0, 1) = ""
        sql = "select * from paymentterm_cls"
        If rsisi.State = 1 Then rsisi.Close
        rsisi.Open sql, Db, adOpenKeyset, adLockOptimistic
        If Not rsisi.EOF And Not rsisi.BOF Then
            rsisi.MoveFirst
            While Not rsisi.EOF
                .AddItem
                .List(.ListCount - 1, 0) = rsisi.Fields("paymentterm_cls")
                .List(.ListCount - 1, 1) = Trim(rsisi.Fields("Description"))
                rsisi.MoveNext
            Wend
        End If
        .ListWidth = 145
        .ColumnWidths = "25 pt;120 pt"
        .ListIndex = 0
    End With
    
    'Packing Type
    Call up_FillCombo(cbopackingtype, "PackingStyle_Cls")
    cbopackingtype.ColumnWidths = "50 pt;90 pt"
    cbopackingtype.ListWidth = 140
    'cbopackingtype.ListIndex = 0
    
    cboPlaceofDestination.AddItem "1"
    cboPlaceofDestination.AddItem "2"
    
    cboPlaceofDestination.Text = "2"
    
End Sub

Private Sub IsiComboPacking()
    
    Dim rsno As New ADODB.Recordset
    
    sql = "select packing_no from packing_master where cust_code = '" & cboCust.Text & "' " & _
    "and etd >= '" & Format(DTDel1.Value, "yyyy-MM-dd") & "' " & _
    " and etd <= '" & Format(DtDel2.Value, "yyyy-MM-dd") & "' "
            If CboPOnO.Text <> "ALL" And CboPOnO <> "" Then
                sql = sql & " AND LIST_DO like'%" & CboPOnO & "%' "
            End If
            sql = sql & " order by packing_no "
            
    rsno.Open sql, Db, adOpenKeyset, adLockOptimistic
    CboPacking.clear
    If Not rsno.EOF Then
        While Not rsno.EOF
            CboPacking.AddItem Trim(rsno.Fields("packing_no"))
            rsno.MoveNext
        Wend
    End If
    rsno.Close
    
End Sub

Private Sub IsiDefaultValue()
    
    Dim adoRs As New ADODB.Recordset
    
    'DtPacking.Value = Date
    DtEtd.Value = Date
    DtEta.Value = Date
    
    'TxtTitle = ""
    TxtForwarder = ""
    TxtVessel.Text = ""
    TxtMotherVessel.Text = ""
    TxtFrom.Text = ""
    TxtTo.Text = ""
    TxtCountry.Text = ""
    TxtFinal.Text = ""
    txtRemarks.Text = ""
    lblfix.Caption = ""
    '''''''''''''''''''''''''''''
    'cbopackingtype.ListIndex = 0
    cbopackingtype.locked = False
    '''''''''''''''''''''''''''''
    
    sql = "Select POPayment_Code, POPayment_Day, POPayment_Terms, POPayment, Transportation_Cls, " & _
        "POCaseMark1, POCaseMark2, POCaseMark3, POCaseMark4, POCaseMark5 " & _
        "From Trade_Master Where Trade_Code = '" & Trim(cboCust.Text) & "'"
    
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        CboTrans.ListIndex = Val(adoRs.Fields("Transportation_Cls") & "")
        cboPaymentCode.ListIndex = Val(adoRs.Fields("POPayment_Code") & "")
        cboPaymentTerm.Text = Trim(adoRs.Fields("POPayment_Terms") & "")
        txtPaymentTerm.Text = adoRs.Fields("POPayment") & ""
        TxtDay.Text = adoRs.Fields("POPayment_Day") & ""
        TxtCaseMark(0).Text = Trim(adoRs.Fields("POCaseMark1") & "")
        TxtCaseMark(1).Text = Trim(adoRs.Fields("POCaseMark2") & "")
        TxtCaseMark(2).Text = Trim(adoRs.Fields("POCaseMark3") & "")
        TxtCaseMark(3).Text = Trim(adoRs.Fields("POCaseMark4") & "")
        TxtCaseMark(4).Text = Trim(adoRs.Fields("POCaseMark5") & "")
    Else
        CboTrans.ListIndex = 0
        cboPaymentCode.ListIndex = 0
        cboPaymentTerm.ListIndex = 0
        txtPaymentTerm.Text = ""
        TxtDay.Text = ""
        TxtCaseMark(0).Text = ""
        TxtCaseMark(1).Text = ""
        TxtCaseMark(2).Text = ""
        TxtCaseMark(3).Text = ""
        TxtCaseMark(4).Text = ""
    End If
    adoRs.Close
    
End Sub

Private Sub GeneratePackingNo()
        
    Dim rsno As New ADODB.Recordset
    Dim VRom(12) As String
    
    VRom(1) = "I": VRom(2) = "II": VRom(3) = "III": VRom(4) = "IV": VRom(5) = "V": VRom(6) = "VI"
    VRom(7) = "VII": VRom(8) = "VIII": VRom(9) = "IX": VRom(10) = "X": VRom(11) = "XI": VRom(12) = "XII"
       
    If rsno.State = 1 Then rsno.Close
    sql = "Select Isnull(Max(left(Packing_No, 3)), 0)  + 1 Nomor " & _
        " From Packing_Master Where Year(Packing_Date) =  " & DtPacking.Year & _
        " And Month(Packing_Date) =  " & DtPacking.Month
        
'    Sql = "Select Isnull(Max(Substring(DO_No, 3, 4)), 0)  + 1 Nomor " & _
'        "From( " & _
'            "Select DO_No From DO_Master Where Year(DO_Date) = " & DtPacking.year & _
'            "Union " & _
'            "Select Packing_No DO_No from Packing_Master Where Year(Packing_Date) =  " & DtPacking.year & _
'        ") a "
    
    If rsno.State <> adStateClosed Then rsno.Close
    rsno.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If rsno.EOF Then
        CboPacking.Text = "001/" & VRom(DtPacking.Month) & "/" & Format(DtPacking, "YY")
    Else
            CboPacking.Text = Format(rsno.Fields("Nomor"), "000") & "/" & VRom(DtPacking.Month) & "/" & Format(DtPacking, "YY")
    End If
    
    If rsno.State <> adStateClosed Then rsno.Close
    TxtPacNo.Text = CboPacking.Text

End Sub

Private Sub IsiDataPacking()
    
    Dim RS As New ADODB.Recordset
    
    RS.Open "select *,(select Description from Transportation_Cls  WHERE Transportation_Cls=a.Transportation_Cls) DESCRIPTION from packing_master a where ltrim(packing_no) = '" & Trim(CboPacking.Text) & "' ", Db, adOpenKeyset, adLockOptimistic
    
    If Not RS.EOF Then
        
        cboDelPlace.Text = Trim(RS.Fields("Consignee"))
        TxtTitle.Text = Trim(RS.Fields("ConsigneeTitle") & "")
        cboNotify.Text = Trim(RS.Fields("Notify_Code"))
        
        DtPacking.Value = Format(RS.Fields("Packing_Date"), "yyyy-MM-dd")
        DtStuffing.Value = Format(RS.Fields("Stuffing_Date"), "yyyy-MM-dd")
        DtEtd.Value = Format(RS.Fields("ETD"), "yyyy-MM-dd")
        DtEta.Value = Format(RS.Fields("ETA"), "yyyy-MM-dd")
        
'        cboTrans.ListIndex = Val(rs.Fields("Transportation_Cls"))
        'AdCboTrans
        
        
        CboTrans.Column(1) = (RS.Fields("Transportation_Cls"))
        CboTrans.Column(0) = IIf(IsNull(RS.Fields("DESCRIPTION")), "", Trim(RS.Fields("DESCRIPTION")))
        'cboTrans = (rs.Fields("DESCRIPTION"))
        cboPaymentCode.ListIndex = Val(RS.Fields("Payment_Code"))
        cboPaymentTerm.Text = Trim(RS.Fields("Payment_Terms") & "")
        
        TxtForwarder.Text = Trim(RS.Fields("Forwarder"))
        TxtVessel.Text = Trim(RS.Fields("Vessel"))
        TxtMotherVessel.Text = Trim(RS.Fields("Mother_Vessel"))
        TxtFrom.Text = Trim(RS.Fields("From_Port"))
        TxtTo.Text = Trim(RS.Fields("To_Port"))
        TxtCountry.Text = Trim(RS.Fields("Country_Origin"))
        TxtFinal.Text = Trim(RS.Fields("Final_Destination"))
        
        txtPaymentTerm.Text = RS.Fields("Payment") & ""
        TxtDay.Text = RS.Fields("Payment_Days")
        
        TxtCaseMark(0).Text = Trim(RS.Fields("POCaseMark1") & "")
        TxtCaseMark(1).Text = Trim(RS.Fields("POCaseMark2") & "")
        TxtCaseMark(2).Text = Trim(RS.Fields("POCaseMark3") & "")
        TxtCaseMark(3).Text = Trim(RS.Fields("POCaseMark4") & "")
        TxtCaseMark(4).Text = Trim(RS.Fields("POCaseMark5") & "")
        txtRemarks.Text = Trim(RS.Fields("Remarks") & "")
        
        cbopackingtype.Text = Trim(RS!PackingStyle_Cls)
        cboPlaceofDestination.Text = IIf(IsNull(RS.Fields("Final_Destination_Cls")), "1", Trim(RS.Fields("Final_Destination_Cls")))
'        cboWH.Text = Trim(rs.Fields("WHCode") & "")
        
        If RS.Fields("Fix_Cls") = "1" Then lblfix.Caption = "Status Fix" Else lblfix.Caption = ""
        TxtPacNo.Text = CboPacking.Text
    
    Else
        
        IsiDefaultValue
    
    End If
    RS.Close
        
End Sub

Private Sub IsiGridHead()

Dim RsIsiG As New ADODB.Recordset
Dim RsIsiD As New ADODB.Recordset
Dim RsCariQty As New ADODB.Recordset
Dim RsQ As New ADODB.Recordset
Dim i, j, k As Long
        
Header
cbopackingtype.locked = False
    
'    #####  Dari DO_Master    & Delivery Order  ######


sql = "  select a.*, isnull(b.grossweight,0) grossweight, isnull(b.netweight,0) netweight,  " & vbCrLf & _
            "   isnull(b.length,0)length, isnull(b.width,0) width, isnull(b.thickness,0) thickness, " & vbCrLf & _
            "   isnull(b.number_entering,0) number_entering  from ( " & vbCrLf & _
            "       select   do_no,  item_code, SerialNoFrom,SerialNoTo,   " & vbCrLf & _
            "               (select makeritem_code from item_master where item_code = tb.item_code) as mki, " & vbCrLf & _
            "               (select im.item_name from item_master as im where im.item_code =  tb.item_code ) as iname, " & vbCrLf & _
            "               (select im.Group_Cls from item_master as im where im.item_code =  tb.item_code ) as Mat_Cls, " & vbCrLf & _
            "               qty,unit_cls,delivery_date,currency_code,Price,DoSeq_no,Po_No, Seq_No from ( " & vbCrLf & _
            "                   select m.cust_code, m.do_no,d.item_code, SerialNoFrom,SerialnoTo,d.qty, " & vbCrLf & _
            "                      d.unit_cls,d.delivery_date,d.currency_code,d.Price,DOSeq_no,d.Po_No,d.Seq_No " & vbCrLf & _
            "                      From DO_Master As m Inner Join Delivery_Order As d on m.Do_no = d.Do_no  "

sql = sql + "                   )tb  where rtrim(do_no) + rtrim(item_code) + cast(DoSeq_no as varchar) in     " & vbCrLf & _
            "                       (select rtrim(de.Do_No) + rtrim(item_code) + cast(DoSeq_no as varchar) " & vbCrLf & _
            "                           from packing_detail as de inner join packing_master as ms  " & vbCrLf & _
            "                               on de.packing_no = ms.packing_no  " & vbCrLf & _
            "                           where ms.cust_code = '" & Trim(cboCust.Text) & "' and ltrim(de.packing_no) = '" & Trim(CboPacking.Text) & "')  " & vbCrLf & _
            "                                   and cust_code = '" & Trim(cboCust.Text) & "' " & vbCrLf & _
            "   union     " & vbCrLf & _
            "       select   do_no,  item_code,SerialNoFrom,SerialNoTo,    " & vbCrLf & _
            "               (select makeritem_code from item_master where item_code = tb.item_code) as mki, " & vbCrLf & _
            "               (select im.item_name from item_master as im where im.item_code =  tb.item_code ) as iname, " & vbCrLf & _
            "               (select im.Group_Cls from item_master as im where im.item_code =  tb.item_code ) as Mat_Cls,qty, "

sql = sql + "               unit_cls,delivery_date,currency_code,Price,DoSeq_no,Po_No,Seq_no   from ( " & vbCrLf & _
            "                   select m.cust_code,m.do_no,d.item_code,SerialNoFrom,SerialNoTo,d.qty,d.unit_cls, " & vbCrLf & _
            "                           d.delivery_date,d.currency_code,d.Price,DoSeq_no,d.Po_No,d.Seq_No " & vbCrLf & _
            "                       From DO_Master As m Inner Join Delivery_Order As d on m.do_no = d.do_no  where fix_cls=1 " & vbCrLf & _
            "                       )tb where cust_code = '" & Trim(cboCust.Text) & "'  and delivery_date >= '" & Format(DTDel1.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
            "                          and delivery_date <= '" & Format(DtDel2.Value, "yyyy-MM-dd") & "'  "
            

If CboPOnO.ListIndex > 0 Then sql = sql & " and Do_no = '" & CboPOnO & "' "
            
sql = sql + "   ) a  " & vbCrLf & _
            "   left outer join (select * from packingitem_master ) b  " & vbCrLf & _
            "       on a.item_code = b.item_code  order by a.Mat_Cls, a.mki  " & vbCrLf

'-----
                      
If RsIsiG.State = 1 Then RsIsiG.Close
RsIsiG.Open sql, Db, adOpenKeyset, adLockOptimistic
If Not RsIsiG.EOF And Not RsIsiG.BOF Then
    RsIsiG.MoveFirst
      i = 2
    j = 2
    
    
    While Not RsIsiG.EOF
        grid.Rows = grid.Rows + 1
        grid.TextMatrix(i, ColAsk) = ""
        grid.TextMatrix(i, ColContainerNo) = ""
        grid.TextMatrix(i, ColDrySize) = ""
        grid.TextMatrix(i, ColOrder) = Trim(RsIsiG.Fields("Do_No"))
        grid.TextMatrix(i, ColProd) = Trim(RsIsiG.Fields("item_code"))
        grid.TextMatrix(i, ColMaker) = Trim(RsIsiG.Fields("mki"))
        grid.TextMatrix(i, ColDesc) = Trim(RsIsiG.Fields("iname"))
        
        If RsCariQty.State = 1 Then RsCariQty.Close
        RsCariQty.Open "select isnull(sum(qty),0) as q from packing_detail where Do_No = '" & RsIsiG.Fields("Do_No") & "' and item_code = '" & RsIsiG.Fields("item_code") & "' and currency_code = '" & RsIsiG.Fields("currency_code") & "' and DoSeq_no = " & RsIsiG.Fields("DoSeq_no") & " and Order_SeqNo=" & RsIsiG.Fields("Seq_no") & " ", Db, adOpenKeyset, adLockOptimistic
        
        If RsQ.State = 1 Then RsQ.Close
        RsQ.Open "select isnull(sum(qty),0) as q from packing_detail where Do_No = '" & RsIsiG.Fields("Do_No") & "' and ltrim(packing_no) = '" & Trim(CboPacking.Text) & "' and item_code = '" & RsIsiG.Fields("item_code") & "' and currency_code = '" & RsIsiG.Fields("currency_code") & "' and DoSeq_no = " & RsIsiG.Fields("DoSeq_no") & " and Order_SeqNo=" & RsIsiG.Fields("Seq_no") & " ", Db, adOpenKeyset, adLockOptimistic
        
        If RsIsiG!number_entering <> 0 Then
            CartonQty = Fix((RsIsiG!Qty - RsCariQty.Fields("q")) / RsIsiG!number_entering)
        Else
            CartonQty = 0
        End If
        
        NetWeight = RsIsiG!NetWeight
        If RsIsiG!number_entering = 0 Then
            NetSisa = 0
        Else
            NetSisa = (((RsIsiG!Qty - RsCariQty.Fields("q")) Mod RsIsiG!number_entering) / RsIsiG!number_entering) * RsIsiG!NetWeight
        End If
        NetWeightPacking = (RsIsiG!NetWeight * CartonQty) + NetSisa
        
        GrossWeight = RsIsiG!GrossWeight
        If RsIsiG!number_entering = 0 Then
            GrossSisa = 0
        Else
            GrossSisa = (((RsIsiG!Qty - RsCariQty.Fields("q")) Mod RsIsiG!number_entering) / RsIsiG!number_entering) * RsIsiG!GrossWeight
        End If
        GrossWeightPacking = (RsIsiG!GrossWeight * CartonQty) + GrossSisa
        
        If RsIsiG!number_entering <> 0 Then
            CartonQty = uf_Ceiling((RsIsiG!Qty - RsCariQty.Fields("q")) / RsIsiG!number_entering)
        Else
            CartonQty = 0
        End If
        Volume = (RsIsiG!Length / 100) * (RsIsiG!Width / 100) * (RsIsiG!Thickness / 100)
        VolumePacking = Format((Volume * CartonQty), gs_formatVolume)
                                 
        grid.TextMatrix(i, ColQty) = Format(0, gs_formatQty)
        grid.TextMatrix(i, colrem) = Format(0, gs_formatQty)
        grid.TextMatrix(i, ColUnit) = uf_GetUnitDescription(RsIsiG("unit_cls"))
        
        If Not IsNull(RsIsiG.Fields("SerialNoFrom")) Then
            grid.TextMatrix(i, ColSerialFrom) = IIf(Trim(RsIsiG.Fields("SerialNoFrom")) = "", "", Trim(RsIsiG.Fields("SerialNoFrom")))
            grid.TextMatrix(i, ColSerialTo) = IIf(Trim(RsIsiG.Fields("SerialNoTo")) = "", "", Trim(RsIsiG.Fields("SerialNoTo")))
        End If
        
        grid.TextMatrix(i, ColWeight) = Format(NetWeight, gs_formatWeight)
        grid.TextMatrix(i, ColWGros) = Format(GrossWeight, gs_formatWeight)
        grid.TextMatrix(i, ColVol) = Format(Volume, gs_formatVolume)
        grid.TextMatrix(i, ColPacking) = Format(RsIsiG!number_entering, gs_formatLength)
        grid.TextMatrix(i, ColCartonNo) = ""
        grid.TextMatrix(i, ColLength) = Format(RsIsiG!Length, gs_formatLength)
        grid.TextMatrix(i, ColWidth) = Format(RsIsiG!Width, gs_formatWidth)
        grid.TextMatrix(i, ColThickness) = Format(RsIsiG!Thickness, gs_formatThickness)
        grid.TextMatrix(i, ColDelDate) = Format(RsIsiG.Fields("delivery_date"), "dd MMM yyyy")
        grid.TextMatrix(i, ColCurr) = uf_GetCurrencyDescription(RsIsiG.Fields("currency_code"))
        grid.TextMatrix(i, ColPrice) = Format(RsIsiG.Fields("Price"), gs_formatPrice)
        grid.TextMatrix(i, ColAmount) = Format(0, gs_formatAmountIDR)
        grid.TextMatrix(i, ColCounter) = "HEAD"
        grid.TextMatrix(i, ColCurrT) = RsIsiG.Fields("currency_code")
        grid.TextMatrix(i, ColUnitT) = RsIsiG.Fields("unit_cls")
        grid.TextMatrix(i, ColRemT) = RsIsiG.Fields("qty") - RsCariQty.Fields("q") + RsQ.Fields("q")
        grid.TextMatrix(i, ColPos) = ""
        grid.TextMatrix(i, ColSeq) = RsIsiG.Fields("DoSeq_no")
        grid.TextMatrix(i, ColQtyPerCarton) = RsIsiG!number_entering
        
        grid.TextMatrix(i, ColPo) = RsIsiG.Fields("Po_No")
        grid.TextMatrix(i, ColpoSeqNo) = RsIsiG.Fields("Seq_no")


        
        grid.Cell(flexcpBackColor, i, ColCtr, i, ColAmount) = &HE0E0E0
        grid.Cell(flexcpBackColor, i, ColAsk) = &HFFFFFF
        
        'ISI CHILD
        If RsIsiD.State = 1 Then RsIsiD.Close
        sql = " select " & vbCrLf & _
              " chk, " & vbCrLf & _
              " Container_No, Container_Size, " & vbCrLf & _
              " (select item_name from item_master where item_code = tb.item_code) as ina, " & vbCrLf & _
              " Do_No, " & vbCrLf & _
              " item_code, SerialNoFrom,SerialNoTo," & vbCrLf & _
              " unit_cls, " & vbCrLf & _
              " qty, " & vbCrLf & _
              " qtyweight_netto, " & vbCrLf & _
              " qtyweight_gross, " & vbCrLf & _
              " detail_cls, " & vbCrLf & _
              " qty_volume, Qty_Ctn, Ctn_No, length, width, thickness, " & vbCrLf & _
              " (select top 1 delivery_date from  Delivery_Order where Do_No = tb.Do_No) as deldate, " & vbCrLf & _
              " currency_code, "

        sql = sql + " price, " & vbCrLf & _
                      " amount, po_no, seq_no " & vbCrLf & _
                      " from ( " & vbCrLf & _
                      "     select  " & vbCrLf & _
                      "     1 as chk, " & vbCrLf & _
                      "     pd.Container_No, pd.Container_Size, " & vbCrLf & _
                      "     pd.Do_No, " & vbCrLf & _
                      "     pd.item_code, SerialNoFrom,SerialNoTo," & vbCrLf & _
                      "     pd.unit_cls, " & vbCrLf & _
                      "     pd.qty, " & vbCrLf & _
                      "     pd.qtyweight_netto, " & vbCrLf & _
                      "     pd.qtyweight_gross, " & vbCrLf & _
                      "     pd.detail_cls, "
        
        sql = sql + "     pd.qty_volume, pd.Qty_Ctn, pd.Ctn_No, pd.length, pd.width, pd.thickness, " & vbCrLf & _
                      "     pd.currency_code, " & vbCrLf & _
                      "     pd.price, " & vbCrLf & _
                      "     pd.amount, pd.order_no as po_no, pd.order_seqno as seq_no" & vbCrLf & _
                      "     from packing_detail as pd inner join packing_master as pm " & vbCrLf & _
                      "     on pd.packing_no = pm.packing_no " & vbCrLf & _
                      "     where ltrim(pm.packing_no) = '" & UCase(CboPacking.Text) & "' " & vbCrLf & _
                      "     and pm.cust_code = '" & UCase(cboCust.Text) & "' " & vbCrLf & _
                      "     and pd.Do_No = '" & UCase(RsIsiG.Fields("Do_No")) & "' " & vbCrLf & _
                      "     and pd.item_code = '" & UCase(RsIsiG.Fields("item_code")) & "'  " & vbCrLf & _
                      "     and pd.currency_code = '" & RsIsiG.Fields("currency_code") & "' " & vbCrLf & _
                      "     and pd.DoSeq_no = " & RsIsiG.Fields("DoSeq_no") & "" & vbCrLf & _
                      "     and pd.Order_Seqno = " & RsIsiG.Fields("Seq_no") & "" & vbCrLf & _
                      "     )tb "
                      
sql = "  select  " & vbCrLf & _
            "  chk,  " & vbCrLf & _
            "  Container_No, Container_Size,  " & vbCrLf & _
            "  (select item_name from item_master where item_code = tb.item_code) as ina,  " & vbCrLf & _
            "  Do_No,  " & vbCrLf & _
            "  item_code, SerialNoFrom,SerialNoTo, " & vbCrLf & _
            "  unit_cls,  " & vbCrLf & _
            "  qty,  " & vbCrLf & _
            "  qtyweight_netto,  " & vbCrLf & _
            "  qtyweight_gross,  " & vbCrLf & _
            "  detail_cls,  "

sql = sql + "  qty_volume, Qty_Ctn, Ctn_No, length, width, thickness,  " & vbCrLf & _
            "  (select top 1 delivery_date from  Delivery_Order where Do_No = tb.Do_No) as deldate,  " & vbCrLf & _
            "  currency_code,  price,  " & vbCrLf & _
            "  amount, po_no, seq_no  " & vbCrLf & _
            "  from (  " & vbCrLf & _
            "      select   " & vbCrLf & _
            "      1 as chk,  " & vbCrLf & _
            "      pd.Container_No, pd.Container_Size,  " & vbCrLf & _
            "      pd.Do_No,  " & vbCrLf & _
            "      pd.item_code, SerialNoFrom,SerialNoTo, " & vbCrLf & _
            "      pd.unit_cls,  "

sql = sql + "      pd.qty,  " & vbCrLf & _
            "      Coalesce(case when pd.qtyweight_netto <=0 then (select NetWeight From PackingItem_Master where Item_code=pd.item_code) end,0)  qtyweight_netto,  " & vbCrLf & _
            "      Coalesce(case when pd.qtyweight_gross <=0 then (select GrossWeight From PackingItem_Master where Item_code=pd.item_code) end,0)  qtyweight_gross,  " & vbCrLf & _
            "      pd.detail_cls,      pd.qty_volume, pd.Qty_Ctn, pd.Ctn_No,  " & vbCrLf & _
            "    Coalesce(case when pd.length <=0 then (select Length From PackingItem_Master where Item_code=pd.item_code) end,0)  length,  " & vbCrLf & _
            "    Coalesce(case when pd.width <=0 then (select Width From PackingItem_Master where Item_code=pd.item_code) end,0)  width,  " & vbCrLf & _
            "    Coalesce(case when pd.thickness <=0 then (select thickness From PackingItem_Master where Item_code=pd.item_code) end,0)  thickness,  " & vbCrLf & _
            "      pd.currency_code,  " & vbCrLf & _
            "      pd.price,  " & vbCrLf & _
            "      pd.amount, pd.order_no as po_no, pd.order_seqno as seq_no " & vbCrLf & _
            "      from packing_detail as pd inner join packing_master as pm  "

sql = sql + "      on pd.packing_no = pm.packing_no  " & vbCrLf & _
            "     where ltrim(pm.packing_no) = '" & UCase(CboPacking.Text) & "' " & vbCrLf & _
            "     and pm.cust_code = '" & UCase(cboCust.Text) & "' " & vbCrLf & _
            "     and pd.Do_No = '" & UCase(RsIsiG.Fields("Do_No")) & "' " & vbCrLf & _
            "     and pd.item_code = '" & UCase(RsIsiG.Fields("item_code")) & "'  " & vbCrLf & _
            "     and pd.currency_code = '" & RsIsiG.Fields("currency_code") & "' " & vbCrLf & _
            "     and pd.DoSeq_no = " & RsIsiG.Fields("DoSeq_no") & "" & vbCrLf & _
            "     and pd.Order_Seqno = " & RsIsiG.Fields("Seq_no") & "" & vbCrLf & _
            "     )tb "

        
        If RsIsiD.State = 1 Then RsIsiD.Close
        RsIsiD.Open sql, Db, adOpenKeyset, adLockOptimistic
        k = 0
        j = i
        If Not RsIsiD.EOF And Not RsIsiD.BOF Then
            If cbopackingtype.locked = False Then cbopackingtype.locked = True
            RsIsiD.MoveFirst
            While Not RsIsiD.EOF
                If RsIsiD.Fields("detail_cls") = "1" Then
                    grid.Rows = grid.Rows + 1
                    j = j + 1
                    k = k + 1
                    grid.Cell(flexcpChecked, j, ColCtr) = flexChecked
                    grid.TextMatrix(j, ColAsk) = ""
                    grid.TextMatrix(j, ColContainerNo) = Trim(RsIsiD.Fields("Container_No"))
                    grid.TextMatrix(j, ColDrySize) = Trim(RsIsiD.Fields("Container_Size"))
                    grid.TextMatrix(j, ColOrder) = ""
                    grid.TextMatrix(j, ColProd) = ""
                    grid.TextMatrix(j, ColDesc) = ""
                    grid.TextMatrix(j, ColQty) = Format(RsIsiD.Fields("qty"), gs_formatQty)
                    grid.TextMatrix(j, colrem) = ""
                    grid.TextMatrix(j, ColUnit) = uf_GetUnitDescription(RsIsiD("unit_cls"))
                    
                    grid.TextMatrix(j, ColSerialFrom) = IIf(IsNull(RsIsiD.Fields("SerialNoFrom")), "", Trim(RsIsiD.Fields("SerialNoFrom")))
                    grid.TextMatrix(j, ColSerialTo) = IIf(IsNull(RsIsiD.Fields("SerialNoTo")), "", Trim(RsIsiD.Fields("SerialNoTo")))
                    
                    grid.TextMatrix(j, ColWeight) = Format(RsIsiD.Fields("qtyweight_netto"), gs_formatWeight)
                    grid.TextMatrix(j, ColWGros) = Format(RsIsiD.Fields("qtyweight_gross"), gs_formatWeight)
                    grid.TextMatrix(j, ColVol) = Format(RsIsiD.Fields("qty_volume"), gs_formatVolume)
                    grid.TextMatrix(j, ColPacking) = Format(RsIsiD.Fields("Qty_Ctn"), gs_formatBox)
                    grid.TextMatrix(j, ColCartonNo) = IIf(IsNull(RsIsiD.Fields("Ctn_No")), "", RsIsiD.Fields("Ctn_No"))
                    grid.TextMatrix(j, ColLength) = Format(RsIsiD.Fields("length"), gs_formatLength)
                    grid.TextMatrix(j, ColWidth) = Format(RsIsiD.Fields("width"), gs_formatWidth)
                    grid.TextMatrix(j, ColThickness) = Format(RsIsiD.Fields("thickness"), gs_formatThickness)
                    grid.TextMatrix(j, ColDelDate) = ""
                    grid.TextMatrix(j, ColCurr) = uf_GetCurrencyDescription(RsIsiG.Fields("currency_code"))
                    grid.TextMatrix(j, ColPrice) = Format(RsIsiD.Fields("Price"), gs_formatPrice)
                    grid.TextMatrix(j, ColAmount) = Format(RsIsiD.Fields("Amount"), gs_formatAmountIDR)
                    grid.TextMatrix(j, ColCounter) = "CHILD"
                    grid.TextMatrix(j, ColCurrT) = RsIsiG.Fields("currency_code")
                    grid.TextMatrix(j, ColUnitT) = RsIsiG.Fields("unit_cls")
                    grid.TextMatrix(j, ColRemT) = "0"
                    grid.TextMatrix(j, ColPos) = k
                    grid.TextMatrix(j, ColPip) = "1"
                    grid.TextMatrix(j, ColQtyPerCarton) = RsIsiG!number_entering
                    grid.TextMatrix(j, ColNetWeight) = NetWeight
                    grid.TextMatrix(j, ColGrossWeight) = GrossWeight
                    
                    grid.TextMatrix(j, ColPo) = RTrim(RsIsiG.Fields("Po_No"))
                    grid.TextMatrix(j, ColpoSeqNo) = RsIsiG.Fields("Seq_no")

                    
                    grid.Cell(flexcpBackColor, j, ColCtr) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColContainerNo, j, ColAmount) = &H80000018
                    grid.Cell(flexcpBackColor, j, ColContainerNo) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColDrySize) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColQty) = &HFFFFFF
                    
                    'Grid.Cell(flexcpBackColor, J, ColSerialFrom) = &HFFFFFF
                    'Grid.Cell(flexcpBackColor, J, ColSerialTo) = &HFFFFFF
                    
                    grid.Cell(flexcpBackColor, j, ColWeight) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColWGros) = &HFFFFFF
                    'grid.Cell(flexcpBackColor, J, ColVol) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColPacking) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColCartonNo) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColLength) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColWidth) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColThickness) = &HFFFFFF
                    
                    grid.TextMatrix(i, ColPos) = k
                    
                    
                    If grid.TextMatrix(j, ColCounter) = "CHILD" Then
                    TotalHead CDbl(ColQty), CDbl(j), CDbl(colrem), 1
'                        TotalHead CDbl(ColWeight), CDbl(j), CDbl(ColWeight), 0
'                        TotalHead CDbl(ColWGros), CDbl(j), CDbl(ColWGros), 0
'                        TotalHead CDbl(ColVol), CDbl(j), CDbl(ColVol), 0
'                        TotalHead CDbl(ColPacking), CDbl(j), CDbl(ColPacking), 0
                    End If
                ElseIf RsIsiD.Fields("detail_cls") = "0" Then
                End If
                RsIsiD.MoveNext
            Wend
            
        Else
            
isitambah:
                k = k + 1
                If RsIsiG.Fields("qty") - RsCariQty.Fields("q") > 0 Then
                    grid.Rows = grid.Rows + 1
                    j = j + 1
                    grid.Cell(flexcpChecked, j, ColCtr) = flexUnchecked
                    grid.TextMatrix(j, ColAsk) = ""
                    grid.TextMatrix(j, ColContainerNo) = ""
                    grid.TextMatrix(j, ColDrySize) = ""
                    grid.TextMatrix(j, ColOrder) = ""
                    grid.TextMatrix(j, ColProd) = ""
                    grid.TextMatrix(j, ColDesc) = ""
                    grid.TextMatrix(j, ColQty) = Format(CDbl(RsIsiG.Fields("qty")) - CDbl(RsCariQty.Fields("q")), gs_formatQty)
                    grid.TextMatrix(j, colrem) = ""
                    grid.TextMatrix(j, ColUnit) = grid.TextMatrix(i, ColUnit)
                    
                    If Not IsNull(RsIsiG.Fields("SerialNoFrom")) Then
                        grid.TextMatrix(j, ColSerialFrom) = IIf(Trim(RsIsiG.Fields("SerialNoFrom")) = "", "", Trim(RsIsiG.Fields("SerialNoFrom")))
                        grid.TextMatrix(j, ColSerialTo) = IIf(Trim(RsIsiG.Fields("SerialNoTo")) = "", "", Trim(RsIsiG.Fields("SerialNoTo")))
                    End If
                    
                    grid.TextMatrix(j, ColWeight) = Format(NetWeightPacking, gs_formatWeight)
                    grid.TextMatrix(j, ColWGros) = Format(GrossWeightPacking, gs_formatWeight)
                    grid.TextMatrix(j, ColVol) = Format(VolumePacking, gs_formatVolume)
                    grid.TextMatrix(j, ColPacking) = Format(CartonQty, gs_formatBox)
                    grid.TextMatrix(j, ColCartonNo) = ""
                    grid.TextMatrix(j, ColLength) = Format(RsIsiG.Fields("length"), gs_formatLength)
                    grid.TextMatrix(j, ColWidth) = Format(RsIsiG.Fields("width"), gs_formatWidth)
                    grid.TextMatrix(j, ColThickness) = Format(RsIsiG.Fields("thickness"), gs_formatThickness)
                    grid.TextMatrix(j, ColDelDate) = ""
                    grid.TextMatrix(j, ColCurr) = uf_GetCurrencyDescription(RsIsiG.Fields("currency_code"))
                    grid.TextMatrix(j, ColPrice) = Format(RsIsiG.Fields("Price"), gs_formatPrice)
                    grid.TextMatrix(j, ColAmount) = Format(0, gs_formatAmountIDR)
                    grid.TextMatrix(j, ColCounter) = "CHILD"
                    grid.TextMatrix(j, ColCurrT) = RsIsiG.Fields("currency_code")
                    grid.TextMatrix(j, ColUnitT) = RsIsiG.Fields("unit_cls")
                    grid.TextMatrix(j, ColRemT) = ""
                    grid.TextMatrix(j, ColPos) = k
                    grid.TextMatrix(j, ColQtyPerCarton) = RsIsiG!number_entering
                    grid.TextMatrix(j, ColNetWeight) = NetWeight
                    grid.TextMatrix(j, ColGrossWeight) = GrossWeight
                    
                    grid.TextMatrix(i, ColPo) = RsIsiG.Fields("Po_No")
                    grid.TextMatrix(i, ColpoSeqNo) = RsIsiG.Fields("Seq_no")
                    
                    
                    grid.Cell(flexcpBackColor, j, ColCtr) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColContainerNo, j, ColAmount) = &H80000018
                    grid.Cell(flexcpBackColor, j, ColContainerNo) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColDrySize) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColQty) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColWeight) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColWGros) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColVol) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColPacking) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColCartonNo) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColLength) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColWidth) = &HFFFFFF
                    grid.Cell(flexcpBackColor, j, ColThickness) = &HFFFFFF
                Else
                    'k = k - 1
                End If
            End If
            
            grid.TextMatrix(i, ColPos) = k
            
            If grid.TextMatrix(j, ColCounter) = "CHILD" Then
            TotalHead CDbl(ColQty), CDbl(j), CDbl(colrem), 1
'            TotalHead CDbl(ColWeight), CDbl(j), CDbl(ColWeight), 0
'            TotalHead CDbl(ColWGros), CDbl(j), CDbl(ColWGros), 0
'            TotalHead CDbl(ColVol), CDbl(j), CDbl(ColVol), 0
'            TotalHead CDbl(ColPacking), CDbl(j), CDbl(ColPacking), 0
            End If
            
            RsIsiG.MoveNext
            i = j
            i = i + 1
        Wend
        
        
    End If

End Sub

Private Sub IsiComboPO()

    Dim rscbo As New ADODB.Recordset
    
    With CboPOnO
        .clear
        
        .columnCount = 3
        .TextColumn = 1
        
        .AddItem strAll
        
sql = " select distinct DM.Do_No, Dm.Cust_Code,isnull(om.Location_Code,DM.Cust_Code) Consignee " & vbCrLf & _
            " from DO_Master DM inner join Delivery_Order DO on DM.Do_No=DO.Do_no " & vbCrLf & _
            " inner join OrderEntry_Master OM on DO.Po_No=OM.Po_No " & vbCrLf & _
            " where DM.cust_code = '" & cboCust & _
            "' and Do_Date >='" & Format(DTDel1, "yyyy-MM-dd") & _
            "' and Do_Date <= '" & Format(DtDel2, "yyyy-MM-dd") & _
            "' and (DM.Fix_Cls = 1) " & _
            "order by DM.Do_NO"
       
        Set rscbo = Db.Execute(sql)
            
        i = 1
        Do While Not (rscbo.EOF)
            .AddItem ""
            .List(i, 0) = Trim(rscbo("Do_No"))
            .List(i, 1) = Trim(rscbo("Cust_Code"))
            .List(i, 1) = Trim(rscbo("Consignee"))
            i = i + 1
            rscbo.MoveNext
        Loop
        
        .Text = ""
        .ListWidth = 150
        .ColumnWidths = "150 pt;0 pt;0 pt "
        .ListIndex = -1
        Set rscbo = Nothing
    End With

End Sub

Private Sub clearmark(strNilai As String)

    Dim k As Long
    If strNilai = "D" Then
        For k = 2 To grid.Rows - 1
            If grid.TextMatrix(k, 1) = "S" Then
                grid.TextMatrix(k, 1) = ""
            End If
        Next
    Else
        For k = 2 To grid.Rows - 1
                grid.TextMatrix(k, 1) = ""
        Next
    End If

End Sub

Private Sub TotalHead(colubah As Long, rowubah As Long, colhasil As Long, chk As Long)

    Dim rowstart As Long
    Dim X As Double
    
Atas:
    rowstart = (rowubah - CDbl(grid.TextMatrix(rowubah, ColPos)))
    
    grid.TextMatrix(rowstart, colubah) = 0
    
    If chk = 1 Then
        grid.TextMatrix(rowstart, ColAmount) = 0
        grid.TextMatrix(rowubah, ColAmount) = uf_Trunc(CDbl(grid.TextMatrix(rowubah, colubah)) * CDbl(grid.TextMatrix(rowubah, ColPrice)), gi_decimalDigitAmountIDR)
        grid.TextMatrix(rowubah, ColAmount) = Format(grid.TextMatrix(rowubah, ColAmount), gs_formatAmountIDR)
    End If
    X = 0
    For i = rowstart To CDbl(rowstart + CDbl(grid.TextMatrix(rowstart, ColPos)))
        
        grid.TextMatrix(rowstart, colubah) = CDbl(grid.TextMatrix(rowstart, colubah)) + CDbl(IIf(Trim(grid.TextMatrix(i, colubah)) = "", 0, grid.TextMatrix(i, colubah)))
        If grid.TextMatrix(i, ColPip) <> "XAXAXA" Then
            X = X + CDbl(IIf(Trim(grid.TextMatrix(i, colubah)) = "", 0, grid.TextMatrix(i, colubah)))
        End If
        If chk = 1 Then
            grid.TextMatrix(rowstart, ColAmount) = CDbl(grid.TextMatrix(rowstart, ColAmount)) + CDbl(grid.TextMatrix(i, ColAmount))
            grid.TextMatrix(rowstart, ColAmount) = Format(grid.TextMatrix(rowstart, ColAmount), gs_formatAmountIDR)
        End If
    Next
    
    If colubah = ColQty Then
        grid.TextMatrix(rowstart, colubah) = Format(grid.TextMatrix(rowstart, colubah), gs_formatQty)
    ElseIf colubah = ColWeight Or colubah = ColWGros Then
        grid.TextMatrix(rowstart, colubah) = Format(grid.TextMatrix(rowstart, colubah), gs_formatWeight)
    ElseIf colubah = ColPacking Then
        grid.TextMatrix(rowstart, colubah) = Format(grid.TextMatrix(rowstart, colubah), gs_formatBox)
    ElseIf colubah = ColVol Then
        grid.TextMatrix(rowstart, colubah) = Format(grid.TextMatrix(rowstart, colubah), gs_formatVolume)
    ElseIf colubah = ColLength Then
        grid.TextMatrix(rowstart, colubah) = Format(grid.TextMatrix(rowstart, colubah), gs_formatLength)
    ElseIf colubah = ColWidth Then
        grid.TextMatrix(rowstart, colubah) = Format(grid.TextMatrix(rowstart, colubah), gs_formatWidth)
    ElseIf colubah = ColThickness Then
        grid.TextMatrix(rowstart, colubah) = Format(grid.TextMatrix(rowstart, colubah), gs_formatThickness)
    End If
    
    If chk = 1 Then
        grid.TextMatrix(rowstart, colhasil) = CDbl(grid.TextMatrix(rowstart, ColRemT)) - X
        grid.TextMatrix(rowstart, colhasil) = Format(grid.TextMatrix(rowstart, colhasil), gs_formatQty)
    End If

End Sub

Private Sub GrandTotal()

    Dim ttam, ttqty, ttctnqty, ttweight, ttweigtg, ttvol As Double
    
    ttam = 0
    ttctnqty = 0
    ttqty = 0
    ttweight = 0
    ttweigtg = 0
    ttvol = 0
    
    For i = 2 To grid.Rows - 1
        If grid.TextMatrix(i, ColCounter) = "CHILD" Then
            If grid.Cell(flexcpChecked, i, ColCtr) = flexChecked Then
                ttam = ttam + CDbl(grid.TextMatrix(i, ColAmount))
                ttqty = ttqty + CDbl(grid.TextMatrix(i, ColQty))
                ttctnqty = ttctnqty + CDbl(grid.TextMatrix(i, ColPacking))
                ttweight = ttweight + CDbl(grid.TextMatrix(i, ColWeight))
                ttweigtg = ttweigtg + CDbl(grid.TextMatrix(i, ColWGros))
                ttvol = ttvol + CDbl(grid.TextMatrix(i, ColVol))
            End If
        End If
    Next
    
    TxtTTAm = Format(ttam, gs_formatAmountIDR)
    TxtTTQty = Format(ttqty, gs_formatQty)
    TxtTTCtnQty = Format(ttctnqty, gs_formatBox)
    TxtTTW = Format(ttweight, gs_formatWeight)
    TxtTTWG = Format(ttweigtg, gs_formatWeight)
    TxtTTV = Format(ttvol, gs_formatVolume)

End Sub

Private Sub ClearPacking()

'    Dim RsD As New ADODB.Recordset
'    If RsD.State = 1 Then RsD.Close
'    RsD.Open "select * from packing_detail where packing_no = '" & TxtPacNo.Text & "'", Db, adOpenKeyset, adLockOptimistic
'    If RsD.EOF And RsD.BOF Then
'        Db.BeginTrans
'        sql = "delete from packing_master where packing_no = '" & TxtPacNo.Text & "'"
'        Db.Execute sql
'        Db.CommitTrans
'    End If

'    Db.BeginTrans
    sql = "delete from packing_master where packing_no not in (select packing_no from packing_detail)"
    Db.Execute sql
'    Db.CommitTrans

End Sub

Private Sub ClearForm()
    
    cboCust.Text = ""
    
End Sub

Private Function CheckPONo(strTempNo As String) As Boolean

'    Dim intRow As Integer

    CheckPONo = True
'    For intRow = 1 To grid.Rows - 1
'        If grid.Cell(flexcpChecked, intRow, ColCtr) = flexChecked And Trim(strTempNo) <> Trim(grid.TextMatrix(intRow - Val(grid.TextMatrix(intRow, ColPos)), ColOrder)) Then
'            LblErrMsg = DisplayMsg("0012")
'            CheckPONo = False
'            Exit For
'        End If
'    Next

End Function

Private Sub CboCust_Change()
    
    LblErrMsg = ""
    DTDel1.Value = Date
    DtDel2.Value = Date
    DtStuffing.Value = Date
    DtPacking.Value = Date
    DtEtd.Value = Date
    DtEta.Value = Date
    cboNotify = Trim(cboCust)
    'cboDelPlace = Trim(cbocust)
    
    If cboCust.MatchFound Then
        lblcust.Text = cboCust.List(cboCust.ListIndex, 1)
        'update by dudi s,by Januari 2009 Untuk mmpercepat proses
        'agar tak semua customer masuk ke fungsi IsiCombopo
        sql = "select distinct Do_No, Cust_Code " & _
            "from DO_Master where " & _
            "cust_code = '" & cboCust & _
            "' and Do_Date >='" & Format(DTDel1, "yyyy-MM-dd") & _
            "' and Do_Date <= '" & Format(DtDel2, "yyyy-MM-dd") & _
            "' and (Fix_Cls = 1) " & _
            "order by Do_NO"
        If CekSql(sql) Then
            IsiComboPO
        Else
            CboPOnO.clear
        End If
        If CboStatus.Text = "Create" Then
            GeneratePackingNo
            IsiDefaultValue
        Else
            IsiComboPacking
        End If
    Else
        lblcust.Text = ""
        CboPOnO.clear
        CboPacking.clear
    End If
    
End Sub

Private Sub cboCust_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CboCust_Change
    End If
End Sub

Private Sub cboDelPlace_Change()

    LblErrMsg = ""
    If cboDelPlace.MatchFound Then
        lblDelPlace.Caption = cboDelPlace.List(cboDelPlace.ListIndex, 1)
        TxtFinal = cboDelPlace.List(cboDelPlace.ListIndex, 2)
    Else
        lblDelPlace.Caption = ""
        TxtFinal = ""
    End If

End Sub

Private Sub cboNotify_Change()
    
    If cboNotify.MatchFound Then
        lblNotify.Text = cboNotify.List(cboNotify.ListIndex, 1)
    Else
        lblNotify.Text = ""
    End If
    
End Sub

Private Sub cboPacking_Change()
    TxtPacNo.Text = CboPacking.Text
    If CboPacking.MatchFound Then
        IsiDataPacking
    Else
        IsiDefaultValue
    End If
    
    Header
    GrandTotal
    
End Sub

Private Sub cboPacking_DblClick(Cancel As MSForms.ReturnBoolean)

    If CboPacking.locked Then CboPacking.locked = False Else CboPacking.locked = True

End Sub

Private Sub cboPaymentTerm_Click()
    
    If cboPaymentTerm.MatchFound Then
        lblTerm.Caption = cboPaymentTerm.List(cboPaymentTerm.ListIndex, 1)
        txtPaymentTerm.Text = cboPaymentTerm.List(cboPaymentTerm.ListIndex, 1)
    Else
        lblTerm.Caption = ""
        txtPaymentTerm.Text = ""
    End If
    
End Sub

Private Sub cboPlaceofDestination_Change()
    If cboPlaceofDestination.ListIndex = 0 Then
        lblPlaceofDestination.Caption = "JAPAN"
    ElseIf cboPlaceofDestination.ListIndex = 1 Then
        lblPlaceofDestination.Caption = "OVERSEAS"
    End If
End Sub

Private Sub CboPOnO_Change()
    If CboPOnO.MatchFound And CboPOnO.ListIndex > 0 And Not CboPacking.MatchFound Then
        cboNotify = CboPOnO.Column(1)
        cboDelPlace = CboPOnO.Column(1)
    End If
End Sub

Private Sub cbopono_Click()
        'Dudi S, Januari 2009, untuk mengecek cbo packing yang adas berdasar DO- Number
            
    
    sql = "select packing_no from packing_master where cust_code = '" & cboCust.Text & "' " & vbCrLf & _
            "and etd >= '" & Format(DTDel1.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
            "and etd <= '" & Format(DtDel2.Value, "yyyy-MM-dd") & "' "
            If CboPOnO.Text <> "ALL" Then
            sql = sql & " AND LIST_DO like '%" & CboPOnO & "%' "
            End If
            sql = sql & " order by packing_no "

        If CekSql(sql) Then
            IsiComboPacking
        Else
            CboPacking.clear
        End If
End Sub

Private Sub cboStatus_Click()
'On Error GoTo MsgError
LblErrMsg = ""
MousePointer = vbHourglass
    
    CmdCreate.Caption = CboStatus.Text
    'cboPacking.locked = (cboStatus.Text = "Create")
        
    If CboStatus.Text = "Create" Then
        ClearPacking
        'GeneratePackingNo
        IsiDefaultValue
        CboPacking.clear
        CboPacking.Text = ""
    Else
        IsiComboPacking
        CboPacking.Text = ""
    End If
    
    Header
    GrandTotal
    MousePointer = vbDefault
Exit Sub
MsgError:
MousePointer = vbDefault
LblErrMsg = err.number & " " & err.Description
End Sub

Private Sub cboTrans_Change()
    Select Case CboTrans.ListIndex
    Case 1
        lblCaption(9).Caption = "First Flight"
        lblCaption(10).Caption = "Second Flight"
    Case 3
        lblCaption(9).Caption = "Feeder"
        lblCaption(10).Caption = "AWB Number"
    Case Else
        lblCaption(9).Caption = "Feeder Vessel"
        lblCaption(10).Caption = "Mother Vessel"
    End Select
End Sub

Private Sub CboTrans_Click()
    Select Case CboTrans.ListIndex
    Case 1
        lblCaption(9).Caption = "First Flight"
        lblCaption(10).Caption = "Second Flight"
    Case 3
        lblCaption(9).Caption = "Feeder"
        lblCaption(10).Caption = "AWB Number"
    Case Else
        lblCaption(9).Caption = "Feeder Vessel"
        lblCaption(10).Caption = "Mother Vessel"
    End Select

End Sub

Private Sub cboWH_Change()
    If cboWH.MatchFound Then
        lblWH.Caption = cboWH.Column(1)
    Else
        lblWH.Caption = ""
    End If
End Sub

Private Sub cmdClear_Click()
    
    cboCust.Text = ""
    TxtPacNo.Text = CboPacking.Text
    If CboStatus.ListIndex = 1 Then CboStatus.ListIndex = 0
    IsiDefaultValue
    
End Sub

Private Sub CmdCreate_Click()
    
    Dim adoRs As New ADODB.Recordset
    
    LblErrMsg = ""
    Me.MousePointer = vbHourglass
    
    If Trim(CboPacking.Text) = "" Then
        LblErrMsg = DisplayMsg(4010)
        Me.MousePointer = vbDefault
        CboPacking.SetFocus
        Exit Sub
    End If
    
    If cboCust.MatchFound = False Then
        LblErrMsg = DisplayMsg(4072)
        cboCust.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    If CboPacking = "" Then
    LblErrMsg = DisplayMsg(8130)
        cboCust.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    
    End If
    If cboPaymentTerm = "" Then
        LblErrMsg = DisplayMsg(8123)
        cboPaymentTerm.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    ElseIf cboPaymentTerm.MatchFound = False Then
        LblErrMsg = DisplayMsg(8123)
        cboPaymentTerm.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    If cboDelPlace.MatchFound = False Then
        LblErrMsg = DisplayMsg(4072)
        cboDelPlace.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    'SAVE HEAD
    If CboStatus.ListIndex = 1 Then 'Create Baru
    
        If Trim(CboPOnO.Text) = "" Then
            LblErrMsg = DisplayMsg(1048)
            Me.MousePointer = vbDefault
            CboPOnO.SetFocus
            Exit Sub
        End If
        If Trim(cbopackingtype.Text) = "" Then
            LblErrMsg = DisplayMsg(8130)
            MousePointer = vbDefault
            cbopackingtype.SetFocus
            Exit Sub
        ElseIf cbopackingtype.MatchFound = False Then
            LblErrMsg = DisplayMsg(8131)
            MousePointer = vbDefault
            CboPacking.SetFocus
            Exit Sub
        End If
        
        sql = "select packing_no from packing_master  where packing_no  = '" & Trim(CboPacking) & "'"
        adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not adoRs.EOF Then
            LblErrMsg = DisplayMsg(1023)
            Me.MousePointer = vbDefault
            adoRs.Close
            Set adoRs = Nothing
            CboPacking.SetFocus
            Exit Sub
        End If
        
        
        adoRs.Close
        Set adoRs = Nothing

        Db.BeginTrans
        
        sql = "Insert Into Packing_Master(" & vbCrLf & _
            "Cust_Code, Consignee, ConsigneeTitle, Notify_Code, Packing_No, Packing_Date, Stuffing_Date, ETD, ETA, " & vbCrLf & _
            "Amount, Total_Qty, TotalWeight_Netto, TotalWeight_Gross, Total_Volume, " & vbCrLf & _
            "Reissue_Cls, Fix_Cls, List_DO, List_DoDate, Payment_Code, Payment_Days, Payment_Terms, Payment, Transportation_Cls, " & vbCrLf & _
            "Forwarder,Vessel, Mother_Vessel, From_Port, Country_Origin, To_Port, Final_Destination, " & vbCrLf & _
            "POCaseMark1, POCaseMark2, POCaseMark3, POCaseMark4, POCaseMark5,Remarks, Packingstyle_cls,Final_Destination_Cls) "
            'WHCode) "
        
        sql = sql & "Values ('" & _
            cboCust.Text & "', '" & _
            cboDelPlace.Text & "', '" & _
            Trim(TxtTitle.Text) & "', '" & _
            cboNotify.Text & "', '" & _
            CboPacking.Text & "', '" & _
            Format(DtPacking.Value, "yyyy-MM-dd") & "', '" & _
            Format(DtStuffing.Value, "yyyy-MM-dd") & "', '" & _
            Format(DtEtd.Value, "yyyy-MM-dd") & "', '" & _
            Format(DtEta.Value, "yyyy-MM-dd") & "', " & _
            "0, 0, 0, 0, 0, '0', '0', '', '', '" & _
            cboPaymentCode.List(cboPaymentCode.ListIndex, 1) & "', " & _
            Val(TxtDay.Text) & ", '" & _
            cboPaymentTerm.Text & "', '" & _
            txtPaymentTerm.Text & "', '" & _
            CboTrans.List(CboTrans.ListIndex, 1) & "', '" & _
            TxtForwarder.Text & "', '" & _
            TxtVessel.Text & "', '" & _
            TxtMotherVessel.Text & "', '" & _
            TxtFrom.Text & "', '" & _
            TxtCountry.Text & "', '" & _
            TxtTo.Text & "', '" & _
            TxtFinal.Text & "', '" & _
            TxtCaseMark(0).Text & "', '" & TxtCaseMark(1).Text & "', '" & TxtCaseMark(2).Text & "', '" & TxtCaseMark(3).Text & "', '" & TxtCaseMark(4).Text & "', '" & _
            Trim(txtRemarks.Text) & "', '" & cbopackingtype.Text & "','" & Trim(cboPlaceofDestination.Text) & "')"
            'cboWH.Text & "')"
        
        Db.Execute sql
        Db.CommitTrans
        
'        cboStatus.ListIndex = 0
 '       cboPacking.Text = TxtPacNo.Text
        
    End If
    
    If CekInvoice Then
        IsiGridHead
        GrandTotal
        LblErrMsg = DisplayMsg(4110)
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    IsiGridHead
    GrandTotal
    Me.MousePointer = vbDefault

End Sub
'menampilkan isi header/packing_Master
Sub IsiHeader()
Dim RHeader As New ADODB.Recordset
Set RHeader = Nothing
sql = "SELECT * FORM Packing_Master WHRE ltrim(Packing_No)='" & CboPacking & "'"

End Sub
Private Sub CmdPreview_Click()
    
    Me.MousePointer = vbHourglass
    
    Call PackingPrintStatus("'" & CboPacking.Text & "'")
    Call PackingPrint("'" & CboPacking.Text & "'")
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cmdserialnumber_Click()

    Me.MousePointer = vbHourglass
    
    Call PackingPrintStatus("'" & CboPacking.Text & "'")
'    Call PackingPrint("'" & cboPacking.Text & "'")
    
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsMain As New Recordset
    Dim rsSub As New Recordset
    Dim RsSubDetail As New Recordset
   
    sql = " Select PM.Packing_No,PM.Packing_Date,PM.Stuffing_Date,PM.ETD,PM.ETA,PM.Payment_Days,PM.Payment, " & vbCrLf & _
                      "     PM.Transportation_Cls,PM.Vessel,PM.Mother_Vessel,PM.From_Port,PM.To_Port,PM.Final_Destination,PM.PackingStyle_Cls, " & vbCrLf & _
                      "     PM.Remarks,PM.Last_User, PM.POCaseMark1,PM.POCaseMark2," & vbCrLf & _
                      "     PM.Cust_Code,TM.Trade_Name Customer_Name,TM.Address1 CsAddress1,TM.Address2 CsAddress2,TM.City CsCity, TM.Country CsCountry, " & vbCrLf & _
                      "     PM.Consignee,PM.ConsigneeTitle,TM1.Trade_Name Consignee_Name,TM1.Address1 CgAddress1,TM1.Address2 CgAddress2,TM1.City CgCity, TM1.Country CgCountry, " & vbCrLf & _
                      "     PM.Payment_Terms,Isnull(PT.Description,'') Payment_Desription, " & vbCrLf & _
                      "     PD.Order_No,PD.Container_No,PD.Container_Size,PD.SerialNoFrom,PD.SerialNoTo,PD.Qty,PD.Unit_Cls, " & vbCrLf & _
                      "     PD.QtyWeight_Netto,Pd.QtyWeight_Gross,Rtrim(PD.Ctn_No) Ctn_No,Qty_Ctn=CASE WHEN PD.Qty_Ctn=0 THEN 1 ELSE PD.Qty_Ctn END,PD.Length,PD.Width,PD.Thickness, " & vbCrLf & _
                      "     PD.Item_Code,IM.Item_Name " & vbCrLf & _
                      "     From Packing_Master PM " & vbCrLf & _
                      "     Inner Join Packing_Detail PD On PM.Packing_No=PD.Packing_No "
    
    sql = sql + "     Inner Join Trade_Master TM on PM.Cust_Code=TM.Trade_Code " & vbCrLf & _
                      "     Inner Join Trade_Master TM1 on PM.Consignee=TM1.Trade_Code " & vbCrLf & _
                      "     Inner Join Item_Master IM on PD.Item_Code=IM.Item_Code " & vbCrLf & _
                      "     Left Join PaymentTerm_Cls PT on PM.Payment_Terms=PT.PaymentTerm_Cls " & vbCrLf & _
                      "     Where ltrim(PM.Packing_No)='" & Trim(CboPacking.Text) & "' " & vbCrLf & _
                      "     Order By PM.Packing_No, PD.PackingSeq_No, PD.Item_Code"


    sqlprint = sql
    sqlprint2 = sql
    rsMain.CursorLocation = adUseClient
    If rsMain.State <> adStateClosed Then rsMain.Close
    rsMain.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If Not rsMain.EOF Then
            
        Set report = application.OpenReport(App.path & "\Reports\PackingList_SerialNumber.rpt")
        report.Database.Tables(1).SetDataSource rsMain
        report.DiscardSavedData

        Dim Rpt As New FrmRpt3

        reportcode = "PackingList_SerialNumber"
        printorient = 1
        packing_no = CboPacking.Text
        With Rpt.CRViewer1
            .ReportSource = report
            .ViewReport
            .Zoom 1
        End With
        With Rpt
            .WindowState = 2
            .Show 1
        End With
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub CmdStuffing_Click()

Dim xlapp As New Excel.application
Dim rsCek As New ADODB.Recordset, sql As String, sql1 As String, sql2 As String, sql3 As String
Dim rsCompany As New ADODB.Recordset
Dim NO As Integer, Idx As Integer, intCount As Integer, TQty As Double, IntPos As Integer

On Error GoTo ErrHandlerExcel
LblErrMsg.Caption = ""
DoEvents


sql = " Select PD.Order_No SI_No,PM.Packing_Date,PM.Final_Destination, PD.Qty," & vbCrLf & _
            "   Container_No,Container_Size,PD.Item_Code,IM.Item_Name,SerialNoFrom,SerialNoTo " & vbCrLf & _
            "       From Packing_Master PM  " & vbCrLf & _
            "           Inner Join Packing_Detail PD On PM.Packing_No=PD.Packing_No " & vbCrLf & _
            "           Inner Join Item_Master IM on PD.Item_Code=Im.Item_Code " & vbCrLf & _
            "       Where ltrim(PM.Packing_No)='" & Trim(CboPacking) & "'"


rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic

If Not rsCek.EOF Then
Screen.MousePointer = vbHourglass
With xlapp

'    sql1 = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax  From company_profile "
'    If rsCompany.State <> adStateClosed Then rsCompany.Close
'    rsCompany.Open sql1, Db, adOpenDynamic, adLockOptimistic
'
'    If rsCompany.EOF Then Screen.MousePointer = vbDefault: Exit Sub
    .Workbooks.Add
'
'    .Range("a2", "h2").Merge
'    .Range("a2") = rsCompany!company_name
'    .Range("a3", "h3").Merge
'    .Range("a3") = rsCompany!address1 & " " & rsCompany!address2 & " " & rsCompany!City & " " & rsCompany!Province & " " & rsCompany!postal_code
'    .Range("a4", "h4").Merge
'    .Range("a4") = "Phone: " & rsCompany!phone1 & " " & rsCompany!phone2 & " Fax: " & rsCompany!fax
    
    .Range("B1") = "STUFFING REPORT"
    .Range("b1", "h1").Merge
    .Range("B1").HorizontalAlignment = xlLeft
    .Range("B1", "H1").Columns.Font.Bold = True
    
    .Range("b3") = "SI / PO No  "
    .Range("b3", "c3").Merge
    .Range("B3").HorizontalAlignment = xlLeft
    
    .Range("b4") = "Export Date  "
    .Range("b4", "c4").Merge
    .Range("B4").HorizontalAlignment = xlLeft
    
    .Range("b5") = "Destination  "
    .Range("b5", "c5").Merge
    .Range("B5").HorizontalAlignment = xlLeft
    

    .Range("d3") = ": " & Trim(rsCek.Fields("SI_No"))
    .Range("d3").HorizontalAlignment = xlLeft
    .Range("d4") = ": " & Format(rsCek.Fields("Packing_Date"), "MMMM YYYY")
    .Range("d4").HorizontalAlignment = xlLeft
    .Range("d5") = ": " & Trim(rsCek.Fields("Final_Destination") & "")
    .Range("d5").HorizontalAlignment = xlLeft
    
    .Range("f3") = "Container/Seal No  "
    .Range("f4") = "Container Size "
    .Range("f5") = "POLICE No"
    
    .Range("g3") = ": " & Trim(rsCek.Fields("Container_No"))
    .Range("g3", "h3").Merge
    .Range("g3").HorizontalAlignment = xlLeft
    
    .Range("g4") = ": " & Trim(rsCek.Fields("Container_Size") & "")
    .Range("g4", "h4").Merge
    .Range("g4").HorizontalAlignment = xlLeft

    .Range("g5") = ": "
    .Range("g5", "h5").Merge
    .Range("g5").HorizontalAlignment = xlLeft
    
    .Range("B3", "G5").Columns.Font.Bold = True
    
    Idx = 7
    
    .Range("B" & Idx) = "NO"
    .Range("C" & Idx) = "Model"
    .Range("C" & Idx, "D" & Idx).Merge
    .Range("E" & Idx) = "Serial No"
    .Range("E" & Idx, "F" & Idx).Merge
    .Range("B" & Idx, "G" & Idx).HorizontalAlignment = xlCenter
    .Range("G" & Idx) = "Qty"
    .Range("G" & Idx).HorizontalAlignment = xlRight
    .Range("B" & Idx, "G" & Idx).Columns.Font.Bold = True
    
    Idx = Idx + 1
    
    Do While Not rsCek.EOF
        
        Idx = Idx
        NO = NO + 1
        .Range("B" & Idx) = NO
        .Range("C" & Idx) = Trim(rsCek.Fields("Item_Code")) & "  " & Trim(rsCek.Fields("Item_Name"))
        .Range("C" & Idx, "D" & Idx).Merge
        .Range("E" & Idx) = Trim(rsCek.Fields("SerialNoFrom")) & " - " & Trim(rsCek.Fields("SerialNoTo"))
        .Range("E" & Idx, "F" & Idx).Merge
        .Range("C" & Idx, "F" & Idx).HorizontalAlignment = xlCenter
        .Range("G" & Idx) = rsCek.Fields("Qty")
        TQty = TQty + rsCek.Fields("Qty")
        Idx = Idx + 1
        rsCek.MoveNext
    Loop
    
    .Range("B7:G" & Idx - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range("B7:G" & Idx - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range("B7:G" & Idx - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range("B7:G" & Idx - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range("B7:G" & Idx - 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Range("B7:G" & Idx - 1).Borders(xlInsideVertical).LineStyle = xlContinuous
     
    .Range("F" & Idx) = "Total"
    .Range("G" & Idx) = TQty
    
    .Range("G7:G" & Idx).Select
    .Selection.NumberFormat = gs_formatQty
    .Range("G7:G" & Idx).HorizontalAlignment = xlRight
    .Range("F" & Idx, "G" & Idx).Columns.Font.Bold = True
    
    Idx = Idx + 2
    
    .Range("B" & Idx) = "Check Sheet"
    .Range("B" & Idx, "D" & Idx).Merge
    .Range("B" & Idx, "H" & Idx).HorizontalAlignment = xlLeft
    .Range("B" & Idx, "H" & Idx).Columns.Font.Bold = True
    
    Idx = Idx + 1
    IntPos = Idx
    
    .Range("B" & Idx) = "NO"
    .Range("C" & Idx) = "Serial Number"
    .Range("D" & Idx) = "Pallet Number"
    .Range("E" & Idx) = "Fumigasi Date"
    .Range("F" & Idx) = "  Load By  "
    .Range("G" & Idx) = "  Load Time  "
    .Range("H" & Idx) = "       Remark       "
    .Range("B" & Idx, "H" & Idx).HorizontalAlignment = xlCenter
    
    .Range("B" & Idx, "H" & Idx).Columns.Font.Bold = True
    
    Idx = Idx + 1
    
    rsCek.MoveFirst
    Do While Not rsCek.EOF
        .Range("C" & Idx) = Trim(rsCek.Fields("Item_Name") & "") & "  " & Trim(rsCek.Fields("SerialNoFrom") & "") & "-" & Trim(rsCek.Fields("SerialNoTo") & "")
        .Range("C" & Idx, "H" & Idx).Merge
        .Range("C" & Idx).HorizontalAlignment = xlLeft
        .Range("B" & Idx, "H" & Idx).Columns.Font.Bold = True
        
        Idx = Idx + 1
        NO = 0
        
        If Trim(rsCek.Fields("SerialNoFrom")) = "" Then GoTo Lanjut
        
        For intCount = Right(Trim(rsCek.Fields("SerialNoFrom") & ""), 6) To Right(Trim(rsCek.Fields("SerialNoTo") & ""), 6)
            NO = NO + 1
            .Range("B" & Idx) = NO
            .Range("C" & Idx) = Left(Trim(rsCek.Fields("SerialNoFrom") & ""), 1) & Format(intCount, "000000")
            Idx = Idx + 1
        Next
Lanjut:
        rsCek.MoveNext
    Loop
    
    .Range("B" & IntPos, "H" & Idx).Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range("B" & IntPos, "H" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range("B" & IntPos, "H" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range("B" & IntPos, "H" & Idx).Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range("B" & IntPos, "H" & Idx).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Range("B" & IntPos, "H" & Idx).Borders(xlInsideVertical).LineStyle = xlContinuous
    
    .Range("B1", "H" & Idx + 3).Columns.Font.Name = "Verdana"
    .Range("B1").Columns.Font.Size = 10
    .Range("B1").Columns.Font.Underline = True
    .Range("B2", "H" & Idx + 3).Columns.Font.Size = 8
       
    .Range("A1").Select
    .Range("A1").ColumnWidth = 1.5
    
    .ActiveSheet.PageSetup.PaperSize = xlPaperA4
    .ActiveSheet.PageSetup.Orientation = 1
    .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.25)
    .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.25)
    .ActiveSheet.PageSetup.TopMargin = application.InchesToPoints(0.3)
    .Range("A:J").Columns.AutoFit
    .WindowState = xlMaximized
    .Visible = True
End With
Else
    'lbl_pesan = DisplayMsg(4006)
End If



Screen.MousePointer = vbDefault
Set rsCek = Nothing

Exit Sub

ErrHandlerExcel:
Screen.MousePointer = vbDefault
If rsCek.State <> adStateClosed Then rsCek.Close
Set rsCek = Nothing
LblErrMsg.Caption = "[" & err.number & "] " & err.Description
err.clear
        
End Sub

Private Sub CmdSubMenu_Click()

    ClearPacking
    DoEvents
    frmMainMenu.Show
    Unload Me

End Sub

Private Sub CmdSubmit_Click()

    Dim i, j, k, X As Long
    Dim rscno As New ADODB.Recordset
    Dim RsTanggal As New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    If lblfix <> "" Then
        LblErrMsg = DisplayMsg(4046)
        Me.MousePointer = vbDefault
        Exit Sub
    End If
     
    If hakUpdate(Me.Name) = 0 Then _
    LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
    If CekInvoice Then
        LblErrMsg = DisplayMsg(4110)
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    '#20071004 Yudha, check data surat jalan yang item nya beda dengan order entry
    If uf_check_SuratJalan_OrderDifferent = False Then
        LblErrMsg = DisplayMsg("0079")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    '20080108 Dudi.. Mengecek apakah sudah ada packing master atau belum
    If Not CekSql("SELECT * FROM Packing_Master WHERE ltrim(Packing_No)='" & Trim(CboPacking.Text) & "'") Then
        LblErrMsg = DisplayMsg("1038")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    'cek Qty di grid gak boleh 0 & cek CartonNo ga boleh kosong
    For X = 2 To grid.Rows - 1
        If grid.Cell(flexcpChecked, X, ColCtr) = flexChecked Then
            If grid.TextMatrix(X, ColCounter) = "CHILD" Then
                If CDbl(grid.TextMatrix(X, ColQty)) = 0 Then
                    LblErrMsg = DisplayMsg(4043) & " 0 !"
                    Me.MousePointer = vbDefault
                    Exit Sub
'                ElseIf Trim(grid.TextMatrix(x, ColCartonNo)) = "" Then
'                    LblErrMsg = DisplayMsg(8107)
'                    grid.Row = x
'                    grid.Col = ColCartonNo
'                    grid.SetFocus
'                    Me.MousePointer = vbDefault
'                    Exit Sub
                End If
            End If
        End If
    Next
    
    Db.BeginTrans
    
    sql = "Update Packing_Master " & _
        "Set Consignee = '" & cboDelPlace.Text & "', " & _
        "ConsigneeTitle = '" & TxtTitle.Text & "', " & _
        "Notify_Code = '" & cboNotify.Text & "', " & _
        "Packing_Date = '" & Format(DtPacking.Value, "yyyy-MM-dd") & "', " & _
        "Stuffing_Date = '" & Format(DtStuffing.Value, "yyyy-MM-dd") & "', " & _
        "ETD ='" & Format(DtEtd.Value, "yyyy-MM-dd") & "', " & _
        "ETA ='" & Format(DtEta.Value, "yyyy-MM-dd") & "', " & _
        "Amount = " & CDbl(TxtTTAm.Text) & ", " & _
        "Total_Qty = " & CDbl(TxtTTQty.Text) & ", " & _
        "TotalWeight_Netto = " & CDbl(TxtTTW.Text) & ", " & _
        "TotalWeight_Gross = " & CDbl(TxtTTWG.Text) & "," & _
        "Total_Volume = " & CDbl(TxtTTV.Text) & ", " & _
        "Payment_Code = '" & cboPaymentCode.List(cboPaymentCode.ListIndex, 0) & "',  " & _
        "Payment_Days = " & Val(TxtDay.Text) & ", " & _
        "Payment_Terms = '" & Trim(cboPaymentTerm.Text) & "', " & _
        "Payment = '" & Trim(txtPaymentTerm) & "',  " & _
        "Transportation_Cls = '" & CboTrans.List(CboTrans.ListIndex, 1) & "', " & _
        "Forwarder = '" & Trim(TxtForwarder.Text) & "', " & _
        "Vessel = '" & Trim(TxtVessel.Text) & "', " & _
        "Mother_Vessel = '" & Trim(TxtMotherVessel.Text) & "', " & _
        "From_Port = '" & Trim(TxtFrom.Text) & "', " & _
        "Country_Origin = '" & Trim(TxtCountry.Text) & "', " & _
        "To_Port = '" & Trim(TxtTo.Text) & "',  " & _
        "Final_Destination = '" & Trim(TxtFinal.Text) & "', "
    
    sql = sql & _
        "POCaseMark1 = '" & TxtCaseMark(0).Text & "', " & _
        "POCaseMark2 = '" & TxtCaseMark(1).Text & "', " & _
        "POCaseMark3 = '" & TxtCaseMark(2).Text & "', " & _
        "POCaseMark4 = '" & TxtCaseMark(3).Text & "', " & _
        "POCaseMark5 = '" & TxtCaseMark(4).Text & "', " & _
        "Remarks = '" & txtRemarks.Text & "', " & _
        "PackingStyle_Cls = '" & cbopackingtype.Text & "', " & _
        "Final_Destination_Cls = '" & Trim(cboPlaceofDestination.Text) & "',"
        '"WHCode = '" & cboWH.Text & "', "
        sql = sql & "Last_Update=getdate()," & _
        "Last_User='" & userLogin & "' " & _
        "Where ltrim(Packing_No) = '" & Trim(CboPacking.Text) & "'"
                  
    Db.Execute sql
    
    If grid.Rows > 2 Then
    
        sql = "delete from packing_detail where ltrim(packing_no) = '" & Trim(CboPacking.Text) & "'"
        Db.Execute sql
        
        j = 1
        k = 1
        For i = 2 To grid.Rows - 1
            
            If grid.TextMatrix(i, ColCounter) = "HEAD" Then k = grid.TextMatrix(i, ColSeq)
            
                If grid.Cell(flexcpChecked, i, ColCtr) = flexChecked Then
     
'                    If Trim(grid.TextMatrix(i, ColContainerNo)) = "" Then
'                        LblErrMsg = DisplayMsg(4103)
'                        Db.RollbackTrans
'                        grid.Row = i
'                        grid.Col = ColContainerNo
'                        grid.SetFocus
'                        Me.MousePointer = vbDefault
'                        Exit Sub
'                    End If
                                        
                    sql = "Insert Into Packing_Detail (" & _
                        "Packing_No, " & _
                        "Container_No, " & _
                        "Container_Size, " & _
                        "PackingSeq_No, " & _
                        "Do_No, " & _
                        "DoSeq_no, " & _
                        "Item_Code, SerialNoFrom,SerialNoTo," & _
                        "MakerItem_Code, " & _
                        "Qty, " & _
                        "Length, " & _
                        "Width, " & _
                        "Thickness, " & _
                        "QtyWeight_Netto, " & _
                        "QtyWeight_Gross, " & _
                        "Qty_Volume,Qty_Ctn, Ctn_No, " & _
                        "Detail_Cls, " & _
                        "Unit_Cls, " & _
                        "Currency_Code, " & _
                        "Price, " & _
                        "Amount, Order_No, Order_SeqNo) "
                        
                    sql = sql & "Values ('" & _
                        CboPacking.Text & "', '" & _
                        Trim(grid.TextMatrix(i, ColContainerNo)) & "', '" & _
                        Trim(grid.TextMatrix(i, ColDrySize)) & "', " & _
                        j & ", '" & _
                        Trim(grid.TextMatrix(CDbl(i) - CDbl(grid.TextMatrix(i, ColPos)), ColOrder)) & "', " & _
                        k & ", '" & _
                        Trim(grid.TextMatrix(CDbl(i) - CDbl(grid.TextMatrix(i, ColPos)), ColProd)) & "', '"
                    sql = sql & _
                        Trim(grid.TextMatrix(i, ColSerialFrom)) & "','" & _
                        Trim(grid.TextMatrix(i, ColSerialTo)) & "','" & _
                        Trim(grid.TextMatrix(CDbl(i) - CDbl(grid.TextMatrix(i, ColPos)), ColMaker)) & "', " & _
                        CDbl(Trim(grid.TextMatrix(i, ColQty))) & ", " & _
                        CDbl(Trim(grid.TextMatrix(i, ColLength))) & ", " & _
                        CDbl(Trim(grid.TextMatrix(i, ColWidth))) & ", " & _
                        CDbl(Trim(grid.TextMatrix(i, ColThickness))) & ", " & _
                        CDbl(Trim(grid.TextMatrix(i, ColWeight))) & ", " & _
                        CDbl(Trim(grid.TextMatrix(i, ColWGros))) & ", " & _
                        CDbl(Trim(grid.TextMatrix(i, ColVol))) & ", " & _
                        CDbl(Trim(grid.TextMatrix(i, ColPacking))) & ", '" & _
                        Trim(grid.TextMatrix(i, ColCartonNo)) & "', '1', '" & _
                        Trim(grid.TextMatrix(CDbl(i) - CDbl(grid.TextMatrix(i, ColPos)), ColUnitT)) & "', '" & _
                        Trim(grid.TextMatrix(CDbl(i) - CDbl(grid.TextMatrix(i, ColPos)), ColCurrT)) & "', " & _
                        CDbl(Trim(grid.TextMatrix(i, ColPrice))) & ", " & _
                        CDbl(Trim(grid.TextMatrix(i, ColAmount))) & ",'" & _
                        Trim(grid.TextMatrix(CDbl(i) - CDbl(grid.TextMatrix(i, ColPos)), ColPo)) & _
                        "'," & Trim(grid.TextMatrix(CDbl(i) - CDbl(grid.TextMatrix(i, ColPos)), ColpoSeqNo)) & ")"
                        
                    Db.Execute sql
                    j = j + 1
                End If
            Next
            
        sql = "Update Packing_Master " & _
            "Set List_Do = '" & listPO & "', " & _
            "List_DoDate = '" & listPODate & "', " & _
            "Last_Update=getdate()," & _
            "Last_User='" & userLogin & "' " & _
            "Where ltrim(Packing_No) = '" & Trim(CboPacking.Text) & "'"
        Db.Execute sql
        
        Db.CommitTrans
        IsiGridHead
        GrandTotal
        LblErrMsg = DisplayMsg(1000)
    Else
        Db.RollbackTrans
        LblErrMsg = DisplayMsg(5012)
    End If
    Me.MousePointer = vbDefault

End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    
    If ErrMsg = "" Then
        Unload Me
    Else
        LblErrMsg.Caption = ErrMsg
    End If

End Sub

Private Sub DTDel1_Change()
    
    If cboCust.MatchFound Then
        IsiComboPO
        If CboStatus.Text = "Update" Then IsiComboPacking
    Else
        CboPOnO.clear
        CboPacking.clear
    End If

End Sub

Private Sub DtDel2_Change()

    If cboCust.MatchFound Then
        IsiComboPO
        If CboStatus.Text = "Update" Then IsiComboPacking
    Else
        CboPOnO.clear
        CboPacking.clear
    End If

End Sub

Private Sub DtEtd_Change()
    DtDel2.Value = DtEtd.Value
End Sub

Private Sub DtPacking_Change()
If CboPacking.Text = "" Then
    'GeneratePackingNo 'Pak toha minta dimanualkan tidak usah outomatic
    DtStuffing.Value = DtPacking.Value
End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    DTDel1.Value = Date
    DtDel2.Value = Date
    DtStuffing.Value = Date
    DtPacking.Value = Date
    
    IsiCombo
'    IsiComboWH
    CboStatus.ListIndex = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode = 0 Then Cancel = 1

End Sub
Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    LblErrMsg = ""
    Dim cr As String
    cr = ""
    cr = grid.TextMatrix(Row, ColCurrT)
    For i = 2 To grid.Rows - 1
        If grid.TextMatrix(i, ColCounter) = "CHILD" Then
            If grid.Cell(flexcpChecked, i, ColCtr) = flexChecked Then
                If grid.TextMatrix(i, ColCurrT) <> cr Then
                    grid.Cell(flexcpChecked, Row, ColCtr) = flexUnchecked
                    LblErrMsg = DisplayMsg(44)
                End If
            End If
        End If
    Next
  
    If grid.Col = ColAsk Then
        If Len(grid.Text) > 1 Then grid.Text = Left(grid.Text, 1)
        If grid.Text = "C" Then
            If grid.TextMatrix(grid.Row, ColRemT) <> 0 Then
                grid.Rows = grid.Rows + 1
                For i = grid.Rows - 1 To grid.Row + 2 Step -1
                    grid.Cell(flexcpChecked, i, ColCtr) = grid.Cell(flexcpChecked, i - 1, ColCtr)
                    grid.TextMatrix(i, ColAsk) = grid.TextMatrix(i - 1, ColAsk)
                    grid.TextMatrix(i, ColContainerNo) = grid.TextMatrix(i - 1, ColContainerNo)
                    grid.TextMatrix(i, ColDrySize) = grid.TextMatrix(i - 1, ColDrySize)
                    grid.TextMatrix(i, ColOrder) = grid.TextMatrix(i - 1, ColOrder)
                    grid.TextMatrix(i, ColProd) = grid.TextMatrix(i - 1, ColProd)
                    grid.TextMatrix(i, ColMaker) = grid.TextMatrix(i - 1, ColMaker)
                    grid.TextMatrix(i, ColDesc) = grid.TextMatrix(i - 1, ColDesc)
                    grid.TextMatrix(i, ColQty) = grid.TextMatrix(i - 1, ColQty)
                    grid.TextMatrix(i, colrem) = grid.TextMatrix(i - 1, colrem)
                    grid.TextMatrix(i, ColUnit) = grid.TextMatrix(i - 1, ColUnit)
                    grid.TextMatrix(i, ColWeight) = grid.TextMatrix(i - 1, ColWeight)
                    grid.TextMatrix(i, ColWGros) = grid.TextMatrix(i - 1, ColWGros)
                    grid.TextMatrix(i, ColVol) = grid.TextMatrix(i - 1, ColVol)
                    grid.TextMatrix(i, ColPacking) = grid.TextMatrix(i - 1, ColPacking)
                    grid.TextMatrix(i, ColCartonNo) = grid.TextMatrix(i - 1, ColCartonNo)
                    grid.TextMatrix(i, ColLength) = grid.TextMatrix(i - 1, ColLength)
                    grid.TextMatrix(i, ColWidth) = grid.TextMatrix(i - 1, ColWidth)
                    grid.TextMatrix(i, ColThickness) = grid.TextMatrix(i - 1, ColThickness)
                    grid.TextMatrix(i, ColDelDate) = grid.TextMatrix(i - 1, ColDelDate)
                    grid.TextMatrix(i, ColCurr) = grid.TextMatrix(i - 1, ColCurr)
                    grid.TextMatrix(i, ColPrice) = grid.TextMatrix(i - 1, ColPrice)
                    grid.TextMatrix(i, ColAmount) = grid.TextMatrix(i - 1, ColAmount)
                    grid.TextMatrix(i, ColCounter) = grid.TextMatrix(i - 1, ColCounter)
                    grid.TextMatrix(i, ColCurrT) = grid.TextMatrix(i - 1, ColCurrT)
                    grid.TextMatrix(i, ColUnitT) = grid.TextMatrix(i - 1, ColUnitT)
                    grid.TextMatrix(i, ColRemT) = grid.TextMatrix(i - 1, ColRemT)
                    grid.TextMatrix(i, ColPos) = grid.TextMatrix(i - 1, ColPos)
                    grid.TextMatrix(i, ColPip) = grid.TextMatrix(i - 1, ColPip)
                    grid.TextMatrix(i, ColSeq) = grid.TextMatrix(i - 1, ColSeq)
                    grid.TextMatrix(i, ColQtyPerCarton) = grid.TextMatrix(i - 1, ColQtyPerCarton)
                    grid.TextMatrix(i, ColNetWeight) = grid.TextMatrix(i - 1, ColNetWeight)
                    grid.TextMatrix(i, ColGrossWeight) = grid.TextMatrix(i - 1, ColGrossWeight)
                    'grid.TextMatrix(i, ColpoSeqNo) = grid.TextMatrix(i - 1, ColpoSeqNo)
                    
                    If grid.TextMatrix(i, ColCounter) = "HEAD" Then
                        grid.Cell(flexcpBackColor, i, ColContainerNo, i, ColAmount) = &H80000018
                        grid.Cell(flexcpBackColor, i, ColCtr, i, ColAmount) = &HE0E0E0
                        grid.Cell(flexcpBackColor, i, ColAsk) = &HFFFFFF
                    ElseIf grid.TextMatrix(i, ColCounter) = "CHILD" Then
                        grid.Cell(flexcpBackColor, i, ColCtr) = &HFFFFFF
                        grid.Cell(flexcpBackColor, i, ColAsk) = &H80000018
                        grid.Cell(flexcpBackColor, i, ColContainerNo, i, ColAmount) = &H80000018
                        grid.Cell(flexcpBackColor, i, ColContainerNo) = &HFFFFFF
                        grid.Cell(flexcpBackColor, i, ColDrySize) = &HFFFFFF
                        grid.Cell(flexcpBackColor, i, ColQty) = &HFFFFFF
                        grid.Cell(flexcpBackColor, i, ColWeight) = &HFFFFFF
                        grid.Cell(flexcpBackColor, i, ColWGros) = &HFFFFFF
                        grid.Cell(flexcpBackColor, i, ColVol) = &HFFFFFF
                        grid.Cell(flexcpBackColor, i, ColPacking) = &HFFFFFF
                        grid.Cell(flexcpBackColor, i, ColCartonNo) = &HFFFFFF
                        grid.Cell(flexcpBackColor, i, ColLength) = &HFFFFFF
                        grid.Cell(flexcpBackColor, i, ColWidth) = &HFFFFFF
                        grid.Cell(flexcpBackColor, i, ColThickness) = &HFFFFFF
                    End If
                Next
                
                grid.Cell(flexcpChecked, grid.Row + 1, ColCtr) = flexUnchecked
                grid.TextMatrix(grid.Row + 1, ColAsk) = ""
                grid.TextMatrix(grid.Row + 1, ColContainerNo) = ""
                grid.TextMatrix(grid.Row + 1, ColDrySize) = ""
                grid.TextMatrix(grid.Row + 1, ColOrder) = ""
                grid.TextMatrix(grid.Row + 1, ColProd) = ""
                grid.TextMatrix(grid.Row + 1, ColDesc) = ""
                grid.TextMatrix(grid.Row + 1, ColQty) = Format(0, gs_formatQty)
                grid.TextMatrix(grid.Row + 1, colrem) = ""
                grid.TextMatrix(grid.Row + 1, ColUnit) = grid.TextMatrix(grid.Row, ColUnit)
                grid.TextMatrix(grid.Row + 1, ColWeight) = Format(0, gs_formatWeight)
                grid.TextMatrix(grid.Row + 1, ColWGros) = Format(0, gs_formatWeight)
                grid.TextMatrix(grid.Row + 1, ColVol) = Format(0, gs_formatVolume)
                grid.TextMatrix(grid.Row + 1, ColPacking) = Format(0, gs_formatBox)
                grid.TextMatrix(grid.Row + 1, ColCartonNo) = ""
                grid.TextMatrix(grid.Row + 1, ColDelDate) = ""
                grid.TextMatrix(grid.Row + 1, ColCurr) = grid.TextMatrix(grid.Row, ColCurr)
                grid.TextMatrix(grid.Row + 1, ColPrice) = grid.TextMatrix(grid.Row, ColPrice)
                grid.TextMatrix(grid.Row + 1, ColAmount) = Format(0, gs_formatAmountIDR)
                grid.TextMatrix(grid.Row + 1, ColCounter) = "CHILD"
                grid.TextMatrix(i, ColCurrT) = grid.TextMatrix(grid.Row, ColCurrT)
                grid.TextMatrix(i, ColUnitT) = grid.TextMatrix(grid.Row, ColUnitT)
                grid.TextMatrix(i, ColUnitT) = grid.TextMatrix(grid.Row, ColUnitT)
                'tambahan dudi, Januari 2009
                grid.TextMatrix(i, ColpoSeqNo) = grid.TextMatrix(grid.Row, ColpoSeqNo)
                grid.TextMatrix(i, ColPos) = grid.TextMatrix(grid.Row, ColPos)
                
                grid.TextMatrix(i, ColRemT) = ""
        
                grid.Cell(flexcpBackColor, i, ColCtr) = &HFFFFFF
        
                For i = 1 To CDbl(grid.TextMatrix(grid.Row, ColPos) + 1)
                    grid.TextMatrix(grid.Row + i, ColPos) = i
                    grid.TextMatrix(grid.Row, ColPos) = i
                Next
                
                clearmark (grid.Text)
            Else
                clearmark (grid.Text)
                LblErrMsg = DisplayMsg("0043")
            End If
        End If
    End If
    
    If grid.Col = ColWeight Or grid.Col = ColWGros Or grid.Col = ColVol Or grid.Col = ColPacking Or grid.Col = ColQtyPerCarton Or _
        grid.Col = ColLength Or grid.Col = ColWidth Or grid.Col = ColThickness Then
        If Trim(grid.Text) = "" Then grid.Text = "0"
        If Not IsNumeric(grid.Text) Then
            grid.Text = Q
        End If
    End If
    
    If grid.Col = ColQty Then
        If CDbl(grid.Text) > gd_MaxQty Then
            grid.Text = Q
            LblErrMsg = DisplayMsg(4045) & " " & gd_MaxQty & " !"
        Else
            If Trim(grid.TextMatrix(Row, ColSerialTo)) <> "" Then _
            grid.TextMatrix(Row, ColSerialTo) = GetSerialTo(grid.TextMatrix(Row, ColSerialFrom), CDbl(grid.Text))
        End If
    ElseIf grid.Col = ColSerialFrom Then
        If Trim(grid.TextMatrix(Row, ColSerialTo)) <> "" Then _
        grid.TextMatrix(Row, ColSerialTo) = GetSerialTo(grid.TextMatrix(Row, ColSerialFrom), CDbl(grid.TextMatrix(Row, ColQty)))
    ElseIf grid.Col = ColWeight Or grid.Col = ColWGros Then
        If CDbl(grid.Text) > gd_MaxWeight Then
            grid.Text = Q
            LblErrMsg = DisplayMsg(8030) & " " & gd_MaxWeight & " !"
        End If
    ElseIf grid.Col = ColVol Then
        If CDbl(grid.Text) > gd_MaxVolume Then
            grid.Text = Q
            LblErrMsg = DisplayMsg("0049") & " " & gd_MaxVolume & " !"
        End If
    ElseIf grid.Col = ColPacking Then
        If CDbl(grid.Text) > gd_MaxBox Then
            grid.Text = Q
            LblErrMsg = DisplayMsg(4037) & " " & gd_MaxBox & " !"
        End If
    ElseIf grid.Col = ColLength Then
        If CDbl(grid.Text) > gd_MaxLength Then
            grid.Text = Q
            LblErrMsg = DisplayMsg(4037) & " " & gd_MaxLength & " !"
        End If
    ElseIf grid.Col = ColWidth Then
        If CDbl(grid.Text) > gd_MaxWidth Then
            grid.Text = Q
            LblErrMsg = DisplayMsg(4037) & " " & gd_MaxWidth & " !"
        End If
    ElseIf grid.Col = ColThickness Then
        If CDbl(grid.Text) > gd_MaxThickness Then
            grid.Text = Q
            LblErrMsg = DisplayMsg(4037) & " " & gd_MaxThickness & " !"
        End If
    End If
    
    If grid.Col = ColQty Then
        If Trim(grid.Text) = "" Then grid.Text = "0"
        TotalHead CDbl(ColQty), CDbl(grid.Row), CDbl(colrem), 1
        If grid.TextMatrix(CDbl(grid.Row) - CDbl(grid.TextMatrix(grid.Row, ColPos)), colrem) < 0 Then
            grid.TextMatrix(grid.Row, ColQty) = Q
            LblErrMsg = DisplayMsg(4045) & " Remaining Qty !"
            TotalHead CDbl(ColQty), CDbl(grid.Row), CDbl(colrem), 1
        Else
            grid.TextMatrix(Row, ColQty) = Format(grid.TextMatrix(Row, ColQty), gs_formatQty)
            If CDbl(grid.TextMatrix(Row, ColQtyPerCarton)) <> 0 Then
                CartonQty = Fix(CDbl(grid.Text) / CDbl(grid.TextMatrix(Row, ColQtyPerCarton)))
            Else
                CartonQty = 0
            End If
                     
            NetWeight = CDbl(grid.TextMatrix(Row, ColNetWeight))
            If CDbl(grid.TextMatrix(Row, ColQtyPerCarton)) = 0 Then
                NetSisa = 0
            Else
                NetSisa = ((CDbl(grid.Text) Mod CDbl(grid.TextMatrix(Row, ColQtyPerCarton))) / _
                    CDbl(grid.TextMatrix(Row, ColQtyPerCarton))) * CDbl(grid.TextMatrix(Row, ColNetWeight))
            End If
            NetWeightPacking = (NetWeight * CartonQty) + NetSisa
            
            GrossWeight = CDbl(grid.TextMatrix(Row, ColGrossWeight))
            If CDbl(grid.TextMatrix(Row, ColQtyPerCarton)) = 0 Then
                GrossSisa = 0
            Else
                GrossSisa = ((CDbl(grid.Text) Mod CDbl(grid.TextMatrix(Row, ColQtyPerCarton))) / _
                CDbl(grid.TextMatrix(Row, ColQtyPerCarton))) * CDbl(grid.TextMatrix(Row, ColGrossWeight))
            End If
            GrossWeightPacking = (GrossWeight * CartonQty) + GrossSisa
            
            If CDbl(grid.TextMatrix(Row, ColQtyPerCarton)) <> 0 Then
                CartonQty = uf_Ceiling(CDbl(grid.Text) / CDbl(grid.TextMatrix(Row, ColQtyPerCarton)))
                grid.TextMatrix(Row, ColPacking) = Format(CartonQty, gs_formatBox)
            Else
                CartonQty = 0
                grid.TextMatrix(Row, ColPacking) = Format(CartonQty, gs_formatBox)
            End If
            
            Volume = (CDbl(grid.TextMatrix(Row, ColLength)) / 1000) * (CDbl(grid.TextMatrix(Row, ColWidth)) / 1000) * (CDbl(grid.TextMatrix(Row, ColThickness)) / 1000)    '* Round(CDbl(ColQty) / CDbl(grid.TextMatrix(Row, ColQtyPerCarton)))
            VolumePacking = Format((Volume * CartonQty), gs_formatVolume)
            
            grid.TextMatrix(Row, ColWeight) = Format(NetWeightPacking, gs_formatWeight)
            grid.TextMatrix(Row, ColWGros) = Format(GrossWeightPacking, gs_formatWeight)
            grid.TextMatrix(Row, ColVol) = Format(VolumePacking, gs_formatVolume)
        End If
    End If
    
    If grid.Col = ColPacking Then
        If Trim(grid.Text) = "" Then grid.Text = "0"
        grid.TextMatrix(Row, ColPacking) = Format(grid.TextMatrix(Row, ColPacking), gs_formatBox)
        CartonQty = CDbl(grid.Text)
        
        NetWeight = CDbl(grid.TextMatrix(Row, ColNetWeight))
        NetWeightPacking = NetWeight * CartonQty
            
        GrossWeight = CDbl(grid.TextMatrix(Row, ColGrossWeight))
        GrossWeightPacking = GrossWeight * CartonQty
        
        Volume = (CDbl(grid.TextMatrix(Row, ColLength)) / 1000) * (CDbl(grid.TextMatrix(Row, ColWidth)) / 1000) * (CDbl(grid.TextMatrix(Row, ColThickness)) / 1000)
        VolumePacking = Format((Volume * CartonQty), gs_formatVolume)
        
        grid.TextMatrix(Row, ColWeight) = Format(NetWeightPacking, gs_formatWeight)
        grid.TextMatrix(Row, ColWGros) = Format(GrossWeightPacking, gs_formatWeight)
        grid.TextMatrix(Row, ColVol) = VolumePacking
    End If
    
    If grid.Col = ColWeight Then
        grid.TextMatrix(Row, ColWeight) = Format(grid.TextMatrix(Row, ColWeight), gs_formatWeight)
    End If
    
    If grid.Col = ColWGros Then
        grid.TextMatrix(Row, ColWGros) = Format(grid.TextMatrix(Row, ColWGros), gs_formatWeight)
    End If
    
    If grid.Col = ColVol Then
        grid.TextMatrix(Row, ColVol) = Format(grid.TextMatrix(Row, ColVol), gs_formatVolume)
    End If
    
    If grid.Col = ColLength Or grid.Col = ColWidth Or grid.Col = ColThickness Then
        grid.TextMatrix(Row, ColVol) = Format(( _
            (CDbl(grid.TextMatrix(Row, ColLength)) / 1000) * _
            (CDbl(grid.TextMatrix(Row, ColWidth)) / 1000) * _
            (CDbl(grid.TextMatrix(Row, ColThickness)) / 1000)) * CDbl(grid.TextMatrix(Row, ColPacking)), gs_formatVolume)
    End If
        
    GrandTotal

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    grid.EditMaxLength = 1
    If grid.Col = ColAsk Then
        If grid.TextMatrix(grid.Row, ColCounter) = "HEAD" Then
            Exit Sub
        End If
    End If
    If grid.Col = ColContainerNo Then
        grid.EditMaxLength = 255
    End If
    If grid.Col = ColDrySize Then
        grid.EditMaxLength = 25
    End If
    If grid.TextMatrix(grid.Row, ColCounter) = "CHILD" Then
        If grid.Col = ColQty Or grid.Col = ColWeight Or grid.Col = ColWGros Or grid.Col = ColPacking _
            Or grid.Col = ColLength Or grid.Col = ColWidth Or grid.Col = ColThickness Then
            Q = IIf(Trim(grid.TextMatrix(grid.Row, grid.Col)) = "", 0, grid.TextMatrix(grid.Row, grid.Col))
            Select Case grid.Col
            Case ColQty, ColPacking: grid.EditMaxLength = 7
            Case ColWeight, ColWGros: grid.EditMaxLength = 10
            Case ColLength, ColWidth, ColThickness: grid.EditMaxLength = 9
            End Select
        End If
        If grid.Col = ColCartonNo Then
            grid.EditMaxLength = 10
        End If
        If grid.Col = ColCtr Then
            If CheckPONo(grid.TextMatrix(Row - grid.TextMatrix(Row, ColPos), ColOrder)) Then Exit Sub
        End If
        If grid.Cell(flexcpChecked, grid.Row, ColCtr) = flexChecked Then
            If grid.Col = ColContainerNo Or grid.Col = ColDrySize _
                Or grid.Col = ColQty Or ColSerialFrom Or grid.Col = ColWeight _
                Or grid.Col = ColWGros Or grid.Col = ColPacking _
                Or grid.Col = ColCartonNo Or grid.Col = ColLength Or grid.Col = ColWidth Or grid.Col = ColThickness _
             Then
                Exit Sub
             End If
        End If
    End If
    Cancel = True

End Sub

Private Sub grid_Click()

    If grid.Row > 1 Then
        grid.FocusRect = flexFocusNone
        If grid.Col = ColAsk Then
            If grid.TextMatrix(grid.Row, ColCounter) = "HEAD" Then
                grid.FocusRect = flexFocusInset
                Exit Sub
            End If
        End If
        If grid.TextMatrix(grid.Row, ColCounter) = "CHILD" Then
            If grid.Col = ColCtr Then grid.FocusRect = flexFocusInset
            If grid.Cell(flexcpChecked, grid.Row, ColCtr) = flexChecked Then
                If grid.Col = ColCtr Or grid.Col = ColContainerNo Or grid.Col = ColDrySize _
                    Or grid.Col = ColQty Or grid.Col = ColWeight _
                    Or grid.Col = ColWGros Or grid.Col = ColPacking _
                    Or grid.Col = ColCartonNo Or grid.Col = ColLength Or grid.Col = ColWidth Or grid.Col = ColThickness _
                    Then
                        grid.FocusRect = flexFocusInset
                End If
            End If
        End If
    End If

End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If grid.TextMatrix(grid.Row, ColCounter) = "HEAD" Then
        If grid.Col = ColAsk Then
            If Len(grid.TextMatrix(grid.Row, ColCtr)) = 1 Then KeyAscii = 0
            If KeyAscii <> Asc("C") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then
                KeyAscii = 0
            End If
            If KeyAscii = Asc(".") Then KeyAscii = 0
        End If
    ElseIf grid.TextMatrix(grid.Row, ColCounter) = "CHILD" Then
        If grid.Col = ColAsk Then
            If Len(grid.TextMatrix(grid.Row, ColCtr)) = 1 Then KeyAscii = 0
            If KeyAscii <> Asc("D") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then
               KeyAscii = 0
            End If
        End If
        If grid.Col = ColQty Or grid.Col = ColLength Or grid.Col = ColWidth Or grid.Col = ColThickness Then If KeyAscii = Asc(".") Then KeyAscii = 0
        If grid.Col <> ColQty And grid.Col <> ColWeight And grid.Col <> ColWGros And grid.Col <> ColVol And grid.Col <> ColPacking _
            And grid.Col <> ColLength And grid.Col <> ColWidth And grid.Col <> ColThickness Then
        Else
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyEscape Then
                KeyAscii = 0
            End If
        End If
    End If

End Sub



Private Sub Label1_Click()

End Sub

Private Sub lblFix_Change()
    
    lblfix.Visible = (lblfix.Caption <> "")
    
End Sub

Private Sub cbopackingtype_Click()
lblpackingtype.Caption = cbopackingtype.List(cbopackingtype.ListIndex, 1)
End Sub

Private Sub cbopackingtype_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub cbopackingtype_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    Call uf_ValidateComboData(cbopackingtype, "0024", LblErrMsg, lblpackingtype)
End If
End Sub

Function listPO() As String
Dim rsIsiPO As New ADODB.Recordset
Dim tampungPO As String
Dim tampungPODate As String

    sql = "Select distinct a.Do_No, b.Do_date from packing_detail a inner join DO_Master b on a.Do_No = b.Do_No where ltrim(packing_no) = '" & Trim(CboPacking.Text) & "'"
    rsIsiPO.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsIsiPO.EOF Then
        listPO = ""
        listPODate = ""
    Else
        tampungPO = ""
        tampungPODate = ""
        Do While Not rsIsiPO.EOF
            tampungPO = tampungPO & Trim(rsIsiPO(0)) & ", "
            tampungPODate = tampungPODate & Format(rsIsiPO(1), "dd-MMM-yyyy") & ", "
            rsIsiPO.MoveNext
        Loop
        listPO = Left(Trim(tampungPO), Len(Trim(tampungPO)) - 1)
        listPODate = Left(Trim(tampungPODate), Len(Trim(tampungPODate)) - 1)
    End If
    rsIsiPO.Close
    Set rsIsiPO = Nothing
End Function

Private Function CekInvoice() As Boolean
    Dim adoRs As New ADODB.Recordset
    sql = "select  packing_no from invoice_detail where ltrim(packing_no) = '" & Trim(CboPacking) & "'"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    CekInvoice = Not adoRs.EOF
    adoRs.Close
    Set adoRs = Nothing
End Function
 Function CekSql(SqlData)
Dim Rdata As New ADODB.Recordset
Set Rdata = Nothing
Rdata.Open SqlData, Db, adOpenDynamic, adLockBatchOptimistic
If Rdata.EOF Then
CekSql = False
Else
CekSql = True
End If
Rdata.Close
End Function

Private Sub IsiComboWH()
    Exit Sub
    Dim adoRs As New ADODB.Recordset
    
    sql = "select wh_code, wh_name from warehouse_master"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With cboWH
        .clear
        .columnCount = 2
        .ColumnWidths = "60 pt; 180 pt"
        .ListWidth = 240
        .ListRows = 15
        While adoRs.EOF = False
            .AddItem ""
            .List(.ListCount - 1, 0) = Trim(adoRs.Fields("wh_code"))
            .List(.ListCount - 1, 1) = Trim(adoRs.Fields("wh_name"))
            adoRs.MoveNext
        Wend
'        .Text = "FG"
    End With
    
    adoRs.Close
    Set adoRs = Nothing
    
End Sub

Private Function uf_check_SuratJalan_OrderDifferent() As Boolean
    Dim Pos As Integer
    Dim ls_sql As String
    Dim found As Boolean
    Dim RS As New ADODB.Recordset
    found = False
    If grid.Rows > 1 Then
        For Pos = 1 To grid.Rows - 1
            If grid.TextMatrix(Pos, ColCounter) = "HEAD" Then
                ls_sql = " select * from Delivery_Order " & _
                    " where Do_No='" & Trim(grid.TextMatrix(Pos, ColOrder)) & "' " & _
                    " and item_code='" & Trim(grid.TextMatrix(Pos, ColProd)) & "' " & _
                    " and DoSeq_no='" & Trim(grid.TextMatrix(Pos, ColSeq)) & "' "
                If RS.State <> adStateClosed Then RS.Close
                RS.CursorLocation = adUseClient
                RS.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
                If RS.EOF = True Then
                    found = True
                    Exit For
                End If
            End If
        Next
        If found = True Then
            uf_check_SuratJalan_OrderDifferent = False
        Else
            uf_check_SuratJalan_OrderDifferent = True
        End If
    Else
        uf_check_SuratJalan_OrderDifferent = True
    End If
    If RS.State <> adStateClosed Then RS.Close
End Function


